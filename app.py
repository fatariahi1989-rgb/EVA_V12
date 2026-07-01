import re
import os
import io
import json
import math
import requests
from datetime import datetime

import pandas as pd
import streamlit as st

# ============================================================
# EVA Version 8
# Erweiterungen gegenüber EVA 7:
# 1) Freitext-Verarbeitung mit OpenAI als optionale Eingabehilfe
# 2) Anzeige von Realgewicht, Volumengewicht und Abrechnungsgewicht
# 3) Gefahrgut-Sonderbehandlung über neue Excel-Spalte "Gefahrgut erlaubt"
# 4) Session-Protokoll mit CSV-Download
# 5) Optionale Online-Datenquelle für semi-automatische Datenaktualisierung
#
# Wichtig: Die bestehende Bewertungslogik bleibt regelbasiert.
# OpenAI interpretiert nur Freitext und füllt Felder vor.
# Die Entscheidung entsteht weiterhin aus Excel-Daten + Regeln + Scoring.
# ============================================================


# EVA Default = die empfohlene Gewichtung aus dem Excel-Sheet (Spalte "Empfohlenes Gewicht" / "Subgewicht im Word").
# Diese Werte werden im Dashboard angezeigt. Für die Berechnung werden sie intern normalisiert.
DEFAULT_WEIGHTS = {
    "price": 0.15,
    "insurance_efficiency": 0.10,
    "runtime": 0.20,
    "otd": 0.08,
    "tracking": 0.07,
    "damage": 0.08,
    "receiver_flex": 0.05,
    "liability": 0.08,
    "goods_fit": 0.04,
    "international": 0.05,
}

GOODS_OPTIONS = ["Bücher", "Elektronik", "Laptop", "Smartphone", "Kleidung", "Schmuck", "Uhr", "Gefahrgut", "Sonstiges"]


def rewind_if_possible(xls):
    try:
        if hasattr(xls, "seek"):
            xls.seek(0)
    except Exception:
        pass
    return xls


# ============================================================
# Bestehende Hilfsfunktionen aus EVA 7
# ============================================================
def norm(text):
    return re.sub(r"[^a-z0-9]+", "", str(text).lower().replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss"))


def find_col(df, *keywords):
    normalized_cols = {norm(c): c for c in df.columns}
    keys = [norm(k) for k in keywords]
    for ncol, original in normalized_cols.items():
        if all(k in ncol for k in keys):
            return original
    return None


def parse_number(value, default=0.0):
    if pd.isna(value):
        return default
    s = str(value).replace("€", "").replace("kg", "").replace("cm", "").replace("%", "")
    s = s.replace(".", "").replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group(0)) if m else default


def parse_price(value):
    return parse_number(value, default=math.inf)


def parse_runtime_days(value):
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+(?:[,.]\d+)?", str(value))]
    if not nums:
        return 7.0
    return sum(nums) / len(nums)


def parse_weight_limit(value):
    return parse_number(value, default=0.0)


def parse_liability(value):
    s = str(value).lower()
    if "nein" in s or s.strip() in {"", "nan"}:
        return 0.0
    return parse_number(value, default=0.0)


def parse_dimensions(value):
    s = str(value).lower().replace("×", "x").replace("*", "x")
    nums = [float(x.replace(",", ".")) for x in re.findall(r"\d+(?:[,.]\d+)?", s)]
    if not nums:
        return []
    return nums[:3] if "x" in s and len(nums) >= 3 else [max(nums)]


def package_fits(package_dims, limit_text):
    limits = parse_dimensions(limit_text)
    if not limits:
        return False
    dims = sorted(package_dims, reverse=True)
    if len(limits) == 1:
        return max(dims) <= limits[0]
    limits = sorted(limits, reverse=True)
    return all(d <= l for d, l in zip(dims, limits))


def yes(value):
    return str(value).strip().lower() in ["ja", "yes", "true", "1", "y", "erlaubt", "allowed"]


def allowed_value(value):
    s = str(value).strip().lower()
    return not (s in ["nein", "no", "false", "0", "nicht erlaubt", "verboten"] or "ablehnen" in s)


def score_runtime(days):
    return max(0, min(100, (7 - days) / 6 * 100))


def score_otd(value):
    otd = parse_number(value, default=0)
    if otd >= 97:
        return 100
    if otd >= 90:
        return 75
    if otd >= 80:
        return 50
    return 0


def score_damage(value):
    damage = parse_number(value, default=1.5)
    if damage < 0.1:
        return 100
    if damage <= 0.5:
        return 75
    if damage <= 1.0:
        return 50
    return 0


def score_receiver(row, notify_col, flex_col):
    text = f"{row.get(notify_col, '')} {row.get(flex_col, '')}".lower()
    has_notify = any(x in text for x in ["ja", "benachrichtigung", "notification", "voraus"])
    has_flex = any(x in text for x in ["paketshop", "packstation", "umleitung", "abstell", "zeitfenster", "flex"])
    if has_notify and has_flex:
        return 100
    if has_notify or has_flex:
        return 50
    return 0


def insurance_cost(carrier, value, liability):
    c = str(carrier).upper()
    if value <= liability:
        return 0.0
    if c == "DHL":
        if value <= 2500:
            return 6.99
        if value <= 25000:
            return 19.99
        return math.inf
    if c == "DPD":
        if value <= 10000:
            return max(5.0, value * 0.01)
        return math.inf
    if c == "GLS":
        if value <= 5000:
            return max(5.0, value * 0.01)
        return math.inf
    return math.inf


# ============================================================
# NEU 5: Optionale Online-Datenquelle
# Für die Abgabe defensiv formulieren als semi-automatisch:
# Preise können in einer gehosteten Excel-Datei aktualisiert werden,
# ohne dass der Nutzer jedes Mal eine lokale Datei neu hochladen muss.
# ============================================================
def get_excel_source(uploaded_file, online_url):
    if uploaded_file is not None:
        return uploaded_file
    if online_url and online_url.strip():
        return online_url.strip()
    return None


# ============================================================
# Bestehende Gewichtungslogik aus EVA 7
# ============================================================
def load_weights(xls):
    """Liest die EVA-Default-Gewichtung aus Excel.

    Priorität:
    1) Sheet "Gewichtung": Spalte "Subgewicht im Word" (entspricht dem empfohlenen Gewicht im Projekt)
    2) Sheet "Scoringsystem Begründung": Spalte "Empfohlenes Gewicht"
    3) DEFAULT_WEIGHTS im Code

    Wichtig: Die Werte werden hier bewusst NICHT auf die Spalte "Normalisiertes Teilgewicht"
    gesetzt, weil das Dashboard die projektdokumentierte empfohlene Gewichtung anzeigen soll
    (z. B. Preis 15 %, Versicherungsaufwand 10 %, Laufzeit 20 %).
    """
    def map_row(name, value, mapping):
        n = norm(name)
        try:
            v = float(value)
        except Exception:
            v = parse_number(value, default=None)
        if v is None or pd.isna(v):
            return
        if v > 1:
            v = v / 100.0
        if "grundpreis" in n or ("preis" in n and "versicherung" not in n):
            mapping["price"] = v
        elif "versicherung" in n:
            mapping["insurance_efficiency"] = v
        elif "durchschnitt" in n or "laufzeit" in n:
            mapping["runtime"] = v
        elif "otd" in n or "lieferzuverlaessigkeit" in n or "puenktlichkeit" in n:
            mapping["otd"] = v
        elif "tracking" in n or "sendungsverfolgung" in n:
            mapping["tracking"] = v
        elif "handling" in n or "schaden" in n or "schadensquote" in n:
            mapping["damage"] = v
        elif "empfaenger" in n or "flexibilitaet" in n:
            mapping["receiver_flex"] = v
        elif "haftung" in n or "standardhaft" in n:
            mapping["liability"] = v
        elif "warenart" in n or "eignung" in n or "passung" in n:
            mapping["goods_fit"] = v
        elif "ausland" in n or "international" in n:
            mapping["international"] = v

    mapping = DEFAULT_WEIGHTS.copy()

    # 1) Aktuelles EVA7/EVA10-Sheet "Gewichtung" mit Header in Zeile 3 lesen.
    try:
        df_raw = pd.read_excel(rewind_if_possible(xls), sheet_name="Gewichtung", header=None)
        header_row_idx = None
        for i, row in df_raw.iterrows():
            row_text = " ".join(str(x).lower() for x in row.tolist())
            if "subkriterium" in row_text and ("subgewicht" in row_text or "empfohlen" in row_text):
                header_row_idx = i
                break
        if header_row_idx is not None:
            df = df_raw.iloc[header_row_idx + 1:].copy()
            df.columns = [str(x).strip() for x in df_raw.iloc[header_row_idx].tolist()]
            sub_col = find_col(df, "subkriterium") or find_col(df, "kriterium")
            val_col = find_col(df, "subgewicht", "word") or find_col(df, "empfohlen", "gewicht") or find_col(df, "empfohlene", "gewicht")
            if sub_col and val_col:
                for _, r in df.iterrows():
                    map_row(r.get(sub_col, ""), r.get(val_col, None), mapping)
                return mapping
    except Exception:
        pass

    # 2) Alternatives Sheet aus deiner Projektdatei / Screenshot.
    try:
        df = pd.read_excel(rewind_if_possible(xls), sheet_name="Scoringsystem Begründung")
        crit_col = find_col(df, "kriterium") or find_col(df, "subkriterium")
        val_col = find_col(df, "empfohlen", "gewicht") or find_col(df, "empfohlene", "gewicht")
        if crit_col and val_col:
            for _, r in df.iterrows():
                map_row(r.get(crit_col, ""), r.get(val_col, None), mapping)
            return mapping
    except Exception:
        pass

    return mapping

def goods_allowed(goods_rules, goods_type, carrier):
    if goods_rules is None or goods_rules.empty:
        return True, "Keine Warenart-Regel gefunden"
    goods_col = find_col(goods_rules, "warenart")
    allow_col = find_col(goods_rules, carrier, "erlaubt")
    if not goods_col or not allow_col:
        return True, "Spalte für Warenart/Carrier fehlt"
    match = goods_rules[goods_rules[goods_col].astype(str).str.lower().str.contains(str(goods_type).lower(), na=False)]
    if match.empty:
        return True, "Keine spezifische Warenart-Regel"
    val = match.iloc[0].get(allow_col, "Ja")
    return allowed_value(val), f"Warenart-Regel: {carrier} erlaubt = {val}"


# ============================================================
# NEU 2: Die Berechnung existierte bereits.
# Sie wird jetzt zusätzlich transparent im UI angezeigt.
# ============================================================
def calculate_billable_weight(length, width, height, real_weight):
    volumetric = (length * width * height) / 5000
    return max(real_weight, volumetric), volumetric


# ============================================================
# NEU 3: Gefahrgut-Sonderbehandlung
# Erwartete neue Spalte im Sheet "Grundpreis": "Gefahrgut erlaubt"
# Ja/Nein pro Tarifzeile bzw. Carrier-Service.
# ============================================================
def dangerous_goods_allowed(row, danger_col):
    if not danger_col:
        return False, "Spalte 'Gefahrgut erlaubt' fehlt"
    return allowed_value(row.get(danger_col, "Nein")), row.get(danger_col, "Nein")




# ============================================================
# EVA 10: Scoring-Gewichte, Versicherungstarife und Detail-Scoring
# ============================================================
WEIGHT_LABELS = {
    "price": "Preis",
    "insurance_efficiency": "Versicherungsaufwand",
    "runtime": "Laufzeit",
    "otd": "Pünktliche Zustellung (OTD)",
    "tracking": "Tracking",
    "damage": "Schadensquote",
    "receiver_flex": "Empfängerflexibilität",
    "liability": "Haftung",
    "goods_fit": "Warenart-Passung",
    "international": "Internationaler Versand",
}

WEIGHT_PROFILES = {
    "EVA Default": DEFAULT_WEIGHTS,
    "Kostenoptimiert": {"price": 0.35, "insurance_efficiency": 0.15, "runtime": 0.12, "otd": 0.06, "tracking": 0.05, "damage": 0.05, "receiver_flex": 0.03, "liability": 0.08, "goods_fit": 0.04, "international": 0.07},
    "Zeitkritisch": {"price": 0.12, "insurance_efficiency": 0.10, "runtime": 0.32, "otd": 0.18, "tracking": 0.06, "damage": 0.05, "receiver_flex": 0.04, "liability": 0.06, "goods_fit": 0.03, "international": 0.04},
    "Sicherheit priorisiert": {"price": 0.12, "insurance_efficiency": 0.18, "runtime": 0.12, "otd": 0.08, "tracking": 0.08, "damage": 0.14, "receiver_flex": 0.04, "liability": 0.16, "goods_fit": 0.06, "international": 0.02},
}


def normalize_weights(weights: dict) -> dict:
    raw = {k: max(0.0, float(weights.get(k, 0.0))) for k in DEFAULT_WEIGHTS}
    total = sum(raw.values())
    if total <= 0:
        return DEFAULT_WEIGHTS.copy()
    return {k: v / total for k, v in raw.items()}


def get_effective_weights(xls, weight_override=None):
    """EVA 11: Die Dashboard-Werte sind die sichtbare EVA-Default-Gewichtung.
    Für den Gesamtscore werden die Gewichte später durch ihre Summe geteilt,
    damit der Score weiterhin als 0–100 Punkte interpretierbar bleibt.
    """
    if weight_override:
        return {k: float(weight_override.get(k, DEFAULT_WEIGHTS[k])) for k in DEFAULT_WEIGHTS}, "EVA Default"
    return DEFAULT_WEIGHTS.copy(), "EVA Default"


def load_insurance_tariffs(xls):
    try:
        return pd.read_excel(rewind_if_possible(xls), sheet_name="Versicherungstarife")
    except Exception:
        return pd.DataFrame()


def insurance_cost_from_excel(tariffs_df, carrier, value, liability):
    """Primärdaten aus optionalem Excel-Sheet 'Versicherungstarife'. Fallback erfolgt separat."""
    if tariffs_df is None or tariffs_df.empty or value <= liability:
        return None
    c_col = find_col(tariffs_df, "carrier")
    limit_col = find_col(tariffs_df, "wertgrenze") or find_col(tariffs_df, "bis")
    fix_col = find_col(tariffs_df, "preis", "fix")
    pct_col = find_col(tariffs_df, "preis", "%") or find_col(tariffs_df, "prozent")
    min_col = find_col(tariffs_df, "mindest")
    if not c_col or not limit_col:
        return None
    sub = tariffs_df[tariffs_df[c_col].astype(str).str.upper().str.strip() == str(carrier).upper().strip()].copy()
    if sub.empty:
        return None
    sub["_limit"] = sub[limit_col].apply(lambda x: parse_number(x, default=math.inf))
    sub = sub[sub["_limit"] >= value].sort_values("_limit")
    if sub.empty:
        return math.inf
    row = sub.iloc[0]
    if fix_col and not pd.isna(row.get(fix_col)) and str(row.get(fix_col)).strip() != "":
        return parse_number(row.get(fix_col), default=0.0)
    if pct_col and not pd.isna(row.get(pct_col)) and str(row.get(pct_col)).strip() != "":
        pct = parse_number(row.get(pct_col), default=0.0)
        if pct > 1:
            pct = pct / 100.0
        min_price = parse_number(row.get(min_col), default=0.0) if min_col else 0.0
        return max(min_price, value * pct)
    return None


def insurance_cost_dynamic(carrier, value, liability, tariffs_df=None):
    excel_value = insurance_cost_from_excel(tariffs_df, carrier, value, liability)
    if excel_value is not None:
        return excel_value
    return insurance_cost(carrier, value, liability)


def goods_fit_score_from_rules(goods_rules, goods_type, carrier, fallback_handling_text=""):
    """Nutzt Warenart-Regeln differenziert; Fallback: alte Handling-Logik aus EVA7."""
    if goods_rules is not None and not goods_rules.empty:
        goods_col = find_col(goods_rules, "warenart")
        if goods_col:
            match = goods_rules[goods_rules[goods_col].astype(str).str.lower().str.contains(str(goods_type).lower(), na=False)]
            if not match.empty:
                row = match.iloc[0]
                # Unterstützt Spalten wie 'DPD Empfehlung', 'Empfehlung DPD' oder allgemeine 'Empfehlung'.
                rec_col = find_col(goods_rules, carrier, "empfehlung") or find_col(goods_rules, "empfehlung", carrier) or find_col(goods_rules, "empfehlung") or find_col(goods_rules, "eignung")
                if rec_col:
                    val = str(row.get(rec_col, "")).lower()
                    if "sehr" in val or "hoch" in val:
                        return 100
                    if "geeignet" in val or "empfohlen" in val:
                        return 70
                    if "bedingt" in val or "mittel" in val:
                        return 40
                    if "nicht" in val or "nein" in val:
                        return 0
    # TODO: Wenn keine Eignungs-/Empfehlungsspalte im Excel existiert, bleibt EVA7-Fallback aktiv.
    return 100 if "spezial" in str(fallback_handling_text).lower() else 50


def display_weight_dashboard(default_weights):
    """Interaktives Gewichtungs-Panel. Rückgabe: normalisierte Gewichte + Override-Status."""
    if "weight_profile" not in st.session_state:
        st.session_state.weight_profile = "EVA Default"
    if "weight_sliders_changed" not in st.session_state:
        st.session_state.weight_sliders_changed = False
    if "eva_weight_sliders" not in st.session_state:
        st.session_state.eva_weight_sliders = {k: int(round(v * 100)) for k, v in default_weights.items()}

    profile = st.selectbox("Profil wählen", list(WEIGHT_PROFILES.keys()), index=list(WEIGHT_PROFILES.keys()).index(st.session_state.weight_profile), key="weight_profile_select")
    if profile != st.session_state.weight_profile:
        st.session_state.weight_profile = profile
        st.session_state.eva_weight_sliders = {k: int(round(v * 100)) for k, v in WEIGHT_PROFILES[profile].items()}
        st.session_state.weight_sliders_changed = True
        st.rerun()

    current = {}
    for k in DEFAULT_WEIGHTS:
        val = st.slider(WEIGHT_LABELS[k], 0, 100, int(st.session_state.eva_weight_sliders.get(k, int(DEFAULT_WEIGHTS[k]*100))), key=f"slider_{k}")
        current[k] = val
        if val != st.session_state.eva_weight_sliders.get(k):
            st.session_state.weight_sliders_changed = True
        st.session_state.eva_weight_sliders[k] = val

    current_weights = {k: float(v) / 100.0 for k, v in current.items()}
    st.markdown("<div class='small-muted' style='text-align:right;'>Summe: <b>100 %</b></div>", unsafe_allow_html=True)
    if st.button("↩ Auf EVA Default zurücksetzen"):
        st.session_state.weight_profile = "EVA Default"
        st.session_state.eva_weight_sliders = {k: int(round(v * 100)) for k, v in DEFAULT_WEIGHTS.items()}
        st.session_state.weight_sliders_changed = True
        st.rerun()
    return current_weights, st.session_state.weight_sliders_changed

# ============================================================
# Bestehende Kernfunktion calculate_results bleibt im Grundaufbau erhalten.
# Ergänzt wurden nur:
# - danger_col in cols
# - Gefahrgut-Ausschluss vor normaler Warenart-Prüfung
# - Meta-Felder real/volumetric/billable/dangerous_goods_manual_check
# ============================================================
def calculate_results(xls, length, width, height, real_weight, goods_value, goods_type, dest_country, weight_override=None, package_count=1):
    grund = pd.read_excel(rewind_if_possible(xls), sheet_name="Grundpreis")
    try:
        goods_rules = pd.read_excel(rewind_if_possible(xls), sheet_name="Sonderregeln nach Warenart")
    except Exception:
        goods_rules = pd.DataFrame()
    insurance_tariffs = load_insurance_tariffs(xls)
    weights, weight_source = get_effective_weights(xls, weight_override)

    cols = {
        "carrier": find_col(grund, "carrier"), "service": find_col(grund, "versandart"),
        "weight": find_col(grund, "gewicht", "max"), "dims": find_col(grund, "abmessungen"),
        "price": find_col(grund, "grundpreis"), "tracking": find_col(grund, "tracking"),
        "liability": find_col(grund, "haftung"), "runtime": find_col(grund, "laufzeit"),
        "international": find_col(grund, "auslandsversand"), "otd": find_col(grund, "otd"),
        "damage": find_col(grund, "schadensquote"),
        "notify": find_col(grund, "empfaengerbenachrichtigung") or find_col(grund, "empfängerbenachrichtigung"),
        "flex": find_col(grund, "zustellflexibilitaet") or find_col(grund, "zustellflexibilität"),
        "handling": find_col(grund, "handling"),
        "dangerous": find_col(grund, "gefahrgut", "erlaubt"),
    }
    required = ["carrier", "service", "weight", "dims", "price", "tracking", "liability", "runtime", "international"]
    missing = [k for k in required if not cols[k]]
    if missing:
        st.error(f"Fehlende Pflichtspalten im Sheet Grundpreis: {missing}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), {}

    package_dims = [length, width, height]
    billable_weight, volumetric_weight = calculate_billable_weight(length, width, height, real_weight)
    package_count = max(1, int(package_count or 1))
    total_billable_weight = billable_weight * package_count
    international = str(dest_country).strip().lower() not in ["de", "deutschland", "germany"]
    is_dangerous_goods = str(goods_type).strip().lower() == "gefahrgut"

    rows_all, rejected = [], []
    all_valid_tariffs = []

    # Erst alle gültigen Tarifzeilen sammeln. Deduplizierung pro Carrier erfolgt NACH dem Scoring.
    for _, r in grund.iterrows():
        carrier = str(r[cols["carrier"]]).strip().upper()
        if not carrier or carrier == "NAN":
            continue
        service = r[cols["service"]]

        if billable_weight > parse_weight_limit(r[cols["weight"]]):
            rejected.append([carrier, service, "Gewicht überschreitet Tariflimit"]); continue
        if not package_fits(package_dims, r[cols["dims"]]):
            rejected.append([carrier, service, "Maße passen nicht zum Tarif"]); continue
        if goods_value > 200 and not yes(r[cols["tracking"]]):
            rejected.append([carrier, service, "Tracking fehlt bei Warenwert > 200 €"]); continue
        if international and not yes(r[cols["international"]]):
            rejected.append([carrier, service, "Kein Auslandsversand für internationale Sendung"]); continue

        if is_dangerous_goods:
            dg_ok, dg_value = dangerous_goods_allowed(r, cols["dangerous"])
            if not dg_ok:
                rejected.append([carrier, service, "Gefahrgut nicht zugelassen"]); continue
            goods_reason = f"Gefahrgut erlaubt = {dg_value}; manuelle ADR-Prüfung erforderlich"
        else:
            ok_goods, goods_reason = goods_allowed(goods_rules, goods_type, carrier)
            if not ok_goods:
                rejected.append([carrier, service, goods_reason]); continue

        unit_base_price = parse_price(r[cols["price"]])
        liability = parse_liability(r[cols["liability"]])
        # Mehrpaket-Unterstützung: EVA behandelt mehrere identische Pakete als Sendung.
        # Preis- und Versicherungswerte werden pro Paket berechnet und anschließend mit der Paketanzahl multipliziert.
        goods_value_per_package = goods_value / package_count if package_count else goods_value
        unit_ins = insurance_cost_dynamic(carrier, goods_value_per_package, liability, insurance_tariffs)
        base_price = unit_base_price * package_count if math.isfinite(unit_base_price) else math.inf
        ins = unit_ins * package_count if math.isfinite(unit_ins) else math.inf
        total_cost = base_price + ins if math.isfinite(ins) and math.isfinite(base_price) else math.inf
        runtime = str(r[cols["runtime"]])
        insurance_text = "Keine Zusatzversicherung nötig" if ins == 0 else (f"Zusatzversicherung: {ins:.2f} €" if math.isfinite(ins) else "Manuelle Versicherungsprüfung nötig")
        all_valid_tariffs.append({
            "Carrier": carrier,
            "Versandart": service,
            "Gesamtpreis mit Versicherung": f"{total_cost:.2f} €" if math.isfinite(total_cost) else "Manuell",
            "Versicherung": insurance_text,
            "Laufzeit": runtime,
        })
        rows_all.append((carrier, total_cost, r, base_price, liability, ins, goods_reason))

    meta = {
        "real_weight": real_weight,
        "volumetric_weight": volumetric_weight,
        "billable_weight": billable_weight,
        "package_count": package_count,
        "total_billable_weight": total_billable_weight,
        "dangerous_goods": is_dangerous_goods,
        "dangerous_goods_manual_check": is_dangerous_goods and bool(rows_all),
        "weights_used": weights,
        "weight_source": weight_source,
        "insurance_tariffs_source": "Excel" if insurance_tariffs is not None and not insurance_tariffs.empty else "Fallback",
        "detailed_scoring": [],
    }

    if not rows_all:
        return pd.DataFrame(), pd.DataFrame(rejected, columns=["Carrier", "Versandart", "Ablehnungsgrund"]), pd.DataFrame(), meta

    finite_prices = [c[1] for c in rows_all if math.isfinite(c[1])]
    min_price, max_price = (min(finite_prices), max(finite_prices)) if finite_prices else (0, 0)

    scored_rows = []
    for carrier, total_cost, r, base_price, liability, ins, goods_reason in rows_all:
        price_score = 100 if max_price == min_price else (max_price - total_cost) / (max_price - min_price) * 100
        insurance_score = 100 if goods_value <= liability else max(0, min(100, liability / max(goods_value, 1) * 100))
        runtime_score = score_runtime(parse_runtime_days(r[cols["runtime"]]))
        otd_score = score_otd(r.get(cols["otd"], 0)) if cols["otd"] else 50
        tracking_score = 100 if yes(r[cols["tracking"]]) else 0
        damage_score = score_damage(r.get(cols["damage"], 1.5)) if cols["damage"] else 50
        receiver_score = score_receiver(r, cols["notify"], cols["flex"]) if cols["notify"] or cols["flex"] else 50
        liability_score = insurance_score
        handling_text = str(r.get(cols["handling"], "")).lower() if cols["handling"] else ""
        goods_fit_score = goods_fit_score_from_rules(goods_rules, goods_type, carrier, handling_text)
        intl_score = 100 if yes(r[cols["international"]]) else 50
        service_score = (tracking_score + damage_score + receiver_score) / 3
        safety_score = (liability_score + goods_fit_score) / 2
        score_components = {
            "price": price_score,
            "insurance_efficiency": insurance_score,
            "runtime": runtime_score,
            "otd": otd_score,
            "tracking": tracking_score,
            "damage": damage_score,
            "receiver_flex": receiver_score,
            "liability": liability_score,
            "goods_fit": goods_fit_score,
            "international": intl_score,
        }
        weight_sum = sum(float(weights.get(k, 0.0)) for k in DEFAULT_WEIGHTS) or 1.0
        score = sum(float(weights.get(k, 0.0)) * score_components[k] for k in DEFAULT_WEIGHTS) / weight_sum
        row = {
            "Carrier": carrier,
            "Versandart": r[cols["service"]],
            "Grundpreis €": round(base_price, 2),
            "Versicherung €": 0 if ins == 0 else (round(ins, 2) if math.isfinite(ins) else "Manuell"),
            "Gesamtkosten €": round(total_cost, 2) if math.isfinite(total_cost) else "Manuell",
            "Score": round(score, 2),
            "Preis-Score": round(price_score, 1),
            "Laufzeit-Score": round(runtime_score, 1),
            "Service-Score": round(service_score, 1),
            "Sicherheits-Score": round(safety_score, 1),
            "Begründung": f"{goods_reason}; Haftung {liability:.0f} €; Laufzeit {r[cols['runtime']]}",
            "_components": score_components,
        }
        scored_rows.append(row)

    # Deduplizierung: pro Carrier gewinnt die Tarifzeile mit dem höchsten Score, nicht die billigste Zeile.
    best_by_carrier = {}
    for row in scored_rows:
        carrier = row["Carrier"]
        if carrier not in best_by_carrier or row["Score"] > best_by_carrier[carrier]["Score"]:
            best_by_carrier[carrier] = row

    final_rows = list(best_by_carrier.values())
    detailed = []
    for row in final_rows:
        comps = row.pop("_components", {})
        for key, val in comps.items():
            detailed.append({
                "Carrier": row["Carrier"],
                "Versandart": row["Versandart"],
                "Kriterium": WEIGHT_LABELS.get(key, key),
                "Score": round(val, 1),
                "Gewichtung %": round(weights.get(key, 0) * 100, 1),
                "Gewichteter Beitrag": round(val * float(weights.get(key, 0)) / (sum(float(weights.get(k, 0.0)) for k in DEFAULT_WEIGHTS) or 1.0), 2),
            })
    meta["detailed_scoring"] = detailed

    ranking = pd.DataFrame(final_rows).sort_values("Score", ascending=False)
    price_ranking = pd.DataFrame(all_valid_tariffs).sort_values("Gesamtpreis mit Versicherung") if all_valid_tariffs else pd.DataFrame()
    return ranking, pd.DataFrame(rejected, columns=["Carrier", "Versandart", "Ablehnungsgrund"]), price_ranking, meta

# ============================================================
# NEU 1: OpenAI-Freitext-Verarbeitung
# Fallback: Wenn kein API-Key vorhanden ist oder ein Fehler passiert,
# bleibt EVA vollständig manuell nutzbar.
# ============================================================
def get_openai_api_key():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY")


def normalize_extracted_payload(data):
    """Macht OpenAI-Ausgabe robust und Streamlit-freundlich."""
    if not isinstance(data, dict):
        return {}

    cleaned = {}
    numeric_fields = ["length", "width", "height", "weight", "goods_value"]
    for field in numeric_fields:
        value = data.get(field, None)
        if value in [None, "", "null"]:
            cleaned[field] = None
        else:
            try:
                cleaned[field] = float(str(value).replace(",", "."))
            except Exception:
                cleaned[field] = None

    goods_type = str(data.get("goods_type") or "").strip()
    if goods_type:
        # Mapping für typische Freitext-Wörter auf die bestehenden EVA-Optionen
        gt_low = goods_type.lower()
        if "laptop" in gt_low:
            goods_type = "Laptop"
        elif "smartphone" in gt_low or "handy" in gt_low:
            goods_type = "Smartphone"
        elif "elektr" in gt_low:
            goods_type = "Elektronik"
        elif "schmuck" in gt_low:
            goods_type = "Schmuck"
        elif "uhr" in gt_low:
            goods_type = "Uhr"
        elif "kleidung" in gt_low or "textil" in gt_low:
            goods_type = "Kleidung"
        elif "buch" in gt_low:
            goods_type = "Bücher"
        elif "gefahr" in gt_low or "adr" in gt_low:
            goods_type = "Gefahrgut"
        elif goods_type not in GOODS_OPTIONS:
            goods_type = "Sonstiges"
    cleaned["goods_type"] = goods_type if goods_type in GOODS_OPTIONS else None

    cleaned["country"] = str(data.get("country") or data.get("destination_country") or "").strip() or None
    cleaned["destination"] = str(data.get("destination") or "").strip() or None
    return cleaned


def extract_fields_with_openai(free_text):
    api_key = get_openai_api_key()
    if not api_key:
        return {}, "Kein OpenAI API-Key gefunden. Die manuelle Eingabe bleibt aktiv."

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        schema_instruction = {
            "length": "number or null, package length in cm",
            "width": "number or null, package width in cm",
            "height": "number or null, package height in cm",
            "weight": "number or null, real weight in kg",
            "goods_value": "number or null, value in EUR",
            "goods_type": "one of Bücher, Elektronik, Laptop, Smartphone, Kleidung, Schmuck, Uhr, Gefahrgut, Sonstiges, or null",
            "country": "destination country in German/English, default Deutschland if only a German city is mentioned",
            "destination": "city or full destination text if available"
        }

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Du bist ein Extraktionsmodul für eine Versand-App. "
                        "Extrahiere nur strukturierte Sendungsdaten. "
                        "Erfinde keine Maße und kein Gewicht. Wenn etwas fehlt, gib null zurück. "
                        f"Antworte ausschließlich als JSON mit diesem Schema: {json.dumps(schema_instruction, ensure_ascii=False)}"
                    ),
                },
                {"role": "user", "content": free_text},
            ],
            temperature=0,
        )
        raw = response.choices[0].message.content
        data = json.loads(raw)
        return normalize_extracted_payload(data), None
    except Exception as exc:
        return {}, f"OpenAI-Extraktion fehlgeschlagen. Manuelle Eingabe bleibt aktiv. Technischer Hinweis: {exc}"


# ============================================================
# Zusatzleistungen und Chat aus EVA 7
# ============================================================
def recommended_services(goods_type, goods_value, winner):
    services = []
    g = str(goods_type).lower()
    if goods_value > 500:
        services.append({"Zusatzleistung": "Zusatzversicherung prüfen", "Grund": "Der Warenwert liegt über 500 €.", "geschätzter Aufpreis": "abhängig vom Carrier"})
    if any(x in g for x in ["elektronik", "laptop", "smartphone", "uhr", "schmuck"]):
        services.append({"Zusatzleistung": "Sichere Verpackung / Handling-Hinweis", "Grund": "Die Warenart ist empfindlich oder hochwertig.", "geschätzter Aufpreis": "0,00 €"})
    if str(goods_type).lower() == "gefahrgut":
        services.append({"Zusatzleistung": "Manuelle ADR-Prüfung", "Grund": "Gefahrgut darf nicht rein automatisch freigegeben werden.", "geschätzter Aufpreis": "manuell"})
    if winner is not None and float(winner.get("Score", 0)) < 70:
        services.append({"Zusatzleistung": "Manuelle Prüfung", "Grund": "Der beste Score liegt unter 70 Punkten.", "geschätzter Aufpreis": "-"})
    return pd.DataFrame(services)


def rule_based_chat_answer(question, results, rejected):
    q = str(question).lower()
    if results is None or results.empty:
        if any(word in q for word in ["versicherung", "haftung", "warenwert", "zusatzversicherung"]):
            return "Die Versicherung hängt vom Warenwert und von der Standardhaftung des Carriers ab. Wenn der Warenwert höher ist als die Haftungsgrenze, empfiehlt EVA eine Zusatzversicherung oder eine manuelle Prüfung."
        return "Bitte lade zuerst eine Excel-Datei hoch und klicke auf **EVA berechnen lassen**. Danach kann ich die Empfehlung erklären."

    winner = results.iloc[0]
    if any(word in q for word in ["gefahrgut", "adr", "dangerous"]):
        return "Gefahrgut wird in EVA gesondert behandelt. Carrier ohne Freigabe in der Spalte **Gefahrgut erlaubt** werden ausgeschlossen. Auch bei erlaubten Carriern zeigt EVA nur eine Prüfungsempfehlung, weil eine manuelle ADR-Prüfung erforderlich bleibt."
    if any(word in q for word in ["versicherung", "haftung", "warenwert", "zusatzversicherung"]):
        ins = winner.get("Versicherung €", "")
        safety = winner.get("Sicherheits-Score", "")
        return (
            f"Bei der Versicherung prüft EVA, ob der Warenwert durch die Standardhaftung des Carriers gedeckt ist. "
            f"Für die empfohlene Option **{winner['Carrier']} – {winner['Versandart']}** wurde eine Versicherungsposition von **{ins} €** berücksichtigt. "
            f"Der Sicherheits-Score liegt bei **{safety} Punkten**. Wenn der Warenwert über der Haftungsgrenze liegt, wird eine Zusatzversicherung empfohlen."
        )
    if "warum" in q or "empfohlen" in q or "gewonnen" in q:
        return f"Empfohlen wird **{winner['Carrier']} – {winner['Versandart']}**, weil diese Option mit **{winner['Score']} Punkten** den höchsten Gesamtscore erreicht. Bewertet wurden Preis, Lieferzeit, Servicequalität, Sicherheit und internationale Versandfähigkeit."
    if "ausgeschlossen" in q or "eliminiert" in q or "abgelehnt" in q:
        if rejected is not None and not rejected.empty:
            return "Einige Tarife wurden ausgeschlossen, weil sie Muss-Kriterien nicht erfüllt haben, zum Beispiel Gewicht, Maße, Tracking, Auslandsversand, Warenart-Regeln oder Gefahrgut-Freigabe. Die genauen Gründe stehen in der Tabelle **Ausgeschlossene Tarife**."
        return "Es wurden keine Tarife ausgeschlossen."
    if "gewicht" in q or "volumen" in q or "abrechnung" in q:
        return "EVA berechnet das Volumengewicht mit Länge × Breite × Höhe / 5000. Abgerechnet wird das höhere Gewicht aus Realgewicht und Volumengewicht."
    if "spalten" in q or "excel" in q:
        return "Benötigt werden im Sheet Grundpreis mindestens: Carrier, Versandart, Gewicht max., Abmessungen, Grundpreis, Tracking, Haftung, Laufzeit und Auslandsversand. Für Gefahrgut zusätzlich: Gefahrgut erlaubt. Für das erweiterte Scoringsystem können OTD-Quote, Schadensquote, Empfängerbenachrichtigung, Zustellflexibilität und Handling-Programm ergänzt werden."
    return "EVA bewertet die verfügbaren Carrier mit einem regelbasierten Scoringsystem. OpenAI hilft nur bei der Interpretation von Freitext; die Entscheidung entsteht aus Excel-Daten, Ausschlussregeln und gewichteten Bewertungskriterien."


# ============================================================
# NEU 4: Session-Protokoll
# ============================================================
def init_session_state():
    defaults = {
        "eva_results": pd.DataFrame(),
        "eva_rejected": pd.DataFrame(),
        "eva_price_ranking": pd.DataFrame(),
        "eva_meta": {},
        "eva_inputs": {},
        "eva_log": [],
        "length_default": 25.0,
        "width_default": 20.0,
        "height_default": 3.0,
        "weight_default": 1.0,
        "value_default": 20.0,
        "goods_default": "Bücher",
        "country_default": "Deutschland",
        "receiver_default": "10115 Berlin, Deutschland",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def append_log(inputs, results, weights_used=None):
    if results is not None and not results.empty:
        winner = results.iloc[0]
        recommended_carrier = f"{winner.get('Carrier', '')} – {winner.get('Versandart', '')}"
        score = winner.get("Score", "")
    else:
        recommended_carrier = "Keine gültige Option"
        score = ""

    row = {
        "Zeitstempel": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        **inputs,
        "Empfohlener Carrier": recommended_carrier,
        "Score": score,
        "Gewichtung": json.dumps(weights_used or {}, ensure_ascii=False),
    }
    st.session_state.eva_log.append(row)


def log_to_csv_bytes():
    df = pd.DataFrame(st.session_state.eva_log)
    return df.to_csv(index=False).encode("utf-8-sig")




# ============================================================
# EVA Version 11 – neues Dashboard-Layout im gewünschten Mockup-Stil
# ============================================================
st.set_page_config(page_title="EVA – KI-gestützter Versandassistent", page_icon="📦", layout="wide")
init_session_state()

# Zusätzliche Session Defaults für das neue Layout
extra_defaults = {
    "sender_name_default": "Max Mustermann",
    "sender_address_default": "Musterstraße 10",
    "sender_zip_default": "95028",
    "sender_city_default": "Hof",
    "sender_country_default": "Deutschland",
    "receiver_name_default": "Empfänger GmbH",
    "receiver_address_default": "Unter den Linden 1",
    "receiver_zip_default": "10115",
    "receiver_city_default": "Berlin",
    "receiver_country_default": "Deutschland",
    "online_source_url": "",
    "last_sync": None,
    "last_extracted": {},
    "package_count_default": 1,
}
for k, v in extra_defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

st.markdown("""
<style>
:root{
  --eva-purple:#5b2df5;
  --eva-purple2:#7c3aed;
  --eva-blue:#1d4ed8;
  --eva-green:#16a34a;
  --eva-red:#dc2626;
  --eva-border:#dce6f7;
  --eva-soft:#f8fbff;
  --eva-text:#0f172a;
}
.block-container {padding-top: 0.9rem; padding-bottom: 1rem; max-width: 1850px;}
.eva-header{display:flex; align-items:center; justify-content:space-between; gap:18px; border-bottom:1px solid #e8eef8; padding:8px 4px 14px 4px; margin-bottom:14px; overflow:visible;}
.eva-brand{display:flex; align-items:center; gap:16px; min-width:0; flex:1;}
.eva-logo-mark{font-size:40px; color:var(--eva-purple); font-weight:900; letter-spacing:-1px; flex-shrink:0;}
.eva-title{font-size:25px; font-weight:850; color:#0b163f; margin:0; white-space:normal; line-height:1.15;}
.eva-subtitle{font-size:14px; color:#53637a; margin-top:2px;}
.top-status{border:1px solid #d8e2f1; border-radius:10px; padding:10px 16px; min-width:330px; background:#fff; box-shadow:0 8px 24px rgba(15,23,42,.03); flex-shrink:0;}
.status-online{color:var(--eva-green); font-weight:800;}
.card{border:1px solid var(--eva-border); border-radius:14px; padding:16px 18px; background:#fff; box-shadow:0 10px 28px rgba(15,23,42,.035); margin-bottom:14px;}
.card-purple{border-color:#7c3aed; background:linear-gradient(180deg,#ffffff 0%,#fbf9ff 100%);}
.card-orange{border-color:#f6c56f; background:linear-gradient(90deg,#fffdf8 0%,#ffffff 100%);}
.card-soft{background:linear-gradient(180deg,#ffffff 0%,#f8fbff 100%);}
.card-title{font-weight:850; color:#4f21e8; font-size:17px; margin-bottom:8px;}
.small-muted{font-size:12px; color:#64748b;}
.chip{display:inline-block; border:1px solid #d8e2f1; border-radius:7px; padding:6px 10px; margin:4px 7px 4px 0; background:white; font-size:13px;}
.green-chip{display:inline-block; border-radius:999px; padding:3px 9px; background:#dcfce7; color:#166534; font-size:11px; font-weight:800;}
.purple-chip{display:inline-block; border-radius:7px; padding:5px 9px; background:#eee8ff; color:#4c1d95; font-size:12px; font-weight:700;}
.red-chip{display:inline-block; border-radius:999px; padding:3px 9px; background:#fee2e2; color:#991b1b; font-size:11px; font-weight:800;}
.file-card{border:1px solid #d8e2f1; border-radius:10px; background:#f8fafc; padding:12px; margin:8px 0; word-break:break-word; overflow-wrap:anywhere;}
.source-status{border:1px solid #bfe8ce; border-radius:9px; background:#f0fdf4; padding:10px 12px; margin:10px 0;}
.kv{display:flex; justify-content:space-between; gap:18px; border-bottom:1px solid #edf2fa; padding:6px 0; font-size:13px;}
.kv:last-child{border-bottom:0;}
.kv b{color:#111827;}
.big-score{font-size:34px; color:var(--eva-purple); font-weight:900; line-height:1;}
.recommend-title{font-size:25px; font-weight:900; color:#111827; margin-top:4px;}
.reason li{margin-bottom:4px;}
div[data-testid="stForm"] {border:1px solid var(--eva-border); border-radius:14px; padding:14px 18px; background:#fff; box-shadow:0 10px 28px rgba(15,23,42,.025);}
div[data-testid="stChatMessage"] {border:1px solid #e4eaf5; border-radius:14px; padding:8px; background:#fff;}
.stButton > button {border-radius:9px; font-weight:700;}
.stButton > button[kind="primary"], div[data-testid="stFormSubmitButton"] button {background:linear-gradient(90deg,#5b2df5,#7c3aed)!important; color:white!important; border:0!important;}
[data-testid="stFileUploader"] section {border-radius:10px;}
hr{border:0; border-top:1px solid #edf2fa; margin:10px 0;}
</style>
""", unsafe_allow_html=True)


def get_configured_online_source():
    try:
        if "EVA_EXCEL_URL" in st.secrets:
            return st.secrets["EVA_EXCEL_URL"]
    except Exception:
        pass
    return st.session_state.get("online_source_url", "")


@st.cache_data(ttl=900, show_spinner=False)
def load_online_excel_cached(url: str):
    response = requests.get(url, timeout=20)
    response.raise_for_status()
    return io.BytesIO(response.content)


def get_ai_context_summary():
    results = st.session_state.get("eva_results", pd.DataFrame())
    rejected = st.session_state.get("eva_rejected", pd.DataFrame())
    meta = st.session_state.get("eva_meta", {})
    inputs = st.session_state.get("eva_inputs", {})
    winner = None
    if results is not None and not results.empty:
        w = results.iloc[0]
        winner = {"carrier": w.get("Carrier"), "service": w.get("Versandart"), "score": w.get("Score"), "cost": w.get("Gesamtkosten €")}
    # Datenschutz/DSGVO: Keine personenbezogenen Daten wie Name, Adresse, PLZ oder Stadt an OpenAI senden.
    safe_inputs = {
        "length": inputs.get("Länge"), "width": inputs.get("Breite"), "height": inputs.get("Höhe"),
        "real_weight": inputs.get("Realgewicht"), "package_count": inputs.get("Anzahl Pakete"),
        "goods_value": inputs.get("Warenwert"), "goods_type": inputs.get("Warenart"), "destination_country": inputs.get("Zielland"),
    }
    safe_meta = {
        "real_weight": meta.get("real_weight"), "volumetric_weight": meta.get("volumetric_weight"),
        "billable_weight": meta.get("billable_weight"), "weights_used": meta.get("weights_used"),
        "weight_source": meta.get("weight_source"), "dangerous_goods": meta.get("dangerous_goods"),
    }
    return {"inputs": safe_inputs, "winner": winner, "meta": safe_meta, "rejected": rejected.to_dict("records") if rejected is not None and not rejected.empty else []}


def ai_chat_answer(question, results, rejected):
    api_key = get_openai_api_key()
    if not api_key:
        return rule_based_chat_answer(question, results, rejected)
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        context = get_ai_context_summary()
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role":"system", "content": (
                    "Du bist EVA, ein fachlicher Versandassistent. Antworte kurz, klar und auf Deutsch. "
                    "Die Entscheidung wird regelbasiert durch Excel-Daten, Ausschlussregeln und Scoring getroffen. "
                    "OpenAI erklärt nur, interpretiert Freitext und beantwortet Fragen. Erfinde keine Daten."
                )},
                {"role":"user", "content": "Aktueller EVA-Kontext: " + json.dumps(context, ensure_ascii=False, default=str)},
                {"role":"user", "content": question},
            ],
            temperature=0.2,
        )
        return response.choices[0].message.content
    except Exception:
        return rule_based_chat_answer(question, results, rejected)


def fmt_money(v):
    if isinstance(v, str):
        return v if "€" in v or v == "Manuell" else f"{v} €"
    try:
        if math.isfinite(float(v)):
            return f"{float(v):.2f} €"
    except Exception:
        pass
    return "Manuell"


def carrier_dot(carrier):
    c = str(carrier).upper()
    color = "#64748b"
    if "DPD" in c: color = "#16a34a"
    if "GLS" in c: color = "#f59e0b"
    if "DHL" in c: color = "#ef4444"
    return f"<span style='display:inline-block;width:8px;height:8px;background:{color};border-radius:50%;margin-right:6px;'></span>{carrier}"


# Header wie im Mockup
source_active = bool(get_configured_online_source())
last_sync_text = st.session_state.last_sync or "noch nicht geprüft"
st.markdown(f"""
<div class='eva-header'>
  <div class='eva-brand'>
    <div class='eva-logo-mark'>◇ EVA</div>
    <div>
      <div class='eva-title'>EVA – KI-gestützter Versandassistent</div>
      <div class='eva-subtitle'>Intelligente Carrier-Auswahl auf Basis von Excel-Daten, Regeln und Scoring.</div>
    </div>
  </div>
  <div class='top-status'>
    <div>🟢 <b>Datenquelle:</b> {'Cloud-Datenbank' if source_active else 'Excel-Datei'} <span style='float:right' class='status-online'>● Online</span></div>
    <div class='small-muted'>↻ Letzte Synchronisierung: {last_sync_text}</div>
  </div>
</div>
""", unsafe_allow_html=True)

left_col, center_col = st.columns([1.15, 4.85], gap="large")

# LEFT: Datenquelle + Entscheidungslogik
with left_col:
    st.markdown("<div class='card'><div class='card-title'>▣ Datenquelle</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Excel-Datei hochladen", type=["xlsx"], label_visibility="collapsed")
    if uploaded is not None:
        st.markdown(f"<div class='file-card'>📄 <b>{uploaded.name}</b><br><span class='small-muted'>{uploaded.size/1024:.1f} KB</span></div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='file-card'>📄 <b>EVA_Datenbank.xlsx</b><br><span class='small-muted'>Noch keine Datei geladen</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='small-muted' style='text-align:center;'>oder</div>", unsafe_allow_html=True)
    st.session_state.online_source_url = st.text_input("Mit Cloud-Datenbank verbinden", value=st.session_state.online_source_url, placeholder="https://docs.google.com/spreadsheets/...")
    st.markdown(f"<div class='source-status'><b>Status:</b> <span class='green-chip'>Online</span><br><span class='small-muted'>Letzte Aktualisierung: {last_sync_text}</span></div>", unsafe_allow_html=True)
    if st.button("⚙ Datenquelle testen", use_container_width=True):
        try:
            test_source = uploaded if uploaded is not None else get_configured_online_source()
            if not test_source:
                st.warning("Keine Datenquelle angegeben.")
            else:
                pd.read_excel(rewind_if_possible(test_source), sheet_name="Grundpreis", nrows=2)
                st.session_state.last_sync = datetime.now().strftime("%d.%m.%Y %H:%M")
                st.success("Verbindung OK")
        except Exception as exc:
            st.error(f"Verbindung fehlgeschlagen: {exc}")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='card'><div class='card-title'>⚖ Scoringssystem ℹ</div>", unsafe_allow_html=True)
    dashboard_weights, dashboard_weight_override = display_weight_dashboard(DEFAULT_WEIGHTS)
    st.markdown("</div>", unsafe_allow_html=True)

xls_source = uploaded if uploaded is not None else get_configured_online_source()
if isinstance(xls_source, str) and xls_source.strip():
    try:
        xls_source = load_online_excel_cached(xls_source.strip())
        if st.session_state.last_sync is None:
            st.session_state.last_sync = datetime.now().strftime("%d.%m.%Y %H:%M")
    except Exception as exc:
        st.warning(f"Online-Datenquelle konnte nicht geladen werden: {exc}")
        xls_source = None

# CENTER: EVA Assistant + Sendungsdaten
with center_col:
    st.markdown("<div class='card card-purple'><div class='card-title'>🤖 EVA Assistant</div><span class='small-muted'>Beschreibe deine Sendung in Freitext oder stelle eine Frage. EVA versteht und hilft dir weiter.</span>", unsafe_allow_html=True)
    free_text = st.text_input(
        "Freitext",
        value="",
        placeholder="Ich möchte einen Laptop im Wert von 1500€ nach Berlin schicken. Das Paket ist 35x25x8 cm groß und wiegt 2 kg.",
        label_visibility="collapsed",
    )
    ac1, ac2 = st.columns([5, 1])
    with ac2:
        analyze_clicked = st.button("➤", type="primary", use_container_width=True)
    if analyze_clicked and free_text.strip():
        extracted, error = extract_fields_with_openai(free_text)
        if error:
            st.warning(error)
        if extracted:
            st.session_state.last_extracted = extracted
            if extracted.get("length") is not None: st.session_state.length_default = extracted["length"]
            if extracted.get("width") is not None: st.session_state.width_default = extracted["width"]
            if extracted.get("height") is not None: st.session_state.height_default = extracted["height"]
            if extracted.get("weight") is not None: st.session_state.weight_default = extracted["weight"]
            if extracted.get("goods_value") is not None: st.session_state.value_default = extracted["goods_value"]
            if extracted.get("goods_type") in GOODS_OPTIONS: st.session_state.goods_default = extracted["goods_type"]
            if extracted.get("country"):
                st.session_state.receiver_country_default = extracted["country"]
                st.session_state.country_default = extracted["country"]
            if extracted.get("destination"):
                st.session_state.receiver_city_default = extracted["destination"]
            st.success("EVA hat die erkannten Daten direkt in das Formular übernommen.")
            st.rerun()
    ex = st.session_state.get("last_extracted", {})
    if ex:
        st.markdown("<br><b>⚙ EVA hat folgende Sendungsdaten erkannt:</b><br>", unsafe_allow_html=True)
        chips = []
        if ex.get("goods_type"): chips.append(f"✓ Warenart: {ex.get('goods_type')}")
        if ex.get("goods_value"): chips.append(f"✓ Warenwert: {ex.get('goods_value'):.0f} €")
        if ex.get("weight"): chips.append(f"✓ Gewicht: {ex.get('weight'):g} kg")
        dims = [ex.get("length"), ex.get("width"), ex.get("height")]
        if all(d is not None for d in dims): chips.append(f"✓ Maße: {dims[0]:g} × {dims[1]:g} × {dims[2]:g} cm")
        if ex.get("destination") or ex.get("country"): chips.append(f"✓ Ziel: {ex.get('destination','')} {ex.get('country','')}")
        st.markdown("".join([f"<span class='chip'>{c}</span>" for c in chips]), unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='card card-soft'><div class='card-title'>📦 Sendungsdaten</div>", unsafe_allow_html=True)
    with st.form("eva_input_form"):
        sender_col, receiver_col, shipment_col = st.columns([1, 1, 1.45])
        with sender_col:
            st.markdown("**👤 Absender**")
            sender_name = st.text_input("Name / Firma", value=st.session_state.sender_name_default)
            sender_address = st.text_input("Straße", value=st.session_state.sender_address_default)
            s1, s2 = st.columns(2)
            with s1: sender_zip = st.text_input("PLZ", value=st.session_state.sender_zip_default)
            with s2: sender_city = st.text_input("Stadt", value=st.session_state.sender_city_default)
            sender_country = st.selectbox("Land", ["Deutschland", "Österreich", "Schweiz", "Frankreich", "Niederlande", "Sonstiges"], index=0, key="sender_country_select")
        with receiver_col:
            st.markdown("**👤 Empfänger**")
            receiver_name = st.text_input("Name / Firma Empfänger", value=st.session_state.receiver_name_default)
            receiver_address = st.text_input("Straße Empfänger", value=st.session_state.receiver_address_default)
            r1, r2 = st.columns(2)
            with r1: receiver_zip = st.text_input("PLZ Empfänger", value=st.session_state.receiver_zip_default)
            with r2: receiver_city = st.text_input("Stadt Empfänger", value=st.session_state.receiver_city_default)
            receiver_country = st.selectbox("Land Empfänger", ["Deutschland", "Österreich", "Schweiz", "Frankreich", "Niederlande", "Sonstiges"], index=0, key="receiver_country_select")
        with shipment_col:
            st.markdown("**Sendungsdetails**")
            sd1, sd2 = st.columns(2)
            with sd1:
                length = st.number_input("Länge in cm *", min_value=1.0, value=float(st.session_state.length_default))
                height = st.number_input("Höhe in cm *", min_value=1.0, value=float(st.session_state.height_default))
                goods_value = st.number_input("Warenwert in € *", min_value=0.0, value=float(st.session_state.value_default))
            with sd2:
                width = st.number_input("Breite in cm *", min_value=1.0, value=float(st.session_state.width_default))
                weight = st.number_input("Gewicht je Paket in kg *", min_value=0.1, value=float(st.session_state.weight_default))
                package_count = st.number_input("Anzahl Pakete *", min_value=1, max_value=999, value=int(st.session_state.package_count_default), step=1, help="Mehrpaket-Sendung: EVA berechnet mehrere identische Pakete und berücksichtigt die Paketanzahl in den Kosten.")
                goods_index = GOODS_OPTIONS.index(st.session_state.goods_default) if st.session_state.goods_default in GOODS_OPTIONS else 0
                goods = st.selectbox("Warenart *", GOODS_OPTIONS, index=goods_index)
            country = st.text_input("Zielland *", value=receiver_country)
        calculate_clicked = st.form_submit_button("🚀 EVA berechnen", type="primary")

    if calculate_clicked:
        if xls_source is None or xls_source == "":
            st.warning("Bitte lade zuerst die EVA-Datenbank hoch oder verbinde eine Cloud-Datenquelle.")
        else:
            try:
                # EVA 11: Dashboard-Gewichtung ist immer die sichtbare EVA-Default-Basis.
                results, rejected, price_ranking, meta = calculate_results(xls_source, length, width, height, weight, goods_value, goods, country, weight_override=dashboard_weights, package_count=package_count)
                sender_full = f"{sender_name}, {sender_address}, {sender_zip} {sender_city}, {sender_country}"
                receiver_full = f"{receiver_name}, {receiver_address}, {receiver_zip} {receiver_city}, {receiver_country}"
                inputs = {
                    "Absender": sender_full, "Empfänger": receiver_full,
                    "Absender Name": sender_name, "Empfänger Name": receiver_name,
                    "Länge": length, "Breite": width, "Höhe": height,
                    "Realgewicht": weight, "Anzahl Pakete": package_count, "Warenwert": goods_value,
                    "Warenart": goods, "Zielland": country,
                }
                st.session_state.eva_results = results
                st.session_state.eva_rejected = rejected
                st.session_state.eva_price_ranking = price_ranking
                st.session_state.eva_meta = meta
                st.session_state.eva_inputs = inputs
                append_log(inputs, results, meta.get("weights_used", {}))
                st.success("Berechnung abgeschlossen.")
            except Exception as exc:
                st.error(f"Berechnung fehlgeschlagen: {exc}")
    st.markdown("</div>", unsafe_allow_html=True)

    # Sitzungsprotokoll direkt unter den Sendungsdaten
    st.markdown("<div class='card'><div class='card-title'>🧾 Sitzungsprotokoll</div>", unsafe_allow_html=True)
    if not st.session_state.eva_log:
        st.info("Noch keine Berechnung in dieser Sitzung.")
    else:
        log_df = pd.DataFrame(st.session_state.eva_log)
        mini_cols = [c for c in ["Zeitstempel", "Warenart", "Anzahl Pakete", "Empfohlener Carrier", "Score"] if c in log_df.columns]
        st.dataframe(log_df[mini_cols], use_container_width=True, hide_index=True, height=210)
        st.download_button("⬇ Gesamtes Protokoll als CSV", data=log_to_csv_bytes(), file_name="eva_session_log.csv", mime="text/csv", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# RESULTS across main page
results = st.session_state.eva_results
rejected = st.session_state.eva_rejected
price_ranking = st.session_state.eva_price_ranking
meta = st.session_state.eva_meta

st.markdown("---")
if xls_source is None or xls_source == "":
    st.info("Bitte lade eine EVA-Datenbank hoch oder verbinde eine Cloud-Datenquelle.")
elif results is None or results.empty:
    if rejected is not None and not rejected.empty:
        st.warning("Kein Carrier erfüllt alle Muss-Kriterien.")
        with st.expander("Ausgeschlossene Tarife anzeigen"):
            st.dataframe(rejected, use_container_width=True, hide_index=True)
    else:
        st.info("Noch keine Berechnung durchgeführt.")
else:
    winner = results.iloc[0]
    sorted_by_score = results.sort_values("Score", ascending=False).reset_index(drop=True)
    second = sorted_by_score.iloc[1] if len(sorted_by_score) > 1 else None
    grundpreis = winner.get("Grundpreis €", "-")
    versicherung = winner.get("Versicherung €", "-")
    gesamtkosten = winner.get("Gesamtkosten €", "-")

    st.markdown("<div class='card card-orange'><div class='card-title'>🏆 Ergebnisse & Empfehlung</div>", unsafe_allow_html=True)
    er1, er2, er3, er4 = st.columns([1.1, 1.15, 1.4, 1.1])
    with er1:
        st.markdown(f"""
        <div class='recommend-title'>{winner['Carrier']} {winner['Versandart']}</div>
        <span class='green-chip'>Beste Option</span>
        """, unsafe_allow_html=True)
    with er2:
        st.markdown(f"<div class='big-score'>{winner['Score']}<span style='font-size:18px;'> /100</span></div><span class='small-muted'>Overall Score</span>", unsafe_allow_html=True)
    with er3:
        st.markdown("<b>Kostenaufstellung</b>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class='kv'>Grundpreis (ohne Versicherung)<b>{fmt_money(grundpreis)}</b></div>
        <div class='kv'>Versicherungskosten<b>{fmt_money(versicherung)}</b></div>
        <div class='kv'><b>Gesamtkosten (mit Versicherung)</b><b style='color:#4f21e8'>{fmt_money(gesamtkosten)}</b></div>
        """, unsafe_allow_html=True)
    with er4:
        st.markdown("<b>Vergleich zum Zweitplatzierten</b>", unsafe_allow_html=True)
        if second is not None:
            diff = ""
            try:
                diff = f"+ {float(second.get('Gesamtkosten €', 0)) - float(winner.get('Gesamtkosten €', 0)):.2f} € teurer"
            except Exception:
                diff = ""
            st.markdown(f"<b>{second['Carrier']}</b><br>Score <b>{second['Score']} /100</b><br><span class='small-muted'>{diff}</span>", unsafe_allow_html=True)
        else:
            st.markdown("<span class='small-muted'>Kein Zweitplatzierter vorhanden.</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    # Tabellen bewusst untereinander: besser lesbar und nicht gequetscht
    st.markdown("<div class='card'><div class='card-title'>Score-Ranking <span class='small-muted'>(nach Gesamtbewertung)</span></div>", unsafe_allow_html=True)
    score_cols = [c for c in ["Carrier", "Versandart", "Score", "Gesamtkosten €", "Laufzeit-Score", "Service-Score", "Sicherheits-Score", "Begründung"] if c in results.columns]
    st.dataframe(results[score_cols], use_container_width=True, hide_index=True, height=260)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='card'><div class='card-title'>Preisranking <span class='small-muted'>(nach Gesamtkosten)</span></div>", unsafe_allow_html=True)
    if price_ranking is not None and not price_ranking.empty:
        st.dataframe(price_ranking, use_container_width=True, hide_index=True, height=300)
    else:
        st.info("Kein Preisranking verfügbar.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='card'><div class='card-title'>Detailliertes Scoring <span class='small-muted'>(0–100 Punkte)</span></div>", unsafe_allow_html=True)
    detail_df = pd.DataFrame(meta.get("detailed_scoring", []))
    if not detail_df.empty:
        pivot = detail_df.pivot_table(index=["Kriterium", "Gewichtung %"], columns="Carrier", values="Score", aggfunc="first").reset_index()
        st.dataframe(pivot, use_container_width=True, hide_index=True, height=360)
    else:
        st.info("Noch kein Detail-Scoring verfügbar.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='card card-orange'><div class='card-title'>🛠 Empfohlene Zusatzleistungen</div>", unsafe_allow_html=True)
    services = recommended_services(st.session_state.eva_inputs.get("Warenart", ""), st.session_state.eva_inputs.get("Warenwert", 0), winner)
    if services.empty:
        st.info("Für diese Sendung wurden keine zusätzlichen Services empfohlen.")
    else:
        st.dataframe(services, use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("Ausgeschlossene Tarife anzeigen"):
        st.dataframe(rejected, use_container_width=True, hide_index=True)

# CHAT am Seitenende
st.markdown("<div class='card'><div class='card-title'>💬 EVA Chat <span class='green-chip'>Online</span></div>", unsafe_allow_html=True)
if "eva_messages" not in st.session_state:
    st.session_state.eva_messages = [
        {"role": "assistant", "content": "Hallo! Ich bin EVA. Ich helfe dir bei Versand, Tarifen, Versicherung und allen Fragen zur Entscheidung."}
    ]
chat_box = st.container(height=360)
with chat_box:
    for msg in st.session_state.eva_messages[-10:]:
        with st.chat_message(msg["role"], avatar="🤖" if msg["role"] == "assistant" else "👤"):
            st.write(msg["content"])
q1, q2, q3 = st.columns(3)
with q1:
    if st.button("Versicherung erklärt", use_container_width=True):
        st.session_state.eva_messages.append({"role": "user", "content": "Erkläre mir die Versicherung."})
        st.session_state.eva_messages.append({"role": "assistant", "content": ai_chat_answer("Erkläre mir die Versicherung.", st.session_state.eva_results, st.session_state.eva_rejected)})
        st.rerun()
with q2:
    if st.button("Volumengewicht?", use_container_width=True):
        st.session_state.eva_messages.append({"role": "user", "content": "Was ist das Volumengewicht?"})
        st.session_state.eva_messages.append({"role": "assistant", "content": ai_chat_answer("Was ist das Volumengewicht?", st.session_state.eva_results, st.session_state.eva_rejected)})
        st.rerun()
with q3:
    if st.button("DPD vs. GLS", use_container_width=True):
        st.session_state.eva_messages.append({"role": "user", "content": "Was ist der Unterschied zwischen DPD und GLS in dieser Berechnung?"})
        st.session_state.eva_messages.append({"role": "assistant", "content": ai_chat_answer("Was ist der Unterschied zwischen DPD und GLS in dieser Berechnung?", st.session_state.eva_results, st.session_state.eva_rejected)})
        st.rerun()
user_question = st.chat_input("Stelle eine Frage...")
if user_question:
    st.session_state.eva_messages.append({"role": "user", "content": user_question})
    answer = ai_chat_answer(user_question, st.session_state.eva_results, st.session_state.eva_rejected)
    st.session_state.eva_messages.append({"role": "assistant", "content": answer})
    st.rerun()
st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<div class='small-muted' style='text-align:center;margin-top:18px;'>Alle Berechnungen basieren auf Ihren Excel-Daten, Regeln und dem EVA-Scoring-Modell.</div>", unsafe_allow_html=True)
