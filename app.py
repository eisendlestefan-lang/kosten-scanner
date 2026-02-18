import io
import re
import math
import pandas as pd
import numpy as np
import streamlit as st
from rapidfuzz import fuzz, process
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4

st.set_page_config(page_title="KI Kosten-Scanner (CSV)", layout="wide")

# ----------------------------
# Helpers
# ----------------------------
EXPECTED_COLS = [
    "date / data",
    "amount_net / importo_netto",
    "vendor_raw / fornitore_originale",
    "account / conto",
    "text / descrizione",
    "vat_rate / aliquota_iva",
    "document_no / numero_documento",
    "cost_center / centro_di_costo",
]

SAVING_RATES = {
    "Leasing/Fahrzeug": 0.05,
    "Marketing/Werbung": 0.15,
    "Energie": 0.10,
    "Software/IT": 0.20,
    "Gebühren/Pflicht": 0.00,
    "Beratung/Dienstleistung": 0.05,
    "Sonstiges": 0.10,
}

def parse_date(s):
    # accepts YYYY-MM-DD or DD.MM.YYYY
    return pd.to_datetime(s, errors="coerce", dayfirst=True)

def norm_vendor(v: str) -> str:
    if not isinstance(v, str):
        return ""
    v = v.strip().lower()
    v = re.sub(r"\s+", " ", v)
    v = v.replace("&", " und ")
    v = re.sub(r"[^a-z0-9äöüßàèéìòù \-\.]", "", v)
    return v.strip()

def categorize(vendor_norm: str, account: str, text: str) -> str:
    n = (vendor_norm or "").lower()
    a = (account or "").lower()
    t = (text or "").lower()

    if "arval" in n or "lease" in n or "leasing" in t:
        return "Leasing/Fahrzeug"
    if "meta" in n or "facebook" in n or "ads" in t:
        return "Marketing/Werbung"
    if "hera" in n or "alperia" in n or "energie" in t or "strom" in t or "gas" in t:
        return "Energie"
    if "aruba" in n or "register" in n or "apple" in n or "microsoft" in n or "google" in n or "adobe" in n:
        return "Software/IT"
    if "gemeinde" in n or "handelskammer" in n or "camera di commercio" in t:
        return "Gebühren/Pflicht"
    if "rst" in n or "steuer" in t or "commercialista" in t:
        return "Beratung/Dienstleistung"

    # Konto-Hints (wenn bei euch vorhanden)
    if a.startswith("71.") or "software" in a:
        return "Software/IT"
    if "versicherung" in t or "assicur" in t:
        return "Sonstiges"

    return "Sonstiges"

def guess_frequency(dates: pd.Series) -> str:
    dates = dates.dropna().sort_values()
    if len(dates) < 3:
        return "unklar"
    deltas = dates.diff().dt.days.dropna()
    if deltas.empty:
        return "unklar"
    med = deltas.median()
    if 25 <= med <= 35:
        return "monatlich"
    if 55 <= med <= 75:
        return "zweimonatlich"
    if 80 <= med <= 110:
        return "quartal"
    if 330 <= med <= 400:
        return "jährlich"
    return "unklar"

def build_pdf(summary_rows: list[tuple], title="Pilotreport – Kosten-Scanner"):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph(title, styles["Heading1"]))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph("Automatisch generierter Kurzbericht (MVP).", styles["Normal"]))
    elements.append(Spacer(1, 10))

    data = [["Kennzahl", "Wert"]] + [[k, v] for (k, v) in summary_rows]
    table = Table(data, colWidths=[260, 260])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    elements.append(table)
    doc.build(elements)
    buf.seek(0)
    return buf

# ----------------------------
# UI
# ----------------------------
st.title("KI Kosten-Scanner – CSV Upload (internes MVP)")

with st.expander("CSV-Anforderungen (DE/IT)", expanded=False):
    st.markdown(
        "- Nutze idealerweise das zweisprachige Template.\n"
        "- Pflichtspalten: **date / data**, **amount_net / importo_netto**, **vendor_raw / fornitore_originale**\n"
        "- Betrag netto als Zahl (Aufwand positiv)."
    )

uploaded = st.file_uploader("CSV hochladen", type=["csv"])

colA, colB = st.columns([2,1])
with colB:
    year_min = st.number_input("Startjahr (Filter)", value=2023, step=1)
    year_max = st.number_input("Endjahr (Filter)", value=2025, step=1)
    alarm_threshold_pct = st.slider("Alarm-Schwelle ± %", 5, 50, 10)
    alarm_min_base = st.number_input("Alarm-Minimum Basis (€)", value=100, step=10)

if not uploaded:
    st.info("Bitte CSV hochladen, um die Analyse zu starten.")
    st.stop()

# Read CSV (auto separator)
raw = uploaded.read()
try:
    df = pd.read_csv(io.BytesIO(raw), sep=None, engine="python")
except Exception:
    st.error("CSV konnte nicht gelesen werden. Bitte Trennzeichen/Encoding prüfen.")
    st.stop()

# Try to auto-map columns if using template
col_map = {}
for c in df.columns:
    lc = c.strip().lower()
    if "date" in lc or "data" in lc:
        col_map["date"] = c
    if "amount" in lc or "importo" in lc:
        col_map["amount"] = c
    if "vendor" in lc or "fornitore" in lc:
        col_map["vendor"] = c
    if "account" in lc or "conto" in lc:
        col_map["account"] = c
    if "text" in lc or "descrizione" in lc:
        col_map["text"] = c

missing = [k for k in ["date","amount","vendor"] if k not in col_map]
if missing:
    st.error(f"Fehlende Spalten (oder nicht erkannt): {missing}. Benenne Spalten wie im Template.")
    st.stop()

# Normalize dataframe
d = pd.DataFrame()
d["date"] = df[col_map["date"]].apply(parse_date)
d["year"] = d["date"].dt.year
d["amount_net"] = pd.to_numeric(df[col_map["amount"]], errors="coerce").fillna(0.0)
d["vendor_raw"] = df[col_map["vendor"]].astype(str)
d["account"] = df[col_map.get("account","")].astype(str) if col_map.get("account") in df.columns else ""
d["text"] = df[col_map.get("text","")].astype(str) if col_map.get("text") in df.columns else ""

# Filter years and expenses
d = d[(d["year"].between(year_min, year_max)) & (d["amount_net"] > 0)].copy()
if d.empty:
    st.warning("Nach Filterung ist keine Datenzeile übrig (Jahre/Beträge prüfen).")
    st.stop()

d["vendor_norm"] = d["vendor_raw"].apply(norm_vendor)
d["category"] = d.apply(lambda r: categorize(r["vendor_norm"], r["account"], r["text"]), axis=1)

# ----------------------------
# 1) Fixkosten / recurring detection
# ----------------------------
vendor_year = d.groupby(["vendor_norm","year"], as_index=False).agg(
    sum_net=("amount_net","sum"),
    count=("amount_net","size"),
)
pivot = vendor_year.pivot(index="vendor_norm", columns="year", values="sum_net").fillna(0.0)
pivot_counts = vendor_year.pivot(index="vendor_norm", columns="year", values="count").fillna(0.0)
years = sorted([y for y in pivot.columns if isinstance(y, (int, np.integer))])

total = pivot.sum(axis=1).rename("total").to_frame()
years_nonzero = (pivot > 0).sum(axis=1).rename("years_nonzero").to_frame()
recurring = total.join(years_nonzero)
recurring["latest_year"] = max(years) if years else None
recurring["latest_cost"] = pivot[recurring["latest_year"]] if years else 0.0
recurring = recurring.sort_values("total", ascending=False)

# add freq guess using raw dates for vendors
freq_map = {}
for v, grp in d.groupby("vendor_norm"):
    freq_map[v] = guess_frequency(grp["date"])
recurring["freq_guess"] = recurring.index.map(freq_map)

# add category (most common)
cat_map = d.groupby("vendor_norm")["category"].agg(lambda x: x.value_counts().index[0])
recurring["category"] = recurring.index.map(cat_map)

# Heuristic recurring flag
recurring["recurring_flag"] = (recurring["years_nonzero"] >= 2) | (recurring["freq_guess"].isin(["monatlich","quartal","zweimonatlich","jährlich"]))
recurring_view = recurring.reset_index().rename(columns={"vendor_norm":"anbieter"})

# ----------------------------
# 2) Trend table (top)
# ----------------------------
trend = recurring_view[["anbieter","category"] + years + ["total","years_nonzero","freq_guess","recurring_flag"]].copy()

# ----------------------------
# 3) Alarms
# ----------------------------
alarms = []
TH = float(alarm_threshold_pct)
MIN = float(alarm_min_base)

for v in pivot.index:
    row = pivot.loc[v]
    for y1, y2 in zip(years[:-1], years[1:]):
        base = float(row.get(y1, 0.0))
        new = float(row.get(y2, 0.0))
        if base == 0 and new >= MIN:
            alarms.append((v, f"{y1}->{y2}", "NEU", base, new, new, None))
        elif base >= MIN and new == 0:
            alarms.append((v, f"{y1}->{y2}", "WEG", base, new, -base, -100.0))
        elif base >= MIN and new > 0:
            pct = (new - base) / base * 100.0
            if abs(pct) >= TH:
                alarms.append((v, f"{y1}->{y2}", "ÄNDERUNG", base, new, new-base, pct))

alarms_df = pd.DataFrame(alarms, columns=["anbieter","periode","typ","basis","neu","delta","pct"])
if not alarms_df.empty:
    alarms_df["category"] = alarms_df["anbieter"].map(cat_map).fillna("Sonstiges")
    alarms_df = alarms_df.sort_values(["periode", "delta"], ascending=[True, False])

# ----------------------------
# 4) Savings potential (latest year)
# ----------------------------
latest_year = max(years) if years else None
savings = vendor_year[vendor_year["year"] == latest_year].copy() if latest_year else vendor_year.copy()
savings["category"] = savings["vendor_norm"].map(cat_map).fillna("Sonstiges")
savings["rate"] = savings["category"].map(SAVING_RATES).fillna(0.10)
savings["potential_eur"] = (savings["sum_net"] * savings["rate"]).round(2)
savings = savings.sort_values("potential_eur", ascending=False)

# Summary metrics
relevant_cost = float(savings["sum_net"].sum()) if latest_year else float(vendor_year["sum_net"].sum())
potential_total = float(savings["potential_eur"].sum()) if latest_year else 0.0
potential_pct = (potential_total / relevant_cost * 100.0) if relevant_cost > 0 else 0.0

# ----------------------------
# UI tabs
# ----------------------------
tab1, tab2, tab3, tab4 = st.tabs(["Fixkosten", "Trends", "Alarme", "Einsparpotenzial"])

with tab1:
    st.subheader("Fixkosten / Wiederkehrende Anbieter")
    st.caption("Heuristik: ≥2 Jahre vorhanden ODER Frequenz-Schätzung (monatlich/quartal/jährlich).")
    st.dataframe(
        trend.sort_values("total", ascending=False).head(50),
        use_container_width=True
    )

with tab2:
    st.subheader("Trendanalyse")
    st.dataframe(
        trend.sort_values("total", ascending=False).head(50),
        use_container_width=True
    )

with tab3:
    st.subheader("Alarm-Liste")
    st.caption(f"Schwelle: ±{alarm_threshold_pct}% ab {alarm_min_base}€ Basis. NEU/WEG ab {alarm_min_base}€.")
    if alarms_df.empty:
        st.info("Keine Alarme nach den aktuellen Schwellenwerten.")
    else:
        st.dataframe(alarms_df, use_container_width=True)

with tab4:
    st.subheader(f"Einsparpotenzial (konservativ, Jahr {latest_year})")
    st.metric("Relevante Kosten", f"{relevant_cost:,.2f} €")
    st.metric("Einsparpotenzial (Schätzung)", f"{potential_total:,.2f} €")
    st.metric("Potenzialquote", f"{potential_pct:.1f} %")
    st.dataframe(savings.rename(columns={
        "vendor_norm":"anbieter",
        "sum_net":"kosten",
        "count":"buchungen"
    })[["anbieter","category","kosten","rate","potential_eur","buchungen"]].head(50), use_container_width=True)

# Downloads
st.divider()
col1, col2, col3 = st.columns([1,1,2])

with col1:
    # Excel export
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        trend.to_excel(writer, index=False, sheet_name="trends_fixkosten")
        if not alarms_df.empty:
            alarms_df.to_excel(writer, index=False, sheet_name="alarme")
        savings.to_excel(writer, index=False, sheet_name="einsparpotenzial")
    out.seek(0)
    st.download_button("Ergebnisse als Excel", data=out, file_name="kosten_scanner_ergebnisse.xlsx")

with col2:
    pdf_buf = build_pdf([
        ("Zeitraum", f"{year_min}–{year_max}"),
        ("Relevante Kosten (letztes Jahr)", f"{relevant_cost:,.2f} €"),
        ("Einsparpotenzial (konservativ)", f"{potential_total:,.2f} €"),
        ("Potenzialquote", f"{potential_pct:.1f} %"),
        ("Alarme", str(len(alarms_df)) if not alarms_df.empty else "0"),
    ], title="Pilotreport – Kosten-Scanner (internes MVP)")
    st.download_button("Kurzreport als PDF", data=pdf_buf, file_name="pilotreport_kosten_scanner.pdf")

with col3:
    st.info(
        "Hinweis: Kategorien/Einsparquoten sind MVP-Heuristiken. "
        "In der Kanzlei-Version sollte es eine Berater-Override-Funktion geben (Mapping & Quoten pro Mandant)."
    )
