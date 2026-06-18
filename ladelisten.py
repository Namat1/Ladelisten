# -*- coding: utf-8 -*-
"""
Tour-/Ladeplan-Generator (NFC / EDEKA Nord)
Logik:
- Blatt 'LADEREIHENFOLGE' (Tour, LF, SAP, Name, Strasse, PLZ, Ort) -> LF1 oben
- Kundennummer aus 'KUNDENDATEN', Telefon aus 'KUNDENLISTE', CSB aus KUNDENLISTE/Depot-Blättern per SAP
- Schlüsselnummer aus CSV wird ausschließlich über die CSB-Nummer gematcht
- Ladenummer/JPG-Nummer aus CSV wird ausschließlich über die CSB-Nummer gematcht
- Depot aus DIREKT/MK/HUPA_NMS/HUPA_MALCHOW
- Tag = 1. Ziffer der Tour (1=Mo .. 6=Sa); pro Tour eine A4-Hochkant-Seite
Design: schwarz/weiß/grau, gleichmäßiges Grid, Kundennummer sichtbar als eigene Spalte, Telefon unter Adresse, Nummern kleiner für mehr Platz in den Mengenfeldern.
"""
import io
import re
from datetime import date
from html import escape

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak

st.set_page_config(page_title="Tour-/Ladeplan", page_icon="🚚", layout="wide")

SHEET = "LADEREIHENFOLGE"
DEPOT_SHEETS = {"DIREKT": "Direkt", "MK": "Marktkauf", "HUPA_NMS": "Neumünster", "HUPA_MALCHOW": "Malchow"}
CSB_SHEETS = ["KUNDENLISTE", "DIREKT", "MK", "HUPA_NMS", "HUPA_MALCHOW"]  # Spalte A=CSB, B=SAP
TAGE = {1: "Montag", 2: "Dienstag", 3: "Mittwoch", 4: "Donnerstag", 5: "Freitag", 6: "Samstag"}
WDCOL = {1: 6, 2: 7, 3: 8, 4: 9, 5: 10, 6: 11}

# ---- Design-Tokens: schwarz/weiß/grau (Akzente: Warnung rot, Kundennummer blau)
INK = colors.HexColor("#16181C")
GRY = colors.HexColor("#3A3F45")
MUTE = colors.HexColor("#6B7075")
LINE = colors.HexColor("#D2D5D9")
SOFT = colors.HexColor("#F5F6F7")
HDR = colors.HexColor("#ECEEF0")
KDC = colors.HexColor("#1357A6")      # Kundennummer
ACC = colors.HexColor("#D81E05")      # nur Leergut
ACC_BG = colors.HexColor("#FBEAE8")

ROW_H_MIN = 12.0 * mm
FIELD_H = 10.8 * mm
FOOT_H = 14.0 * mm


# ------------------------------------------------------------------ Hilfsfunktionen

def _clean(value) -> str:
    """Lesbaren Zellinhalt zurückgeben, ohne nan/None und ohne .0 bei ganzen Zahlen."""
    if value is None or pd.isna(value):
        return ""
    txt = str(value).strip().replace("\ufeff", "")
    if txt.lower() in {"nan", "none", "nat"}:
        return ""
    txt = txt.replace("\u00a0", " ").strip()
    try:
        f = float(txt.replace(",", "."))
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return txt


def _norm_num(value) -> str:
    """Nummern für SAP/CSB-Vergleiche normalisieren: 40725, 40725.0 und ' 040725 ' werden gleich behandelt."""
    txt = _clean(value)
    if not txt:
        return ""
    digits = re.sub(r"\D", "", txt)
    if not digits:
        return ""
    try:
        return str(int(digits))
    except Exception:
        return digits.lstrip("0") or "0"


def _html(value) -> str:
    return escape(_clean(value), quote=True)


# ------------------------------------------------------------------ Daten
@st.cache_data(show_spinner=False)
def lade_daten(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    df = pd.read_excel(xls, SHEET, usecols="A:G")
    df.columns = ["Tour", "LF", "SAP", "Name", "Strasse", "PLZ", "Ort"]
    df = df.dropna(subset=["Tour"]).copy()
    df["Tour"] = pd.to_numeric(df["Tour"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["Tour"]).copy()
    df["Tour"] = df["Tour"].astype(int)
    df["LF"] = pd.to_numeric(df["LF"], errors="coerce")
    df["SAP"] = df["SAP"].apply(_norm_num)
    for c in ["Name", "Strasse", "Ort"]:
        df[c] = df[c].apply(_clean)
    df["PLZ"] = df["PLZ"].apply(lambda x: "" if pd.isna(x) else str(int(x)) if str(x).replace('.0', '').isdigit() else str(x).strip())
    df["TagZiffer"] = df["Tour"].apply(lambda t: int(str(t)[0]))

    # Kundennummer (KUNDENDATEN: A=Kundennummer, B=SAP)
    sap2kd = {}
    if "KUNDENDATEN" in xls.sheet_names:
        kd = pd.read_excel(xls, "KUNDENDATEN", usecols="A:B")
        kd.columns = ["KdNr", "SAP"]
        for _, r in kd.dropna(subset=["KdNr"]).iterrows():
            sap = _norm_num(r["SAP"])
            kdnr = _norm_num(r["KdNr"])
            if sap and kdnr:
                sap2kd.setdefault(sap, kdnr)

    # Telefon (KUNDENLISTE: Spalte B=SAP, Spalte G=Telefon)
    sap2tel = {}
    if "KUNDENLISTE" in xls.sheet_names:
        d = pd.read_excel(xls, "KUNDENLISTE")
        for _, r in d.iterrows():
            sap = _norm_num(r.iloc[1]) if d.shape[1] > 1 else ""
            tel = _clean(r.iloc[6]) if d.shape[1] > 6 else ""
            if sap and tel:
                sap2tel.setdefault(sap, tel)

    # SAP -> CSB (Spalte A=CSB, B=SAP in KUNDENLISTE + Depot-Blättern)
    sap2csb = {}
    for sh in CSB_SHEETS:
        if sh not in xls.sheet_names:
            continue
        d = pd.read_excel(xls, sh)
        if d.shape[1] < 2:
            continue
        for _, r in d.iterrows():
            csb = _norm_num(r.iloc[0])
            sap = _norm_num(r.iloc[1])
            if sap and csb and sap not in sap2csb:
                sap2csb[sap] = csb

    # Depot
    tour2dep = {}
    for sh, lab in DEPOT_SHEETS.items():
        if sh not in xls.sheet_names:
            continue
        d = pd.read_excel(xls, sh, header=0)
        for ci in WDCOL.values():
            if ci < d.shape[1]:
                for v in pd.to_numeric(d.iloc[:, ci], errors="coerce").dropna().astype(int):
                    tour2dep.setdefault(v, lab)
    return df, sap2kd, sap2tel, sap2csb, tour2dep


@st.cache_data(show_spinner=False)
def lade_nummern(csv_bytes: bytes):
    """
    Schlüsseldatei einlesen.
    Erwartung: erste echte Datenspalte = CSB-Nummer, zweite echte Datenspalte = Schlüsselnummer.
    Kopfzeilen wie 'Schlüsselkunden-NFC' oder 'Knd-Nr.;Schlüssel-Nr.' werden automatisch ignoriert.
    Rückgabe: {CSB: Schlüsselnummer}
    """
    d = None
    for sep in [";", ",", "\t"]:
        try:
            probe = pd.read_csv(
                io.BytesIO(csv_bytes),
                sep=sep,
                dtype=str,
                header=None,
                encoding="utf-8-sig",
                engine="python",
                on_bad_lines="skip",
            )
            if probe.shape[1] >= 2:
                d = probe
                break
        except Exception:
            d = None
    if d is None or d.shape[1] < 2:
        return {}

    # Leere Spalten entfernen, aber Reihenfolge behalten.
    keep_cols = []
    for col in d.columns:
        has_value = d[col].fillna("").astype(str).str.strip().ne("").any()
        if has_value:
            keep_cols.append(col)
    d = d[keep_cols]
    if d.shape[1] < 2:
        return {}

    m = {}
    for _, r in d.iterrows():
        csb = _norm_num(r.iloc[0])
        schluessel = _clean(r.iloc[1])
        if not csb or not schluessel:
            continue
        # Nur numerische CSB-Werte übernehmen. Dadurch werden Text-Kopfzeilen sicher ausgelassen.
        if not csb.isdigit():
            continue
        if schluessel.lower().startswith(("schlüssel", "nummer", "nr")):
            continue
        m.setdefault(csb, schluessel)
    return m


@st.cache_data(show_spinner=False)
def lade_ladenummern(csv_bytes: bytes):
    """
    Ladenummer-/JPG-Zuordnung einlesen.
    Erwartung bevorzugt: 'CSB Nummer;Ladenummer;...'
    Rückgabe: {CSB: Ladenummer}
    """
    d = None
    for sep in [";", ",", "\t"]:
        try:
            probe = pd.read_csv(
                io.BytesIO(csv_bytes),
                sep=sep,
                dtype=str,
                encoding="utf-8-sig",
                engine="python",
                on_bad_lines="skip",
            )
            if probe.shape[1] >= 2:
                d = probe
                break
        except Exception:
            d = None
    if d is None or d.shape[1] < 2:
        return {}

    d.columns = [_clean(c).strip().lstrip("\ufeff") for c in d.columns]
    lower_cols = {c: c.lower() for c in d.columns}

    csb_col = next((c for c, lc in lower_cols.items() if "csb" in lc), d.columns[0])
    lade_col = next(
        (c for c, lc in lower_cols.items() if "ladenummer" in lc or "lade" in lc or "jpg" in lc or "nummer" == lc),
        d.columns[1],
    )

    m = {}
    for _, r in d.iterrows():
        csb = _norm_num(r.get(csb_col, ""))
        ladenummer = _clean(r.get(lade_col, ""))
        if not csb or not ladenummer:
            continue
        if not csb.isdigit():
            continue
        if ladenummer.lower() in {"ladenummer", "jpg", "jpg nummer", "nummer"}:
            continue
        m.setdefault(csb, ladenummer)
    return m


# ------------------------------------------------------------------ Styles
def _styles():
    return {
        "title": ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=13.2, textColor=INK, leading=14.5),
        "code": ParagraphStyle("code", fontName="Helvetica", fontSize=7.2, textColor=MUTE, alignment=2, leading=8.5),
        "tour": ParagraphStyle("tour", fontName="Helvetica-Bold", fontSize=20, textColor=INK, leading=20.5),
        "tourlbl": ParagraphStyle("tourlbl", fontName="Helvetica-Bold", fontSize=6.6, textColor=MUTE, leading=7.5),
        "chip": ParagraphStyle("chip", fontName="Helvetica-Bold", fontSize=8.6, textColor=INK, leading=9.5),
        "chiplbl": ParagraphStyle("chiplbl", fontName="Helvetica", fontSize=6.1, textColor=MUTE, leading=6.9),
        "flbl": ParagraphStyle("flbl", fontName="Helvetica-Bold", fontSize=6.0, textColor=MUTE, leading=7.0),
        "warn": ParagraphStyle("warn", fontName="Helvetica-Bold", fontSize=8.5, textColor=ACC, leading=10.0),
        "thead": ParagraphStyle("thead", fontName="Helvetica-Bold", fontSize=5.7, textColor=INK, leading=6.4, alignment=1),
        "theadL": ParagraphStyle("theadL", fontName="Helvetica-Bold", fontSize=5.9, textColor=INK, leading=6.6),
        "lf": ParagraphStyle("lf", fontName="Helvetica-Bold", fontSize=9.6, textColor=INK, alignment=1, leading=10.2),
        "key": ParagraphStyle("key", fontName="Helvetica-Bold", fontSize=9.4, textColor=INK, alignment=1, leading=10.0),
        "shop": ParagraphStyle("shop", fontName="Helvetica-Bold", fontSize=9.0, textColor=INK, alignment=1, leading=9.8),
        "csb": ParagraphStyle("csb", fontName="Helvetica-Bold", fontSize=8.2, textColor=INK, alignment=1, leading=8.8),
        "cust": ParagraphStyle("cust", fontName="Helvetica", fontSize=8.1, textColor=GRY, leading=9.6),
    }


def _flabel(s, label):
    label_txt = escape(label.upper()).replace(" ", "&nbsp;")
    return Paragraph(f"<font name=Helvetica-Bold size=5.8 color='#6B7075'>{label_txt}</font>", s["flbl"])


def _meta_line(label: str, value: str) -> str:
    if not value:
        return ""
    return f"<font name=Helvetica-Bold color='#6B7075'>{escape(label)}</font>&nbsp;{_html(value)}"

def _nowrap_num(value, style):
    """Nummernzelle ohne Umbruch darstellen. Wichtig für Kundennummer und Ladenummer."""
    txt = _html(value)
    return Paragraph(f"<nobr>{txt}</nobr>", style)


def _estimate_customer_row_height(k) -> float:
    """Dynamische Zeilenhöhe, damit lange Markt-Namen/Adressen nicht in die nächste Zeile laufen."""
    text_len = max(len(_clean(k.get("name", ""))), len(_clean(k.get("adr", ""))))
    tel_extra = 1 if _clean(k.get("tel", "")) else 0
    # Grundhöhe für Name + Adresse + optional Telefon, danach Zuschläge für lange Inhalte
    lines = 2 + tel_extra
    if text_len > 34:
        lines += 1
    if text_len > 62:
        lines += 1
    if text_len > 88:
        lines += 1
    return max(ROW_H_MIN, (4.0 + lines * 3.05) * mm)



def tour_block(tour, depot, tagname, datum_txt, kunden, s, W):
    el = []

    band = Table([[Paragraph("TOUR&nbsp;/&nbsp;LADEPLAN", s["title"]),
                   Paragraph("Fleischwerk EDEKA Nord · NFC", s["code"])]],
                 colWidths=[W * 0.62, W * 0.38])
    band.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (0, 0), 0), ("RIGHTPADDING", (1, 0), (1, 0), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LINEBELOW", (0, 0), (-1, -1), 1.1, INK),
    ]))
    el.append(band)

    def chip(lbl, val):
        return [Paragraph(lbl.upper(), s["chiplbl"]), Paragraph(_html(val), s["chip"])]

    cells = [chip("Wochentag", tagname), chip("Depot", depot or "—"),
             chip("Datum", datum_txt), chip("Kunden", str(len(kunden)))]
    chips = Table(
        [[Paragraph("TOUR", s["tourlbl"]), *[c[0] for c in cells]],
         [Paragraph(str(tour), s["tour"]), *[c[1] for c in cells]]],
        colWidths=[W * 0.22, W * 0.205, W * 0.205, W * 0.20, W * 0.17])
    chips.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, -1), (-1, -1), 1.0, INK),
        ("LINEAFTER", (0, 0), (0, -1), 0.55, LINE),
        ("LEFTPADDING", (1, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, 0), 3), ("BOTTOMPADDING", (0, -1), (-1, -1), 3),
        ("TOPPADDING", (0, 1), (-1, 1), 0),
    ]))
    el.append(chips)
    el.append(Spacer(1, 3))

    # Gleichmäßiges Grid oben: 3 Reihen mit 4 gleich breiten Feldern.
    fields = [
        ["Name Fahrer", "Kennzeichen", "Tor", "LKW"],
        ["Kilometer Start", "Kilometer Ende", "Start Arbeitszeit", "Ende Arbeitszeit"],
        ["Rolli Rückgabe", "Markt-Schlüssel", "Sonstiges", ""],
    ]
    fdata = [[_flabel(s, f) if f else "" for f in row] for row in fields]
    ftab = Table(fdata, colWidths=[W / 4] * 4, rowHeights=[FIELD_H] * 3)
    ftab.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.55, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
    ]))
    el.append(ftab)
    el.append(Spacer(1, 3))

    warn = Table([[Paragraph("ACHTUNG — zwingend gesamtes Leergut abräumen.", s["warn"])]], colWidths=[W])
    warn.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ACC_BG),
        ("LINEBEFORE", (0, 0), (0, -1), 3, ACC),
        ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    el.append(warn)
    el.append(Spacer(1, 4))

    # Kundentabelle: Kundennummer sichtbar als eigene, ruhige Spalte.
    # Telefon steht direkt unter der Adresse.
    # Nummernspalten kleiner, damit rechts mehr Platz für die Mengen bleibt.
    # Kundenzeilen wachsen dynamisch, damit lange Namen und Adressen nicht überlappen.
    cw_mm = [6.5, 15, 11.5, 13.5, None, 12.5, 12.5, 12.5, 12.5, 12.5, 10, 10]
    fixed = sum(x for x in cw_mm if x is not None) * mm
    cw_mm[4] = max(50, (W - fixed) / mm)
    cw = [x * mm for x in cw_mm]

    head0 = [
        Paragraph("LF", s["thead"]),
        Paragraph("MARKT-<br/>SCHLÜSSEL", s["thead"]),
        Paragraph("LADENR.", s["thead"]),
        Paragraph("KUNDEN-<br/>NUMMER", s["thead"]),
        Paragraph("KUNDE / ADRESSE / TELEFON", s["theadL"]),
        Paragraph("TRANSPORT", s["thead"]), "", "", "", "",
        Paragraph("ZEIT", s["thead"]), "",
    ]
    head1 = ["", "", "", "", "",
             Paragraph("PA/TKT", s["thead"]), Paragraph("RO", s["thead"]),
             Paragraph("E2", s["thead"]), Paragraph("E1", s["thead"]),
             Paragraph("KT", s["thead"]),
             Paragraph("von", s["thead"]), Paragraph("bis", s["thead"])]
    data = [head0, head1]
    row_heights = [None, None]

    for k in kunden:
        name = f"<font name=Helvetica-Bold size=8.3 color='#16181C'>{_html(k['name'])}</font>"
        adr = f"<font name=Helvetica size=7.0 color='#3A3F45'>{_html(k['adr'])}</font>"
        tel = ""
        if k.get("tel", ""):
            tel = f"<br/><font name=Helvetica size=6.8 color='#6B7075'>Tel. {_html(k['tel'])}</font>"
        kunde_cell = Paragraph(f"{name}<br/>{adr}{tel}", s["cust"])

        data.append([
            _nowrap_num(k["lf"], s["lf"]),
            _nowrap_num(k.get("nr", ""), s["key"]),
            _nowrap_num(k.get("ladenr", ""), s["shop"]),
            _nowrap_num(k.get("csb", ""), s["csb"]),
            kunde_cell,
            "", "", "", "", "", "", "",
        ])
        row_heights.append(_estimate_customer_row_height(k))

    tab = Table(data, colWidths=cw, rowHeights=row_heights, repeatRows=2)
    tab.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 1), HDR),
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (1, 1)),
        ("SPAN", (2, 0), (2, 1)),
        ("SPAN", (3, 0), (3, 1)),
        ("SPAN", (4, 0), (4, 1)),
        ("SPAN", (5, 0), (9, 0)),
        ("SPAN", (10, 0), (11, 0)),
        ("VALIGN", (0, 0), (-1, 1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 1), 3), ("BOTTOMPADDING", (0, 0), (-1, 1), 3),
        ("LINEBELOW", (0, 1), (-1, 1), 0.8, INK),
        ("VALIGN", (0, 2), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 2), (-1, -1), 2),
        ("RIGHTPADDING", (0, 2), (-1, -1), 2),
        ("LEFTPADDING", (4, 2), (4, -1), 4),
        ("TOPPADDING", (0, 2), (-1, -1), 2), ("BOTTOMPADDING", (0, 2), (-1, -1), 2),
        ("LINEBELOW", (0, 1), (-1, -1), 0.5, LINE),
        ("LINEAFTER", (0, 0), (4, -1), 0.5, LINE),
        ("INNERGRID", (5, 2), (11, -1), 0.5, LINE),
        ("LINEAFTER", (9, 0), (9, -1), 0.5, LINE),
        ("BOX", (0, 0), (-1, -1), 0.8, INK),
    ]))
    el.append(tab)
    el.append(Spacer(1, 4))

    foot = Table([[_flabel(s, "Summe Rollis"), _flabel(s, "Unterschrift Fahrer")]],
                 colWidths=[W * 0.28, W * 0.72], rowHeights=[FOOT_H])
    foot.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.55, LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
    ]))
    el.append(foot)
    return el


def baue_pdf(df_tag, sap2kd, sap2tel, sap2csb, csb2num, csb2laden, tour2dep, tagname, datum_txt):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=4 * mm, rightMargin=4 * mm,
                            topMargin=5 * mm, bottomMargin=5 * mm, title=f"Ladeplan {tagname}")
    s = _styles()
    W = doc.width
    story = []
    touren = sorted(df_tag["Tour"].unique())
    for i, tour in enumerate(touren):
        sub = df_tag[df_tag["Tour"] == tour].sort_values(["LF", "SAP"], na_position="last")
        kunden = []
        for _, r in sub.iterrows():
            sap = _norm_num(r["SAP"])
            csb = sap2csb.get(sap, "")
            plz_ort = " ".join([x for x in [_clean(r["PLZ"]), _clean(r["Ort"])] if x]).strip()
            adr = " · ".join([x for x in [_clean(r["Strasse"]), plz_ort] if x]).strip()
            kunden.append({
                "lf": "" if pd.isna(r["LF"]) else str(int(r["LF"])),
                "nr": csb2num.get(csb, ""),        # Markt-Schlüssel per Kundennummer/CSB
                "ladenr": csb2laden.get(csb, ""),  # Ladenummer/JPG-Nummer per Kundennummer/CSB
                "csb": csb,                        # sichtbar als Kundennummer
                "kd": sap2kd.get(sap, ""),         # nur intern / Debug
                "tel": sap2tel.get(sap, ""),
                "name": _clean(r["Name"]),
                "adr": adr,
            })
        story += tour_block(tour, tour2dep.get(tour, ""), tagname, datum_txt, kunden, s, W)
        if i < len(touren) - 1:
            story.append(PageBreak())
    doc.build(story)
    return buf.getvalue()


# ------------------------------------------------------------------ UI
st.title("🚚 Tour-/Ladeplan-Generator")
st.caption("Kompakte Version: nutzt fast die ganze Seitenbreite, kleinere Abstände, kleinere Nummern und automatische Zeilenhöhen bei langen Kundeneinträgen.")

up = st.file_uploader("1) Quelldatei (.xlsx)", type=["xlsx"], key="quelldatei_upload")
csv_up = st.file_uploader("2) Schlüsseldatei (.csv) — optional", type=["csv"], key="schluesseldatei_upload")
laden_up = st.file_uploader("3) Ladenummer-/JPG-Zuordnung (.csv) — optional", type=["csv"], key="ladenummer_upload")
if not up:
    st.info("Quelldatei mit Blatt **LADEREIHENFOLGE** hochladen.")
    st.stop()

df, sap2kd, sap2tel, sap2csb, tour2dep = lade_daten(up.getvalue())
csb2num = lade_nummern(csv_up.getvalue()) if csv_up else {}
csb2laden = lade_ladenummern(laden_up.getvalue()) if laden_up else {}
saps = df["SAP"].apply(_norm_num)
csb_gefunden = saps.map(lambda sap: sap2csb.get(sap, "") != "").sum()
if csv_up:
    hit = saps.map(lambda sap: csb2num.get(sap2csb.get(sap, ""), "") != "").sum()
    st.caption(f"Schlüssel-Matching per CSB-Nummer: {hit}/{len(df)} Kunden · Schlüsseldatei: {len(csb2num)} Zuordnungen")
    if len(csb2num) == 0:
        st.warning("In der Schlüsseldatei wurden keine verwertbaren Zuordnungen gefunden. Erwartet wird: erste Datenspalte CSB-Nummer, zweite Datenspalte Schlüsselnummer.")
if laden_up:
    hit_laden = saps.map(lambda sap: csb2laden.get(sap2csb.get(sap, ""), "") != "").sum()
    st.caption(f"Ladenummer-Matching per CSB-Nummer: {hit_laden}/{len(df)} Kunden · CSB im Kundenstamm gefunden: {csb_gefunden}/{len(df)} · Ladenummerdatei: {len(csb2laden)} Zuordnungen")
    if len(csb2laden) == 0:
        st.warning("In der Ladenummer-Datei wurden keine verwertbaren Zuordnungen gefunden. Erwartet wird zum Beispiel: CSB Nummer;Ladenummer;Kundenname;Ort")

verf = sorted(df["TagZiffer"].unique())
opt = {f"{TAGE[z]}  ({df.loc[df['TagZiffer'] == z, 'Tour'].nunique()} Touren)": z for z in verf}
c1, c2 = st.columns([2, 1])
with c1:
    aus = st.selectbox("Tag (= 1. Ziffer der Tour)", list(opt.keys()), key="tag_select")
z = opt[aus]
with c2:
    datum = st.date_input("Druckdatum", value=date.today(), format="DD.MM.YYYY", key="druckdatum")
df_tag = df[df["TagZiffer"] == z]
st.caption(f"{df_tag['Tour'].nunique()} Touren · {len(df_tag)} Kunden für **{TAGE[z]}**")

if st.button("📄 Tour-/Ladepläne erzeugen", type="primary", use_container_width=True, key="pdf_erzeugen"):
    with st.spinner("Erzeuge PDF …"):
        pdf = baue_pdf(df_tag, sap2kd, sap2tel, sap2csb, csb2num, csb2laden, tour2dep, TAGE[z], datum.strftime("%d.%m.%Y"))
    st.success(f"{df_tag['Tour'].nunique()} Tour-Seiten erzeugt.")
    st.download_button("⬇️ PDF herunterladen", data=pdf,
                       file_name=f"Ladeplan_{TAGE[z]}_{datum.strftime('%Y%m%d')}.pdf",
                       mime="application/pdf", use_container_width=True)
