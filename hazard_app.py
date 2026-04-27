"""
PPS Rail — NWR Hazard Directory, Access Points & Signal Box Contacts
Streamlit Web Application
=========================================================
Auto-loads all hazard directory CSV files from the data subfolder.
Enter an ELR and mileage range — get hazards, access points,
and signal box contacts in one search. Download each as PDF.
"""

import streamlit as st
import pandas as pd
import re
import os
import io
import glob
from datetime import datetime

# PDF generation
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus.doctemplate import PageTemplate, BaseDocTemplate, Frame
import base64

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PPS Rail — Hazard Directory",
    page_icon="⚠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Brand colours ─────────────────────────────────────────────────────────────
COLOURS = {
    'bg':       '#0a0a0a',
    'surface':  '#141414',
    'surface2': '#1e1e1e',
    'border':   '#2a2a2a',
    'red':      '#CC2200',
    'amber':    '#F5A800',
    'green':    '#00A651',
    'nwr_blue': '#003366',
    'text':     '#e0e0e0',
    'muted':    '#888888',
    'white':    '#ffffff',
}

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700&family=Barlow:wght@300;400;500&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Barlow', sans-serif;
    background-color: {COLOURS['bg']};
    color: {COLOURS['text']};
  }}

  header[data-testid="stHeader"] {{
    background-color: {COLOURS['bg']} !important;
    color: {COLOURS['bg']} !important;
  }}

  .stDeployButton {{
    display: none !important;
  }}

  section[data-testid="stSidebar"] {{
    background-color: {COLOURS['surface']} !important;
    border-right: 1px solid {COLOURS['border']};
  }}
  section[data-testid="stSidebar"] * {{
    color: {COLOURS['text']} !important;
  }}

  .main .block-container {{
    background-color: {COLOURS['bg']};
    padding-top: 1rem;
  }}

  [data-testid="stAppViewContainer"],
  [data-testid="stAppViewBlockContainer"],
  [data-testid="block-container"] {{
    background-color: {COLOURS['bg']} !important;
  }}

  .stApp {{
    background-color: {COLOURS['bg']} !important;
  }}

  h1, h2, h3, h4 {{
    font-family: 'Barlow Condensed', sans-serif !important;
    letter-spacing: 0.04em;
    color: {COLOURS['white']} !important;
  }}

  .pps-card {{
    background: {COLOURS['surface']};
    border: 1px solid {COLOURS['border']};
    border-radius: 4px;
    padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
  }}
  .pps-card-green {{ border-left: 3px solid {COLOURS['green']}; }}
  .pps-card-amber {{ border-left: 3px solid {COLOURS['amber']}; }}
  .pps-card-grey {{ border-left: 3px solid {COLOURS['muted']}; }}

  .badge {{
    display: inline-block;
    padding: 2px 10px;
    border-radius: 2px;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }}
  .badge-green {{ background: {COLOURS['green']}22; color: {COLOURS['green']}; border: 1px solid {COLOURS['green']}55; }}
  .badge-amber {{ background: {COLOURS['amber']}22; color: {COLOURS['amber']}; border: 1px solid {COLOURS['amber']}55; }}
  .badge-grey {{ background: {COLOURS['muted']}22; color: {COLOURS['muted']}; border: 1px solid {COLOURS['muted']}55; }}

  .metric-box {{
    background: {COLOURS['surface']};
    border: 1px solid {COLOURS['border']};
    border-radius: 4px;
    padding: 0.8rem 1rem;
    text-align: center;
  }}
  .metric-num {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    line-height: 1;
  }}
  .metric-lbl {{
    font-size: 0.72rem;
    color: {COLOURS['muted']};
    text-transform: uppercase;
    letter-spacing: 0.08em;
    margin-top: 0.2rem;
  }}

  .pps-divider {{
    border: none;
    border-top: 1px solid {COLOURS['border']};
    margin: 1rem 0;
  }}

  .section-header {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: {COLOURS['white']};
    letter-spacing: 0.04em;
    padding: 0.6rem 1rem;
    background: {COLOURS['surface']};
    border-left: 3px solid {COLOURS['nwr_blue']};
    margin: 1.5rem 0 0.8rem 0;
  }}

  .stTextInput input {{
    background-color: {COLOURS['surface2']} !important;
    color: {COLOURS['white']} !important;
    border: 1px solid {COLOURS['border']} !important;
    border-radius: 3px !important;
  }}
  .stTextInput input::placeholder {{
    color: {COLOURS['muted']} !important;
    opacity: 1 !important;
  }}
  .stTextInput label {{
    color: {COLOURS['white']} !important;
    font-size: 0.85rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.04em !important;
  }}

  .stButton > button {{
    background-color: {COLOURS['nwr_blue']} !important;
    color: white !important;
    border: none !important;
    border-radius: 3px !important;
    font-family: 'Barlow Condensed', sans-serif !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    padding: 0.5rem 2rem !important;
    text-transform: uppercase !important;
  }}
  .stButton > button:hover {{
    background-color: #002244 !important;
  }}

  label, .stMarkdown p {{
    color: {COLOURS['text']} !important;
  }}

  p, span, div, li {{
    color: {COLOURS['text']};
  }}

  .stDataFrame {{
    color: {COLOURS['white']} !important;
  }}
</style>
""", unsafe_allow_html=True)


# ── Mileage helpers ──────────────────────────────────────────────────────────
def mileage_to_decimal(s):
    """Convert '182m 10ch' to decimal miles (182.0220) for hazard CSV comparison."""
    if not s:
        return None
    m = re.search(r'(\d+)\s*m\s*([\d.]+)\s*c(?:h)?', str(s), re.IGNORECASE)
    if m:
        miles = int(m.group(1))
        chains = float(m.group(2))
        yards = chains * 22
        return miles + yards / 10000
    return None


def decimal_to_miles_chains(dec):
    """Convert decimal miles.yards (182.0259) to 'Xm Ych' display format."""
    if dec is None or dec == '':
        return ''
    try:
        dec = float(dec)
        miles = int(dec)
        yards = round((abs(dec) - abs(miles)) * 10000)
        chains = round(yards / 22)
        return f"{miles}m {chains:02d}ch"
    except (ValueError, TypeError):
        return str(dec)


# ── Auto-load hazard CSVs ────────────────────────────────────────────────────
@st.cache_data
def load_all_hazard_csvs():
    """Load all CSV files from the data subfolder and combine into one DataFrame."""
    data_dir = os.path.join(os.path.dirname(__file__), 'data')
    if not os.path.exists(data_dir):
        return None, [], 0

    csv_files = glob.glob(os.path.join(data_dir, '*.csv'))
    if not csv_files:
        return None, [], 0

    dfs = []
    loaded_files = []
    for f in sorted(csv_files):
        try:
            df = pd.read_csv(f, quotechar='"', encoding='utf-8')
            if 'ELR' in df.columns:
                dfs.append(df)
                loaded_files.append(os.path.basename(f))
        except Exception:
            pass

    if not dfs:
        return None, [], 0

    combined = pd.concat(dfs, ignore_index=True)
    if 'Hazard ID' in combined.columns:
        combined = combined.drop_duplicates(subset=['Hazard ID'])

    return combined, loaded_files, len(csv_files)


# ── Auto-load signal box contacts ────────────────────────────────────────────
@st.cache_data
def load_signal_box_contacts():
    """Load signal box contacts CSV from the data subfolder."""
    data_dir = os.path.join(os.path.dirname(__file__), 'data')
    csv_path = os.path.join(data_dir, 'signal_box_contacts.csv')
    if not os.path.exists(csv_path):
        return None
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        return df
    except Exception:
        return None


# ── Auto-load A&E departments ────────────────────────────────────────────────
@st.cache_data
def load_ae_departments():
    """Load Type 1 A&E departments CSV from the data subfolder."""
    data_dir = os.path.join(os.path.dirname(__file__), 'data')
    csv_path = os.path.join(data_dir, 'ae_departments_type1.csv')
    if not os.path.exists(csv_path):
        return None
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        return df
    except Exception:
        return None


def find_nearest_ae(lat, lon, ae_df, n=3):
    """Find the n nearest Type 1 A&E departments to a given lat/lon using haversine."""
    import math
    R = 3959  # miles

    def haversine(lat1, lon1, lat2, lon2):
        dlat = math.radians(lat2 - lat1)
        dlon = math.radians(lon2 - lon1)
        a = (math.sin(dlat / 2) ** 2 +
             math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) *
             math.sin(dlon / 2) ** 2)
        return R * 2 * math.asin(math.sqrt(a))

    results = []
    for _, row in ae_df.iterrows():
        try:
            d = haversine(lat, lon, float(row['latitude']), float(row['longitude']))
            results.append({
                'Hospital': row['hospital_name'],
                'Trust': row['trust_name'],
                'Address': row['address'],
                'Postcode': row['postcode'],
                'Distance (miles)': round(d, 1),
            })
        except (ValueError, TypeError):
            pass
    results.sort(key=lambda x: x['Distance (miles)'])
    return results[:n]


# Approximate UK postcode coordinates for common railway areas
# Used when access point postcodes are found in search results
POSTCODE_AREA_COORDS = {
    'AB': (57.15, -2.11), 'AL': (51.75, -0.34), 'B': (52.48, -1.89),
    'BA': (51.38, -2.36), 'BB': (53.75, -2.48), 'BD': (53.80, -1.76),
    'BH': (50.72, -1.88), 'BL': (53.58, -2.43), 'BN': (50.83, -0.14),
    'BR': (51.39, 0.05), 'BS': (51.45, -2.59), 'CA': (54.89, -2.93),
    'CB': (52.20, 0.12), 'CF': (51.48, -3.18), 'CH': (53.19, -2.89),
    'CM': (51.73, 0.47), 'CO': (51.89, 0.90), 'CR': (51.37, -0.10),
    'CT': (51.28, 1.08), 'CV': (52.41, -1.51), 'CW': (53.10, -2.44),
    'DA': (51.44, 0.21), 'DD': (56.46, -2.97), 'DE': (52.92, -1.47),
    'DH': (54.78, -1.57), 'DL': (54.52, -1.55), 'DN': (53.52, -1.13),
    'DT': (50.71, -2.44), 'DY': (52.51, -2.08), 'E': (51.55, -0.06),
    'EC': (51.52, -0.09), 'EH': (55.95, -3.19), 'EN': (51.65, -0.08),
    'EX': (50.72, -3.53), 'FK': (56.00, -3.78), 'FY': (53.82, -3.05),
    'G': (55.86, -4.25), 'GL': (51.86, -2.24), 'GU': (51.24, -0.77),
    'HA': (51.58, -0.34), 'HD': (53.65, -1.78), 'HG': (54.00, -1.54),
    'HP': (51.75, -0.74), 'HR': (52.06, -2.72), 'HU': (53.74, -0.33),
    'HX': (53.73, -1.86), 'IG': (51.56, 0.08), 'IP': (52.06, 1.16),
    'KT': (51.38, -0.30), 'KY': (56.20, -3.15), 'L': (53.41, -2.98),
    'LA': (54.05, -2.80), 'LD': (52.25, -3.38), 'LE': (52.63, -1.13),
    'LL': (53.23, -3.83), 'LN': (53.23, -0.54), 'LS': (53.80, -1.55),
    'LU': (51.88, -0.42), 'M': (53.48, -2.24), 'ME': (51.39, 0.54),
    'MK': (52.04, -0.76), 'ML': (55.77, -3.99), 'N': (51.57, -0.10),
    'NE': (55.00, -1.60), 'NG': (52.95, -1.15), 'NN': (52.24, -0.90),
    'NP': (51.59, -3.00), 'NR': (52.63, 1.30), 'NW': (51.55, -0.17),
    'OL': (53.54, -2.10), 'OX': (51.75, -1.26), 'PA': (55.85, -4.44),
    'PE': (52.57, -0.24), 'PL': (50.37, -4.14), 'PO': (50.80, -1.09),
    'PR': (53.76, -2.70), 'RG': (51.45, -1.00), 'RH': (51.17, -0.19),
    'RM': (51.57, 0.18), 'S': (53.38, -1.47), 'SA': (51.65, -3.94),
    'SE': (51.48, -0.06), 'SG': (51.90, -0.20), 'SK': (53.39, -2.16),
    'SL': (51.51, -0.60), 'SM': (51.37, -0.17), 'SN': (51.56, -1.78),
    'SO': (50.93, -1.40), 'SP': (51.07, -1.80), 'SR': (54.89, -1.38),
    'SS': (51.54, 0.71), 'ST': (52.98, -2.18), 'SW': (51.46, -0.17),
    'SY': (52.71, -2.75), 'TA': (51.01, -3.10), 'TD': (55.60, -2.43),
    'TF': (52.68, -2.49), 'TN': (51.13, 0.26), 'TQ': (50.47, -3.53),
    'TR': (50.26, -5.05), 'TS': (54.57, -1.23), 'TW': (51.45, -0.34),
    'UB': (51.53, -0.44), 'W': (51.51, -0.18), 'WA': (53.39, -2.60),
    'WC': (51.52, -0.12), 'WD': (51.66, -0.40), 'WF': (53.68, -1.50),
    'WN': (53.55, -2.63), 'WR': (52.19, -2.22), 'WS': (52.58, -1.97),
    'WV': (52.59, -2.13), 'YO': (53.96, -1.08),
}


def extract_postcode_from_text(text):
    """Extract a UK postcode from free text. Returns (outward_code, full_postcode) or (None, None)."""
    if not text or not isinstance(text, str):
        return None, None
    # Full postcode pattern
    m = re.search(r'([A-Z]{1,2}\d{1,2}[A-Z]?)\s*(\d[A-Z]{2})', text.upper())
    if m:
        return m.group(1).rstrip(), f"{m.group(1)} {m.group(2)}"
    return None, None


def get_coords_from_postcode_area(postcode):
    """Get approximate lat/lon from a postcode's area code."""
    if not postcode:
        return None, None
    pc = postcode.upper().strip()
    # Try 2-letter area first, then 1-letter
    area2 = re.match(r'^([A-Z]{2})', pc)
    area1 = re.match(r'^([A-Z])', pc)
    if area2 and area2.group(1) in POSTCODE_AREA_COORDS:
        return POSTCODE_AREA_COORDS[area2.group(1)]
    if area1 and area1.group(1) in POSTCODE_AREA_COORDS:
        return POSTCODE_AREA_COORDS[area1.group(1)]
    return None, None


# ── Filter helpers ───────────────────────────────────────────────────────────
def filter_access_points(df):
    """Filter DataFrame to only access point rows."""
    if 'Hazard Description' not in df.columns:
        return df.iloc[0:0]
    mask = df['Hazard Description'].str.contains('Access Point', case=False, na=False)
    return df[mask].copy()


def filter_hazards_only(df):
    """Filter DataFrame to exclude access point rows (hazards only)."""
    if 'Hazard Description' not in df.columns:
        return df
    mask = ~df['Hazard Description'].str.contains('Access Point', case=False, na=False)
    return df[mask].copy()


def query_by_elr_mileage(df, elr_from, elr_to, from_dec, to_dec):
    """Query a DataFrame by ELR(s) and mileage range with overlap logic."""
    elrs_to_query = [elr_from]
    if elr_to and elr_to != elr_from:
        elrs_to_query.append(elr_to)

    frames = []
    for elr_q in elrs_to_query:
        elr_data = df[df['ELR'] == elr_q].copy()
        if not elr_data.empty:
            elr_data['mile_from'] = pd.to_numeric(
                elr_data['Mileage  From'], errors='coerce')
            elr_data['mile_to'] = pd.to_numeric(
                elr_data['Mileage To'], errors='coerce')

            if elr_q == elr_from and elr_q == elr_to:
                filt = elr_data[
                    ~(elr_data['mile_to'] < from_dec) &
                    ~(elr_data['mile_from'] > to_dec)
                ]
            elif elr_q == elr_from:
                filt = elr_data[elr_data['mile_from'] >= from_dec]
            else:
                filt = elr_data[elr_data['mile_to'] <= to_dec]

            frames.append(filt)

    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame()


# ── PDF generation ───────────────────────────────────────────────────────────
NWR_BLUE_PDF = HexColor("#003366")
ALT_ROW_PDF = HexColor("#EEF2F7")
GRID_GREY_PDF = HexColor("#CCCCCC")


class HazardDocTemplate(BaseDocTemplate):
    def __init__(self, filename, **kwargs):
        self.elr = kwargs.pop('elr', '')
        self.issue_date = kwargs.pop('issue_date', '')
        self.from_mil = kwargs.pop('from_mil', '')
        self.to_mil = kwargs.pop('to_mil', '')
        self.pdf_title = kwargs.pop('pdf_title', 'NWR Hazard Directory')
        BaseDocTemplate.__init__(self, filename, **kwargs)


def on_page(canvas, doc):
    canvas.saveState()
    w, h = landscape(A4)
    margins = 10 * mm

    bar_h = 14 * mm
    canvas.setFillColor(NWR_BLUE_PDF)
    canvas.rect(0, h - bar_h, w, bar_h, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont('Helvetica-Bold', 10)
    title = f"{doc.pdf_title}  \u2013  ELR: {doc.elr}  \u2013  Issue Date: {doc.issue_date}"
    canvas.drawCentredString(w / 2, h - bar_h + 4 * mm, title)

    canvas.setFillColor(black)
    canvas.setFont('Helvetica-Bold', 8)
    range_text = f"Extract: {doc.from_mil} to {doc.to_mil}"
    canvas.drawCentredString(w / 2, h - bar_h - 4 * mm, range_text)

    canvas.setFillColor(black)
    canvas.setFont('Helvetica', 7)
    canvas.drawString(margins, 6 * mm,
                      f"PPS Rail  |  {doc.issue_date}")
    canvas.drawRightString(w - margins, 6 * mm,
                           f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


def generate_pdf(filtered_df, elr, from_mil, to_mil, col_config, pdf_title='NWR Hazard Directory'):
    """Generate PDF in memory and return bytes."""
    buf = io.BytesIO()
    issue_date = datetime.now().strftime("%d/%m/%Y")

    page_w, page_h = landscape(A4)
    margins = 10 * mm
    usable_w = page_w - 2 * margins

    col_widths = [usable_w * c['width_pct'] for c in col_config]

    header_style = ParagraphStyle(
        'HeaderCell', fontName='Helvetica-Bold', fontSize=7.5,
        textColor=white, leading=9, alignment=TA_LEFT,
    )
    cell_style = ParagraphStyle(
        'DataCell', fontName='Helvetica', fontSize=7,
        textColor=black, leading=8.5, alignment=TA_LEFT,
        wordWrap='CJK',
    )

    doc = HazardDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=margins, rightMargin=margins,
        topMargin=22 * mm, bottomMargin=14 * mm,
        elr=elr, issue_date=issue_date,
        from_mil=from_mil, to_mil=to_mil,
        pdf_title=pdf_title,
    )

    frame = Frame(margins, 14 * mm, usable_w, page_h - 36 * mm, id='main')
    template = PageTemplate(id='main', frames=[frame], onPage=on_page)
    doc.addPageTemplates([template])

    header_row = [Paragraph(c['header'], header_style) for c in col_config]
    table_data = [header_row]

    for _, row in filtered_df.iterrows():
        data_row = [
            Paragraph(str(row.get(c['field'], '')).replace('\n', '<br/>'), cell_style)
            for c in col_config
        ]
        table_data.append(data_row)

    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), NWR_BLUE_PDF),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7.5),
        ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.4, GRID_GREY_PDF),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]

    for i in range(1, len(table_data)):
        if i % 2 == 0:
            style_commands.append(
                ('BACKGROUND', (0, i), (-1, i), ALT_ROW_PDF))

    table.setStyle(TableStyle(style_commands))
    doc.build([table])
    buf.seek(0)
    return buf


# ── Column configs for PDFs ──────────────────────────────────────────────────
HAZARD_COLS = [
    {'header': 'ELR',         'field': 'ELR',                'width_pct': 0.055},
    {'header': 'ELR Name',    'field': 'ELR Name',           'width_pct': 0.18},
    {'header': 'Start',       'field': 'Mileage  From',      'width_pct': 0.055},
    {'header': 'End',         'field': 'Mileage To',         'width_pct': 0.055},
    {'header': 'Description', 'field': 'Hazard Description', 'width_pct': 0.12},
    {'header': 'Local Name',  'field': 'Local Name',         'width_pct': 0.10},
    {'header': 'Track',       'field': 'Track',              'width_pct': 0.07},
    {'header': 'Free Text',   'field': 'Free Text',          'width_pct': 0.365},
]

ACCESS_COLS = [
    {'header': 'ELR',         'field': 'ELR',                'width_pct': 0.07},
    {'header': 'ELR Name',    'field': 'ELR Name',           'width_pct': 0.18},
    {'header': 'Mileage',     'field': 'Mileage  From',      'width_pct': 0.08},
    {'header': 'Type',        'field': 'Hazard Description', 'width_pct': 0.18},
    {'header': 'Local Name',  'field': 'Local Name',         'width_pct': 0.18},
    {'header': 'Track',       'field': 'Track',              'width_pct': 0.10},
    {'header': 'Details',     'field': 'Free Text',          'width_pct': 0.21},
]

SIGNALBOX_COLS = [
    {'header': 'Signal Box / Panel / Workstation', 'field': 'Signal Box',         'width_pct': 0.40},
    {'header': 'External Telephone',               'field': 'External Telephone', 'width_pct': 0.25},
    {'header': 'Signal Prefix',                    'field': 'Signal Prefix',      'width_pct': 0.15},
    {'header': 'Route',                            'field': 'Source',             'width_pct': 0.20},
]


# ── APP LAYOUT ───────────────────────────────────────────────────────────────

# Top bar
logo_path = os.path.join(os.path.dirname(__file__), 'PPS-logo-ol.png')
logo_b64 = ''
if os.path.exists(logo_path):
    with open(logo_path, 'rb') as f:
        logo_b64 = base64.b64encode(f.read()).decode()

st.markdown(f"""
<div style="
  display: flex;
  align-items: center;
  gap: 1.5rem;
  padding: 1rem 0 1.2rem 0;
  border-bottom: 1px solid {COLOURS['border']};
  margin-bottom: 1.5rem;
  background: {COLOURS['bg']};
">
  {'<img src="data:image/png;base64,' + logo_b64 + '" style="height:60px;width:auto;border-radius:3px;" />' if logo_b64 else ''}
  <div>
    <div style="font-family: Barlow Condensed, sans-serif; font-size: 1.6rem; font-weight: 700; color: white; letter-spacing: 0.05em;">
      Worksite Intelligence</div>
    <div style="font-size: 0.8rem; color: {COLOURS['muted']}; letter-spacing: 0.1em; text-transform: uppercase;">
      Hazards &bull; Access Points &bull; Signal Box Contacts &bull; Nearest A&amp;E</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Load data ────────────────────────────────────────────────────────────────
hazard_df, loaded_files, total_files = load_all_hazard_csvs()
signalbox_df = load_signal_box_contacts()
ae_df = load_ae_departments()

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    if hazard_df is not None and not hazard_df.empty:
        n_hazards = len(hazard_df)
        n_elrs = hazard_df['ELR'].nunique()
        n_access = len(filter_access_points(hazard_df))
        n_sb = len(signalbox_df) if signalbox_df is not None else 0
        n_ae = len(ae_df) if ae_df is not None else 0

        st.markdown("### Data Loaded")
        st.markdown(f"""
        <div class="pps-card pps-card-green">
          <span class="badge badge-green">&#10003; {len(loaded_files)} files loaded</span>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="metric-box" style="margin-bottom:0.5rem">
          <div class="metric-num" style="color:{COLOURS['green']}">{n_hazards:,}</div>
          <div class="metric-lbl">Hazards</div>
        </div>
        <div class="metric-box" style="margin-bottom:0.5rem">
          <div class="metric-num" style="color:{COLOURS['amber']}">{n_access:,}</div>
          <div class="metric-lbl">Access Points</div>
        </div>
        <div class="metric-box" style="margin-bottom:0.5rem">
          <div class="metric-num" style="color:{COLOURS['text']}">{n_sb:,}</div>
          <div class="metric-lbl">Signal Boxes</div>
        </div>
        <div class="metric-box" style="margin-bottom:0.5rem">
          <div class="metric-num" style="color:{COLOURS['red']}">{n_ae:,}</div>
          <div class="metric-lbl">A&amp;E Depts</div>
        </div>
        <div class="metric-box">
          <div class="metric-num" style="color:{COLOURS['text']}">{n_elrs:,}</div>
          <div class="metric-lbl">ELRs</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<hr class="pps-divider">', unsafe_allow_html=True)
        st.markdown("### Available ELRs")
        elrs = sorted(hazard_df['ELR'].dropna().unique().tolist())
        for e in elrs[:50]:
            count = len(hazard_df[hazard_df['ELR'] == e])
            st.markdown(f"<small><b>{e}</b> <span style='color:{COLOURS['muted']}'>"
                        f"({count})</span></small>",
                        unsafe_allow_html=True)
        if len(elrs) > 50:
            st.markdown(f"<small style='color:{COLOURS['muted']}'>...and {len(elrs)-50} more</small>",
                        unsafe_allow_html=True)
    else:
        st.markdown("### No Data")
        st.markdown(f"""
        <div class="pps-card pps-card-amber">
          <span class="badge badge-amber">No CSV files found</span>
          <div style="margin-top:0.5rem;font-size:0.85rem;color:{COLOURS['muted']}">
            Place hazard directory CSV files in the <b>data</b> subfolder.
          </div>
        </div>
        """, unsafe_allow_html=True)


# ── Main ─────────────────────────────────────────────────────────────────────
if hazard_df is None or hazard_df.empty:
    st.markdown(f"""
    <div class="pps-card pps-card-grey">
      <span class="badge badge-grey">No data loaded</span>
      <span style="margin-left:1rem;font-size:0.85rem;color:#888">
        Place your hazard directory CSV files in the <b>data</b> subfolder and restart the app.
      </span>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("#### Enter your worksite mileage range")

    c1, c2, c3, c4 = st.columns([2, 2, 2, 2])
    with c1:
        elr_from = st.text_input("ELR FROM", placeholder="e.g. CGJ3").upper()
    with c2:
        mil_from = st.text_input("Mileage FROM", placeholder="e.g. 182m 10ch")
    with c3:
        elr_to = st.text_input("ELR TO", placeholder="e.g. CGJ3 (or different)").upper()
    with c4:
        mil_to = st.text_input("Mileage TO", placeholder="e.g. 182m 30ch")

    # Default ELR TO to ELR FROM if left blank
    if not elr_to and elr_from:
        elr_to = elr_from

    search = st.button("🔍  SEARCH WORKSITE")

    if search:
        if not elr_from:
            st.warning("Please enter an ELR FROM.")
        elif not mil_from or not mil_to:
            st.warning("Please enter mileage FROM and TO.")
        else:
            from_dec = mileage_to_decimal(mil_from)
            to_dec = mileage_to_decimal(mil_to)

            if from_dec is None or to_dec is None:
                st.error("Invalid mileage format. Use e.g. 182m 10ch")
            else:
                elr_label = elr_from if elr_from == elr_to else f"{elr_from} to {elr_to}"
                timestamp = datetime.now().strftime("%Y%m%d_%H%M")

                # ─── SECTION 1: HAZARDS ──────────────────────────────────
                st.markdown(f'<div class="section-header">⚠️  HAZARDS — {elr_label} {mil_from} to {mil_to}</div>',
                            unsafe_allow_html=True)

                all_in_range = query_by_elr_mileage(hazard_df, elr_from, elr_to, from_dec, to_dec)
                hazards = filter_hazards_only(all_in_range) if not all_in_range.empty else pd.DataFrame()

                if hazards.empty:
                    st.markdown(f"""
                    <div class="pps-card pps-card-green">
                      <span class="badge badge-green">&#10003; Clear</span>
                      <span style="margin-left:1rem;font-size:0.85rem">No hazards found</span>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    display_cols = ['ELR', 'ELR Name', 'Mileage  From',
                                    'Mileage To', 'Hazard Description',
                                    'Local Name', 'Track', 'Free Text']
                    hz_display = hazards[
                        [c for c in display_cols if c in hazards.columns]
                    ].fillna('')

                    mc1, mc2 = st.columns([1, 5])
                    with mc1:
                        st.markdown(f"""
                        <div class="metric-box">
                          <div class="metric-num" style="color:{COLOURS['amber']}">{len(hz_display)}</div>
                          <div class="metric-lbl">Hazards</div>
                        </div>""", unsafe_allow_html=True)

                    st.dataframe(hz_display, use_container_width=True, hide_index=True)

                    pdf_buf = generate_pdf(hz_display, elr_label, mil_from, mil_to,
                                           HAZARD_COLS, 'NWR Hazard Directory')
                    st.download_button(
                        "⬇  Download Hazard PDF", data=pdf_buf,
                        file_name=f"Hazards_{elr_from}_{timestamp}.pdf",
                        mime="application/pdf", key="hz_pdf")

                # ─── SECTION 2: ACCESS POINTS ────────────────────────────
                st.markdown(f'<div class="section-header">🚪  ACCESS POINTS — {elr_label} {mil_from} to {mil_to}</div>',
                            unsafe_allow_html=True)

                ap_all = filter_access_points(hazard_df)
                access_pts = query_by_elr_mileage(ap_all, elr_from, elr_to, from_dec, to_dec) if not ap_all.empty else pd.DataFrame()

                if access_pts.empty:
                    st.markdown(f"""
                    <div class="pps-card pps-card-amber">
                      <span class="badge badge-amber">No access points</span>
                      <span style="margin-left:1rem;font-size:0.85rem">No access points found in range</span>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    display_cols_ap = ['ELR', 'ELR Name', 'Mileage  From',
                                       'Hazard Description', 'Local Name',
                                       'Track', 'Free Text']
                    ap_display = access_pts[
                        [c for c in display_cols_ap if c in access_pts.columns]
                    ].fillna('')

                    ap_show = ap_display.rename(columns={
                        'Mileage  From': 'Mileage',
                        'Hazard Description': 'Type',
                        'Free Text': 'Details',
                    })
                    ap_show['Mileage'] = ap_show['Mileage'].apply(decimal_to_miles_chains)

                    # Also convert for PDF
                    ap_display['Mileage  From'] = ap_display['Mileage  From'].apply(decimal_to_miles_chains)

                    n_ped = len(access_pts[access_pts['Hazard Description'].str.contains('Pedestrian', na=False)])
                    n_veh = len(access_pts[access_pts['Hazard Description'].str.contains('Vehicle', na=False)])
                    n_rr = len(access_pts[access_pts['Hazard Description'].str.contains('Road-Rail|Machines', na=False)])

                    mc1, mc2, mc3, mc4 = st.columns(4)
                    with mc1:
                        st.markdown(f"""
                        <div class="metric-box">
                          <div class="metric-num" style="color:{COLOURS['green']}">{len(access_pts)}</div>
                          <div class="metric-lbl">Total</div>
                        </div>""", unsafe_allow_html=True)
                    with mc2:
                        st.markdown(f"""
                        <div class="metric-box">
                          <div class="metric-num" style="color:{COLOURS['text']}">{n_ped}</div>
                          <div class="metric-lbl">Pedestrian</div>
                        </div>""", unsafe_allow_html=True)
                    with mc3:
                        st.markdown(f"""
                        <div class="metric-box">
                          <div class="metric-num" style="color:{COLOURS['amber']}">{n_veh}</div>
                          <div class="metric-lbl">Vehicle</div>
                        </div>""", unsafe_allow_html=True)
                    with mc4:
                        st.markdown(f"""
                        <div class="metric-box">
                          <div class="metric-num" style="color:{COLOURS['red']}">{n_rr}</div>
                          <div class="metric-lbl">Road-Rail</div>
                        </div>""", unsafe_allow_html=True)

                    st.dataframe(ap_show, use_container_width=True, hide_index=True)

                    pdf_buf = generate_pdf(ap_display, elr_label, mil_from, mil_to,
                                           ACCESS_COLS, 'Access Points')
                    st.download_button(
                        "⬇  Download Access Points PDF", data=pdf_buf,
                        file_name=f"Access_Points_{elr_from}_{timestamp}.pdf",
                        mime="application/pdf", key="ap_pdf")

                # ─── SECTION 3: SIGNAL BOX CONTACTS ──────────────────────
                st.markdown(f'<div class="section-header">📞  SIGNAL BOX CONTACTS</div>',
                            unsafe_allow_html=True)

                if signalbox_df is None or signalbox_df.empty:
                    st.markdown(f"""
                    <div class="pps-card pps-card-grey">
                      <span class="badge badge-grey">Not available</span>
                      <span style="margin-left:1rem;font-size:0.85rem">
                        Place <b>signal_box_contacts.csv</b> in the data subfolder.
                      </span>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    # Find signal boxes by searching for signal prefixes
                    # that appear in the hazards/access points results
                    st.markdown(f"""
                    <div style="font-size:0.85rem;color:{COLOURS['muted']};margin-bottom:0.8rem">
                      Search by signal box name or signal prefix to find contact details for this area.
                    </div>
                    """, unsafe_allow_html=True)

                    sb1, sb2 = st.columns([3, 3])
                    with sb1:
                        sb_name = st.text_input("Search by name", placeholder="e.g. Rugby, Colwich, Crewe", key="sb_name")
                    with sb2:
                        sb_prefix = st.text_input("Search by signal prefix", placeholder="e.g. CE, SW, AJ", key="sb_prefix").upper()

                    if sb_name or sb_prefix:
                        results = signalbox_df.copy()

                        if sb_name:
                            results = results[results['Signal Box'].str.contains(sb_name, case=False, na=False)]
                        if sb_prefix:
                            results = results[results['Signal Prefix'].str.upper().eq(sb_prefix)]

                        if results.empty:
                            st.markdown(f"""
                            <div class="pps-card pps-card-amber">
                              <span class="badge badge-amber">No results</span>
                              <span style="margin-left:1rem;font-size:0.85rem">
                                No signal boxes found matching your search.
                              </span>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            display_sb = results[['Signal Box', 'External Telephone', 'Signal Prefix', 'Source']].copy()
                            display_sb = display_sb.fillna('')
                            display_sb = display_sb.rename(columns={
                                'Signal Box': 'Name',
                                'External Telephone': 'Phone',
                                'Signal Prefix': 'Prefix',
                                'Source': 'Route',
                            })

                            st.markdown(f"""
                            <div class="metric-box" style="max-width:200px;margin-bottom:0.8rem">
                              <div class="metric-num" style="color:{COLOURS['green']}">{len(results)}</div>
                              <div class="metric-lbl">Contacts found</div>
                            </div>""", unsafe_allow_html=True)

                            st.dataframe(display_sb, use_container_width=True, hide_index=True)

                            search_desc = sb_name or sb_prefix
                            pdf_buf = generate_pdf(results, search_desc, '', '',
                                                   SIGNALBOX_COLS, 'Signal Box Contacts')
                            st.download_button(
                                "⬇  Download Contacts PDF", data=pdf_buf,
                                file_name=f"Signal_Box_Contacts_{search_desc}_{timestamp}.pdf",
                                mime="application/pdf", key="sb_pdf")

                # ─── SECTION 4: NEAREST A&E ──────────────────────────────
                st.markdown(f'<div class="section-header">🏥  NEAREST A&amp;E — 24hr Type 1 Emergency Departments</div>',
                            unsafe_allow_html=True)

                if ae_df is None or ae_df.empty:
                    st.markdown(f"""
                    <div class="pps-card pps-card-grey">
                      <span class="badge badge-grey">Not available</span>
                      <span style="margin-left:1rem;font-size:0.85rem">
                        Place <b>ae_departments_type1.csv</b> in the data subfolder.
                      </span>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    # Try to auto-detect location from access point postcodes
                    auto_lat, auto_lon = None, None
                    auto_source = None

                    if not access_pts.empty and 'Free Text' in access_pts.columns:
                        for _, ap_row in access_pts.iterrows():
                            outward, full_pc = extract_postcode_from_text(str(ap_row.get('Free Text', '')))
                            if outward:
                                auto_lat, auto_lon = get_coords_from_postcode_area(outward)
                                if auto_lat:
                                    auto_source = full_pc
                                    break

                    if auto_lat and auto_lon:
                        nearest = find_nearest_ae(auto_lat, auto_lon, ae_df, n=3)
                        st.markdown(f"""
                        <div style="font-size:0.85rem;color:{COLOURS['muted']};margin-bottom:0.8rem">
                          Auto-detected from access point postcode: <b>{auto_source}</b>
                        </div>
                        """, unsafe_allow_html=True)

                        for i, hosp in enumerate(nearest):
                            border_col = COLOURS['green'] if i == 0 else COLOURS['border']
                            label = 'NEAREST' if i == 0 else f'#{i+1}'
                            badge_class = 'badge-green' if i == 0 else 'badge-grey'
                            st.markdown(f"""
                            <div class="pps-card" style="border-left:3px solid {border_col}">
                              <span class="badge {badge_class}">{label} — {hosp['Distance (miles)']} miles</span>
                              <div style="margin-top:0.5rem;font-size:1rem;color:{COLOURS['white']};font-weight:600">
                                {hosp['Hospital']}</div>
                              <div style="font-size:0.85rem;color:{COLOURS['muted']};margin-top:0.2rem">
                                {hosp['Address']}, {hosp['Postcode']}</div>
                              <div style="font-size:0.78rem;color:{COLOURS['muted']};margin-top:0.1rem">
                                {hosp['Trust']}</div>
                            </div>
                            """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div style="font-size:0.85rem;color:{COLOURS['muted']};margin-bottom:0.8rem">
                          No postcode found in access points — search by town or postcode to find nearest A&amp;E.
                        </div>
                        """, unsafe_allow_html=True)

                    # Manual search fallback / override
                    ae_search = st.text_input("Search A&E by town, hospital name or postcode",
                                              placeholder="e.g. Derby, Reading, LS1",
                                              key="ae_search")

                    if ae_search:
                        search_upper = ae_search.upper().strip()
                        # Check if it looks like a postcode area
                        pc_lat, pc_lon = get_coords_from_postcode_area(search_upper)

                        if pc_lat and pc_lon:
                            nearest = find_nearest_ae(pc_lat, pc_lon, ae_df, n=5)
                            ae_results = pd.DataFrame(nearest)
                            st.dataframe(ae_results, use_container_width=True, hide_index=True)
                        else:
                            # Text search on hospital name, address, trust
                            matches = ae_df[
                                ae_df['hospital_name'].str.contains(ae_search, case=False, na=False) |
                                ae_df['address'].str.contains(ae_search, case=False, na=False) |
                                ae_df['trust_name'].str.contains(ae_search, case=False, na=False) |
                                ae_df['postcode'].str.contains(search_upper, na=False)
                            ].copy()

                            if matches.empty:
                                st.markdown(f"""
                                <div class="pps-card pps-card-amber">
                                  <span class="badge badge-amber">No results</span>
                                  <span style="margin-left:1rem;font-size:0.85rem">
                                    Try a different town name or postcode area.
                                  </span>
                                </div>
                                """, unsafe_allow_html=True)
                            else:
                                display_ae = matches[['hospital_name', 'address', 'postcode', 'trust_name']].copy()
                                display_ae.columns = ['Hospital', 'Address', 'Postcode', 'Trust']
                                st.dataframe(display_ae, use_container_width=True, hide_index=True)
