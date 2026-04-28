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


# ── Auto-load access points with coordinates ─────────────────────────────────
@st.cache_data
def load_access_points_coords():
    """Load access points with lat/lon from the data subfolder."""
    data_dir = os.path.join(os.path.dirname(__file__), 'data')
    csv_path = os.path.join(data_dir, 'access_points_with_coords.csv')
    if not os.path.exists(csv_path):
        return None
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        return df
    except Exception:
        return None


# ── Auto-load signal box areas ───────────────────────────────────────────────
@st.cache_data
def load_signal_box_areas():
    """Load signal box areas CSV from the data subfolder."""
    data_dir = os.path.join(os.path.dirname(__file__), 'data')
    csv_path = os.path.join(data_dir, 'signal_box_areas.csv')
    if not os.path.exists(csv_path):
        return None
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
        return df
    except Exception:
        return None


# ── Auto-load line names ─────────────────────────────────────────────────────
@st.cache_data
def load_line_names():
    """Load line names CSV from the data subfolder."""
    data_dir = os.path.join(os.path.dirname(__file__), 'data')
    csv_path = os.path.join(data_dir, 'line_names.csv')
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


def find_worksite_coords(elr_from, elr_to, from_dec, to_dec, ap_coords_df):
    """Find the best lat/lon for a worksite by matching access points with coordinates."""
    if ap_coords_df is None or ap_coords_df.empty:
        return None, None, None

    mid_dec = (from_dec + to_dec) / 2

    # Search both ELRs
    elrs = [elr_from]
    if elr_to and elr_to != elr_from:
        elrs.append(elr_to)

    best_dist = float('inf')
    best_lat, best_lon, best_name = None, None, None

    for elr_q in elrs:
        matches = ap_coords_df[ap_coords_df['elr'] == elr_q].copy()
        if matches.empty:
            continue
        matches['dist'] = (matches['mileage_decimal'] - mid_dec).abs()
        closest = matches.nsmallest(1, 'dist').iloc[0]
        if closest['dist'] < best_dist:
            best_dist = closest['dist']
            best_lat = closest['latitude']
            best_lon = closest['longitude']
            best_name = f"{closest['name']} ({closest['mileage_ch']})"

    if best_lat and best_lon:
        return best_lat, best_lon, best_name
    return None, None, None


def enrich_access_points_with_coords(access_pts_df, ap_coords_df):
    """Add lat/lon and Google Maps link to access points by matching ELR + mileage."""
    if ap_coords_df is None or ap_coords_df.empty or access_pts_df.empty:
        return access_pts_df

    lats, lons, links = [], [], []

    for _, row in access_pts_df.iterrows():
        elr = row.get('ELR', '')
        mil_dec = pd.to_numeric(row.get('Mileage  From', None), errors='coerce')

        if pd.isna(mil_dec) or not elr:
            lats.append('')
            lons.append('')
            links.append('')
            continue

        matches = ap_coords_df[ap_coords_df['elr'] == elr].copy()
        if matches.empty:
            lats.append('')
            lons.append('')
            links.append('')
            continue

        matches['dist'] = (matches['mileage_decimal'] - mil_dec).abs()
        closest = matches.nsmallest(1, 'dist').iloc[0]

        # Only match if within ~0.5 decimal miles (about 0.5 miles)
        if closest['dist'] < 0.5:
            lat = closest['latitude']
            lon = closest['longitude']
            lats.append(f"{lat:.5f}")
            lons.append(f"{lon:.5f}")
            links.append(f"https://maps.google.com/?q={lat:.4f},{lon:.4f}")
        else:
            lats.append('')
            lons.append('')
            links.append('')

    result = access_pts_df.copy()
    result['Lat'] = lats
    result['Lon'] = lons
    result['Google Maps'] = links
    return result


def format_phone_numbers(raw_phone):
    """Format phone numbers that may be concatenated into separate lines.

    Splits strings like '033 085 41095 (emergency only) 033 085 41096 01270 255 582'
    into individual numbers, each on its own line.
    """
    if not raw_phone or raw_phone == 'nan':
        return ''
    raw_phone = str(raw_phone).strip()
    if not raw_phone:
        return ''

    # Find all UK phone numbers: start with 0, then digits/spaces totalling 10-11 digits
    # Optionally followed by a parenthetical note like '(emergency only)'
    numbers = re.findall(
        r'(0(?:\d[\s]?){9,10}(?:\s*\([^)]+\))?)', raw_phone
    )
    if len(numbers) <= 1:
        return raw_phone
    return '<br/>'.join(n.strip() for n in numbers)


def find_signal_boxes_for_mileage(elr_from, elr_to, from_ch, to_ch, sba_df, signalbox_df):
    """Find signal boxes covering a mileage range and match with contact details."""
    if sba_df is None or sba_df.empty:
        return []

    elrs = [elr_from]
    if elr_to and elr_to != elr_from:
        elrs.append(elr_to)

    found_boxes = []
    seen_names = set()

    for elr_q in elrs:
        matches = sba_df[
            (sba_df['elr'] == elr_q) &
            (sba_df['mileage_from_ch'] <= to_ch) &
            (sba_df['mileage_to_ch'] >= from_ch)
        ]
        # Fallback: if no mileage overlap found, show all signal boxes on this ELR
        if matches.empty:
            matches = sba_df[sba_df['elr'] == elr_q]
        for _, row in matches.iterrows():
            box_name = str(row['signal_box_name']).strip()
            if box_name in seen_names or len(box_name) < 3:
                continue
            seen_names.add(box_name)

            phone = ''
            prefix = str(row.get('signal_prefix', '')).strip()
            if prefix == 'nan':
                prefix = ''
            eco_name = str(row.get('eco_name', '')).strip()
            if eco_name == 'nan':
                eco_name = ''
            eco_type = str(row.get('eco_type', '')).strip()
            if eco_type == 'nan':
                eco_type = ''

            if signalbox_df is not None and not signalbox_df.empty and box_name:
                first_word = box_name.split()[0] if box_name else ''
                contact_match = signalbox_df[
                    signalbox_df['Signal Box'].str.contains(first_word, case=False, na=False)
                ]
                if not contact_match.empty:
                    phone = str(contact_match.iloc[0].get('External Telephone', '')).strip()
                if (not phone or phone == 'nan') and prefix:
                    prefix_match = signalbox_df[
                        signalbox_df['Signal Prefix'].str.upper().eq(prefix.upper())
                    ]
                    if not prefix_match.empty:
                        phone = str(prefix_match.iloc[0].get('External Telephone', '')).strip()
            if phone == 'nan':
                phone = ''

            eco_str = f"{eco_name.rstrip('. ')} {eco_type}".strip()

            # Look up ECR phone number from contacts
            eco_phone = ''
            if eco_name and signalbox_df is not None and not signalbox_df.empty:
                # Search for ECR/ECO entry matching the eco_name
                eco_search = eco_name.split()[0] if eco_name else ''
                if eco_search:
                    ecr_match = signalbox_df[
                        signalbox_df['Signal Box'].str.contains(eco_search, case=False, na=False) &
                        signalbox_df['Signal Box'].str.contains('ECR|ECO|Electrical', case=False, na=False, regex=True)
                    ]
                    if not ecr_match.empty:
                        eco_phone = str(ecr_match.iloc[0].get('External Telephone', '')).strip()
                    if not eco_phone or eco_phone == 'nan':
                        # Try matching just the location name + ECR
                        ecr_match2 = signalbox_df[
                            signalbox_df['Signal Box'].str.contains(f'{eco_search}.*ECR', case=False, na=False, regex=True)
                        ]
                        if not ecr_match2.empty:
                            eco_phone = str(ecr_match2.iloc[0].get('External Telephone', '')).strip()
                if eco_phone == 'nan':
                    eco_phone = ''

            found_boxes.append({
                'Signal Box': box_name,
                'Prefix': prefix,
                'Phone': phone,
                'ECO': eco_str,
                'ECO Phone': eco_phone,
                'ELR': elr_q,
                'Coverage': f"{row.get('mileage_from_raw', '')} to {row.get('mileage_to_raw', '')}",
            })

    return found_boxes


def find_line_names_for_mileage(elr_from, elr_to, from_ch, to_ch, ln_df):
    """Find line names covering a mileage range."""
    if ln_df is None or ln_df.empty:
        return []

    elrs = [elr_from]
    if elr_to and elr_to != elr_from:
        elrs.append(elr_to)

    found_lines = []
    seen = set()

    for elr_q in elrs:
        matches = ln_df[
            (ln_df['elr'] == elr_q) &
            (ln_df['mileage_from_ch'] <= to_ch) &
            (ln_df['mileage_to_ch'] >= from_ch)
        ]
        # Fallback: if no mileage overlap found, show all line names on this ELR
        if matches.empty:
            matches = ln_df[ln_df['elr'] == elr_q]
        for _, row in matches.iterrows():
            abbr = str(row['abbreviation']).strip()
            full = str(row['full_name']).strip()
            key = f"{abbr}_{full}"
            if key not in seen:
                seen.add(key)
                found_lines.append({
                    'Abbreviation': abbr,
                    'Line Name': full,
                })

    return found_lines


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

    link_style = ParagraphStyle(
        'LinkCell', fontName='Helvetica', fontSize=6,
        textColor=HexColor("#0000EE"), leading=7.5, alignment=TA_LEFT,
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
        data_row = []
        for c in col_config:
            val = str(row.get(c['field'], '')).replace('\n', '<br/>')
            if val.startswith('https://'):
                data_row.append(Paragraph(f'<a href="{val}" color="blue"><u>Map Link</u></a>', link_style))
            else:
                data_row.append(Paragraph(val, cell_style))
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
    {'header': 'ELR',         'field': 'ELR',                'width_pct': 0.06},
    {'header': 'ELR Name',    'field': 'ELR Name',           'width_pct': 0.15},
    {'header': 'Mileage',     'field': 'Mileage  From',      'width_pct': 0.07},
    {'header': 'Type',        'field': 'Hazard Description', 'width_pct': 0.14},
    {'header': 'Local Name',  'field': 'Local Name',         'width_pct': 0.14},
    {'header': 'Track',       'field': 'Track',              'width_pct': 0.08},
    {'header': 'Details',     'field': 'Free Text',          'width_pct': 0.18},
    {'header': 'Google Maps', 'field': 'Google Maps',        'width_pct': 0.18},
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
      Hazards &bull; Access Points &bull; Signal Boxes &bull; Lines &bull; Nearest A&amp;E</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Load data ────────────────────────────────────────────────────────────────
hazard_df, loaded_files, total_files = load_all_hazard_csvs()
signalbox_df = load_signal_box_contacts()
ae_df = load_ae_departments()
ap_coords_df = load_access_points_coords()
sba_df = load_signal_box_areas()
ln_df = load_line_names()

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

    # Persist search state so A&E text input doesn't reset results
    if search:
        st.session_state['ws_search'] = True
        st.session_state['ws_elr_from'] = elr_from
        st.session_state['ws_elr_to'] = elr_to
        st.session_state['ws_mil_from'] = mil_from
        st.session_state['ws_mil_to'] = mil_to

    if st.session_state.get('ws_search'):
        # Use stored values so results persist when A&E search triggers rerun
        elr_from = st.session_state.get('ws_elr_from', elr_from)
        elr_to = st.session_state.get('ws_elr_to', elr_to)
        mil_from = st.session_state.get('ws_mil_from', mil_from)
        mil_to = st.session_state.get('ws_mil_to', mil_to)

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

                    # Enrich with coordinates and Google Maps links
                    ap_display = enrich_access_points_with_coords(ap_display, ap_coords_df)

                    ap_show = ap_display.rename(columns={
                        'Mileage  From': 'Mileage',
                        'Hazard Description': 'Type',
                        'Free Text': 'Details',
                    })
                    ap_show['Mileage'] = ap_show['Mileage'].apply(decimal_to_miles_chains)

                    # Also convert mileage for PDF
                    ap_display['Mileage  From'] = ap_display['Mileage  From'].apply(decimal_to_miles_chains)

                    n_ped = len(access_pts[access_pts['Hazard Description'].str.contains('Pedestrian', na=False)])
                    n_veh = len(access_pts[access_pts['Hazard Description'].str.contains('Vehicle', na=False)])
                    n_rr = len(access_pts[access_pts['Hazard Description'].str.contains('Road-Rail|Machines', na=False)])
                    n_coords = len(ap_display[ap_display['Google Maps'] != ''])

                    mc1, mc2, mc3, mc4, mc5 = st.columns(5)
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
                    with mc5:
                        st.markdown(f"""
                        <div class="metric-box">
                          <div class="metric-num" style="color:{COLOURS['nwr_blue']}">{n_coords}</div>
                          <div class="metric-lbl">Map Links</div>
                        </div>""", unsafe_allow_html=True)

                    # On-screen: show clickable Google Maps links
                    ap_screen = ap_show.copy()
                    if 'Google Maps' in ap_screen.columns:
                        ap_screen['Location'] = ap_screen['Google Maps'].apply(
                            lambda x: f'[Map]({x})' if x else '')
                        ap_screen = ap_screen.drop(columns=['Google Maps', 'Lat', 'Lon'])

                    st.dataframe(ap_screen, use_container_width=True, hide_index=True,
                                 column_config={'Location': st.column_config.LinkColumn('Location', display_text='📍 Map')})

                    pdf_buf = generate_pdf(ap_display, elr_label, mil_from, mil_to,
                                           ACCESS_COLS, 'Access Points')
                    st.download_button(
                        "⬇  Download Access Points PDF", data=pdf_buf,
                        file_name=f"Access_Points_{elr_from}_{timestamp}.pdf",
                        mime="application/pdf", key="ap_pdf")

                # ─── SECTION 3: SIGNAL BOX CONTACTS & LINE NAMES ─────────
                st.markdown(f'<div class="section-header">📞  SIGNAL BOX CONTACTS — {elr_label} {mil_from} to {mil_to}</div>',
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
                    # Convert mileage to chains for signal box lookup
                    from_ch = int(from_dec) * 80 + round((from_dec - int(from_dec)) * 10000 / 22)
                    to_ch = int(to_dec) * 80 + round((to_dec - int(to_dec)) * 10000 / 22)

                    # Auto-detect signal boxes from mileage
                    auto_boxes = find_signal_boxes_for_mileage(
                        elr_from, elr_to, from_ch, to_ch, sba_df, signalbox_df)

                    if auto_boxes:
                        st.markdown(f"""
                        <div class="metric-box" style="max-width:200px;margin-bottom:0.8rem">
                          <div class="metric-num" style="color:{COLOURS['green']}">{len(auto_boxes)}</div>
                          <div class="metric-lbl">Signal boxes found</div>
                        </div>""", unsafe_allow_html=True)

                        for box in auto_boxes:
                            phone_display = f" — <b>{box['Phone']}</b>" if box['Phone'] else ""
                            prefix_display = f" ({box['Prefix']})" if box['Prefix'] else ""
                            # Build ECR display with phone if available
                            eco_parts = []
                            if box.get('ECO'):
                                eco_label = box['ECO']
                                eco_phone = format_phone_numbers(box.get('ECO Phone', ''))
                                if eco_phone:
                                    eco_parts.append(f"ECR: {eco_label} — <b>{eco_phone}</b>")
                                else:
                                    eco_parts.append(f"ECR: {eco_label}")
                            eco_display = f"<br/><span style='font-size:0.78rem'>{'<br/>'.join(eco_parts)}</span>" if eco_parts else ""
                            st.markdown(f"""
                            <div class="pps-card pps-card-green">
                              <div style="font-size:1rem;color:{COLOURS['white']};font-weight:600">
                                {box['Signal Box']}{prefix_display}{phone_display}</div>
                              <div style="font-size:0.82rem;color:{COLOURS['muted']};margin-top:0.2rem">
                                {box['ELR']} {box['Coverage']}{eco_display}</div>
                            </div>
                            """, unsafe_allow_html=True)
                    else:
                        st.markdown(f"""
                        <div style="font-size:0.85rem;color:{COLOURS['muted']};margin-bottom:0.8rem">
                          No signal box areas found for this ELR/mileage — use manual search below.
                        </div>
                        """, unsafe_allow_html=True)

                    # Manual search fallback
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

                            st.dataframe(display_sb, use_container_width=True, hide_index=True)

                            search_desc = sb_name or sb_prefix
                            pdf_buf = generate_pdf(results, search_desc, '', '',
                                                   SIGNALBOX_COLS, 'Signal Box Contacts')
                            st.download_button(
                                "⬇  Download Contacts PDF", data=pdf_buf,
                                file_name=f"Signal_Box_Contacts_{search_desc}_{timestamp}.pdf",
                                mime="application/pdf", key="sb_pdf")

                # ─── SECTION 3b: LINE NAMES ──────────────────────────────
                auto_lines = find_line_names_for_mileage(
                    elr_from, elr_to,
                    int(from_dec) * 80 + round((from_dec - int(from_dec)) * 10000 / 22),
                    int(to_dec) * 80 + round((to_dec - int(to_dec)) * 10000 / 22),
                    ln_df)

                if auto_lines:
                    st.markdown(f'<div class="section-header">🚂  LINES AT SITE — {elr_label} {mil_from} to {mil_to}</div>',
                                unsafe_allow_html=True)

                    st.markdown(f"""
                    <div class="metric-box" style="max-width:200px;margin-bottom:0.8rem">
                      <div class="metric-num" style="color:{COLOURS['text']}">{len(auto_lines)}</div>
                      <div class="metric-lbl">Lines</div>
                    </div>""", unsafe_allow_html=True)

                    lines_df = pd.DataFrame(auto_lines)
                    st.dataframe(lines_df, use_container_width=True, hide_index=True)
                else:
                    st.markdown(f'<div class="section-header">🚂  LINES AT SITE — {elr_label} {mil_from} to {mil_to}</div>',
                                unsafe_allow_html=True)
                    st.markdown(f"""
                    <div style="font-size:0.85rem;color:{COLOURS['muted']};margin-bottom:0.8rem">
                      No line name data found for this ELR/mileage.
                    </div>
                    """, unsafe_allow_html=True)

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
                    # Auto-detect location from access points coordinates file
                    auto_lat, auto_lon, auto_source = find_worksite_coords(
                        elr_from, elr_to, from_dec, to_dec, ap_coords_df)

                    if auto_lat and auto_lon:
                        nearest = find_nearest_ae(auto_lat, auto_lon, ae_df, n=3)
                        st.markdown(f"""
                        <div style="font-size:0.85rem;color:{COLOURS['muted']};margin-bottom:0.8rem">
                          Location from nearest access point: <b>{auto_source}</b>
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
                          No coordinate data for this ELR — search by town or hospital name below.
                        </div>
                        """, unsafe_allow_html=True)

                    # Manual search fallback / override
                    ae_search = st.text_input("Search A&E by town or hospital name",
                                              placeholder="e.g. Derby, Reading, Warrington",
                                              key="ae_search")

                    if ae_search:
                        matches = ae_df[
                            ae_df['hospital_name'].str.contains(ae_search, case=False, na=False) |
                            ae_df['address'].str.contains(ae_search, case=False, na=False) |
                            ae_df['trust_name'].str.contains(ae_search, case=False, na=False) |
                            ae_df['postcode'].str.contains(ae_search.upper().strip(), na=False)
                        ].copy()

                        if matches.empty:
                            st.markdown(f"""
                            <div class="pps-card pps-card-amber">
                              <span class="badge badge-amber">No results</span>
                              <span style="margin-left:1rem;font-size:0.85rem">
                                Try a different town name or hospital name.
                              </span>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            display_ae = matches[['hospital_name', 'address', 'postcode', 'trust_name']].copy()
                            display_ae.columns = ['Hospital', 'Address', 'Postcode', 'Trust']
                            st.dataframe(display_ae, use_container_width=True, hide_index=True)
