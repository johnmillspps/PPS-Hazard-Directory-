"""
PPS Rail — NWR Hazard Directory & Access Points
Streamlit Web Application
=========================================================
Auto-loads all hazard directory CSV files from the data\ subfolder.
Tab 1: Extract hazards for a mileage range. Download as PDF.
Tab 2: Find access points for a mileage range. Download as PDF.
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

  h1, h2, h3 {{
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
    # Drop duplicates based on Hazard ID if present
    if 'Hazard ID' in combined.columns:
        combined = combined.drop_duplicates(subset=['Hazard ID'])

    return combined, loaded_files, len(csv_files)


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

    # Header bar
    bar_h = 14 * mm
    canvas.setFillColor(NWR_BLUE_PDF)
    canvas.rect(0, h - bar_h, w, bar_h, fill=1, stroke=0)
    canvas.setFillColor(white)
    canvas.setFont('Helvetica-Bold', 10)
    title = f"{doc.pdf_title}  \u2013  ELR: {doc.elr}  \u2013  Issue Date: {doc.issue_date}"
    canvas.drawCentredString(w / 2, h - bar_h + 4 * mm, title)

    # Mileage range subtitle
    canvas.setFillColor(black)
    canvas.setFont('Helvetica-Bold', 8)
    range_text = f"Extract: {doc.from_mil} to {doc.to_mil}"
    canvas.drawCentredString(w / 2, h - bar_h - 4 * mm, range_text)

    # Footer
    canvas.setFillColor(black)
    canvas.setFont('Helvetica', 7)
    canvas.drawString(margins, 6 * mm,
                      f"PPS Rail  |  {doc.issue_date}")
    canvas.drawRightString(w - margins, 6 * mm,
                           f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


def generate_pdf(filtered_df, elr, from_mil, to_mil, col_config, pdf_title='NWR Hazard Directory'):
    """Generate PDF in memory and return bytes.
    col_config: list of dicts with 'header', 'field', 'width_pct' keys.
    """
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
        buf,
        pagesize=landscape(A4),
        leftMargin=margins,
        rightMargin=margins,
        topMargin=22 * mm,
        bottomMargin=14 * mm,
        elr=elr,
        issue_date=issue_date,
        from_mil=from_mil,
        to_mil=to_mil,
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


# ── Access point filter ──────────────────────────────────────────────────────
ACCESS_POINT_TYPES = [
    'Access Point',
    'Authorised Access Point',
    'Authorised Access Point - Pedestrian',
    'Authorised Access Point - Road-Rail Machines',
    'Authorised Access Point - Vehicle',
    'Access Point - Platform End',
]


def filter_access_points(df):
    """Filter DataFrame to only access point rows."""
    if 'Hazard Description' not in df.columns:
        return df.iloc[0:0]
    mask = df['Hazard Description'].str.contains('Access Point', case=False, na=False)
    return df[mask].copy()


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
      NWR Hazard Directory & Access Points</div>
    <div style="font-size: 0.8rem; color: {COLOURS['muted']}; letter-spacing: 0.1em; text-transform: uppercase;">
      Safe Work Pack Tool</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Load data ────────────────────────────────────────────────────────────────
hazard_df, loaded_files, total_files = load_all_hazard_csvs()

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    if hazard_df is not None and not hazard_df.empty:
        n_hazards = len(hazard_df)
        n_elrs = hazard_df['ELR'].nunique()
        n_access = len(filter_access_points(hazard_df))

        st.markdown("### 📡 Data Loaded")
        st.markdown(f"""
        <div class="pps-card pps-card-green">
          <span class="badge badge-green">✓ {len(loaded_files)} files loaded</span>
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
        <div class="metric-box">
          <div class="metric-num" style="color:{COLOURS['text']}">{n_elrs:,}</div>
          <div class="metric-lbl">ELRs</div>
        </div>
        """, unsafe_allow_html=True)

        # Show loaded files
        st.markdown('<hr class="pps-divider">', unsafe_allow_html=True)
        st.markdown("### 📂 Loaded Files")
        for f in loaded_files:
            st.markdown(f"<small style='color:{COLOURS['muted']}'>{f}</small>",
                        unsafe_allow_html=True)

        # Show available ELRs
        st.markdown('<hr class="pps-divider">', unsafe_allow_html=True)
        st.markdown("### 🔍 Available ELRs")
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
        st.markdown("### ⚠️ No Data")
        st.markdown(f"""
        <div class="pps-card pps-card-amber">
          <span class="badge badge-amber">No CSV files found</span>
          <div style="margin-top:0.5rem;font-size:0.85rem;color:{COLOURS['muted']}">
            Place hazard directory CSV files in the <b>data</b> subfolder:
            <br/><code>Hazard App\\data\\</code>
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
    # ── Tabs ─────────────────────────────────────────────────────────────
    tab1, tab2 = st.tabs(["⚠️  HAZARD DIRECTORY", "🚪  ACCESS POINTS"])

    # ═══════════════════════════════════════════════════════════════════
    # TAB 1 — HAZARD DIRECTORY
    # ═══════════════════════════════════════════════════════════════════
    with tab1:
        st.markdown("#### Extract hazards for a mileage range")

        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            hz_elr = st.text_input("ELR", placeholder="e.g. CGJ3", key="hz_elr").upper()
        with c2:
            hz_from = st.text_input("Mileage FROM", placeholder="e.g. 182m 10ch", key="hz_from")
        with c3:
            hz_to = st.text_input("Mileage TO", placeholder="e.g. 182m 30ch", key="hz_to")

        hz_search = st.button("🔍  EXTRACT HAZARDS", key="hz_btn")

        if hz_search:
            if not hz_elr:
                st.warning("Please enter an ELR.")
            elif not hz_from or not hz_to:
                st.warning("Please enter mileage FROM and TO.")
            else:
                from_dec = mileage_to_decimal(hz_from)
                to_dec = mileage_to_decimal(hz_to)

                if from_dec is None or to_dec is None:
                    st.error("Invalid mileage format. Use e.g. 182m 10ch")
                else:
                    elr_data = hazard_df[hazard_df['ELR'] == hz_elr].copy()

                    if elr_data.empty:
                        st.markdown(f"""
                        <div class="pps-card pps-card-amber">
                          <span class="badge badge-amber">No results</span>
                          <span style="margin-left:1rem;font-size:0.85rem">
                            ELR <b>{hz_elr}</b> not found in loaded data.
                          </span>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        elr_data['mile_from'] = pd.to_numeric(
                            elr_data['Mileage  From'], errors='coerce')
                        elr_data['mile_to'] = pd.to_numeric(
                            elr_data['Mileage To'], errors='coerce')

                        filtered = elr_data[
                            ~(elr_data['mile_to'] < from_dec) &
                            ~(elr_data['mile_from'] > to_dec)
                        ].copy()

                        display_cols = ['ELR', 'ELR Name', 'Mileage  From',
                                        'Mileage To', 'Hazard Description',
                                        'Local Name', 'Track', 'Free Text']
                        filtered = filtered[
                            [c for c in display_cols if c in filtered.columns]
                        ].fillna('')

                        st.markdown('<hr class="pps-divider">', unsafe_allow_html=True)

                        if filtered.empty:
                            st.markdown(f"""
                            <div class="pps-card pps-card-green">
                              <span class="badge badge-green">✓ Clear</span>
                              <span style="margin-left:1rem;font-size:0.85rem">
                                No hazards found for <b>{hz_elr}</b>
                                {hz_from} — {hz_to}
                              </span>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            n = len(filtered)

                            mc1, mc2, mc3 = st.columns(3)
                            with mc1:
                                st.markdown(f"""
                                <div class="metric-box">
                                  <div class="metric-num" style="color:{COLOURS['amber']}">{n}</div>
                                  <div class="metric-lbl">Hazards found</div>
                                </div>""", unsafe_allow_html=True)
                            with mc2:
                                n_types = filtered['Hazard Description'].nunique() \
                                    if 'Hazard Description' in filtered.columns else 0
                                st.markdown(f"""
                                <div class="metric-box">
                                  <div class="metric-num" style="color:{COLOURS['text']}">{n_types}</div>
                                  <div class="metric-lbl">Hazard types</div>
                                </div>""", unsafe_allow_html=True)
                            with mc3:
                                st.markdown(f"""
                                <div class="metric-box">
                                  <div class="metric-num" style="color:{COLOURS['green']}">{hz_elr}</div>
                                  <div class="metric-lbl">ELR</div>
                                </div>""", unsafe_allow_html=True)

                            st.markdown(f"**Hazards for {hz_elr} — {hz_from} to {hz_to}:**")
                            st.dataframe(filtered, use_container_width=True,
                                         hide_index=True)

                            pdf_buf = generate_pdf(filtered, hz_elr, hz_from, hz_to,
                                                   HAZARD_COLS, 'NWR Hazard Directory')
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                            st.download_button(
                                "⬇  Download Hazard PDF",
                                data=pdf_buf,
                                file_name=f"NWR_Hazard_Directory_{hz_elr}_{timestamp}.pdf",
                                mime="application/pdf",
                                key="hz_pdf")

    # ═══════════════════════════════════════════════════════════════════
    # TAB 2 — ACCESS POINTS
    # ═══════════════════════════════════════════════════════════════════
    with tab2:
        st.markdown("#### Find access points for a mileage range")

        ac1, ac2, ac3 = st.columns([2, 2, 2])
        with ac1:
            ap_elr = st.text_input("ELR", placeholder="e.g. CGJ3", key="ap_elr").upper()
        with ac2:
            ap_from = st.text_input("Mileage FROM", placeholder="e.g. 182m 10ch", key="ap_from")
        with ac3:
            ap_to = st.text_input("Mileage TO", placeholder="e.g. 182m 30ch", key="ap_to")

        ap_search = st.button("🔍  FIND ACCESS POINTS", key="ap_btn")

        if ap_search:
            if not ap_elr:
                st.warning("Please enter an ELR.")
            elif not ap_from or not ap_to:
                st.warning("Please enter mileage FROM and TO.")
            else:
                from_dec = mileage_to_decimal(ap_from)
                to_dec = mileage_to_decimal(ap_to)

                if from_dec is None or to_dec is None:
                    st.error("Invalid mileage format. Use e.g. 182m 10ch")
                else:
                    # Filter to access points only, then by ELR
                    ap_data = filter_access_points(hazard_df)
                    ap_data = ap_data[ap_data['ELR'] == ap_elr].copy()

                    if ap_data.empty:
                        st.markdown(f"""
                        <div class="pps-card pps-card-amber">
                          <span class="badge badge-amber">No results</span>
                          <span style="margin-left:1rem;font-size:0.85rem">
                            No access points found for ELR <b>{ap_elr}</b>.
                          </span>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        ap_data['mile_from'] = pd.to_numeric(
                            ap_data['Mileage  From'], errors='coerce')
                        ap_data['mile_to'] = pd.to_numeric(
                            ap_data['Mileage To'], errors='coerce')

                        filtered_ap = ap_data[
                            ~(ap_data['mile_to'] < from_dec) &
                            ~(ap_data['mile_from'] > to_dec)
                        ].copy()

                        display_cols_ap = ['ELR', 'ELR Name', 'Mileage  From',
                                           'Hazard Description', 'Local Name',
                                           'Track', 'Free Text']
                        filtered_ap = filtered_ap[
                            [c for c in display_cols_ap if c in filtered_ap.columns]
                        ].fillna('')

                        # Rename for cleaner display
                        display_rename = {
                            'Mileage  From': 'Mileage',
                            'Hazard Description': 'Type',
                            'Free Text': 'Details',
                        }
                        display_df = filtered_ap.rename(columns=display_rename)

                        # Convert mileage to miles and chains
                        display_df['Mileage'] = display_df['Mileage'].apply(decimal_to_miles_chains)

                        # Also convert in filtered_ap for PDF
                        filtered_ap['Mileage  From'] = filtered_ap['Mileage  From'].apply(decimal_to_miles_chains)

                        st.markdown('<hr class="pps-divider">', unsafe_allow_html=True)

                        if filtered_ap.empty:
                            st.markdown(f"""
                            <div class="pps-card pps-card-amber">
                              <span class="badge badge-amber">No access points</span>
                              <span style="margin-left:1rem;font-size:0.85rem">
                                No access points found for <b>{ap_elr}</b>
                                {ap_from} — {ap_to}
                              </span>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            n_ap = len(filtered_ap)

                            # Count by type
                            n_ped = len(filtered_ap[filtered_ap['Hazard Description'].str.contains('Pedestrian', na=False)])
                            n_veh = len(filtered_ap[filtered_ap['Hazard Description'].str.contains('Vehicle', na=False)])
                            n_rr = len(filtered_ap[filtered_ap['Hazard Description'].str.contains('Road-Rail|Machines', na=False)])
                            n_other = n_ap - n_ped - n_veh - n_rr

                            mc1, mc2, mc3, mc4 = st.columns(4)
                            with mc1:
                                st.markdown(f"""
                                <div class="metric-box">
                                  <div class="metric-num" style="color:{COLOURS['green']}">{n_ap}</div>
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

                            st.markdown(f"**Access points for {ap_elr} — {ap_from} to {ap_to}:**")
                            st.dataframe(display_df, use_container_width=True,
                                         hide_index=True)

                            pdf_buf = generate_pdf(filtered_ap, ap_elr, ap_from, ap_to,
                                                   ACCESS_COLS, 'Access Points')
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                            st.download_button(
                                "⬇  Download Access Points PDF",
                                data=pdf_buf,
                                file_name=f"Access_Points_{ap_elr}_{timestamp}.pdf",
                                mime="application/pdf",
                                key="ap_pdf")
