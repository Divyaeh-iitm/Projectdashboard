import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from scipy.stats import skew
from io import BytesIO
import itertools
from st_aggrid import AgGrid, GridOptionsBuilder
from PIL import Image
import plotly.graph_objects as go
from pathlib import Path
import base64

logo_path = Path("assets/logo.png")         # adjust if you renamed / moved
st.sidebar.image(logo_path, use_column_width=True)

with open(logo_path, "rb") as f:
    logo_bytes = f.read()
    logo_b64   = base64.b64encode(logo_bytes).decode()

excel_path = Path("data/template.xlsx")
df_template = pd.read_excel(excel_path)
with open(excel_path, "rb") as f:
    excel_bytes = f.read()
excel_buffer = BytesIO(excel_bytes)


###############################################################################
# 0) Page Config & Custom CSS
###############################################################################
st.set_page_config(
    page_title="D.I.V.Y.A – Data Interface Visualization for Your Alternatives",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── 2) Global CSS ────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
      /* ─── 1) Uniform page padding ─────────────────────────────────────── */
      .main .block-container {
        padding-top:    0rem !important;
        padding-bottom: 0.5rem !important;
        padding-left:   2rem !important;
        padding-right:  2rem !important;
      }
      hr { display: none !important; }

       /* a single class for all your page titles */
      .page-title {
          font-size: 24px !important;    /* pick a value smaller than “Welcome” now, 
                                             but larger than your Step‑2/3 default */
          font-weight: 600 !important;
          margin: 0.25rem 0 0.75rem 0 !important;
      }

      /* ─── 2) Sidebar expander styling ───────────────────────────────── */
      [data-testid="stSidebar"] .stExpander > div:first-child {
        background-color: #1f77b4 !important;
        color:            #fff !important;
        font-weight:      600;
        border-radius:    0.5rem;
      }

      /* ─── 3) Tooltip styling ───────────────────────────────────────── */
      [data-testid="stTooltip"] {
        background:    #444 !important;
        color:         #fff !important;
        border-radius: 4px;
      }

      /* ─── 4) Canvas wrapper + canvas styling ─────────────────────── */
      .canvas-wrapper {
        margin-top: 0rem !important;
      }
      .canvas-wrapper .canvas {
        border:     3px solid #444 !important;
        border-top: none       !important;
        padding:    0rem       !important;
        box-sizing: border-box;
      }
      .canvas-wrapper .canvas .section {
        padding:        0rem;
        margin-bottom:  0.75rem;
        box-sizing:     border-box;
      }
      .canvas-wrapper .canvas .top-label {
        font-size:   1.2rem;
        font-weight: bold;
        color:       #fff;
      }

      /* ─── 5) Card styling everywhere ─────────────────────────────── */
      .card {
        background:    #fff;
        border-radius: 8px;
        padding:       0rem;
        margin-bottom: 1rem;
        box-shadow:    0 2px 6px rgba(0,0,0,0.08);
      }
      .top-label {
        color:       #fff;
        font-weight: 600;
        text-align:  center;
        padding:     0rem;
        border-radius: 4px;
      }

      /* ─── 6) DataFrame tweaks ───────────────────────────────────── */
      .stDataFrame table {
        table-layout: auto !important;
        word-wrap:    break-word;
      }
      .stDataFrame table tbody tr th:first-child,
      .stDataFrame table thead tr th:first-child {
        position: sticky;
        left:     0;
        background:#fff;
        z-index:  2;
      }

      /* ─── 7) Multiselect token styling ─────────────────────────── */
      div[data-testid="stMultiSelect"] span {
        background-color: #1f77b4 !important;
        color:            #fff !important;
        border-radius:    4px !important;
        padding:          2px 6px !important;
        margin:           2px !important;
        font-size:        14px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown("""
<script>
const width = window.innerWidth;
if (width < 600) {
  document.body.classList.add('mobile');
}
</script>
<style>
body.mobile .stTabs { flex-direction: column; }
</style>
""", unsafe_allow_html=True)
# ─── Your Plotly template constant ───────────────────────────────────────────────
COLOR_TEMPLATE = "plotly_white"

col_empty_left, col_center, col_empty_right = st.columns([1,6,1], gap="small")

with col_center:
    st.markdown(
        f"""
        <div style="
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 1rem;
            margin-bottom: 1.5rem;
        ">
          <img
            src="data:image/png;base64,{LOGO_BASE64}"
            style="width:60px; height:auto;"
          />
          <div style="text-align: left;">
            <h1 style="
                margin: 0;
                font-size: 2.25rem;
                color: #2c3e50;
                line-height: 1;
            ">D.I.V.Y.A</h1>
            <p style="
                margin: 0;
                font-size: 1rem;
                color: #7f8c8d;
                line-height: 1.2;
            ">Data Interface Visualization for Your Alternatives.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
def render_step_progress(current_step: int):
    """
    Draw a 4‐segment progress bar,
    marking completed steps with ✓, the current step with →,
    and future steps in grey.
    """
    labels = ["Step 1","Step 2","Step 3","Results Canvas"]
    html = ['<div style="display:flex; margin-bottom:1rem;">']
    for i, label in enumerate(labels):
        if i < current_step - 1:
            icon = "✓"
            color = "#2c3e50"
        elif i == current_step - 1:
            icon = "→"
            color = "#f39c12"
        else:
            icon = ""
            color = "#aaa"
        html.append(
            f'''<div style="
                    flex:1;
                    text-align:center;
                    padding:0.5rem;
                    border-bottom:3px solid {color};
                    color:{color};
                    font-weight:bold;
                ">
                  {icon} {label}
               </div>'''
        )
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)

for key, val in {
    "page":             0,
    "manual_override": False,
    "work_package":    None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = val
###############################################################################
# 1) Custom Rounding & Helper Functions
###############################################################################
def custom_round(x):
    try:
        val = float(x)
        if 0 < abs(val) < 1:
            return round(val, 2)
        else:
            return int(round(val))
    except:
        return x

def remove_outliers_inclusive(series):
    if series.empty:
        return series
    Q1 = series.quantile(0.25, interpolation="midpoint")
    Q3 = series.quantile(0.75, interpolation="midpoint")
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    result = series[(series >= lower) & (series <= upper)]
    if result.empty:
        return series
    return result

def rename_columns(df):
    df.columns = df.columns.str.strip().str.lower()
    return df

def compute_uncertainty_for_row(row):
    """
    Example logic for computing an uncertainty factor from row data.
    Adjust or remove if not needed.
    """
    if "uncertainity zepd" in row and pd.notna(row["uncertainity zepd"]):
        val_str = str(row["uncertainity zepd"]).strip().lower()
        if val_str != "not declared":
            val_str = val_str.replace("%", "").strip()
            try:
                val_num = float(val_str)
                return val_num / 100 if val_num > 1 else val_num
            except:
                pass
    # Fallback approach
    Z_M, Z_F, Z_P = 0.20, 0.20, 0.20
    if "specificity" in row and pd.notna(row["specificity"]):
        spec = str(row["specificity"]).strip().lower()
        if spec == "product specific":
            Z_M, Z_F, Z_P = 0.02, 0.02, 0.02
        elif spec == "plant specific":
            Z_M, Z_F, Z_P = 0.02, 0.02, 0.20
        elif spec == "manufacturer specific":
            Z_M, Z_F, Z_P = 0.02, 0.20, 0.20
    Z_T = 0.20
    if "time representativeness" in row and pd.notna(row["time representativeness"]):
        time_str = str(row["time representativeness"]).lower()
        days = 365
        nums = re.findall(r"\d+", time_str)
        if nums:
            num_val = float(nums[0])
            if "day" in time_str:
                days = num_val
            elif "year" in time_str:
                days = num_val * 365
        Z_T = 0.02 if days < 90 else 0.20
    Z_S = 0.10
    return np.sqrt(Z_M**2 + Z_F**2 + Z_P**2 + Z_T**2 + Z_S**2)

def load_predefined_mapping(mapping_file, work_package):
    """
    Reads the Excel file (mapping_file) and returns a dictionary of 
    PM -> {SM1..SM5, plus SM2 dependencies}.
    """
    xl = pd.ExcelFile(mapping_file)
    mapping_dict = {}
    for sheet in xl.sheet_names:
        if sheet not in ["Packages", "Instructions-Steps"]:
            df_map = pd.read_excel(mapping_file, sheet_name=sheet, header=None)
            # Check cell B1 for the selected work package
            sheet_work = str(df_map.iloc[0, 1]).strip()
            if sheet_work == work_package:
                # Parse from row 4 => index=3
                for idx, row in df_map.iloc[3:].iterrows():
                    pm = row[1]  # column B
                    if pd.isna(pm):
                        continue
                    pm = str(pm).strip()
                    mapping_dict[pm] = {}
                    # SM1..SM5 => columns C..G => indexes 2..6
                    for i, smcat in enumerate(["SM1", "SM2", "SM3", "SM4", "SM5"]):
                        col_idx = 2 + i
                        cell_val = row[col_idx] if col_idx in row.index else None
                        if pd.isna(cell_val):
                            mapping_dict[pm][smcat] = []
                        else:
                            items = [x.strip() for x in str(cell_val).split(",") if x.strip()]
                            if smcat == "SM2":
                                # list of dicts => "material":..., "dependency":None
                                mapping_dict[pm][smcat] = [{"material": it, "dependency": None} for it in items]
                            else:
                                # SM1, SM3..SM5 => list of dicts => "material": item
                                mapping_dict[pm][smcat] = [{"material": it} for it in items]
                    # SM2 dependency => columns H..I => indexes 7..8
                    sm2_dep_item = row[7] if 7 in row.index else None
                    sm2_dep_value = row[8] if 8 in row.index else None
                    if pd.notna(sm2_dep_item) and str(sm2_dep_item).strip() != "":
                        sm2_dep_item = str(sm2_dep_item).strip()
                        sm2_dep_value_str = None
                        if pd.notna(sm2_dep_value) and str(sm2_dep_value).strip() != "":
                            sm2_dep_value_str = str(sm2_dep_value).strip()
                        if "SM2" in mapping_dict[pm]:
                            for obj in mapping_dict[pm]["SM2"]:
                                if obj["material"] == sm2_dep_item:
                                    obj["dependency"] = sm2_dep_value_str
                break
    return mapping_dict

###############################################################################
# 2) Analyze Each Material (Sheet) for Case 1 & Case 2
###############################################################################
def analyze_material(df, material_name):
    df = rename_columns(df)
    # Check for required columns
    for col in ["embodied energy", "embodied carbon"]:
        if col not in df.columns:
            raise KeyError(f"Sheet '{material_name}' missing '{col}'")
    ee_list_c1 = []
    ec_list_c1 = []
    ee_list_c2 = []
    ec_list_c2 = []
    debug_rows = []

    # Build row-by-row lists
    for i, row in df.iterrows():
        if pd.isna(row["embodied energy"]) or pd.isna(row["embodied carbon"]):
            continue
        ee_val = float(row["embodied energy"])
        ec_val = float(row["embodied carbon"])

        # CASE 1
        ee_list_c1.append(ee_val)
        ec_list_c1.append(ec_val)

        # CASE 2 => add uncertainty factor to carbon
        ee_list_c2.append(ee_val)
        factor = compute_uncertainty_for_row(row)
        if ec_val >= 0:
            adj_ec = ec_val * (1 + factor)
        else:
            adj_ec = ec_val * (1 - factor)
        ec_list_c2.append(adj_ec)

        debug_rows.append({
            "RowIndex": i,
            "Material": material_name,
            "Raw EC": ec_val,
            "Factor": round(factor, 4),
            "Adjusted EC": round(adj_ec, 4),
        })

    # Convert to Series
    s_ee_c1 = pd.Series(ee_list_c1)
    s_ec_c1 = pd.Series(ec_list_c1)
    s_ee_c2 = pd.Series(ee_list_c2)
    s_ec_c2 = pd.Series(ec_list_c2)

    # CASE 1 => median, range, etc.
    s_ee_c1_clean = remove_outliers_inclusive(s_ee_c1)
    s_ec_c1_clean = remove_outliers_inclusive(s_ec_c1)
    ee_median = s_ee_c1_clean.median() if not s_ee_c1_clean.empty else np.nan
    ec_median = s_ec_c1_clean.median() if not s_ec_c1_clean.empty else np.nan
    ee_min = s_ee_c1_clean.min() if not s_ee_c1_clean.empty else np.nan
    ee_max = s_ee_c1_clean.max() if not s_ee_c1_clean.empty else np.nan
    ec_min = s_ec_c1_clean.min() if not s_ec_c1_clean.empty else np.nan
    ec_max = s_ec_c1_clean.max() if not s_ec_c1_clean.empty else np.nan
    ee_q1 = s_ee_c1_clean.quantile(0.25, interpolation="midpoint") if not s_ee_c1_clean.empty else np.nan
    ee_q3 = s_ee_c1_clean.quantile(0.75, interpolation="midpoint") if not s_ee_c1_clean.empty else np.nan
    ec_q1 = s_ec_c1_clean.quantile(0.25, interpolation="midpoint") if not s_ec_c1_clean.empty else np.nan
    ec_q3 = s_ec_c1_clean.quantile(0.75, interpolation="midpoint") if not s_ec_c1_clean.empty else np.nan
    from scipy.stats import skew
    ee_skew_val = skew(s_ee_c1) if len(s_ee_c1) > 0 else np.nan
    ec_skew_val = skew(s_ec_c1) if len(s_ec_c1) > 0 else np.nan

    stats_c1 = {
        "EE Skew": custom_round(ee_skew_val),
        "EE Median": custom_round(ee_median),
        "EE Q1": custom_round(ee_q1),
        "EE Q3": custom_round(ee_q3),
        "EE Range": f"[{custom_round(ee_min)}, {custom_round(ee_max)}]",
        "EC Skew": custom_round(ec_skew_val),
        "EC Median": custom_round(ec_median),
        "EC Q1": custom_round(ec_q1),
        "EC Q3": custom_round(ec_q3),
        "EC Range": f"[{custom_round(ec_min)}, {custom_round(ec_max)}]"
    }

    # CASE 2 => worst-case approach
    s_ee_c2_clean = remove_outliers_inclusive(s_ee_c2)
    s_ec_c2_clean = remove_outliers_inclusive(s_ec_c2)
    if s_ee_c2_clean.empty:
        s_ee_c2_clean = s_ee_c2
    if s_ec_c2_clean.empty:
        s_ec_c2_clean = s_ec_c2
    ee_max_c2 = s_ee_c2_clean.max() if not s_ee_c2_clean.empty else np.nan
    ec_max_c2 = s_ec_c2_clean.max() if not s_ec_c2_clean.empty else np.nan
    stats_c2 = {
        "EE Max": custom_round(ee_max_c2),
        "EE Range (Case2)": f"[{custom_round(s_ee_c2_clean.min())}, {custom_round(s_ee_c2_clean.max())}]" if not s_ee_c2_clean.empty else "N/A",
        "EC Max": custom_round(ec_max_c2),
        "EC Range (Case2)": f"[{custom_round(s_ec_c2_clean.min())}, {custom_round(s_ec_c2_clean.max())}]" if not s_ec_c2_clean.empty else "N/A"
    }

    # Optional box plots
    ee_box = px.box(
        df, y='embodied energy', points="all", template=COLOR_TEMPLATE,
        color_discrete_sequence=px.colors.qualitative.Dark2,
        title=f"{material_name} – EE Distribution (Case 1)",
        labels={'embodied energy': 'MJ'}
    )
    ec_box = px.box(
        df, y='embodied carbon', points="all", template=COLOR_TEMPLATE,
        color_discrete_sequence=px.colors.qualitative.Dark2,
        title=f"{material_name} – EC Distribution (Case 1)",
        labels={'embodied carbon': 'kg CO₂eq.'}
    )
    # Add annotations
    ee_box.add_annotation(
        text=(
            f"Median: {custom_round(ee_median)}<br>"
            f"Q1: {custom_round(ee_q1)} & Q3: {custom_round(ee_q3)}<br>"
            f"Range: [{custom_round(ee_min)}, {custom_round(ee_max)}]"
        ),
        xref="paper", yref="paper", x=0.5, y=0.95, showarrow=False,
        font=dict(color="black", size=12)
    )
    ec_box.add_annotation(
        text=(
            f"Median: {custom_round(ec_median)}<br>"
            f"Q1: {custom_round(ec_q1)} & Q3: {custom_round(ec_q3)}<br>"
            f"Range: [{custom_round(ec_min)}, {custom_round(ec_max)}]"
        ),
        xref="paper", yref="paper", x=0.5, y=0.95, showarrow=False,
        font=dict(color="black", size=12)
    )

    debug_df = pd.DataFrame(debug_rows)
    return {
        "stats": stats_c1,
        "stats_worst": stats_c2,
        "ee_box": ee_box,
        "ec_box": ec_box,
        "debug_df": debug_df
    }
def compute_predefined_mapping():
    """
    Find the sheet in mapping_file whose cell B1 matches st.session_state['work_package'],
    then load its predefined mapping and stash primary_materials + mappings.
    """
    if st.session_state.get("manual_override", False):
        return
    mapping_file = st.session_state.get("mapping_file", None)
    work_package = st.session_state.get("work_package", "")

    # defaults if we can’t do anything
    st.session_state["primary_materials"] = []
    st.session_state["mappings"] = {}

    if mapping_file is None or not work_package:
        return

    try:
        xl = pd.ExcelFile(mapping_file)
    except Exception as e:
        st.warning(f"Couldn’t open mapping file: {e}")
        return

    # look for the sheet whose B1 matches our work_package
    target_sheet = None
    for sheet in xl.sheet_names:
        if sheet in ("Instructions‑steps", "Packages"):
            continue
        # read only the top‐left corner
        df0 = pd.read_excel(mapping_file, sheet_name=sheet, header=None, nrows=1, usecols="A:B")
        if len(df0) > 0 and str(df0.iat[0,1]).strip() == work_package:
            target_sheet = sheet
            break

    if target_sheet is None:
        st.warning(f"Work package '{work_package}' not found in any sheet’s B1.")
        return

    # now delegate to your loader (you already wrote this utility)
    predefined = load_predefined_mapping(mapping_file, work_package)

    # stash results
    primaries = list(predefined.keys())
    st.session_state["primary_materials"] = primaries
    st.session_state["mappings"]          = predefined

def override_page():
    """
    Step 2 (override): ask exactly for
    1) which primary materials
    2) define SM1…SM5 + their purposes
    3) map each SM to each primary
    """

    st.header("Step 4: Manual System-Mapping Override")

    analysis_dict = st.session_state.get("analysis_dict", {})
    if not analysis_dict:
        st.error("No data loaded. Please go back and upload your Excel.")
        return
    primary_materials = st.multiselect(
        "Select Primary Materials:",
        options=list(analysis_dict.keys()),
        key="primary_select"
    )

    # collect all SM definitions
    secondary_materials = {}
    used = set()
    for i in range(1, 6):
        sm = f"SM{i}"
        purpose = st.text_input(
            f"Purpose for {sm}", 
            key=f"purpose_{sm}",
            placeholder="enter N/A to stop"
        ).strip()
        if purpose.lower() == "n/a": 
            break

        available = [m for m in analysis_dict if m not in primary_materials and m not in used]
        chosen = st.multiselect(
            f"Select items for {sm} (Purpose: {purpose}):",
            options=available,
            key=f"list_{sm}"
        )
        secondary_materials[sm] = {"purpose": purpose, "materials": chosen}
        used.update(chosen)

    # now map each PM → SM
    mappings = {}
    for pm in primary_materials:
        mappings[pm] = {}
        st.markdown(f"**Mapping for {pm}**")
        for sm, data in secondary_materials.items():
            chosen = st.multiselect(
                f"- {sm} ({data['purpose']}):",
                options=data["materials"],
                key=f"map_{pm}_{sm}"
            )

            entries = []
            # special SM2 dependency logic
            if sm == "SM2" and secondary_materials.get("SM1"):
                for item in chosen:
                    dep = None
                    chk = st.checkbox(
                        f"Does '{item}' need a dependency?",
                        key=f"depchk_{pm}_{item}"
                    )
                    if chk:
                        dep = st.selectbox(
                            f"Dependency for '{item}':",
                            options=secondary_materials["SM1"]["materials"],
                            key=f"dep_{pm}_{item}"
                        )
                    entries.append({"material": item, "dependency": dep})
            else:
                entries = [{"material": it, "dependency": None} for it in chosen]

            mappings[pm][sm] = entries

    if st.button("💾 Save & Continue"):
        st.session_state["primary_materials"] = primary_materials
        st.session_state["mappings"]          = mappings
        st.session_state["system_data"]       = mappings
        st.session_state.page                  = 4
        return
    col1, col2 = st.columns(2)
    if col1.button("← Previous"):
        st.session_state.page = 2
    

def landing_page():
    render_step_progress(1)
    st.markdown(
    '<div class="page-title"> Welcome to the Comparison visual board</div>',
    unsafe_allow_html=True,
    )
    st.write("Please enter your name to get started.")

    # Show the text_input, seeded from session_state if present
    name = st.text_input(
        "Enter your name",
        value=st.session_state.get("user_name", ""),
        key="user_name"
    )

# 2) Project details
def project_details_page():
    # ─── load work_packages from mapping file or fallback ─────────────────────────
    mapping_file = st.session_state.get("mapping_file")
    work_packages = []
    if mapping_file:
        try:
            df_pk = pd.read_excel(mapping_file, sheet_name="Packages", header=None)
            work_packages = df_pk.iloc[3:, 1].dropna().astype(str).tolist()
        except Exception:
            st.warning("Couldn't read 'Packages' sheet; using default list.")
    if not work_packages:
        work_packages = [
            "Shuttering","Concreting","Flooring","Painting",
            "False Ceiling","CP & Sanitary Works","Railing and Metal Works",
            "Windows, Doors"
        ]

    render_step_progress(2)
    st.markdown(
        '<div class="page-title">Step 2: Project Details & Work Package</div>',
        unsafe_allow_html=True
    )

    left_col, right_col = st.columns([3, 1], gap="medium")

    # ─── LEFT: compact form ────────────────────────────────────────────────────────
    with left_col, st.form("step2_form"):
        # first row: project name + area
        r1c1, r1c2 = st.columns(2, gap="small")
        with r1c1:
            project_name = st.text_input(
                "Project Name",
                help="Enter the Project name for which the analysis is required."
            )
        with r1c2:
            project_area = st.text_input(
                "📐 Area (sq ft)",
                help="Total Super Built-up Area."
            )

        # second row: location + unit
        r2c1, r2c2 = st.columns(2, gap="small")
        with r2c1:
            project_location = st.text_input(
                "📍 Location",
                help="City where the project is located."
            )
        with r2c2:
            declared_unit = st.text_input(
                "📏 Declared Unit",
                value="sqm",
                key="declared_unit",
                help="E.g. sqm, m³, etc."
            )

        # third row: mode + override + work package (if detailed)
        r3c1, r3c2 = st.columns([3, 1], gap="small")
        with r3c1:
            mode = st.radio(
                "🔎 Assessment Mode",
                ["Detailed alternative assessment", "Project assessment"],
                index=0,
                key="assessment_mode",
                horizontal=True,
                help="Pick ‘Detailed’ to expose the Work Package dropdown."
            )
        with r3c2:
            st.checkbox(
                "🔄 Manual Override",
                key="manual_override",
                help="If checked, you’ll pick mappings yourself."
            )

        # only show Work Package dropdown when in Detailed mode
        work_pkg = None
        if st.session_state.assessment_mode == "Detailed alternative assessment":
            work_pkg = st.selectbox(
                "Work Package",
                options=work_packages,
                key="work_package_widget",
                help="Select which package to map."
            )

        # final submit button
        submitted = st.form_submit_button("Next →")
        if submitted:
            st.session_state.update({
                "project_name":     project_name,
                "project_area":     project_area,
                "project_location": project_location,
            })
            if work_pkg is not None:
                st.session_state["work_package"] = work_pkg

            st.session_state.page = 2
            return
        

    # ─── RIGHT: logo + collapsible tips ───────────────────────────────────────────
    with right_col:
        logo = Image.open(BytesIO(base64.b64decode(LOGO_BASE64)))
        st.image(logo, width=100)
        with st.expander("💡 Points to Note", expanded=True):
            st.markdown(
                """
                - Expand the sidebar for the User Manual  
                - Hover the ❓ icons for quick instructions  
                - Only “Detailed” mode exposes Work Package  
                - Click **Next →** to continue
                """,
                unsafe_allow_html=True
            )

# 3) Upload
def upload_page():
    render_step_progress(3)
    st.markdown(
    '<div class="page-title">Step 3: Upload material data file</div>',
    unsafe_allow_html=True,
    )
    override_flag = st.session_state.get("manual_override", False)
    st.write("⚙️ manual_override =", override_flag)
    uploaded_file = st.file_uploader("Upload an Excel file of the materials data:", type=["xlsx"])
    st.session_state["manual_override"] = override_flag
    if not uploaded_file:
        st.error("Please upload an Excel file to proceed.")
        return
    st.session_state["uploaded_file"]   = uploaded_file
    all_data = pd.read_excel(uploaded_file, sheet_name=None)
    sheet_names = list(all_data.keys())
    
    # Initialize local variables
    analysis_dict = {}
    debug_info = {}
    sheet_names = []
    
    if uploaded_file:
        st.success("File uploaded successfully!")
        st.session_state["uploaded_file"] = uploaded_file
        with st.spinner("Reading & analyzing…"):
            all_data = pd.read_excel(uploaded_file, sheet_name=None)
            sheet_names = list(all_data.keys())
        
        
    # Initialize summary lists
    summary_list_case1 = []
    summary_list_case2 = []
    
    # Loop through sheets and perform analysis
    for mat in sheet_names:
        try:
            analysis = analyze_material(all_data[mat], mat)
            analysis_dict[mat] = analysis
            debug_info[mat] = analysis["debug_df"]
            s1 = analysis["stats"]
            s2 = analysis["stats_worst"]

            summary_list_case1.append({
                "Material": mat,
                "EE Skew": s1["EE Skew"],
                "EE Median": s1["EE Median"],
                "EE Q1": s1["EE Q1"],
                "EE Q3": s1["EE Q3"],
                "EE Range": s1["EE Range"],
                "EC Skew": s1["EC Skew"],
                "EC Median": s1["EC Median"],
                "EC Q1": s1["EC Q1"],
                "EC Q3": s1["EC Q3"],
                "EC Range": s1["EC Range"]
            })
            summary_list_case2.append({
                "Material": mat,
                "EE Max": s2["EE Max"],
                "EE Range": s2["EE Range (Case2)"],
                "EC Max": s2["EC Max"],
                "EC Range": s2["EC Range (Case2)"]
            })
        except Exception as e:
            st.warning(f"Skipping {mat}: {e}")
        
    # Create DataFrames for the summaries
    df_case1_analysis = pd.DataFrame(summary_list_case1)
    df_case1_analysis.reset_index(drop=True, inplace=True)
    df_case1_analysis.insert(0, "Sl. No.", range(1, len(df_case1_analysis) + 1))
    
    df_case2_analysis = pd.DataFrame(summary_list_case2)
    df_case2_analysis.reset_index(drop=True, inplace=True)
    df_case2_analysis.insert(0, "Sl. No.", range(1, len(df_case2_analysis) + 1))
    
    # Store variables in session_state for later steps
    st.session_state["analysis_dict"] = analysis_dict
    st.session_state["debug_info"] = debug_info
    st.session_state["sheet_names"] = sheet_names
    st.session_state["df_case1_analysis"] = df_case1_analysis
    st.session_state["df_case2_analysis"] = df_case2_analysis

    with st.expander("Show Statistical Analysis", expanded=False):
        choice = st.radio(
            "Select an analysis:",
            ["Case 1: Balanced", "Case 2: Worst-case", "Uncertainty Analysis", "Box Plots"]
        )
        if choice == "Case 1: Balanced":
            st.subheader("Material Analysis Summary (Case 1: Balanced)")
            st.dataframe(df_case1_analysis)
        elif choice == "Case 2: Worst-case":
            st.subheader("Material Analysis Summary (Case 2: Worst-case)")
            st.dataframe(df_case2_analysis)
        elif choice == "Uncertainty Analysis":
            for mat in sheet_names:
                if mat in debug_info:
                    st.subheader(f"{mat} – Row-by-Row Data")
                    st.dataframe(debug_info[mat])
        elif choice == "Box Plots":
            st.subheader("Material Box Plots (Case 1: Balanced)")
            cols = st.columns(3)
            for idx, mat in enumerate(sheet_names):
                with cols[idx % 3]:
                    st.plotly_chart(analysis_dict[mat]["ee_box"], use_container_width=True)
                    st.plotly_chart(analysis_dict[mat]["ec_box"], use_container_width=True)
        
    col1, col2 = st.columns(2)
    if col1.button("← Previous"):
        st.session_state.page = 1
    if col2.button("Next →"):
        if override_flag:
            st.session_state.page = 3
        else:
            pmapping = load_predefined_mapping(
                st.session_state.mapping_file,
                st.session_state.get("work_package")
            )
            st.session_state["primary_materials"] = list(pmapping.keys())
            st.session_state["mappings"]          = pmapping
            st.session_state.page = 4
    
def material_system_results(primary_materials, mappings, compute_only=False):
    # ─────────────────────────────────────────────────────────────────────────────
    # 1) Grab inputs (no Streamlit calls here)
    uploaded_file     = st.session_state.get("uploaded_file")
    analysis_dict     = st.session_state.get("analysis_dict", {})
    primary_materials = st.session_state.get("primary_materials", [])
    mappings          = st.session_state.get("mappings", {})
    declared_unit     = st.session_state.get("declared_unit", "sqm")

    # ─────────────────────────────────────────────────────────────────────────────
    # 2) Compute combinations & accumulate EE/EC
    system_results_case1 = []
    system_results_case2 = []
    counter = 1
    # --- if we’re in compute_only and no one ever set primaries, default them from analysis_dict
    if compute_only and not primary_materials and analysis_dict:
        primary_materials = list(analysis_dict.keys())
        st.session_state["primary_materials"] = primary_materials

    # --- similarly, if mappings is empty you could set an empty map for each primary
    if compute_only and not mappings and primary_materials:
        mappings = {pm: {} for pm in primary_materials}
        st.session_state["mappings"] = mappings
    if compute_only or (uploaded_file and primary_materials and mappings):
        sm_cats = ["SM1","SM2","SM3","SM4","SM5"]
        for primary in primary_materials:
            # prepare options per SM category
            mapping_lists = []
            for cat in sm_cats:
                entries = mappings.get(primary, {}).get(cat, [])
                if cat=="SM2" and entries and isinstance(entries[0], dict):
                    opts = entries
                else:
                    if entries and isinstance(entries[0], dict):
                        opts = [[d["material"]] for d in entries]
                    else:
                        opts = [[x] for x in entries] if entries else [[None]]
                mapping_lists.append((cat, opts))

            cats = [t[0] for t in mapping_lists]
            opts = [t[1] for t in mapping_lists]

            for combo in itertools.product(*opts):
                # base EE/EC
                ee1 = analysis_dict[primary]["stats"]["EE Median"]
                ec1 = analysis_dict[primary]["stats"]["EC Median"]
                ee2 = analysis_dict[primary]["stats_worst"]["EE Max"]
                ec2 = analysis_dict[primary]["stats_worst"]["EC Max"]

                sm_map = {c: "N/A" for c in cats}
                sm1_list = []

                # map chosen → sm_map & collect SM1 deps
                for i, c in enumerate(cats):
                    chosen = combo[i]
                    if c=="SM2" and isinstance(chosen, dict):
                        mat = chosen.get("material"); dep = chosen.get("dependency")
                        if mat: sm_map["SM2"] = mat
                        if dep: sm1_list.append(dep)
                    else:
                        vals = [v for v in chosen if v is not None]
                        if c=="SM1":
                            sm1_list.extend(vals)
                        elif vals:
                            sm_map[c] = ", ".join(vals)

                if sm1_list:
                    sm_map["SM1"] = ", ".join(sorted(set(sm1_list)))

                # accumulate sub-material stats
                for c in cats:
                    mat_str = sm_map[c]
                    if mat_str!="N/A":
                        for m in mat_str.split(", "):
                            if m in analysis_dict:
                                ee1 += analysis_dict[m]["stats"]["EE Median"]
                                ec1 += analysis_dict[m]["stats"]["EC Median"]
                                ee2 += analysis_dict[m]["stats_worst"]["EE Max"]
                                ec2 += analysis_dict[m]["stats_worst"]["EC Max"]

                # record Case 1 & Case 2
                rec1 = {
                    "System": f"S{counter}",
                    "Primary": primary,
                    **{c: sm_map[c] for c in cats},
                    f"EE (MJ/{declared_unit})": custom_round(ee1),
                    f"EC (kg CO₂eq./{declared_unit})": custom_round(ec1)
                }
                rec2 = {
                    "System": f"S{counter}",
                    "Primary": primary,
                    **{c: sm_map[c] for c in cats},
                    f"EE (MJ/{declared_unit})": custom_round(ee2),
                    f"EC (kg CO₂eq./{declared_unit})": custom_round(ec2)
                }
                system_results_case1.append(rec1)
                system_results_case2.append(rec2)
                counter += 1

    # ─────────────────────────────────────────────────────────────────────────────
    # 3) Convert to DataFrames & store
    df1 = pd.DataFrame(system_results_case1)
    df2 = pd.DataFrame(system_results_case2)
    st.session_state["df_systems_case1"] = df1
    st.session_state["df_systems_case2"] = df2

    # ─────────────────────────────────────────────────────────────────────────────
    # 4) Create sorted tables for ascending EE and EC, store in session
    ee_col = f"EE (MJ/{declared_unit})"
    ec_col = f"EC (kg CO₂eq./{declared_unit})"

    # Case‑1 EE‑sorted
    if ee_col in df1.columns:
        st.session_state["df1_sorted_ee"] = df1.sort_values(by=ee_col).reset_index(drop=True)
    else:
        st.session_state["df1_sorted_ee"] = df1.copy()

    # Case‑1 EC‑sorted
    if ec_col in df1.columns:
        st.session_state["df1_sorted_ec"] = df1.sort_values(by=ec_col).reset_index(drop=True)
    else:
        st.session_state["df1_sorted_ec"] = df1.copy()

    # Case‑2 EE‑sorted
    if ee_col in df2.columns:
        st.session_state["df2_sorted_ee"] = df2.sort_values(by=ee_col).reset_index(drop=True)
    else:
        st.session_state["df2_sorted_ee"] = df2.copy()

    # Case‑2 EC‑sorted
    if ec_col in df2.columns:
        st.session_state["df2_sorted_ec"] = df2.sort_values(by=ec_col).reset_index(drop=True)
    else:
        st.session_state["df2_sorted_ec"] = df2.copy()

    # ─────────────────────────────────────────────────────────────────────────────
    # 5) If compute_only, exit before any UI
    if compute_only:
        return

    # ─────────────────────────────────────────────────────────────────────────────
    # 6) UI: render header, warnings, selectbox, then display & chart based on df1/df2
    st.header("Step 3: Material System Results")
    if not uploaded_file:
        st.warning("No file uploaded in Step 1.")
    if not primary_materials:
        st.warning("No primary materials found. Please select them in Step 2.")
    if not mappings:
        st.warning("No mappings found from Step 2.")

    display_option = st.selectbox(
        "Select Analysis Display Option:",
        ["Case 1 Only (Median)", "Case 2 Only (Worst-case)",
         "Both Side-by-Side", "Both Stacked", "Comparison Dashboard"]
    )

    # helper to show a case table and bar charts
    colors    = px.colors.qualitative.Dark2
    color_map = {p: colors[i % len(colors)] for i, p in enumerate(primary_materials)}

    def show_case(df_case, title):
        st.subheader(title)
        st.dataframe(df_case)
        if "System" in df_case:
            fig = px.bar(
                df_case, x="System",
                y=[ee_col, ec_col],
                title=title,
                color="Primary",
                color_discrete_map=color_map,
                template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)

    df1 = st.session_state["df_systems_case1"]
    df2 = st.session_state["df_systems_case2"]

    if display_option == "Case 1 Only (Median)":
        show_case(df1, "Case 1: Median")
    elif display_option == "Case 2 Only (Worst-case)":
        show_case(df2, "Case 2: Worst-case")
    elif display_option == "Both Side-by-Side":
        c1, c2 = st.columns(2)
        with c1: show_case(df1, "Case 1")
        with c2: show_case(df2, "Case 2")
    elif display_option == "Both Stacked":
        show_case(df1, "Case 1")
        show_case(df2, "Case 2")
    else:
        merged = df1.merge(df2, on="System", suffixes=(" (Case 1)", " (Case 2)"))
        st.subheader("Comparison Dashboard")
        st.dataframe(merged)
        
def base_case_comparison(compute_only: bool = False):
    # ─────────────────────────────────────────────────────────────────────────────
    # 1) Grab inputs (no st.xxx calls here)
    declared_unit    = st.session_state.get("declared_unit", "sqm")
    uploaded_file    = st.session_state.get("uploaded_file")
    df_case1         = st.session_state.get("df_systems_case1", pd.DataFrame())

    # ─────────────────────────────────────────────────────────────────────────────
    # 2) Compute df_display (with EE/EC change) and store sorted versions
    ee_col = f"EE (MJ/{declared_unit})"
    ec_col = f"EC (kg CO₂eq./{declared_unit})"

    # default empty
    df_display = pd.DataFrame()

    if uploaded_file and not df_case1.empty and "System" in df_case1.columns:
        # ensure SM columns exist
        for col in ["Primary", "SM1", "SM2", "SM3", "SM4", "SM5"]:
            if col not in df_case1.columns:
                df_case1[col] = "N/A"

        # build DisplayName if missing
        if "DisplayName" not in df_case1.columns:
            def _mk_name(r):
                parts = [f"{r['System']}: PM={r['Primary']}"]
                for sm in ["SM1","SM2","SM3","SM4","SM5"]:
                    if r.get(sm,"N/A")!="N/A":
                        parts.append(f"{sm}={r[sm]}")
                return " | ".join(parts)
            df_case1["DisplayName"] = df_case1.apply(_mk_name, axis=1)

        # compute % change
        # pick a default base (first row) in case user hasn't selected yet
        base_system = st.session_state.get("base_case_selection", df_case1["System"].iloc[0])
        base_row    = df_case1[df_case1["System"]==base_system].iloc[0]
        base_ee     = base_row[ee_col]
        base_ec     = base_row[ec_col]

        def _pct(x, base): 
            return custom_round(((x-base)/base)*100) if base!=0 else 0

        df_case1["EE Change (%)"] = df_case1[ee_col].apply(lambda x: _pct(x, base_ee))
        df_case1["EC Change (%)"] = df_case1[ec_col].apply(lambda x: _pct(x, base_ec))

        # drop SM columns if all "N/A"
        df_display = df_case1.copy()
        for sm in ["SM1","SM2","SM3","SM4","SM5"]:
            if df_display[sm].eq("N/A").all():
                df_display.drop(columns=[sm], inplace=True)

    # store for downstream
    st.session_state["df_base_display"]     = df_display
    st.session_state["df_base_sorted_ee"]   = (
        df_display.sort_values(by=ee_col).reset_index(drop=True)
        if not df_display.empty else pd.DataFrame()
    )
    st.session_state["df_base_sorted_ec"]   = (
        df_display.sort_values(by=ec_col).reset_index(drop=True)
        if not df_display.empty else pd.DataFrame()
    )

    # ─────────────────────────────────────────────────────────────────────────────
    # 3) Bail out before any st.xxx when compute_only
    if compute_only:
        return

    # ─────────────────────────────────────────────────────────────────────────────
    # 4) UI rendering
    st.header("Step 4: Base Case Comparison")

    if not uploaded_file:
        st.warning("No file was uploaded in Step 1. Please go back and upload a file.")
        return
    if df_display.empty:
        st.warning("No valid Case 1 systems found. Please complete Step 3 to generate systems.")
        return

    # select base system
    base_choice = st.selectbox(
        "Select a base system for comparison (Case 1 results):",
        options=df_display["DisplayName"].tolist(),
        key="base_choice_label"
    )
    if base_choice:
        # persist the selection
        system = df_display.loc[df_display["DisplayName"]==base_choice, "System"].iloc[0]
        st.session_state["base_case_selection"] = system

    # show sorted EE table
    st.subheader("Case 1 Systems Sorted by EE")
    st.dataframe(st.session_state["df_base_sorted_ee"])

    # show sorted EC table
    st.subheader("Case 1 Systems Sorted by EC")
    st.dataframe(st.session_state["df_base_sorted_ec"])

    # bar charts for % change
    colors    = px.colors.qualitative.Dark2
    prims     = df_display["Primary"].unique().tolist()
    cmap      = {p: colors[i%len(colors)] for i,p in enumerate(prims)}

    fig_ee = px.bar(
        df_display, x="System", y="EE Change (%)",
        title="EE Change (%) Relative to Base",
        color="Primary", color_discrete_map=cmap, template="plotly_white"
    )
    st.plotly_chart(fig_ee, use_container_width=True)

    fig_ec = px.bar(
        df_display, x="System", y="EC Change (%)",
        title="EC Change (%) Relative to Base",
        color="Primary", color_discrete_map=cmap, template="plotly_white"
    )
    st.plotly_chart(fig_ec, use_container_width=True)

def total_work_calculation(compute_only: bool = False):
    # ─────────────────────────────────────────────────────────────────────────────
    # 1) Grab inputs & defaults (no st.xxx calls here)
    df_case1       = st.session_state.get("df_systems_case1", pd.DataFrame())
    base_system    = st.session_state.get("base_case_selection", None)
    declared_unit  = st.session_state.get("declared_unit", "sqm")
    prev_qty       = st.session_state.get("total_qty", 100.0)

    ee_col = f"EE (MJ/{declared_unit})"
    ec_col = f"EC (kg CO₂eq./{declared_unit})"

    # ─────────────────────────────────────────────────────────────────────────────
    # 2) Compute per-system totals using the last stored quantity
    df_totals = pd.DataFrame()
    if not df_case1.empty and ee_col in df_case1 and ec_col in df_case1:
        df = df_case1.copy()
        df["Total EE (GJ)"]        = df[ee_col] * prev_qty / 1000
        df["Total EC (TonCO₂ eq.)"] = df[ec_col] * prev_qty / 1000
        # carry over DisplayName if it exists
        if "DisplayName" in df_case1.columns:
            df["DisplayName"] = df_case1["DisplayName"]
        else:
            df["DisplayName"] = df["System"]
        st.session_state["df_totals"] = df

    # ─────────────────────────────────────────────────────────────────────────────
    # 3) Bail before UI when computing for the canvas
    if compute_only:
        return

    # ─────────────────────────────────────────────────────────────────────────────
    # 4) UI: Step 5 header and validation
    st.header("Step 5: Total Work Calculation and Comparison")
    if df_case1.empty or "System" not in df_case1.columns:
        st.warning("No valid Case 1 systems found. Please complete Step 3 first.")
        return

    # 5) Let user enter (or update) total quantity
    total_qty = st.number_input(
        f"Enter the total quantity of work (in {declared_unit}):",
        min_value=0.0,
        value=prev_qty,
        step=1.0,
        key="step5_total_qty_input"
    )
    st.session_state["total_qty"] = total_qty

    # 6) Recompute df_totals with the newly entered quantity
    df = df_case1.copy()
    df["Total EE (GJ)"]         = df[ee_col] * total_qty / 1000
    df["Total EC (TonCO₂ eq.)"] = df[ec_col] * total_qty / 1000
    if "DisplayName" in df_case1.columns:
        df["DisplayName"] = df_case1["DisplayName"]
    else:
        df["DisplayName"] = df["System"]
    st.session_state["df_totals"] = df

    # ─────────────────────────────────────────────────────────────────────────────
    # 7) Project‐level comparison summary
    #    Use base_system (fallback to first) and let user pick a comparison system
    if base_system is None or base_system not in df["System"].values:
        base_system = df["System"].iloc[0]

    base_row = df[df["System"] == base_system].iloc[0]
    base_ee  = base_row["Total EE (GJ)"]
    base_ec  = base_row["Total EC (TonCO₂ eq.)"]

    st.subheader("Select Another System for Comparison")
    comp_choice = st.selectbox(
        "Comparison system:",
        options=df["DisplayName"].tolist(),
        key="comp_choice"
    )
    comp_row = df[df["DisplayName"] == comp_choice].iloc[0]
    comp_ee  = comp_row["Total EE (GJ)"]
    comp_ec  = comp_row["Total EC (TonCO₂ eq.)"]

    # 8) Build summary table
    def diff_and_pct(new, base):
        d   = new - base
        pct = (d / base * 100) if base != 0 else 0
        return d, pct

    ee_diff, ee_pct = diff_and_pct(comp_ee, base_ee)
    ec_diff, ec_pct = diff_and_pct(comp_ec, base_ec)

    summary = pd.DataFrame({
        "Scenario": ["Base System", "Comparison System", "Difference (Comp - Base)"],
        "Total EE (GJ)": [base_ee, comp_ee, ee_diff],
        "Total EC (TonCO₂ eq.)": [base_ec, comp_ec, ec_diff],
        "EE Diff (%)": [None, None, custom_round(ee_pct)],
        "EC Diff (%)": [None, None, custom_round(ec_pct)]
    })
    st.subheader("Project‑Level Comparison Summary")
    st.dataframe(summary)

    # ─────────────────────────────────────────────────────────────────────────────
    # 9) Optional bar charts for all systems’ totals
    fig1 = px.bar(
        df, x="System", y="Total EE (GJ)",
        title="Total EE (GJ) for Each System",
        template="plotly_white"
    )
    st.plotly_chart(fig1, use_container_width=True)

    fig2 = px.bar(
        df, x="System", y="Total EC (TonCO₂ eq.)",
        title="Total EC (TonCO₂ eq.) for Each System",
        template="plotly_white"
    )
    st.plotly_chart(fig2, use_container_width=True)

def extended_project(compute_only: bool = False):
    # ─────────────────────────────────────────────────────────────────────────────
    # 1) Pull in your stored totals & previous selections
    df_totals        = st.session_state.get("df_totals", pd.DataFrame())
    base_system      = st.session_state.get("base_case_selection")
    comparison_choice= st.session_state.get("comparison_choice")
    debug_info       = st.session_state.get("debug_info", {})

    # ─────────────────────────────────────────────────────────────────────────────
    # 2) Compute user_items & summary DataFrame (no st calls here)
    user_items = []
    # If someone stored `user_items` from a prior run, pick that up:
    if "user_items" in st.session_state:
        user_items = st.session_state["user_items"]

    # Build summary only if df_totals is present
    df_summary = pd.DataFrame()
    df_pie     = pd.DataFrame()
    if not df_totals.empty:
        # 2a) Rest totals from user_items
        if user_items:
            rest_total_ee = sum(item["EE (GJ)"] for item in user_items)
            rest_total_ec = sum(item["EC (TonCO₂ eq.)"] for item in user_items)
        else:
            rest_total_ee = 0.0
            rest_total_ec = 0.0

        # 2b) Determine comparison system
        if comparison_choice and "DisplayName" in df_totals:
            comp_row = df_totals[df_totals["DisplayName"] == comparison_choice].iloc[0]
        else:
            # default to first
            df_totals["DisplayName"] = df_totals.get("DisplayName", df_totals["System"])
            comp_row = df_totals.iloc[0]
            comparison_choice = comp_row["DisplayName"]

        comp_total_ee = comp_row["Total EE (GJ)"]
        comp_total_ec = comp_row["Total EC (TonCO₂ eq.)"]

        # 2c) Base system totals
        if base_system and base_system in df_totals["System"].values:
            base_row = df_totals[df_totals["System"] == base_system].iloc[0]
        else:
            base_row = df_totals.iloc[0]
            base_system = base_row["System"]
        base_ee = base_row["Total EE (GJ)"]
        base_ec = base_row["Total EC (TonCO₂ eq.)"]

        # 2d) Compute grand totals and differences
        final_total_ee      = rest_total_ee + comp_total_ee
        final_total_ec      = rest_total_ec + comp_total_ec
        base_final_total_ee = rest_total_ee + base_ee
        base_final_total_ec = rest_total_ec + base_ec

        def diff_and_pct(new, base):
            d   = new - base
            pct = (d/base*100) if base!=0 else 0
            return d, pct

        diff_ee, diff_ee_pct = diff_and_pct(final_total_ee, base_final_total_ee)
        diff_ec, diff_ec_pct = diff_and_pct(final_total_ec, base_final_total_ec)

        # 2e) Build summary DataFrame
        df_summary = pd.DataFrame({
            "Scenario": [
                "Additional Project Only (Rest)",
                f"Chosen System ({comp_row['System']})",
                "Grand Total (Chosen)",
                f"Base System ({base_system})",
                "Grand Total (Base)",
                "Difference (Chosen vs Base)"
            ],
            "EE (GJ)": [
                rest_total_ee,
                comp_total_ee,
                final_total_ee,
                base_ee,
                base_final_total_ee,
                diff_ee
            ],
            "EC (TonCO₂ eq.)": [
                rest_total_ec,
                comp_total_ec,
                final_total_ec,
                base_ec,
                base_final_total_ec,
                diff_ec
            ],
            "EE Diff (%)": [None,None,None,None,None,diff_ee_pct],
            "EC Diff (%)": [None,None,None,None,None,diff_ec_pct]
        })

        # 2f) Build pie‐chart DataFrame
        pie_slices = user_items.copy()
        pie_slices.append({
            "Name": f"Chosen: {comp_row['System']}",
            "EE (GJ)": comp_total_ee,
            "EC (TonCO₂ eq.)": comp_total_ec
        })
        df_pie = pd.DataFrame(pie_slices)

    # 2g) Store for downstream or for canvas
    st.session_state["extended_summary"] = df_summary
    st.session_state["extended_pie"]     = df_pie

    # 2h) Prepare Excel if needed
    def make_excel():
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            if not df_totals.empty:
                df_totals.to_excel(writer, sheet_name="Totals", index=False)
            if not df_summary.empty:
                df_summary.to_excel(writer, sheet_name="Extended_Summary", index=False)
            for mat, dbg in debug_info.items():
                if dbg is not None and not dbg.empty:
                    dbg.to_excel(writer, sheet_name=f"Debug_{mat[:20]}", index=False)
        return out.getvalue()
    st.session_state["extended_excel"] = make_excel() if not df_totals.empty else None

    # ─────────────────────────────────────────────────────────────────────────────
    # 3) Bail if compute_only
    if compute_only:
        return

    # ─────────────────────────────────────────────────────────────────────────────
    # 4) UI rendering
    st.header("Step 6: Extended Project‑Level Impact & Comparison")

    if df_totals.empty:
        st.warning("No project totals available. Please complete Step 5 first.")
        return

    st.markdown("""
    **Project‑Level Impact:**  
    Now add additional EE & EC (excluding this work) and see combined totals.
    """)

    # A) Rest inputs
    mode = st.selectbox(
        "How to add additional EE & EC:",
        ["Single total", "Multiple items"]
    )

    if mode == "Single total":
        rest_ee = st.number_input("Project EE (GJ) excluding this work:", min_value=0.0, value=0.0, step=1.0)
        rest_ec = st.number_input("Project EC (TonCO₂ eq.) excl. this work:", min_value=0.0, value=0.0, step=1.0)
        user_items = [{"Name":"Rest of Project","EE (GJ)":rest_ee,"EC (TonCO₂ eq.)":rest_ec}]
    else:
        n = st.number_input("Number of additional items:", min_value=1, max_value=10, value=1, step=1, key="ext_n")
        user_items = []
        for i in range(int(n)):
            cols = st.columns([2,1,1])
            name = cols[0].text_input(f"Item #{i+1} Name:", key=f"ext_name_{i}")
            ee   = cols[1].number_input(f"EE (GJ) #{i+1}:", min_value=0.0, value=0.0, key=f"ext_ee_{i}")
            ec   = cols[2].number_input(f"EC (TonCO₂ eq.) #{i+1}:", min_value=0.0, value=0.0, key=f"ext_ec_{i}")
            user_items.append({"Name": name or f"Item {i+1}", "EE (GJ)": ee, "EC (TonCO₂ eq.)": ec})

    # B) Show summary table
    st.subheader("Extended Impact Summary")
    st.dataframe(df_summary)

    # C) Pie charts
    p1, p2 = st.columns(2)
    with p1:
        fig1 = px.pie(df_pie, names="Name", values="EE (GJ)", title="EE Breakdown")
        st.plotly_chart(fig1, use_container_width=True)
    with p2:
        fig2 = px.pie(df_pie, names="Name", values="EC (TonCO₂ eq.)", title="EC Breakdown")
        st.plotly_chart(fig2, use_container_width=True)

    # D) Download all results
    if st.session_state.get("extended_excel"):
        st.download_button(
            "Download All Results (Excel)",
            data=st.session_state["extended_excel"],
            file_name="Extended_Project_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def draw_horizontal_heat_strip(df, metric_col, title, colorscale):
    # Sort ascending
    sorted_df = df.sort_values(metric_col)
    values = sorted_df[metric_col].tolist()
    names  = sorted_df["DisplayName"].tolist()
    n = len(values)

    # One-row heatmap
    fig = go.Figure(
        go.Heatmap(
            z=[values],               # shape = (1, n)
            x=list(range(n)),         # positions 0…n-1
            y=[metric_col],
            text=[                       # also 1×N!
                [f"{names[i]}<br>{metric_col}: {values[i]:.1f}"
                for i in range(n)]
            ],
            hoverinfo="text",
            colorscale=colorscale,    # e.g. "Greens" or px.colors.sequential.Greens
            showscale=False,
        )
    )
    
    # Layout tweaks
    fig.update_layout(
        title=title,
        height=100,
        margin=dict(l=40, r=40, t=40, b=30),
        xaxis=dict(
            tickmode="array",
            tickvals=[0, n - 1],
            ticktext=["Least (prefer)", "Greatest (avoid)"],
            showgrid=False,
            zeroline=False,
            showline=False,
        ),
        yaxis=dict(visible=False),
        plot_bgcolor="white",
    )
    return fig


def display_canvas_dashboard():
    st.markdown("""
    <style>
    /* ─── bump the canvas down so it doesn’t get cut off ─────────────────── */
    .canvas-wrapper {
        margin-top: 1rem !important;
    }

    /* ─── minimal outer padding inside the block-container ──────────────── */
    .block-container {
        padding: 1rem 2rem !important;
    }

    /* ─── card styling ─────────────────────────────────────────────────── */
    .card {
        background: #fff;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    }

    /* ─── top‑labels inside cards ───────────────────────────────────────── */
    .top-label {
        color: #fff;
        font-weight: 600;
        text-align: center;
        padding: 0.5rem;
        border-radius: 4px;
    }
    </style>
    """, unsafe_allow_html=True)

    # 0) Ensure quantity is initialized
    if "total_qty" not in st.session_state:
        st.session_state["total_qty"] = 1.0
    wp        = st.session_state.get("work_package", "")
    uploaded = st.session_state.get("uploaded_file")
    if not (wp and uploaded):
        return
    
    try:
        xls    = pd.ExcelFile(uploaded)
        sheets = {s.strip() for s in xls.sheet_names}
    except Exception:
        st.error("❗ Could not read your Excel file. Please upload a valid .xlsx.")
        return
    
    pms = st.session_state.get("primary_materials", [])
    primaries = [str(x).strip() for x in pms if isinstance(x, (str, bytes))]
    maps = st.session_state.get("mappings", {})
    matches = sheets & set(primaries)

    if len(matches) < 2:
        st.markdown(
            f"""
            <div style="
                background-color: #FFF0F0;
                border-left: 6px solid #E53935;
                padding: 16px;
                border-radius: 8px;
                margin-bottom: 24px;
            ">
            <p style="margin:0 0 8px 0;
                        font-size:1.1rem;
                        font-weight:600;
                        color:#B71C1C;">
                ❗ Insufficient Material Tabs
            </p>
            <p style="margin:0 0 12px 0; color:#333;">
                Only {len(matches)} of the primary materials 
                <code>{sorted(primaries)}</code>
                were found in your uploaded file’s sheets&nbsp;
                <code>{sorted(sheets)}</code>.
            </p>
            <ul style="margin:0 0 0 16px; padding:0; color:#333;">
                <li>Please select the correct work package.</li>
                <li>Or upload an Excel with at least two of those primary materials as sheet names.</li>
            </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return
    
    material_system_results(    pms, maps, compute_only=True )
    base_case_comparison(compute_only=True)
    total_work_calculation(compute_only=True)
    extended_project(compute_only=True)
    
    df = st.session_state.get("df_totals", pd.DataFrame())
    if df.empty:
        st.error("No system data — please complete Steps 1–5.")
        return

    # 2) Pull context
    total_qty = st.session_state["total_qty"]
    unit      = st.session_state.get("declared_unit", "sqm")
    project   = st.session_state.get("project_name", "")
    
    ee_col    = f"EE (MJ/{unit})"
    ec_col    = f"EC (kg CO₂eq./{unit})"

    dark2 = px.colors.qualitative.Dark2
    cmap  = {pm: dark2[i % len(dark2)] for i, pm in enumerate(df["Primary"].unique())}
    
    # 4) Begin the single bordered canvas
    st.markdown('<div class="canvas">', unsafe_allow_html=True)

    # ─── Top bar
    cols = st.columns([3,3,2], gap="small")
    for col, txt, bg in zip(cols,
                            [f"Project: {project}", f"Work Package: {wp}", f"# Systems: {df['System'].nunique()}"],
                            [dark2[0], dark2[1], dark2[2]]):
        col.markdown(
            f'<div class="top-label" style="background:{bg}">{txt}</div>',
            unsafe_allow_html=True
        )

    # ─── Materials + Quantity
    st.markdown('<div class="card">', unsafe_allow_html=True)
    mat_col, qty_col = st.columns([4,1], gap="small")
    with mat_col:
        selected = st.multiselect("Primary Materials:", options=pms, default=pms, key="canvas_pm")
    with qty_col:
        total_qty = st.number_input(
            f"Quantity ({unit}):",
            min_value=1.0, step=1.0,
            key="total_qty"
        )
    st.markdown('</div>', unsafe_allow_html=True)

    # 3) Compute totals
    total_qty = st.session_state["total_qty"]
    df["Total EE (GJ)"] = df[f"EE (MJ/{unit})"] * total_qty / 1000
    df["Total EC (tCO₂)"] = df[f"EC (kg CO₂eq./{unit})"] * total_qty / 1000

    df = df[df["Primary"].isin(selected)]

    # ─── Total EE & EC bars
    st.markdown('<div class="card">', unsafe_allow_html=True)
    b1, b2 = st.columns(2, gap="small")
    with b1:
        st.plotly_chart(
            px.bar(df, x="System", y="Total EE (GJ)",
                   color="Primary", color_discrete_map=cmap,
                   title="Total Embodied Energy (GJ)",
                   template="plotly_white"),
            use_container_width=True
        )
    with b2:
        st.plotly_chart(
            px.bar(df, x="System", y="Total EC (tCO₂)",
                   color="Primary", color_discrete_map=cmap,
                   title="Total Embodied Carbon (tCO₂)",
                   template="plotly_white"),
            use_container_width=True
        )
    st.markdown('</div>', unsafe_allow_html=True)

    df1_ee = st.session_state.get("df1_sorted_ee", pd.DataFrame())
    df1_ec = st.session_state.get("df1_sorted_ec", pd.DataFrame())
    df2_ee = st.session_state.get("df2_sorted_ee", pd.DataFrame())
    df2_ec = st.session_state.get("df2_sorted_ec", pd.DataFrame())
    table_slot = st.empty()

    if st.button("📊 View Detailed Tables"):
    # when clicked, fill the placeholder with an expander + tabs
        with table_slot.expander("🔎 Quick Table Views", expanded=True):
            if st.button("✖ Close Tables"):
                table_slot.empty()
            tabs = st.tabs([
                "⚡ Case 1 EE",
                "🌿 Case 1 EC",
                "⚡ Case 2 EE",
                "🌿 Case 2 EC",
            ])
            with tabs[0]:
                st.subheader("Case 1 – sorted by EE")
                st.dataframe(df1_ee, use_container_width=True)
            with tabs[1]:
                st.subheader("Case 1 – sorted by EC")
                st.dataframe(df1_ec, use_container_width=True)
            with tabs[2]:
                st.subheader("Case 2 – sorted by EE")
                st.dataframe(df2_ee, use_container_width=True)
            with tabs[3]:
                st.subheader("Case 2 – sorted by EC")
                st.dataframe(df2_ec, use_container_width=True)
    
    if not df.empty and "DisplayName" in df.columns:
        ee_col = f"EE (MJ/{unit})"
        ec_col = f"EC (kg CO₂eq./{unit})"

        st.markdown("### Quick recommendation", unsafe_allow_html=True)

        ee_strip = draw_horizontal_heat_strip(df, ee_col,
                        " Embodied Energy", "Blues")
        ec_strip = draw_horizontal_heat_strip(df, ec_col,
                        " Embodied Carbon", "Greens")

        st.plotly_chart(ee_strip, use_container_width=True)
        st.plotly_chart(ec_strip, use_container_width=True)
        
        recs = (
            df.groupby("DisplayName")[[ee_col, ec_col]]
            .sum()
            .reset_index()
        )
        recs["score"] = recs[ee_col] + recs[ec_col]

        # pick best/worst
        ee_best  = recs.loc[recs[ee_col].idxmin(), "DisplayName"]
        ec_best  = recs.loc[recs[ec_col].idxmin(), "DisplayName"]
        ee_worst = recs.loc[recs[ee_col].idxmax(), "DisplayName"]
        ec_worst = recs.loc[recs[ec_col].idxmax(), "DisplayName"]
        both_best = (recs
            .assign(rank_sum=lambda d: d[ee_col].rank() + d[ec_col].rank())
            .nsmallest(1, "rank_sum")["DisplayName"]
            .iat[0]
        )

        # ultra-compact summary
        st.markdown("## Preferred & Cautionary Picks", unsafe_allow_html=True)

        # Row 1: best EE & EC
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            <div style="
                background:#E8F5E9;
                border:1px solid #C8E6C9;
                border-radius:8px;
                padding:16px;
                text-align:center;
            ">
                <h4 style="margin:0; color:#2E7D32;">⚡️ Energy-Efficient System</h4>
                <p style="margin:8px 0 0; font-weight:600;">{ee_best}</p>
            </div>
            """.format(ee_best=ee_best), unsafe_allow_html=True)
        with col2:
            st.markdown("""
            <div style="
                background:#E3F2FD;
                border:1px solid #BBDEFB;
                border-radius:8px;
                padding:16px;
                text-align:center;
            ">
                <h4 style="margin:0; color:#1565C0;">🌍 Emission-Saver System</h4>
                <p style="margin:8px 0 0; font-weight:600;">{ec_best}</p>
            </div>
            """.format(ec_best=ec_best), unsafe_allow_html=True)

        st.markdown("")

        # Row 2: worst EE & EC
        col3, col4 = st.columns(2)
        with col3:
            st.markdown("""
            <div style="
                background:#FFEBEE;
                border:1px solid #FFCDD2;
                border-radius:8px;
                padding:16px;
                text-align:center;
            ">
                <h4 style="margin:0; color:#C62828;">⚠️ Highest-Impact (EE)</h4>
                <p style="margin:8px 0 0; font-weight:600;">{ee_worst}</p>
            </div>
            """.format(ee_worst=ee_worst), unsafe_allow_html=True)
        with col4:
            st.markdown("""
            <div style="
                background:#FFF8E1;
                border:1px solid #FFECB3;
                border-radius:8px;
                padding:16px;
                text-align:center;
            ">
                <h4 style="margin:0; color:#EF6C00;">⚠️ Highest-Impact (EC)</h4>
                <p style="margin:8px 0 0; font-weight:600;">{ec_worst}</p>
            </div>
            """.format(ec_worst=ec_worst), unsafe_allow_html=True)

        st.markdown("---")

        # Final single‐line “Preferred Sustainable System”
        st.markdown(f"""
        <div style="
            background:#F1F8E9;
            border-left:4px solid #8BC34A;
            padding:12px 16px;
            border-radius:4px;
            font-size:1.1rem;
        ">
        🌱 <strong>Preferred Sustainable System:</strong> {both_best}
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("Run through Steps 1–5 to unlock impact heat strips.")

    # ─── Base vs Compare
    st.markdown('<div class="card">', unsafe_allow_html=True)
    bc, cc = st.columns(2, gap="small")
    with bc:
        base = st.selectbox("Base-Case System:",    df["DisplayName"], key="canvas_base")
    with cc:
        comp = st.selectbox("Compare-Case System:", df["DisplayName"], key="canvas_compare")
    st.markdown('</div>', unsafe_allow_html=True)
    
    try:
        brow = df.loc[df["DisplayName"] == base].iloc[0]
    except IndexError:
        st.error(
            f"❗ No row with DisplayName == “{base}” was found.  "
            "Check your mapping/excel and chosen work package."
        )
        return

    brow    = df.loc[df["DisplayName"] == base].iloc[0]
    crow    = df.loc[df["DisplayName"] == comp].iloc[0]
    delta_ee = crow["Total EE (GJ)"] - brow["Total EE (GJ)"]
    delta_ec = crow["Total EC (tCO₂)"] - brow["Total EC (tCO₂)"]
    pct_ee   = (delta_ee / brow["Total EE (GJ)"] * 100) if brow["Total EE (GJ)"] else 0
    pct_ec   = (delta_ec / brow["Total EC (tCO₂)"] * 100) if brow["Total EC (tCO₂)"] else 0

    def fmt_raw(v):
        """No sign for raw numbers: one decimal if 0<|v|<1, else integer."""
        if abs(v) < 1:
            return f"{v:.1f}"
        else:
            return f"{v:.0f}"

    def fmt_signed(v):
        """Always show +/– and integer (we only use this for the delta absolute)."""
        return f"{v:+.0f}"

    def fmt_pct(v, base):
        """Compute percent = v/base*100; always show +/–, integer, and %."""
        if base == 0:
            return "0%"
        p = v / base * 100
        if abs(p) < 1:
            return f"{p:+.1f}%"
        else:
            return f"{p:+.0f}%"

    base_ee_str      = fmt_raw(brow["Total EE (GJ)"])
    base_ec_str      = fmt_raw(brow["Total EC (tCO₂)"])
    compare_ee_str   = fmt_raw(crow["Total EE (GJ)"])
    compare_ec_str   = fmt_raw(crow["Total EC (tCO₂)"])

    delta_ee_abs_str = fmt_signed(delta_ee)
    delta_ec_abs_str = fmt_signed(delta_ec)
    delta_ee_pct_str = fmt_pct(delta_ee, brow["Total EE (GJ)"])
    delta_ec_pct_str = fmt_pct(delta_ec, brow["Total EC (tCO₂)"])

    delta_ee_combined = f"{delta_ee_abs_str} ({delta_ee_pct_str})"
    delta_ec_combined = f"{delta_ec_abs_str} ({delta_ec_pct_str})"

    st.markdown('<div class="card">', unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4, gap="small")

    m1.metric("Base EE (GJ)",  base_ee_str)
    m2.metric("Base EC (tCO₂)", base_ec_str)
    m3.metric(
        "Compare EE (GJ)",
        compare_ee_str,
        delta=delta_ee_combined,
        delta_color="inverse"
    )
    m4.metric(
        "Compare EC (tCO₂)",
        compare_ec_str,
        delta=delta_ec_combined,
        delta_color="inverse"
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ─── % Change bars
    df["EE Change (%)"] = (df["Total EE (GJ)"] - brow["Total EE (GJ)"]) / brow["Total EE (GJ)"] * 100
    df["EC Change (%)"] = (df["Total EC (tCO₂)"] - brow["Total EC (tCO₂)"]) / brow["Total EC (tCO₂)"] * 100

    st.markdown('<div class="card">', unsafe_allow_html=True)
    t1, t2 = st.columns(2, gap="small")
    with t1:
        st.plotly_chart(
            px.bar(df, x="System", y="EE Change (%)",
                   color="Primary", color_discrete_map=cmap,
                   title="EE % Change vs Base", template="plotly_white"),
            use_container_width=True
        )
    with t2:
        st.plotly_chart(
            px.bar(df, x="System", y="EC Change (%)",
                   color="Primary", color_discrete_map=cmap,
                   title="EC % Change vs Base", template="plotly_white"),
            use_container_width=True
        )
    st.markdown('</div>', unsafe_allow_html=True)

    # ─── Pie charts
    st.markdown('<div class="card">', unsafe_allow_html=True)
    mode = st.selectbox(
        "How to calculate project overall EE & EC?",
        ["Single total", "Multiple items"],
        key="canvas_ext_mode"
    )

    user_items = []
    if mode == "Single total":
        rest_ee = st.number_input(
            "Project EE (GJ) excl. this work:",
            min_value=0.0, value=0.0, step=1.0, key="canvas_rest_ee"
        )
        rest_ec = st.number_input(
            "Project EC (TonCO₂ eq.) excl. this work:",
            min_value=0.0, value=0.0, step=1.0, key="canvas_rest_ec"
        )
        user_items = [
            {"Name": "Rest of Project", "EE (GJ)": rest_ee, "EC (TonCO₂ eq.)": rest_ec}
        ]
    else:
        n = st.number_input(
            "Number of additional items:",
            min_value=1, max_value=10, value=1, step=1, key="canvas_ext_n"
        )

        # single header row
        hdr_name, hdr_ee, hdr_ec = st.columns([2, 1, 1], gap="small")
        hdr_name.markdown("**Name**")
        hdr_ee.markdown("**EE (GJ)**")
        hdr_ec.markdown("**EC (TonCO₂ eq.)**")

        # one row per item
        for i in range(int(n)):
            c0, c1, c2 = st.columns([2, 1, 1], gap="small")
            name = c0.text_input(
                label=f"", 
                placeholder=f"Item {i+1} name",
                key=f"canvas_ext_name_{i}"
            )
            ee = c1.number_input(
                label="",
                min_value=0.0, value=0.0, step=1.0,
                key=f"canvas_ext_ee_{i}"
            )
            ec = c2.number_input(
                label="",
                min_value=0.0, value=0.0, step=1.0,
                key=f"canvas_ext_ec_{i}"
            )
            user_items.append({
                "Name": name.strip() or f"Item {i+1}",
                "EE (GJ)": ee,
                "EC (TonCO₂ eq.)": ec
            })
    df_pie = pd.DataFrame(user_items)
    crow = df.loc[df["DisplayName"] == comp].iloc[0]
    chosen_slice = pd.DataFrame([{
        "Name":            f"Chosen System: {crow['System']}",
        "EE (GJ)":         crow["Total EE (GJ)"],
        "EC (TonCO₂ eq.)": crow["Total EC (tCO₂)"]
    }])
    df_pie = pd.concat([df_pie, chosen_slice], ignore_index=True)
    p1, p2 = st.columns(2, gap="small")
    with p1:
        fig_ee_pie = px.pie(
            df_pie,
            names="Name",
            values="EE (GJ)",
            title="EE Breakdown (GJ)",
            template="plotly_white"
        )
        st.plotly_chart(fig_ee_pie, use_container_width=True)

    with p2:
        fig_ec_pie = px.pie(
            df_pie,
            names="Name",
            values="EC (TonCO₂ eq.)",
            title="EC Breakdown (TonCO₂ eq.)",
            template="plotly_white"
        )
        st.plotly_chart(fig_ec_pie, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    if col1.button("← Previous"):
        st.session_state.page = 3
    if col2.button("Finish"):
        st.success("🎉 All done!")

###############################################################################
# 9) Sidebar – Predefined Mapping Configuration
###############################################################################
DEFAULT_MAPPING_PATH = Path("data/Pre-defined systems.xlsx")
with st.sidebar.expander("Predefined Mapping Configuration", expanded=False):
    st.write("Below is the default or last-uploaded pre-defined systems file.")

    # 1) Ensure we have "mapping_file" in session_state
    if "mapping_file" not in st.session_state:
        # Initialize to the embedded default
        st.session_state["mapping_file"] = DEFAULT_MAPPING_PATH

    # 2) Provide a download button for the current file in memory
    mapping_src = st.session_state["mapping_file"]
    if mapping_src is not None:
        # Get bytes *once* no matter where they come from
        if isinstance(mapping_src, BytesIO):
            bytes_data = mapping_src.getvalue()
        else:  # pathlib.Path or str ⇒ read from disk
            bytes_data = mapping_src.read_bytes()
        st.download_button(
            label="Download Mapping File",
            data=bytes_data,
            file_name="Pre-defined systems.xlsx",
            key="download_mapping_file"
        )
   
    # 3) Let user upload a new file to override
    uploaded_mapping = st.file_uploader(
        "Re-upload Updated Mapping Excel File:",
        type=["xlsx"],
        key="mapping_upload"
    )
    if uploaded_mapping is not None:
        st.session_state["mapping_file"] = BytesIO(uploaded_mapping.getvalue())
        st.success("Using your newly uploaded mapping file for this session.")
    else:
        st.info("Currently using the embedded default or the last uploaded file.")

    # 4) Attempt to read the "Packages" sheet from the current file
    mapping_file = st.session_state["mapping_file"]
    if mapping_src is not None:
        try:
            # Reset the cursor if we're dealing with BytesIO
            if isinstance(mapping_src, BytesIO):
                mapping_src.seek(0)
            mapping_xl = pd.ExcelFile(mapping_src)
            if "Packages" in mapping_xl.sheet_names:
                df_packages = pd.read_excel(mapping_src, sheet_name="Packages", header=None)
                work_packages = (
                    df_packages.iloc[3:, 1]   # col B, starting row 4
                    .dropna()
                    .astype(str)
                    .tolist()
                )
                # Present them as a bulleted list instead of raw Python list
                if work_packages:
                    st.markdown("**Available Work Packages:**")
                    for pkg in work_packages:
                        st.markdown(f"- {pkg}")
                else:
                    st.warning("No work packages found in 'Packages' sheet.")

            else:
                work_packages = None
                st.warning("No 'Packages' sheet found in the mapping file.")
        except Exception as e:
            st.error(f"Error reading mapping file: {e}")
            work_packages = None
    else:
        st.warning("No mapping file found or uploaded.")
        work_packages = None

# ─── 4) Sidebar “User Manual” ────────────────────────────────────────────────────
with st.sidebar.expander("📖 User Manual", expanded=False):
    st.markdown("""
    **How to use this dashboard**  
    1. **Step 1**: Upload your Excel with material EE/EC data.  
    2. **Step 2**: Fill in project name, area, location & pick a Work Package.  
    3. **Step 3**: (Auto) generate all system alternatives.  
    4. **Step 4**: Select a base system and compare others.  
    5. **Step 5**: Enter total qty (sqm) to get project‐level totals.  
    6. **Step 6**: (Optional) Add other project items and see pie charts.  

    _Tips:_  
    - Hover over any label to see more info.  
    - You can re‐collapse this panel at any time.  
    """)
       
    
if "page" not in st.session_state:
    st.session_state.page = 0

if   st.session_state.page == 0:
    landing_page()
    if st.button("Next →"):
        st.session_state.page = 1

elif st.session_state.page == 1:
    project_details_page()
    
elif st.session_state.page == 2:
    upload_page()
    
elif st.session_state.page == 3:
    override_page()               

elif st.session_state.page == 4:   
    display_canvas_dashboard()
