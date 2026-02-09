import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
import datetime
import time
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# --- SAFE IMPORT FOR SPELLCHECKER ---
try:
    from spellchecker import SpellChecker
    SPELLCHECK_AVAILABLE = True
except ImportError:
    SPELLCHECK_AVAILABLE = False

# --- CONFIGURATION ---
st.set_page_config(page_title="Manifest Master", layout="wide", page_icon="🚢")

# --- SESSION STATE INITIALIZATION ---
if 'verified_words' not in st.session_state:
    st.session_state.verified_words = set()
if 'run_clicked' not in st.session_state:
    st.session_state.run_clicked = False

# --- 1. SETUP ONLINE CHECKER ---
geolocator = Nominatim(user_agent="my_logistics_checker_cool_v1")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

COUNTRY_ALIASES = {
    "USA": ["UNITED STATES", "US", "U.S.A."],
    "UK": ["UNITED KINGDOM", "GREAT BRITAIN", "GB"],
    "KOREA": ["SOUTH KOREA", "KR", "REPUBLIC OF KOREA"],
    "VIETNAM": ["VIET NAM", "VN"],
    "CHINA": ["PRC", "P.R.CHINA", "CN"],
    "CANADA": ["CA"],
    "JAPAN": ["JP"],
    "GERMANY": ["DE", "DEUTSCHLAND"],
    "AUSTRALIA": ["AU"],
    "INDIA": ["IN"],
    "FRANCE": ["FR"],
    "ITALY": ["IT"],
    "SPAIN": ["ES"],
    "TAIWAN": ["TW"],
    "THAILAND": ["TH"],
    "MALAYSIA": ["MY"],
    "SINGAPORE": ["SG"],
    "INDONESIA": ["ID"],
    "CAMBODIA": ["KH"]
}

ISO_MAPPING = {
    "USA": "US", "UNITED STATES": "US", "US": "US",
    "VIETNAM": "VN", "VN": "VN",
    "CHINA": "CN", "CN": "CN", "PRC": "CN",
    "KOREA": "KR", "SOUTH KOREA": "KR", "KR": "KR",
    "JAPAN": "JP", "JP": "JP",
    "CANADA": "CA", "CA": "CA",
    "UK": "GB", "UNITED KINGDOM": "GB", "GREAT BRITAIN": "GB", "GB": "GB",
    "GERMANY": "DE", "DE": "DE",
    "FRANCE": "FR", "FR": "FR",
    "ITALY": "IT", "IT": "IT",
    "SPAIN": "ES", "ES": "ES",
    "INDIA": "IN", "IN": "IN",
    "AUSTRALIA": "AU", "AU": "AU",
    "TAIWAN": "TW", "TW": "TW",
    "THAILAND": "TH", "TH": "TH",
    "MALAYSIA": "MY", "MY": "MY",
    "SINGAPORE": "SG", "SG": "SG",
    "INDONESIA": "ID", "ID": "ID",
    "CAMBODIA": "KH", "KH": "KH",
    "HONG KONG": "HK", "HK": "HK"
}

# --- 2. SETUP SPELL CHECKER ---
if SPELLCHECK_AVAILABLE:
    spell = SpellChecker()
    LOGISTICS_WORDS = [
        'KGS', 'KG', 'CBM', 'PKGS', 'PKG', 'PCS', 'PC', 'SET', 'SETS', 
        'G.W', 'N.W', 'GW', 'NW', 'CNTR', 'CONT', 'CTR', 'FCL', 'LCL', 
        'COC', 'SOC', 'NOS', 'QTY', 'PALLET', 'PALLETS', 'CTN', 'CTNS', 
        'CARTON', 'CARTONS', 'DRUM', 'DRUMS', 'BAG', 'BAGS', 'ROLL', 'ROLLS',
        'GENERAL', 'CARGO', 'FREIGHT', 'WOODEN', 'CASE', 'CASES', 'PARTS',
        'ACCESSORIES', 'MACHINERY', 'EQUIPMENT', 'TOOLS', 'PLASTIC', 'RUBBER',
        'GARMENT', 'GARMENTS', 'FABRIC', 'TEXTILE', 'FURNITURE', 'PERSONAL', 
        'EFFECTS', 'HOUSEHOLD', 'GOODS', 'MERCHANDISE', 'STC', 'SAID', 'CONTAIN',
        'PNH', 'CAT', 'LAI', 'VGM', 'YARN', 'FABRICS', 'APPAREL', 'SHOES', 'RICE',
        'FROZEN', 'SEAFOOD', 'FISH', 'SHRIMP', 'COFFEE', 'BEANS', 'PEPPER', 'CASHEW'
    ]
    spell.word_frequency.load_words(LOGISTICS_WORDS)

# --- 3. CUSTOM CSS (GEMINI VIBE) ---
def inject_custom_css():
    st.markdown("""
        <style>
        /* 1. BACKGROUND: Deep Charcoal */
        .stApp {
            background-color: #131314;
            color: #E3E3E3;
        }
        
        /* 2. CARDS: Soft Dark Grey */
        .stDataFrame, .stTable, div[data-testid="stFileUploader"], div[class*="stMetric"] {
            background-color: #1E1F20;
            border-radius: 16px;
            padding: 20px;
            border: 1px solid #333;
            box-shadow: none;
        }
        
        /* 3. HEADERS: Gradient Text */
        h1, h2, h3 {
            background: linear-gradient(90deg, #4285F4, #9B72CB, #D96570);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-family: 'Google Sans', sans-serif;
            font-weight: 700;
        }
        
        /* 4. TEXT & LABELS */
        h4, h5, p, label, li, .stMarkdown {
            color: #E3E3E3 !important;
            font-family: 'Google Sans', sans-serif;
        }
        
        /* 5. BUTTONS: Gradient Sparkle */
        .stButton>button {
            background: linear-gradient(135deg, #1c7ef0, #a855f7, #ec4899);
            color: white;
            border: none;
            border-radius: 20px;
            padding: 10px 24px;
            font-weight: 600;
            letter-spacing: 0.5px;
            transition: all 0.3s;
        }
        .stButton>button:hover {
            box-shadow: 0 0 15px rgba(168, 85, 247, 0.6);
            transform: scale(1.02);
        }
        
        /* 6. UPLOADER STYLING */
        div[data-testid="stFileUploader"] {
            border: 2px dashed #555;
        }
        div[data-testid="stFileUploader"]:hover {
            border-color: #a855f7;
        }

        /* 7. METRICS */
        div[data-testid="stMetricValue"] {
            color: #FFFFFF !important;
            font-size: 24px !important;
        }
        div[data-testid="stMetricLabel"] {
            color: #AAAAAA !important;
        }
        
        /* 8. ALERTS */
        .stAlert {
            background-color: #1E1F20;
            color: #E3E3E3;
            border: 1px solid #444;
            border-radius: 12px;
        }
        </style>
    """, unsafe_allow_html=True)

def shipping_loader():
    return """
    <div style="text-align: center; margin: 40px 0;">
        <div style="font-size: 60px; display: inline-block; animation: float 3s ease-in-out infinite; filter: drop-shadow(0 0 10px #4285F4);">
            🚢
        </div>
        <h4 style="margin-top: 15px; background: linear-gradient(90deg, #4285F4, #D96570); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
            Analyzing Manifest Data...
        </h4>
        <style>
            @keyframes float {
                0% { transform: translateY(0px); }
                50% { transform: translateY(-10px); }
                100% { transform: translateY(0px); }
            }
        </style>
    </div>
    """

inject_custom_css()

# --- 4. VALIDATION LOGIC ---
@st.cache_data
def check_location_online(destination, expected_country):
    try:
        location = geolocator.geocode(destination, language='en')
        if not location: return False, "Location not found on map"
        real_address = location.address.upper()
        target_country = expected_country.upper().strip()
        
        if target_country in real_address: return True, "Match"
        if target_country in COUNTRY_ALIASES:
            for alias in COUNTRY_ALIASES[target_country]:
                if alias in real_address: return True, "Match (Alias)"
        return False, f"Map location: '{real_address}'"
    except: return True, "Internet Error"

def is_port_code_match(dest, country):
    clean_dest = str(dest).strip().upper()
    clean_country = str(country).strip().upper()
    if len(clean_dest) == 5 and clean_dest.isalpha():
        prefix = clean_dest[:2]
        if clean_country == prefix: return True
        if clean_country in ISO_MAPPING:
            if ISO_MAPPING[clean_country] == prefix: return True
    return False

def determine_grouping(row):
    owner = str(row.iloc[6]).strip()
    bl_val = row.iloc[8]
    if pd.isna(bl_val): return owner, "Owner"
    bl_str = str(bl_val).strip().upper()
    if bl_str in ["#N/A", "IN CY", "NAN", ""]: return owner, "Owner"
    if isinstance(bl_val, (datetime.datetime, datetime.date, pd.Timestamp)): return owner, "Owner"
    if "00:00:00" in bl_str: return owner, "Owner"
    return bl_str, "B/L No"

def calculate_summary(df):
    """Calculates sums for Count, TEUs, Weight, VGM"""
    df_calc = df.copy()
    
    # Cleanup to stop at Total
    total_idx = None
    for idx, row in df_calc.iterrows():
        check_vals = [str(row.iloc[i]).upper() for i in range(min(3, len(row)))]
        if any("TOTAL" in v for v in check_vals):
            total_idx = idx
            break
    if total_idx is not None and total_idx in df_calc.index:
         int_loc = df_calc.index.get_loc(total_idx)
         if isinstance(int_loc, slice): int_loc = int_loc.start
         if hasattr(int_loc, '__iter__'): int_loc = int_loc[0]
         df_calc = df_calc.iloc[:int_loc]
         
    # TEU Logic
    def get_teu(val):
        s = str(val).strip()
        if '20' in s: return 1
        if '40' in s: return 2
        if '45' in s: return 2
        return 0 
    
    total_cnt = len(df_calc)
    total_teus = df_calc.iloc[:, 3].apply(get_teu).sum()
    
    try: total_wgt = pd.to_numeric(df_calc.iloc[:, 11], errors='coerce').sum()
    except: total_wgt = 0
    
    try: total_vgm = pd.to_numeric(df_calc.iloc[:, 20], errors='coerce').sum()
    except: total_vgm = 0
    
    return total_cnt, total_teus, total_wgt, total_vgm

def validate_data(df, use_internet):
    errors = []
    potential_typos = set()
    checks_passed = {
        "Container (Col C)": "Pending",
        "Seals (Col F)": "Pending",
        "PKG (Col K)": "Pending",
        "VGM vs Weight": "Pending",
        "Location": "Pending",
        "Description (Col M)": "Pending"
    }
    
    if len(df.columns) < 13:
        errors.append({"Ref (Col B)": "CRITICAL", "Column": "Structure", "Error": f"Columns missing. Found only {len(df.columns)} columns."})
        checks_passed["Structure"] = "❌ Failed"
        return errors, potential_typos, checks_passed

    duplicate_seals_list = []
    if len(df.columns) > 5:
        all_seals = df.iloc[:, 5].astype(str).str.strip()
        all_seals = all_seals[~all_seals.isin(['nan', 'NaN', ''])]
        duplicate_seals_list = all_seals[all_seals.duplicated()].unique()

    location_issues = {}
    if use_internet and len(df.columns) > 15:
        try:
            unique_pairs = df.iloc[:, [13, 15]].drop_duplicates().values
            for cntry, dest in unique_pairs:
                if "TOTAL" in str(cntry).upper() or "TOTAL" in str(dest).upper(): continue
                if pd.notna(cntry) and pd.notna(dest):
                    if is_port_code_match(dest, cntry): continue
                    is_valid, msg = check_location_online(str(dest), str(cntry))
                    if not is_valid: location_issues[(str(cntry), str(dest))] = msg
        except: pass

    num_cols = len(df.columns)
    
    has_cont_error = False
    has_seal_error = False
    has_pkg_error = False
    has_vgm_error = False
    has_loc_error = False
    has_desc_error = False

    for index, row in df.iterrows():
        def get_val(col_idx):
            if col_idx < num_cols: return row.iloc[col_idx]
            return None

        check_vals = [str(get_val(i)).upper() for i in range(3) if get_val(i) is not None]
        if any("TOTAL" in v for v in check_vals): break 

        ref_val = get_val(1)
        if pd.isna(ref_val) or str(ref_val).strip() == "" or str(ref_val).lower() == 'nan': continue
        ref_id = str(ref_val).replace(".0", "")
        
        def log(col, msg): errors.append({"Ref (Col B)": f"No {ref_id}", "Column": col, "Error": msg})

        cont = get_val(2)
        if pd.notna(cont) and str(cont).strip() != "":
            if not re.match(r'^[A-Z]{4}\d{7}$', str(cont).strip().upper()):
                log("C (Container)", "Invalid format")
                has_cont_error = True

        seal = get_val(5)
        if pd.notna(seal) and str(seal).strip() != "":
            if str(seal).strip() in duplicate_seals_list: 
                log("F (Seal)", "Duplicate Seal")
                has_seal_error = True

        pkg = get_val(10)
        if pd.notna(pkg) and str(pkg).strip() != "":
            try:
                val = float(pkg)
                if val % 1 != 0: 
                    log("K (PKG)", "Decimal found")
                    has_pkg_error = True
            except ValueError: 
                log("K (PKG)", "Not a number")
                has_pkg_error = True

        vgm = get_val(20)
        wgt = get_val(11)
        if pd.notna(vgm) and pd.notna(wgt):
            try:
                if float(vgm) <= float(wgt): 
                    log("U (VGM)", f"VGM ({vgm}) <= Wgt ({wgt})")
                    has_vgm_error = True
            except: pass
            
        country = get_val(13)
        dest = get_val(15)
        if pd.notna(country) and pd.notna(dest):
             c_str, d_str = str(country), str(dest)
             if is_port_code_match(d_str, c_str): pass
             elif use_internet:
                 if (c_str, d_str) in location_issues: 
                     log("N vs P", location_issues[(c_str, d_str)])
                     has_loc_error = True
             elif not use_internet:
                 if c_str.upper() not in d_str.upper() and d_str.upper() not in c_str.upper():
                     log("N vs P", f"Check '{c_str}' vs '{d_str}'")
                     has_loc_error = True

        if SPELLCHECK_AVAILABLE:
            desc = get_val(12)
            if pd.notna(desc) and str(desc).strip() != "":
                desc_text = str(desc).strip()
                if not desc_text.isdigit() and len(desc_text) > 2:
                    clean_text = re.sub(r'[^a-zA-Z\s]', '', desc_text)
                    words = clean_text.split()
                    misspelled = spell.unknown(words)
                    for w in misspelled:
                        if len(w) > 2: 
                            potential_typos.add(w.upper())
                            has_desc_error = True

    checks_passed["Container (Col C)"] = "❌ Failed" if has_cont_error else "✅ OK"
    checks_passed["Seals (Col F)"] = "❌ Failed" if has_seal_error else "✅ OK"
    checks_passed["PKG (Col K)"] = "❌ Failed" if has_pkg_error else "✅ OK"
    checks_passed["VGM vs Weight"] = "❌ Failed" if has_vgm_error else "✅ OK"
    checks_passed["Location"] = "❌ Failed" if has_loc_error else "✅ OK"
    checks_passed["Description (Col M)"] = "⚠️ Review" if has_desc_error else "✅ OK"

    return errors, potential_typos, checks_passed

# --- 5. BREAKDOWN & CONVERSION ---
def generate_breakdown(df):
    df_temp = df.copy()
    total_idx = None
    for idx, row in df_temp.iterrows():
        check_vals = [str(row.iloc[i]).upper() for i in range(min(3, len(row)))]
        if any("TOTAL" in v for v in check_vals):
            total_idx = idx
            break
    if total_idx is not None and total_idx in df_temp.index:
         int_loc = df_temp.index.get_loc(total_idx)
         if isinstance(int_loc, slice): int_loc = int_loc.start
         if hasattr(int_loc, '__iter__'): int_loc = int_loc[0]
         df_temp = df_temp.iloc[:int_loc]

    df_temp['GroupName'], df_temp['GroupType'] = zip(*df_temp.apply(determine_grouping, axis=1))
    
    stats = []
    for (name, g_type), group in df_temp.groupby(['GroupName', 'GroupType']):
        try: w_sum = pd.to_numeric(group.iloc[:, 11], errors='coerce').sum()
        except: w_sum = 0
        stats.append({
            "Sheet Name": name,
            "Grouped By": g_type,
            "Containers": len(group),
            "Total Weight": f"{w_sum:,.2f}"
        })
    return pd.DataFrame(stats)

def convert_to_template(df, header_info, logo_bytes):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})
    
    font_name = 'Times New Roman'
    fmt_bold_left = workbook.add_format({'bold': True, 'font_name': font_name, 'font_size': 11, 'align': 'left'})
    fmt_title = workbook.add_format({'bold': True, 'font_name': font_name, 'font_size': 16, 'align': 'center', 'valign': 'vcenter'})
    fmt_table_header = workbook.add_format({'bold': True, 'font_name': font_name, 'font_size': 10, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_data_center = workbook.add_format({'font_name': font_name, 'font_size': 11, 'border': 1, 'align': 'center'})
    fmt_data_left   = workbook.add_format({'font_name': font_name, 'font_size': 11, 'border': 1, 'align': 'left'})
    fmt_data_bold   = workbook.add_format({'font_name': font_name, 'font_size': 11, 'border': 1, 'bold': True, 'align': 'center'})

    df_clean = df.copy()
    total_idx = None
    for idx, row in df_clean.iterrows():
        check_vals = [str(row.iloc[i]).upper() for i in range(min(3, len(row)))]
        if any("TOTAL" in v for v in check_vals):
            total_idx = idx
            break
    if total_idx is not None and total_idx in df_clean.index:
         int_loc = df_clean.index.get_loc(total_idx)
         if isinstance(int_loc, slice): int_loc = int_loc.start
         if hasattr(int_loc, '__iter__'): int_loc = int_loc[0]
         df_clean = df_clean.iloc[:int_loc]

    if len(df_clean.columns) < 9: return None

    df_clean['GroupName'], df_clean['GroupType'] = zip(*df_clean.apply(determine_grouping, axis=1))
    unique_groups = df_clean['GroupName'].unique()

    def clean_val(val):
        if pd.isna(val) or str(val).lower() == 'nan': return ""
        return val

    def get_val(row, idx):
        if idx < len(row): return row.iloc[idx]
        return ""

    for group_name in unique_groups:
        sheet_name = str(group_name).replace("/", "-").replace("\\", "-").replace("?", "").replace("*", "").replace("[", "").replace("]", "").replace(":", "")[:30]
        ws = workbook.add_worksheet(sheet_name)
        
        ws.set_column('A:A', 5)
        ws.set_column('B:B', 15)
        ws.set_column('C:C', 6)
        ws.set_column('D:D', 15)
        ws.set_column('E:G', 6)
        ws.set_column('H:H', 10)
        ws.set_column('I:I', 10)
        ws.set_column('J:J', 6)
        ws.set_column('K:N', 10)

        if logo_bytes is not None:
            try:
                ws.insert_image('A1', 'logo.png', {'image_data': logo_bytes, 'x_scale': 0.8, 'y_scale': 0.8, 'x_offset': 5, 'y_offset': 5})
            except: pass 

        ws.write('C1', "CÔNG TY CP TÂN CẢNG CYPRESS", fmt_bold_left)
        ws.write('C2', "Cat Lai port, Nguyen Thi Dinh st, Dist. 2, Ho Chi Minh City.", fmt_bold_left)
        ws.write('C3', "Tel: (084) 08 3 7425200,               Fax: (084) 08 3 425202", fmt_bold_left)
        ws.merge_range('A4:N4', "CONTAINER LOADING LIST", fmt_title)
        ws.write('C5', "Cat Lai port, Nguyen Thi Dinh st, Dist. 2, Ho Chi Minh City.", fmt_bold_left)

        ws.write('A7', "TO", fmt_bold_left)
        ws.write('D7', "CAT LAI", fmt_bold_left)
        ws.write('A8', "MV", fmt_bold_left)
        ws.write('D8', header_info.get('mv', ''), fmt_bold_left)
        ws.write('I8', "VOY:", fmt_bold_left)
        ws.write('J8', header_info.get('voy', ''), fmt_bold_left)
        ws.write('A9', "Port of Discharge Loading", fmt_bold_left)
        ws.write('D9', "CAT LAI", fmt_bold_left)
        ws.write('I9', "ETD:", fmt_bold_left)
        ws.write('J9', header_info.get('etd', ''), fmt_bold_left)

        headers = ["Seq", "Conts Number", "Size", "Seal", "BILL NO.", "OPR", "Type", "VGM", "WeightKgs", "PKGS", "CBM", "HANDLING", "VSL", "DES"]
        for col, txt in enumerate(headers):
            ws.write(9, col, txt, fmt_table_header)

        group_data = df_clean[df_clean['GroupName'] == group_name]
        row_idx = 10
        total_weight = 0
        seq_counter = 1

        for _, row in group_data.iterrows():
            try: 
                w_val = float(get_val(row, 11))
                if pd.isna(w_val): w_val = 0
            except: w_val = 0
            total_weight += w_val
            
            ws.write(row_idx, 0, seq_counter, fmt_data_center) 
            seq_counter += 1
            ws.write(row_idx, 1, clean_val(get_val(row, 2)), fmt_data_center) 
            ws.write(row_idx, 2, clean_val(get_val(row, 3)), fmt_data_center) 
            ws.write(row_idx, 3, clean_val(get_val(row, 5)), fmt_data_left)   
            ws.write(row_idx, 4, "", fmt_data_center)
            ws.write(row_idx, 5, "", fmt_data_center)
            ws.write(row_idx, 6, "", fmt_data_center)
            ws.write(row_idx, 7, "", fmt_data_center) 
            ws.write(row_idx, 8, w_val, fmt_data_center)                   
            ws.write(row_idx, 9, clean_val(get_val(row, 10)), fmt_data_center) 
            ws.write(row_idx, 10, "", fmt_data_center)
            ws.write(row_idx, 11, "", fmt_data_center)
            ws.write(row_idx, 12, "", fmt_data_center)
            ws.write(row_idx, 13, "", fmt_data_center)
            row_idx += 1

        ws.write(row_idx, 1, "TOTAL:", fmt_data_bold)
        ws.write(row_idx, 8, total_weight, fmt_data_bold) 

    workbook.close()
    output.seek(0)
    return output

# --- 6. UI LOGIC ---
st.title("✨ Manifest Master")
st.markdown("### Intelligent Validation & Processing System")

c1, c2 = st.columns([1, 1])
uploaded_file = c1.file_uploader("1. Upload Excel Manifest 📂", type=["xlsx", "xls"])
logo_file = c2.file_uploader("2. Logo (Optional) 🖼️", type=["png", "jpg", "jpeg"])
enable_internet = st.checkbox("Enable Deep Location Check (Slower) 🌐", value=False)

if st.button("🔄 Reset / Check New File"):
    st.session_state.run_clicked = False
    st.session_state.verified_words = set()
    st.rerun()

if uploaded_file:
    if st.button("✨ Run Analysis"):
        st.session_state.run_clicked = True
    
    if st.session_state.run_clicked:
        try:
            # 1. READ RAW
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            valid_sheet_name = None
            detected_header_idx = None
            
            for sheet in sheet_names:
                df_temp = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, nrows=50)
                for idx, row in df_temp.iterrows():
                    row_str = " ".join([str(x) for x in row.values if pd.notna(x)]).upper()
                    if "CONTAINER" in row_str and ("SEQ" in row_str or "BOOKING" in row_str or "SEAL" in row_str):
                        detected_header_idx = idx
                        valid_sheet_name = sheet
                        break
                if valid_sheet_name: break 
            
            if valid_sheet_name is None:
                st.error("Could not find a valid loading list.")
                st.stop()

            # 2. READ REAL
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, sheet_name=valid_sheet_name, header=detected_header_idx)
            df.columns = [str(c) if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
            
            # 3. META
            header_info = {'mv': '', 'voy': '', 'etd': ''}
            try:
                if detected_header_idx > 0:
                    uploaded_file.seek(0)
                    df_meta = pd.read_excel(uploaded_file, sheet_name=valid_sheet_name, header=None, nrows=detected_header_idx)
                    for idx, row in df_meta.iterrows():
                        row_str = " ".join([str(x) for x in row.values if pd.notna(x)]).upper()
                        if "M/V" in row_str or "MV:" in row_str:
                            for i, val in enumerate(row.values):
                                if pd.notna(val) and ("M/V" in str(val).upper() or "MV:" in str(val).upper()):
                                    for offset in [1, 2]:
                                        if i+offset < len(row) and pd.notna(row.iloc[i+offset]):
                                            header_info['mv'] = str(row.iloc[i+offset])
                                            break
                        if "VOY" in row_str:
                            for i, val in enumerate(row.values):
                                if pd.notna(val) and "VOY" in str(val).upper():
                                    for offset in [1, 2]:
                                        if i+offset < len(row) and pd.notna(row.iloc[i+offset]):
                                            header_info['voy'] = str(row.iloc[i+offset])
                                            break
                        if "ETD" in row_str:
                            for i, val in enumerate(row.values):
                                if pd.notna(val) and "ETD" in str(val).upper():
                                    for offset in [1, 2]:
                                        if i+offset < len(row) and pd.notna(row.iloc[i+offset]):
                                            header_info['etd'] = str(row.iloc[i+offset]).replace("00:00:00", "").strip()
                                            break
            except: pass
            
            st.divider()

            # 4. RUN VALIDATION
            loader = st.empty()
            with loader:
                st.markdown(shipping_loader(), unsafe_allow_html=True)
                errors, typos, checks_passed = validate_data(df, enable_internet)
            loader.empty() # Clear Animation
            
            # 5. NEW SUMMARY DASHBOARD
            total_cnt, total_teus, total_wgt, total_vgm = calculate_summary(df)
            
            st.markdown("### 📊 Manifest Summary")
            
            def get_status(key, success_msg="Checked"):
                return f"✅ {success_msg}" if "OK" in checks_passed.get(key, "") else "⚠️ Issues Found"
            
            # Key Metrics
            st.info(f"**📦 Total {total_cnt} Cont = {total_teus} TEUs**")
            
            c1, c2 = st.columns(2)
            with c1:
                st.write(f"**🔒 Seals:** {get_status('Seals (Col F)', 'No Duplicated')}")
                st.write(f"**📝 Description:** {get_status('Description (Col M)', 'Checked')}")
                st.write(f"**📦 PKG:** {get_status('PKG (Col K)', 'Checked')}")
                st.write(f"**🌍 Location:** {get_status('Location', 'Checked')}")
            with c2:
                st.write(f"**⚖️ Weight:** {total_wgt:,.2f} kgs")
                st.write(f"**📏 VGM:** {total_vgm:,.2f} kgs")
            
            # 6. SPELLING UI
            unchecked_typos_as_errors = []
            if typos:
                st.warning("⚠️ Unrecognized Terminology Detected")
                typo_list = sorted(list(typos))
                editor_data = pd.DataFrame({
                    "Term": typo_list,
                    "Verified": [word in st.session_state.verified_words for word in typo_list]
                })
                edited_df = st.data_editor(
                    editor_data, 
                    column_config={"Verified": st.column_config.CheckboxColumn(required=True)},
                    disabled=["Term"],
                    hide_index=True,
                    key="spelling_editor"
                )
                verified_now = set(edited_df[edited_df["Verified"] == True]["Term"])
                st.session_state.verified_words.update(verified_now)
                unchecked_words = set(edited_df[edited_df["Verified"] == False]["Term"])
                if unchecked_words:
                    for w in unchecked_words:
                        unchecked_typos_as_errors.append({
                            "Ref (Col B)": "Multiple", "Column": "M (Desc)", 
                            "Error": f"Unknown Term: '{w}'"
                        })

            all_errors = errors + unchecked_typos_as_errors
            
            if all_errors:
                st.error(f"❌ Found {len(all_errors)} Issues")
                st.table(pd.DataFrame(all_errors))
            else:
                st.success("✅ All Checks Passed!")
                
                # PREVIEW BREAKDOWN
                st.markdown("### 📄 Sheet Breakdown")
                try:
                    stats_df = generate_breakdown(df)
                    st.dataframe(stats_df, use_container_width=True)
                except: pass

                st.markdown("---")
                logo_bytes = io.BytesIO(logo_file.read()) if logo_file else None
                excel_data = convert_to_template(df, header_info, logo_bytes)
                
                if excel_data:
                    st.download_button(
                        label="📥 Download Processed Manifest",
                        data=excel_data,
                        file_name="Processed_Manifest.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Conversion failed. Check file columns.")

        except Exception as e:
            st.error(f"System Error: {e}")
