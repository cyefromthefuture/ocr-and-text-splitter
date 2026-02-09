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

# --- CUSTOM CSS & ANIMATION ---
def inject_custom_css():
    st.markdown("""
        <style>
        /* 1. ANIMATED OCEAN BACKGROUND */
        .stApp {
            background: linear-gradient(-45deg, #021B79, #0575E6, #00F260, #0575E6);
            background-size: 400% 400%;
            animation: gradient 15s ease infinite;
            color: white;
        }
        @keyframes gradient {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }

        /* 2. GLASSMORPHISM CONTAINERS */
        .stDataFrame, .stTable, div[data-testid="stFileUploader"] {
            background: rgba(255, 255, 255, 0.1);
            backdrop-filter: blur(10px);
            border-radius: 15px;
            padding: 20px;
            border: 1px solid rgba(255, 255, 255, 0.2);
        }
        
        /* 3. HEADERS */
        h1, h2, h3 {
            color: #FFFFFF !important;
            text-shadow: 2px 2px 4px #000000;
        }
        p, label {
            color: #E0E0E0 !important;
            font-weight: 500;
        }

        /* 4. BUTTONS */
        .stButton>button {
            background-color: #FF4B4B;
            color: white;
            border-radius: 20px;
            border: none;
            padding: 10px 24px;
            font-weight: bold;
            transition: all 0.3s;
        }
        .stButton>button:hover {
            transform: scale(1.05);
            background-color: #FF2B2B;
            box-shadow: 0px 0px 15px rgba(255, 75, 75, 0.7);
        }
        </style>
    """, unsafe_allow_html=True)

# --- CUSTOM SHIPPING LOADER ---
def shipping_loader():
    """Returns the HTML for a shipping animation"""
    return """
    <div style="display: flex; justify-content: center; align-items: center; margin: 20px 0;">
        <div class="loader">
            <style>
                .loader {
                    width: 100px;
                    height: 100px;
                    position: relative;
                }
                .ship {
                    font-size: 50px;
                    position: absolute;
                    bottom: 20px;
                    left: 20px;
                    animation: sail 2s infinite ease-in-out;
                }
                .wave {
                    position: absolute;
                    bottom: 0;
                    left: 0;
                    width: 100%;
                    height: 10px;
                    background: rgba(255,255,255,0.6);
                    border-radius: 10px;
                    animation: wave 1s infinite linear;
                }
                @keyframes sail {
                    0%, 100% { transform: rotate(0deg) translateY(0); }
                    25% { transform: rotate(5deg) translateY(-5px); }
                    75% { transform: rotate(-5deg) translateY(5px); }
                }
                @keyframes wave {
                    0% { transform: translateX(-5px); }
                    50% { transform: translateX(5px); }
                    100% { transform: translateX(-5px); }
                }
            </style>
            <div class="ship">🚢</div>
            <div class="wave"></div>
        </div>
        <h3 style="margin-left: 20px;">Processing Manifest...</h3>
    </div>
    """

# Initialize CSS
inject_custom_css()

st.title("🚢 Manifest Validator & Converter")

# --- SESSION STATE ---
if 'verified_words' not in st.session_state:
    st.session_state.verified_words = set()
if 'run_clicked' not in st.session_state:
    st.session_state.run_clicked = False

# --- 1. SETUP ONLINE CHECKER ---
geolocator = Nominatim(user_agent="my_logistics_checker_v13")
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

# --- 3. VALIDATION ENGINE ---
def validate_data(df, use_internet):
    errors = []
    potential_typos = set()
    
    if len(df.columns) < 5:
        errors.append({"Ref (Col B)": "CRITICAL", "Column": "Structure", "Error": f"Columns missing. Found only {len(df.columns)} columns. Header detection failed."})
        return errors, potential_typos

    duplicate_seals_list = []
    if len(df.columns) > 5:
        all_seals = df.iloc[:, 5].astype(str).str.strip()
        all_seals = all_seals[~all_seals.isin(['nan', 'NaN', ''])]
        duplicate_seals_list = all_seals[all_seals.duplicated()].unique()

    location_issues = {}
    if use_internet and len(df.columns) > 15:
        with st.empty():
            st.markdown(shipping_loader(), unsafe_allow_html=True) # SHOW ANIMATION
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
                log("C (Container)", "Invalid format (Must be 4 Letters + 7 Digits)")

        seal = get_val(5)
        if pd.notna(seal) and str(seal).strip() != "":
            if str(seal).strip() in duplicate_seals_list: log("F (Seal)", "Duplicate Seal Number found")

        pkg = get_val(10)
        if pd.notna(pkg) and str(pkg).strip() != "":
            try:
                val = float(pkg)
                if val % 1 != 0: log("K (PKG)", f"Value '{pkg}' has decimals. Must be an Integer.")
            except ValueError: log("K (PKG)", "Not a valid number")

        vgm = get_val(20)
        wgt = get_val(11)
        if pd.notna(vgm) and pd.notna(wgt):
            try:
                if float(vgm) <= float(wgt): log("U (VGM)", f"VGM ({vgm}) must be > Weight ({wgt})")
            except: pass
            
        country = get_val(13)
        dest = get_val(15)
        if pd.notna(country) and pd.notna(dest):
             c_str, d_str = str(country), str(dest)
             if is_port_code_match(d_str, c_str): pass
             elif use_internet:
                 if (c_str, d_str) in location_issues: log("N vs P", location_issues[(c_str, d_str)])
             elif not use_internet:
                 if c_str.upper() not in d_str.upper() and d_str.upper() not in c_str.upper():
                     log("N vs P", f"Check Country '{c_str}' vs Dest '{d_str}'")

        if SPELLCHECK_AVAILABLE:
            desc = get_val(12)
            if pd.notna(desc) and str(desc).strip() != "":
                desc_text = str(desc).strip()
                if not desc_text.isdigit() and len(desc_text) > 2:
                    clean_text = re.sub(r'[^a-zA-Z\s]', '', desc_text)
                    words = clean_text.split()
                    misspelled = spell.unknown(words)
                    for w in misspelled:
                        if len(w) > 2: potential_typos.add(w.upper())

    return errors, potential_typos

# --- 4. CONVERSION ENGINE ---
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

    # 1. Clean Data & Slice at TOTAL
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

    if len(df_clean.columns) < 9: 
        return None

    # --- 2. SMART GROUPING LOGIC ---
    # Case 1: Valid B/L -> Group by B/L
    # Case 2: Invalid B/L (#N/A, IN CY, Date) -> Group by Owner
    def get_sheet_group(row):
        owner = str(row.iloc[6]).strip()
        bl_val = row.iloc[8]
        
        # Check null/empty
        if pd.isna(bl_val):
            return owner
            
        bl_str = str(bl_val).strip().upper()
        
        # Keyword checks
        if bl_str in ["#N/A", "IN CY", "NAN", ""]:
            return owner
            
        # Date checks
        if isinstance(bl_val, (datetime.datetime, datetime.date, pd.Timestamp)):
            return owner
        if "00:00:00" in bl_str: # String timestamp format
            return owner
            
        return bl_str # Return B/L No

    df_clean['_SheetGroup'] = df_clean.apply(get_sheet_group, axis=1)
    unique_groups = df_clean['_SheetGroup'].unique()

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

        if logo_bytes:
            ws.insert_image('A1', 'logo.png', {'image_data': logo_bytes, 'x_scale': 0.8, 'y_scale': 0.8, 'x_offset': 5, 'y_offset': 5})

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

        group_data = df_clean[df_clean['_SheetGroup'] == group_name]
        
        row_idx = 10
        total_weight = 0
        seq_counter = 1 # RESET SEQ FOR EACH SHEET

        for _, row in group_data.iterrows():
            try: 
                w_val = float(get_val(row, 11))
                if pd.isna(w_val): w_val = 0
            except: w_val = 0
            total_weight += w_val
            
            # Use Auto-Increment SEQ
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

# --- UI LOGIC ---
c1, c2 = st.columns([1, 1])
uploaded_file = c1.file_uploader("1. Upload Loading List (Excel)", type=["xlsx", "xls"])
logo_file = c2.file_uploader("2. Upload Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
enable_internet = st.checkbox("Enable Internet Location Check", value=False)

if uploaded_file:
    if st.button("Run Validation"):
        st.session_state.run_clicked = True
    
    if st.session_state.run_clicked:
        try:
            # 1. FIND CORRECT SHEET AND HEADER
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            valid_sheet_name = None
            detected_header_idx = None
            
            for sheet in sheet_names:
                df_temp = pd.read_excel(uploaded_file, sheet_name=sheet, header=None, nrows=50)
                
                for idx, row in df_temp.iterrows():
                    row_str = " ".join([str(x) for x in row.values if pd.notna(x)]).upper()
                    if "CONTAINER" in row_str and ("SEQ" in row_str or "BOOKING" in row_str or "SEAL" in row_str or "SIZE" in row_str):
                        detected_header_idx = idx
                        valid_sheet_name = sheet
                        break
                
                if valid_sheet_name:
                    break 
            
            if valid_sheet_name is None:
                st.error("Could not find a valid loading list in any sheet (looking for 'Container', 'Seal', etc).")
                st.stop()

            # 2. READ REAL DATA
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, sheet_name=valid_sheet_name, header=detected_header_idx)
            df.columns = [str(c) if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
            
            # 3. META INFO
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

            # 4. RUN VALIDATION (WITH ANIMATION)
            with st.empty():
                st.markdown(shipping_loader(), unsafe_allow_html=True) # SHOW ANIMATION
                time.sleep(1.5) # Fake delay for effect
                errors, typos = validate_data(df, enable_internet)
            
            # 5. SPELLING UI
            unchecked_typos_as_errors = []
            if typos:
                st.info("ℹ️ Spelling Review")
                st.write("Unrecognized words found. Check box if correct.")
                
                typo_list = sorted(list(typos))
                editor_data = pd.DataFrame({
                    "Word": typo_list,
                    "Is Correct?": [word in st.session_state.verified_words for word in typo_list]
                })
                
                edited_df = st.data_editor(
                    editor_data, 
                    column_config={"Is Correct?": st.column_config.CheckboxColumn(required=True)},
                    disabled=["Word"],
                    hide_index=True,
                    key="spelling_editor"
                )
                
                verified_now = set(edited_df[edited_df["Is Correct?"] == True]["Word"])
                st.session_state.verified_words.update(verified_now)
                
                unchecked_words = set(edited_df[edited_df["Is Correct?"] == False]["Word"])
                if unchecked_words:
                    for w in unchecked_words:
                        unchecked_typos_as_errors.append({
                            "Ref (Col B)": "Multiple", "Column": "M (Desc)", 
                            "Error": f"Spelling Error: '{w}' (Unchecked)"
                        })

            all_errors = errors + unchecked_typos_as_errors
            
            if all_errors:
                st.error(f"❌ Validation Failed: {len(all_errors)} Errors.")
                st.warning("Fix errors in Excel OR check spelling boxes.")
                st.table(pd.DataFrame(all_errors))
            else:
                st.success("✅ All Checks Passed!")
                logo_bytes = io.BytesIO(logo_file.read()) if logo_file else None
                
                excel_data = convert_to_template(df, header_info, logo_bytes)
                
                if excel_data:
                    st.download_button(
                        label="📥 Download UP_Converted.xlsx",
                        data=excel_data,
                        file_name="UP_Converted.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Conversion failed. Check file columns.")

        except Exception as e:
            st.error(f"System Error: {e}")
