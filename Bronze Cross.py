import pandas as pd
from pypdf import PdfReader, PdfWriter
import os
import math

# --- CONFIGURATION ---
INPUT_PDF = "95tsbronzecross2020_fillable.pdf"
INPUT_CSV = "roster.csv"
OUTPUT_FOLDER = "filled_forms/"

# --- INVOICING DATA (HOST & FACILITY) ---
HOST_DATA = {
    "host_name": "City of Markham",
    "host_area_code": "905",
    "host_phone_num": "4703590 EXT 4342",
    "host_addr": "8600 McCowan Road",
    "host_city": "Markham",
    "host_prov": "ON",
    "host_postal": "L3P 3M2",
    "facility_name": "Centennial C.C."
}

# --- FIELD MAPPING FOR INVOICE ---
HOST_FIELD_MAP = {
    "host_name": "Text19",      
    "host_area_code": "Text20", 
    "host_phone_num": "Text21", 
    "host_addr": "Text22",      
    "host_city": "Text23",      
    "host_prov": "Text24",      
    "host_postal": "Text25",    
    "facility_name": "Text29"   
}

# --- CANDIDATE MAPPING ---
candidate_map = [
    # === PAGE 1 (Candidates 1-6) ===
    # These use the standard "Address1.0", "Address1.1.0" etc.
    {"p": "", "s": ".0"},           # 1
    {"p": "", "s": ".1.0"},         # 2
    {"p": "", "s": ".1.1.0"},       # 3 
    {"p": "", "s": ".1.1.1.0"},     # 4
    {"p": "", "s": ".1.1.1.1.0"},   # 5
    {"p": "", "s": ".1.1.1.1.1"},   # 6

    # === PAGE 2 (Candidates 7-13) ===
    # These use prefixes "7", "8", etc.
    {"p": "7", "s": ".0"},          # 7
    {"p": "8", "s": ".1.0"},        # 8
    
    # 9: FIXED
    # Removed "Address1.0", "Address1.1.0", "Address1.1.1.0" (Which belong to Cand 1, 2, 3)
    # Added "9Address1.1.1.0" (Logical pattern matches 7 & 8)
    {"p": "9", "s": ".1.1.0", "addr_override": [
        "9Address1.1.1.0",  # The logical field name for 9
        "Address1.1.1.0X",  # The specific ghost field
    ]},
    
    {"p": "10", "s": ".1.1.1.0", "name_override": "10"}, # 10
    {"p": "11", "s": ".1.1.1.1.0"},     # 11
    {"p": "12", "s": ".1.1.1.1.1"},     # 12
    {"p": "13", "s": ".1.1.1.1.1"},     # 13
]

def clean_name(raw_name):
    if pd.isna(raw_name): return ""
    raw_name = str(raw_name)
    if "," in raw_name:
        parts = raw_name.split(",")
        if len(parts) >= 2:
            return f"{parts[1].strip()} {parts[0].strip()}"
    return raw_name

def fill_pdf(batch_df, batch_num):
    if not os.path.exists(INPUT_PDF):
        print(f"ERROR: Could not find {INPUT_PDF}")
        return

    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader)
    
    data_map = {}
    
    # --- 1. APPLY HOST & FACILITY DATA ---
    for key, pdf_field in HOST_FIELD_MAP.items():
        data_map[pdf_field] = HOST_DATA[key]

    # --- 2. APPLY CANDIDATE DATA ---
    for i, (idx, row) in enumerate(batch_df.iterrows()):
        if i >= len(candidate_map): break
        
        slot = candidate_map[i]
        p = slot.get("p", "") 
        s = slot.get("s", "") 
        
        prefix_str = p if p else ""
        
        # Build Standard Field Names
        f_name = f"{prefix_str}Name1{s}"
        f_addr = f"{prefix_str}Address1{s}"
        f_city = f"{prefix_str}City1{s}"
        f_zip  = f"{prefix_str}Postal1{s}"
        f_email = f"{prefix_str}Email1{s}"
        f_phone = f"{prefix_str}Phone1{s}"
        f_dd = f"{prefix_str}DOBD1{s}"
        f_mm = f"{prefix_str}DOBM1{s}"
        f_yy = f"{prefix_str}DOBY1{s}"

        # Handle Name Override
        if "name_override" in slot:
            f_name = slot["name_override"]

        # Data Preparation
        full_name = clean_name(row.get("AttendeeName", ""))
        address_val = str(row.get("Street", ""))
        
        raw_dob = row.get("DateOfBirth", "")
        dd, mm, yy = "", "", ""
        if pd.notna(raw_dob):
            try:
                dt = pd.to_datetime(raw_dob, dayfirst=True)
                dd = str(dt.day).zfill(2)
                mm = str(dt.month).zfill(2)
                yy = str(dt.year)[-2:] # 2 Digit Year
            except: pass

        # Map Standard Data
        data_map[f_name] = full_name
        data_map[f_city] = str(row.get("City", ""))
        data_map[f_zip] = str(row.get("PostalCode", ""))
        data_map[f_email] = str(row.get("E-mail", ""))
        data_map[f_phone] = str(row.get("AttendeePhone", ""))
        data_map[f_dd] = dd
        data_map[f_mm] = mm
        data_map[f_yy] = yy

        # --- ADDRESS HANDLING ---
        if "addr_override" in slot:
            override = slot["addr_override"]
            if isinstance(override, list):
                # Write to all fields in the safe list
                for field in override:
                    data_map[field] = address_val
            else:
                data_map[override] = address_val
        else:
            data_map[f_addr] = address_val

    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    output_filename = f"{OUTPUT_FOLDER}Bronze_Cross_Test_Sheet_{batch_num}.pdf"
    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"Generated: {output_filename}")

# --- MAIN EXECUTION ---
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

print("Reading roster.csv...")
if os.path.exists(INPUT_CSV):
    df = pd.read_csv(INPUT_CSV, dtype=str).fillna("")

    BATCH_SIZE = 13
    total_batches = math.ceil(len(df) / BATCH_SIZE)

    print(f"Processing {len(df)} candidates into {total_batches} batch(es)...")

    for i in range(total_batches):
        batch = df.iloc[i * BATCH_SIZE : (i + 1) * BATCH_SIZE]
        fill_pdf(batch, i + 1)

    print("Done.")