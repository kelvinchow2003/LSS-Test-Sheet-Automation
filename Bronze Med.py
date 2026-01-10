import pandas as pd
from pypdf import PdfReader, PdfWriter
import os
import math

# --- CONFIGURATION ---
INPUT_PDF = "95tsbronzemedallion2020_fillable.pdf"
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
    "facility_name": "Centennial C.C.",
    
    # EXAM TELEPHONE (Same as Host)
    # Using the Text30/Text31 slots which follow the Facility Name
    "exam_area_code": "905",
    "exam_phone_num": "4703590 EXT 4342"
}

# --- FIELD MAPPING ---
# Text19 = Host Name
# Text20 = Area Code
# Text21 = Phone Number
# Text22 = Address
# ...
# Text29 = Facility Name
# Text30 = Facility Area Code (deduced)
# Text31 = Facility Phone Num (deduced)
HOST_FIELD_MAP = {
    "host_name": "Text19",      
    "host_area_code": "Text20", 
    "host_phone_num": "Text21", 
    "host_addr": "Text22",      
    "host_city": "Text23",      
    "host_prov": "Text24",      
    "host_postal": "Text25",    
    "facility_name": "Text29",
    "exam_area_code": "Text30", # New
    "exam_phone_num": "Text31"  # New
}

# --- CANDIDATE MAPPING (UNCHANGED) ---
candidate_map = [
    # === PAGE 1 (Candidates 1-6) ===
    {"base": "1", "s": ".0"},           # 1
    {"base": "1", "s": ".1.0"},         # 2
    {"base": "1", "s": ".1.1.0"},       # 3
    {"base": "1", "s": ".1.1.1.0"},     # 4
    {"base": "1", "s": ".1.1.1.1.0"},   # 5
    {"base": "1", "s": ".1.1.1.1.1"},   # 6

    # === PAGE 2 (Candidates 7-13) ===
    {"base": "", "s": ".0.0"},          # 7
    {"base": "", "s": ".0.1.0"},        # 8
    {"base": "", "s": ".0.1.1.0"},      # 9
    {"base": "", "s": ".0.1.1.1.0"},    # 10
    {"base": "", "s": ".0.1.1.1.1.0"},  # 11
    {"base": "", "s": ".0.1.1.1.1.1.0"},# 12
    {"base": "", "s": ".0.1.1.1.1.1.1"} # 13
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
        b = slot["base"]
        s = slot["s"]
        
        # Build field names
        f_name = f"Name{b}{s}"
        f_addr = f"Address{b}{s}"
        f_city = f"City{b}{s}"
        f_zip  = f"Postal{b}{s}"
        f_email = f"Email{b}{s}"
        f_phone = f"Phone{b}{s}"
        f_dd = f"DOBD{b}{s}"
        f_mm = f"DOBM{b}{s}"
        f_yy = f"DOBY{b}{s}"

        # Clean Name
        full_name = clean_name(row.get("AttendeeName", ""))
        
        # Parse Date
        raw_dob = row.get("DateOfBirth", "")
        dd, mm, yy = "", "", ""
        if pd.notna(raw_dob):
            try:
                dt = pd.to_datetime(raw_dob, dayfirst=True)
                dd = str(dt.day).zfill(2)
                mm = str(dt.month).zfill(2)
                yy = str(dt.year)[-2:] 
            except: pass

        # Map Data
        data_map[f_name] = full_name
        data_map[f_addr] = str(row.get("Street", ""))
        data_map[f_city] = str(row.get("City", ""))
        data_map[f_zip] = str(row.get("PostalCode", ""))
        data_map[f_email] = str(row.get("E-mail", ""))
        data_map[f_phone] = str(row.get("AttendeePhone", ""))
        data_map[f_dd] = dd
        data_map[f_mm] = mm
        data_map[f_yy] = yy

    # Apply to all pages
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    output_filename = f"{OUTPUT_FOLDER}Bronze_Medallion_Test_Sheet_{batch_num}.pdf"
    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"Generated: {output_filename}")

# --- MAIN EXECUTION ---
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

print("Reading roster.csv...")
df = pd.read_csv(INPUT_CSV, dtype=str).fillna("")

BATCH_SIZE = 13
total_batches = math.ceil(len(df) / BATCH_SIZE)

print(f"Processing {len(df)} candidates into {total_batches} batch(es)...")

for i in range(total_batches):
    batch = df.iloc[i * BATCH_SIZE : (i + 1) * BATCH_SIZE]
    fill_pdf(batch, i + 1)

print("Done.")