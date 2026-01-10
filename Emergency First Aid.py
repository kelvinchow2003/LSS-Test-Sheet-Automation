import pandas as pd
from pypdf import PdfReader, PdfWriter
import os
import math

# --- CONFIGURATION ---
INPUT_PDF = "95efa_on2014.pdf" 
INPUT_CSV = "roster.csv"
OUTPUT_FOLDER = "filled_forms/"

# --- CONSTANT DATA (HOST & FACILITY) ---
# The dump confirms these are SPLIT fields (Area Code + Number).
# We map the specific pieces of data to the exact names found in your dump.
HOST_DATA = {
    "Host Name": "City of Markham",
    "Host Address": "8600 McCowan Road",
    "Host City": "Markham",
    "Host Province": "ON",
    "Host Postal Code": "L3P 3M2",
    
    # SPLIT PHONES (Matches your dump exactly)
    "Host Area Code": "905",
    "Host Number": "470-3590 EXT 4342",
    
    "Facility Name": "Centennial C.C.",
    "Facility Area Code": "905",
    "Facility Number": "470-3590 EXT 4342",
    
    # SHOTGUN FALLBACKS (In case there are hidden fields)
    "Host Phone": "905-470-3590",
    "Facility Phone": "905-470-3590",
    "Telephone": "905-470-3590",
    "Phone": "905-470-3590"
}

# --- THE MAPPING ---
candidate_map = []

# Generate the map for candidates 1 to 10
for i in range(1, 11):
    suffix = str(i)
    
    entry = {
        "name": f"Name {suffix}",
        "addr": f"Address {suffix}",
        "apt":  f"apt {suffix}",      # Lowercase "apt" from dump
        "city": f"City {suffix}",
        "zip":  f"Postal {suffix}",
        "email": f"Email {suffix}",
        "phone": f"Phone {suffix}",
        "dd": f"Day {suffix}",
        "mm": f"Month {suffix}",
        "yy": f"Year {suffix}"
    }
    
    # SPECIAL CASE: Candidate 10's Name field is just "10"
    if i == 10:
        entry["name"] = "10"
        
    candidate_map.append(entry)

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
        print(f"ERROR: Could not find {INPUT_PDF}.")
        return

    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader)
    
    data_map = {}
    
    # 1. APPLY HOST & FACILITY DATA
    for field, value in HOST_DATA.items():
        data_map[field] = value

    # 2. APPLY CANDIDATE DATA
    for i, (idx, row) in enumerate(batch_df.iterrows()):
        if i >= len(candidate_map): break
        
        fields = candidate_map[i]
        
        full_name = clean_name(row.get("AttendeeName", ""))
        
        raw_dob = row.get("DateOfBirth", "")
        dd, mm, yy = "", "", ""
        if pd.notna(raw_dob):
            try:
                dt = pd.to_datetime(raw_dob, dayfirst=True)
                dd = str(dt.day).zfill(2)
                mm = str(dt.month).zfill(2)
                yy = str(dt.year)[-2:] 
            except: pass

        data_map[fields["name"]] = full_name
        data_map[fields["addr"]] = str(row.get("Street", ""))
        data_map[fields["city"]] = str(row.get("City", ""))
        data_map[fields["zip"]] = str(row.get("PostalCode", ""))
        data_map[fields["email"]] = str(row.get("E-mail", ""))
        data_map[fields["phone"]] = str(row.get("AttendeePhone", ""))
        data_map[fields["dd"]] = dd
        data_map[fields["mm"]] = mm
        data_map[fields["yy"]] = yy

    # Apply to all pages
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    output_filename = f"{OUTPUT_FOLDER}EFA_Test_Sheet_{batch_num}.pdf"
    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"Generated: {output_filename}")

# --- MAIN EXECUTION ---
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

print("Reading roster.csv...")
if os.path.exists(INPUT_CSV):
    df = pd.read_csv(INPUT_CSV, dtype=str).fillna("")
    BATCH_SIZE = 10
    total_batches = math.ceil(len(df) / BATCH_SIZE)

    for i in range(total_batches):
        batch = df.iloc[i * BATCH_SIZE : (i + 1) * BATCH_SIZE]
        fill_pdf(batch, i + 1)
    print("Done.")
else:
    print(f"ERROR: {INPUT_CSV} not found.")