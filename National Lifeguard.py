import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os
import math

# --- CONFIGURATION ---
INPUT_PDF = "95nlpool 2022_tsfillable 20250819 x.pdf"
INPUT_CSV = "roster.csv"
OUTPUT_FOLDER = "filled_forms/"

# --- CONSTANT DATA ---
HOST_DATA = {
    "Host Name": "City of Markham",
    "Host Area": "905",
    "Host Phone": "4703590 EXT 4342",
    "Host Street": "8600 McCowan Road",
    "Host City": "Markham",
    "Host Prov": "ON",
    "Host Postal": "L3P 3M2",
    "Exam Facility": "Centennial C.C.",
    "Exam Area": "905",
    "Exam Phone": "4703590 EXT 4342",
}

# --- DYNAMIC MAPPING GENERATOR (1 to 8) ---
# Updated based on user feedback:
# Address=X.5, City=X.6, Prov=X.7, Postal=X.8
# Inferred: Name=X, First=X.4, Email=X.9, Phone=X.10, DOB=X.11/12/13
candidate_map = []
for i in range(1, 9): # Candidates 1-8
    p = str(i) # The prefix (1, 2, 3...)
    entry = {
        "last_name":  p,           # Field "1"
        "first_name": f"{p}.4",    # Field "1.4"
        "addr":       f"{p}.5",    # Field "1.5" (Was 1.1)
        "city":       f"{p}.6",    # Field "1.6" (Was 1.2)
        "prov":       f"{p}.7",    # Field "1.7"
        "zip":        f"{p}.8",    # Field "1.8" (Was 1.3)
        "email":      f"{p}.9",    # Field "1.9"
        "phone":      f"{p}.10",   # Field "1.10"
        "yy":         f"{p}.11",   # Field "1.11" (Year)
        "mm":         f"{p}.12",   # Field "1.12" (Month)
        "dd":         f"{p}.13"    # Field "1.13" (Day)
    }
    candidate_map.append(entry)

def parse_name(raw_name):
    """
    Splits 'Ausar, Lautaro' into ('Ausar', 'Lautaro')
    Returns (Last, First)
    """
    if pd.isna(raw_name): return "", ""
    raw_name = str(raw_name).strip()
    
    # If format is "Last, First"
    if "," in raw_name:
        parts = raw_name.split(",")
        if len(parts) >= 2:
            return parts[0].strip(), parts[1].strip()
            
    # If format is "First Last" (fallback)
    parts = raw_name.split(" ")
    if len(parts) >= 2:
        return parts[-1], " ".join(parts[:-1])
        
    return raw_name, "" # Return whole string as Last Name if unsure

def fill_pdf(batch_df, batch_num):
    if not os.path.exists(INPUT_PDF):
        print(f"ERROR: Could not find {INPUT_PDF}")
        return

    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader)
    
    data_map = {}

    # 1. HOST DATA
    for field, value in HOST_DATA.items():
        data_map[field] = value
    
    # 2. CANDIDATE DATA
    for i, (idx, row) in enumerate(batch_df.iterrows()):
        if i >= len(candidate_map): break
        
        fields = candidate_map[i]
        
        # Name Split
        last, first = parse_name(row.get("AttendeeName", ""))
        
        # Date Parsing
        raw_dob = row.get("DateOfBirth", "")
        dd, mm, yy = "", "", ""
        if pd.notna(raw_dob):
            try:
                dt = pd.to_datetime(raw_dob, dayfirst=True)
                dd = str(dt.day).zfill(2)
                mm = str(dt.month).zfill(2)
                yy = str(dt.year)
            except: pass

        # Map to fields
        data_map[fields["last_name"]] = last
        data_map[fields["first_name"]] = first
        data_map[fields["addr"]] = str(row.get("Street", ""))
        data_map[fields["city"]] = str(row.get("City", ""))
        data_map[fields["prov"]] = str(row.get("Province", "ON")) # Default to ON if missing
        data_map[fields["zip"]] = str(row.get("PostalCode", ""))
        data_map[fields["email"]] = str(row.get("E-mail", ""))
        data_map[fields["phone"]] = str(row.get("AttendeePhone", ""))
        data_map[fields["dd"]] = dd
        data_map[fields["mm"]] = mm
        data_map[fields["yy"]] = yy

    # Apply to all pages
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    # --- FIXES ---
    if "/AcroForm" not in writer.root_object:
        writer.root_object.update({NameObject("/AcroForm"): DictionaryObject()})
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)
    
    if "/OCProperties" in reader.root_object:
        writer.root_object[NameObject("/OCProperties")] = \
            reader.root_object["/OCProperties"].clone(writer)

    output_filename = f"{OUTPUT_FOLDER}NL_Pool_Sheet_{batch_num}.pdf"
    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"Generated: {output_filename}")

# --- MAIN EXECUTION ---
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

print("Reading roster.csv...")
if os.path.exists(INPUT_CSV):
    df = pd.read_csv(INPUT_CSV, dtype=str).fillna("")
    BATCH_SIZE = 8
    total_batches = math.ceil(len(df) / BATCH_SIZE)

    print(f"Processing {len(df)} candidates into {total_batches} batch(es)...")
    for i in range(total_batches):
        batch = df.iloc[i * BATCH_SIZE : (i + 1) * BATCH_SIZE]
        fill_pdf(batch, i + 1)
    print("Done.")