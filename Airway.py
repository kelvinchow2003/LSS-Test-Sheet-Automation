import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os
import math

# --- CONFIGURATION ---
INPUT_PDF = "95airwaymanagement2022-fillable.pdf"
INPUT_CSV = "roster.csv"
OUTPUT_FOLDER = "filled_forms/"

# --- INVOICING DATA (HOST & FACILITY) ---
HOST_FIELD_MAP = {
    # FRONT PAGE
    "Host Name": "City of Markham",
    "Host Area Code": "905",
    "Host Telephone #": "4703590 EXT 4342",
    "Host Address": "8600 McCowan Road",
    "Host City": "Markham",
    "Host Prov": "ON",
    "Host Postal Code": "L3P 3M2",
    "Facility Name": "Centennial C.C.",
    "Facility Area Code": "905",
    "Facility Telephone #": "4703590 EXT 4342",

    # REVERSE PAGE
    "Host Name Reverse": "City of Markham",
    "Host Area Code Reverse": "905",
    "Host Telephone # Reverse": "4703590 EXT 4342",
    "Facility Name Reverse": "Centennial C.C.",
    "Facility Area Code Reverse": "905",
    "Facility Telephone # Reverse": "4703590 EXT 4342"
}

# --- CANDIDATE MAPPING ---
candidate_map = []

# Generate map for 1-10
for i in range(1, 11):
    s = str(i)
    
    # HANDLE TYPO IN PDF: Candidate 5 has "postal code5" (no space)
    if i == 5:
        p_code = "postal code5"
    else:
        p_code = f"postal code {s}"

    entry = {
        "name": f"Name {s}",
        "addr": f"address {s}", 
        "apt":  f"apt# {s}",
        "city": f"city {s}",
        "zip":  p_code,
        "email": f"email {s}",
        "phone": f"phone {s}",
        "dd": f"day {s}",
        "mm": f"month {s}",
        "yy": f"year {s}"
    }
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
        print(f"ERROR: Could not find {INPUT_PDF}")
        return

    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader)
    
    data_map = {}
    
    # 1. APPLY HOST & FACILITY DATA
    for field_name, value in HOST_FIELD_MAP.items():
        data_map[field_name] = value

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

    # Apply data to all pages
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    # =========================================================
    # CRITICAL FIX 1: Fix "Floating Text" / Font Issues
    # =========================================================
    # This forces the PDF viewer to regenerate the field appearance
    # using the native form fonts rather than generic text.
    if "/AcroForm" not in writer.root_object:
        writer.root_object.update({
            NameObject("/AcroForm"): DictionaryObject()
        })
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    # =========================================================
    # CRITICAL FIX 2: Fix "French text on top of English"
    # =========================================================
    # This copies the Layer Visibility settings (OCProperties) from
    # the original file. Without this, hidden layers (like the French
    # translation) become visible and stack on top of everything.
    if "/OCProperties" in reader.root_object:
        writer.root_object[NameObject("/OCProperties")] = \
            reader.root_object["/OCProperties"].clone(writer)
    # =========================================================

    output_filename = f"{OUTPUT_FOLDER}Airway_Mgmt_Test_Sheet_{batch_num}.pdf"
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

    print(f"Processing {len(df)} candidates into {total_batches} batch(es)...")

    for i in range(total_batches):
        batch = df.iloc[i * BATCH_SIZE : (i + 1) * BATCH_SIZE]
        fill_pdf(batch, i + 1)

    print("Done.")
else:
    print(f"Error: {INPUT_CSV} not found.")