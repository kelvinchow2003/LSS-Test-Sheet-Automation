import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os
import re

# --- CONFIGURATION ---
INPUT_PDF = "basicfirstaidwcprcandaed_20260619_fillable.pdf"
INPUT_CSV = "EventRosterReportSeries (2).csv" # Update to your CSV filename
OUTPUT_FOLDER = "filled_forms/"

# --- CONSTANT DATA ---
HOST_DATA = {
    "Invoicing Info - Host Name 1": "City of Markham",
    "Invoicing Info - Phone Number 1": "905-470-3590 EXT 4342",
    "Invoicing Info - Address 1": "8600 McCowan Road",
    "Invoicing Info - City 1": "Markham",
    "Invoicing Info - Province 1": "ON",
    "Invoicing Info - Postal Code 1": "L3P 3M2",
    
    "Exam Info - Facility Name 1": "Centennial C.C.",
    "Exam Info - Phone Number 1": "905-470-3590 EXT 4342",
}

# --- HELPER: MAP ROW TO SLOT ---
def get_slot_data(row, slot_number):
    """Maps a CSV row to the PDF fields for a specific slot (1-8)."""
    p = str(slot_number)
    
    # 1. Name & Address Logic
    full_name = str(row.get("AttendeeName", "")).strip()
    street = str(row.get("Street", ""))
    city = str(row.get("City", ""))
    prov = str(row.get("Province", "ON"))
    postal = str(row.get("PostalCode", ""))
    email = str(row.get("E-mail", ""))
    
    # 2. Phone Formatting (Area Code Separated)
    raw_phone = str(row.get("AttendeePhone", ""))
    digits = re.sub(r'\D', '', raw_phone) # Extract only numbers
    if len(digits) >= 10:
        # Formats to: 905-470-3590
        phone = f"{digits[:3]}-{digits[3:6]}-{digits[6:10]}"
    else:
        phone = raw_phone # Fallback if phone number is too short
    
    # 3. DOB Formatting (YY/MM/DD)
    raw_dob = row.get("DateOfBirth", "")
    dob_formatted = ""
    if pd.notna(raw_dob):
        try:
            dt = pd.to_datetime(raw_dob, dayfirst=True)
            dd = str(dt.day).zfill(2)
            mm = str(dt.month).zfill(2)
            yy = str(dt.year)[-2:] # YY format
            dob_formatted = f"{yy}/{mm}/{dd}"
        except: pass

    # 4. AGGRESSIVE MAPPING (Accounts for form typos and Acrobat copy-paste errors)
    data = {}
    
    # Standard expected variations (catches spacing/capitalization typos)
    variations = [
        (f"Name {p}", full_name),
        (f"DOB {p}", dob_formatted),
        (f"D.O.B. {p}", dob_formatted),
        (f"Address {p}", street),
        (f"City {p}", city),
        (f"Province {p}", prov),
        
        # Postal Code Fallbacks
        (f"Postal Code {p}", postal),
        (f"Postal code {p}", postal),
        (f"PostalCode {p}", postal),
        
        # Phone Fallbacks
        (f"Phone Number {p}", phone),
        (f"Phone {p}", phone),
        
        (f"Email {p}", email),
    ]
    
    # Acrobat Page 2 Duplication Fallbacks (Fixes Candidates 5-8)
    if 5 <= slot_number <= 8:
        base = slot_number - 4  # Maps 5->1, 6->2, etc.
        variations.extend([
            (f"Name {base}_2", full_name),
            (f"DOB {base}_2", dob_formatted),
            (f"Address {base}_2", street),
            (f"City {base}_2", city),
            (f"Province {base}_2", prov),
            (f"Postal Code {base}_2", postal),
            (f"Postal code {base}_2", postal),
            (f"Phone Number {base}_2", phone),
            (f"Email {base}_2", email),
        ])
        
    # Apply all variations to the data map (PDF ignores the ones that don't exist)
    for key, val in variations:
        data[key] = val

    return data

# --- SAVE FUNCTION ---
def _finalize_and_save(writer, reader, data_map, index, suffix):
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    # Force Adobe/Chrome to re-render text
    if "/AcroForm" not in writer.root_object:
        writer.root_object.update({NameObject("/AcroForm"): DictionaryObject()})
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    # Clone Layer Settings
    if "/OCProperties" in reader.root_object:
        writer.root_object[NameObject("/OCProperties")] = \
            reader.root_object["/OCProperties"].clone(writer)

    output_filename = f"{OUTPUT_FOLDER}BasicFirstAid_{index}_{suffix}.pdf"
    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"Generated: {output_filename}")

# --- MASTER FILE GENERATOR (Candidates 1-8) ---
def create_master_file(batch_df, file_index):
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader) 

    data_map = HOST_DATA.copy()
    
    for i, (idx, row) in enumerate(batch_df.iterrows()):
        current_num = i + 1  # Slots 1-8
        data_map.update(get_slot_data(row, slot_number=current_num))

    _finalize_and_save(writer, reader, data_map, file_index, "Master")

# --- CONTINUATION FILE GENERATOR (Candidates 9+) ---
def create_continuation_file(batch_df, file_index):
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader) # Duplicate full form (Front + Back)

    data_map = HOST_DATA.copy()

    for i, (idx, row) in enumerate(batch_df.iterrows()):
        slot_id = i + 1  # Reuse PDF slots 1-8 for the new candidates
        data_map.update(get_slot_data(row, slot_number=slot_id))

    _finalize_and_save(writer, reader, data_map, file_index, "Continuation")

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    print(f"Reading {INPUT_CSV}...")
    if os.path.exists(INPUT_CSV):
        df = pd.read_csv(INPUT_CSV, dtype=str).fillna("")
        
        total_candidates = len(df)
        print(f"Found {total_candidates} candidates.")

        BATCH_SIZE = 8 # Basic First Aid fits 8 candidates (4 front, 4 back)

        # 1. Process Master Sheet (First 8)
        batch1 = df.iloc[0:BATCH_SIZE]
        if not batch1.empty:
            create_master_file(batch1, 1)

        # 2. Process Continuation Sheets (Remaining in groups of 8)
        start_index = BATCH_SIZE
        batch_counter = 2
        
        while start_index < total_candidates:
            end_index = start_index + BATCH_SIZE
            batch_next = df.iloc[start_index:end_index]
            
            create_continuation_file(batch_next, batch_counter)
            
            start_index += BATCH_SIZE
            batch_counter += 1

        print("Done.")
    else:
        print(f"ERROR: {INPUT_CSV} not found.")