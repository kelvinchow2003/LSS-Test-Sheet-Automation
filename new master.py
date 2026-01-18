import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os

# --- CONFIGURATION ---
INPUT_PDF = "leadershipmastersheet_on_20250219_fillable.pdf"
INPUT_CSV = "roster.csv"
OUTPUT_FOLDER = "filled_forms/"

# --- HOST DATA ---
HOST_DATA = {
    "Host Name": "City of Markham",
    "Host Area": "905",
    "Host Phone": "4703590 EXT 4342",
    "Host Street": "8600 McCowan Road",
    "Host City": "Markham",
    "Host Province": "ON",
    "Host Postal": "L3P 3M2",
    "Host Facility": "Centennial C.C.",
    "Host Facility Area": "905",
    "Host Facility Phone": "4703590 EXT 4342",
    "Exam Fees Attached": "/Yes"
}

# --- HELPER: MAP ROW TO SLOT ---
def get_slot_data(row, field_id, visible_number):
    """
    field_id: The hardcoded PDF field name (e.g., '4', '5'...)
    visible_number: The actual candidate number to write (e.g., '10', '11'...)
    """
    p = str(field_id)      # The PDF field ID (4, 5, 6...)
    num_str = str(visible_number) # The text to type (10, 11, 12...)
    
    full_name = str(row.get("AttendeeName", ""))
    street = str(row.get("Street", ""))
    city = str(row.get("City", ""))
    zip_code = str(row.get("PostalCode", ""))
    full_address = f"{street}, {city} {zip_code}".strip(", ")

    # DOB Formatting: YY/MM/DD
    raw_dob = row.get("DateOfBirth", "")
    formatted_dob = ""
    if pd.notna(raw_dob):
        try:
            dt = pd.to_datetime(raw_dob, dayfirst=True)
            formatted_dob = dt.strftime("%y/%m/%d")
        except: pass

    data = {
        f"{p}.1": full_name,
        f"{p}.2": full_address,
        f"{p}.3": str(row.get("AttendeePhone", "")),
        f"{p}.4": str(row.get("E-mail", "")),
        f"{p}.5": formatted_dob
    }

    # --- LOGIC UPDATE: Write the GLOBAL candidate number ---
    # We write '10', '11', '12' into the box named '4.0', '5.0', etc.
    # We only write this for slots > 3 (The back page).
    if int(field_id) > 3:
        data[f"{p}.0"] = num_str

    return data

# --- SAVE FUNCTION ---
def _finalize_and_save(writer, reader, data_map, index, suffix):
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    # Force Fonts & Layers
    if "/AcroForm" not in writer.root_object:
        writer.root_object.update({NameObject("/AcroForm"): DictionaryObject()})
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    if "/OCProperties" in reader.root_object:
        writer.root_object[NameObject("/OCProperties")] = \
            reader.root_object["/OCProperties"].clone(writer)

    output_filename = f"{OUTPUT_FOLDER}Leadership_{index}_{suffix}.pdf"
    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"Generated: {output_filename}")

# --- MASTER FILE (Candidates 1-9) ---
def create_master_file(batch_df, file_index, total_count):
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader) 

    data_map = HOST_DATA.copy()
    data_map["Total Enrolled"] = str(total_count)

    for i, (idx, row) in enumerate(batch_df.iterrows()):
        current_num = i + 1  # 1, 2, 3... 9
        
        # For the Master sheet, Field ID and Candidate Number are the same
        data_map.update(get_slot_data(row, field_id=current_num, visible_number=current_num))

    _finalize_and_save(writer, reader, data_map, file_index, "Master")

# --- CONTINUATION FILE (Candidates 10+) ---
def create_continuation_file(batch_df, file_index, start_number, total_count):
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    
    # Copy PDF and remove Page 1
    writer.append(reader)
    del writer.pages[0]

    data_map = HOST_DATA.copy()
    data_map["Total Enrolled"] = str(total_count)

    # Loop through the batch (up to 6 people)
    for i, (idx, row) in enumerate(batch_df.iterrows()):
        # The PDF Field IDs are HARDCODED to 4, 5, 6, 7, 8, 9 on the back page
        pdf_field_id = i + 4 
        if pdf_field_id > 9: break 
        
        # The Visible Number continues counting (10, 11, 12...)
        actual_candidate_num = start_number + i
        
        data_map.update(get_slot_data(row, field_id=pdf_field_id, visible_number=actual_candidate_num))

    _finalize_and_save(writer, reader, data_map, file_index, "Continuation")

# --- MAIN EXECUTION ---
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

print("Reading roster.csv...")
if os.path.exists(INPUT_CSV):
    df = pd.read_csv(INPUT_CSV, dtype=str).fillna("")
    
    total_candidates = len(df)
    print(f"Found {total_candidates} candidates.")

    # 1. First 9 Candidates -> Master Sheet
    batch1 = df.iloc[0:9]
    if not batch1.empty:
        create_master_file(batch1, 1, total_candidates)

    # 2. Remaining Candidates -> Continuation Sheets (6 at a time)
    start_index = 9
    batch_counter = 2
    
    while start_index < total_candidates:
        end_index = start_index + 6
        batch_next = df.iloc[start_index:end_index]
        
        # Current Global Number is start_index + 1 (e.g., Index 9 is Candidate 10)
        current_start_number = start_index + 1
        
        create_continuation_file(batch_next, batch_counter, current_start_number, total_candidates)
        
        start_index += 6
        batch_counter += 1

    print("Done.")
else:
    print(f"ERROR: {INPUT_CSV} not found.")