import pandas as pd
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os

# --- CONFIGURATION ---
INPUT_PDF = "95tsbronzecrossr20250605 fillable.pdf"
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

# --- THE "COPY-PASTE" FIELD MAP ---
FIELD_MAP = [
    {"name": "Name 1.0.0", "addr": "Address1.0.0", "city": "City1.0.0", "postal": "Postal1.0.0", "email": "Email1.0.0", "phone": "Phone1.0.0"},
    {"name": "Name 1.0.1.0", "addr": "Address1.0.1.0", "city": "City1.0.1.0", "postal": "Postal1.0.1.0", "email": "Email1.0.1.0", "phone": "Phone1.0.1.0"},
    {"name": "Name1.0.1.1.0", "addr": "Address1.0.1.1.0", "city": "City1.0.1.1.0", "postal": "Postal1.0.1.1.0", "email": "Email1.0.1.1.0", "phone": "Phone1.0.1.1.0"},
    {"name": "Name1.0.1.1.1.0", "addr": "Address1.0.1.1.1.0", "city": "City1.0.1.1.1.0", "postal": "Postal1.0.1.1.1.0", "email": "Email1.0.1.1.1.0", "phone": "Phone1.0.1.1.1.0"},
    {"name": "Name 1.0.1.1.1.1.0", "addr": "Address1.0.1.1.1.1.0", "city": "City1.0.1.1.1.1.0", "postal": "Postal1.0.1.1.1.1.0", "email": "Email1.0.1.1.1.1.0", "phone": "Phone1.0.1.1.1.1.0"},
    {"name": "Name1.0.1.1.1.1.1", "addr": "Address1.0.1.1.1.1.1", "city": "City1.0.1.1.1.1.1", "postal": "Postal1.0.1.1.1.1.1", "email": "Email1.0.1.1.1.1.1", "phone": "Phone1.0.1.1.1.1.1"},
    {"name": "7Name1.0.0", "addr": "7Address1.0.0", "city": "7City1.0.0", "postal": "7Postal1.0.0", "email": "7Email1.0.0", "phone": "7Phone1.0.0"},
    {"name": "8Name1.0.1.0", "addr": "8Address1.0.1.0", "city": "8City1.0.1.0", "postal": "8Postal1.0.1.0", "email": "8Email1.0.1.0", "phone": "8Phone1.0.1.0"},
    {"name": "9Name1.0.1.1.0", "addr": "9Address1.0.1.1.0", "city": "9City1.0.1.1.0", "postal": "9Postal1.0.1.1.0", "email": "9Email1.0.1.1.0", "phone": "9Phone1.0.1.1.0"},
    {"name": "10Name1.0.1.1.1.0", "addr": "10Address1.0.1.1.1.0", "city": "10City1.0.1.1.1.0", "postal": "10Postal1.0.1.1.1.0", "email": "10Email1.0.1.1.1.0", "phone": "10Phone1.0.1.1.1.0"},
    {"name": "11Name1.0.1.1.1.1.0", "addr": "11Address1.0.1.1.1.1.0", "city": "11City1.0.1.1.1.1.0", "postal": "11Postal1.0.1.1.1.1.0", "email": "11Email1.0.1.1.1.1.0", "phone": "11Phone1.0.1.1.1.1.0"},
    {"name": "12Name1.0.1.1.1.1.1.0", "addr": "12Address1.0.1.1.1.1.1.0", "city": "12City1.0.1.1.1.1.1.0", "postal": "12Postalt.0.1.1.1110", "email": "12Emailt 0.1.1.1.1.1.0", "phone": "12Phone1.0.1.1.1.1.1.0"},
    {"name": "13Name1.0.1.1.1.1.1.1", "addr": "13Address1.0.1.1.1.1.1.1", "city": "13City1.0.1.1.1.1.1.1", "postal": "13Pasalt 0.1.1.1.1.3.3", "email": "13 Email 1.0.1.1.1.1.1.1", "phone": "13Phone1.0.1.1.1.1.1.1"}
]

# --- HELPER: MAP ROW TO SLOT ---
def get_slot_data(row, slot_index):
    full_name = str(row.get("AttendeeName", "")).strip()

    street = str(row.get("Street", ""))
    city = str(row.get("City", ""))
    prov = str(row.get("Province", "ON"))
    postal = str(row.get("PostalCode", ""))
    email = str(row.get("E-mail", ""))
    phone = str(row.get("AttendeePhone", ""))
    city_prov = f"{city}, {prov}".strip(", ")

    # DOB Formatting
    raw_dob = row.get("DateOfBirth", "")
    dd, mm, yy = "", "", ""
    if pd.notna(raw_dob):
        try:
            dt = pd.to_datetime(raw_dob, dayfirst=True)
            dd = str(dt.day).zfill(2)
            mm = str(dt.month).zfill(2)
            yy = str(dt.year)[-2:] 
        except: pass

    fields = FIELD_MAP[slot_index]

    data = {
        fields["name"]: full_name,
        fields["addr"]: street,
        fields["city"]: city_prov,
        fields["postal"]: postal,
        fields["email"]: email,
        fields["phone"]: phone,
        
        # Adding the broken DOB fields to try and force them in
        "DO": f"{yy}/{mm}/{dd}",
        "Year": yy
    }

    return data

# --- SAVE FUNCTION ---
def _finalize_and_save(writer, reader, data_map, index, suffix):
    for page in writer.pages:
        writer.update_page_form_field_values(page, data_map)

    if "/AcroForm" not in writer.root_object:
        writer.root_object.update({NameObject("/AcroForm"): DictionaryObject()})
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    if "/OCProperties" in reader.root_object:
        writer.root_object[NameObject("/OCProperties")] = \
            reader.root_object["/OCProperties"].clone(writer)

    output_filename = f"{OUTPUT_FOLDER}BronzeCross_Recert_{index}_{suffix}.pdf"
    with open(output_filename, "wb") as f:
        writer.write(f)
    print(f"Generated: {output_filename}")

# --- MASTER FILE GENERATOR (Candidates 1-13) ---
def create_master_file(batch_df, file_index):
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader) 

    data_map = HOST_DATA.copy()
    
    for i, (idx, row) in enumerate(batch_df.iterrows()):
        data_map.update(get_slot_data(row, slot_index=i))

    _finalize_and_save(writer, reader, data_map, file_index, "Master")

# --- CONTINUATION FILE GENERATOR (Candidates 14+) ---
def create_continuation_file(batch_df, file_index):
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader) 

    data_map = HOST_DATA.copy()

    for i, (idx, row) in enumerate(batch_df.iterrows()):
        data_map.update(get_slot_data(row, slot_index=i))

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

        BATCH_SIZE = 13 

        batch1 = df.iloc[0:BATCH_SIZE]
        if not batch1.empty:
            create_master_file(batch1, 1)

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