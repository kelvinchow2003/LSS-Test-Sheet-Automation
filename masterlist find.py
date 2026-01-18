from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os

# --- CONFIGURATION ---
INPUT_PDF = "leadershipmastersheet_on_20250219_fillable.pdf"
OUTPUT_PDF = "DEBUG_LEADERSHIP_MAP.pdf"

if not os.path.exists(INPUT_PDF):
    print(f"Error: {INPUT_PDF} not found")
else:
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader)

    # Dictionary to hold the debug data
    debug_data = {}

    # Loop through all fields
    if reader.get_fields():
        for field_name, field_data in reader.get_fields().items():
            # Write the Field Name inside the Field Box
            debug_data[field_name] = field_name

    # Apply to pages
    for page in writer.pages:
        writer.update_page_form_field_values(page, debug_data)

    # Force appearance (NeedAppearances)
    if "/AcroForm" not in writer.root_object:
        writer.root_object.update({NameObject("/AcroForm"): DictionaryObject()})
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    with open(OUTPUT_PDF, "wb") as f:
        writer.write(f)

    print(f"Created {OUTPUT_PDF}. Open it and check the boxes for Candidate 1.")