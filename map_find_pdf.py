from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os

INPUT_PDF = "95nlpool 2022_tsfillable 20250819 x.pdf"
OUTPUT_PDF = "DEBUG_MAP.pdf"

if not os.path.exists(INPUT_PDF):
    print(f"Error: {INPUT_PDF} not found")
else:
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader)

    # Dictionary to hold the debug data
    debug_data = {}

    # Loop through all fields in the PDF
    if reader.get_fields():
        for field_name, field_data in reader.get_fields().items():
            # We only want to write into text fields, not checkboxes (which look like /Btn)
            # But writing text to a checkbox usually just fails silently, which is fine.
            debug_data[field_name] = field_name

    # Apply the names to the pages
    for page in writer.pages:
        writer.update_page_form_field_values(page, debug_data)

    # Force the text to appear (NeedAppearances)
    if "/AcroForm" not in writer.root_object:
        writer.root_object.update({NameObject("/AcroForm"): DictionaryObject()})
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    with open(OUTPUT_PDF, "wb") as f:
        writer.write(f)

    print(f"Created {OUTPUT_PDF}. Open it to see the field IDs written in the boxes.")