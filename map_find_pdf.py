from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, DictionaryObject
import os

INPUT_PDF = "basicfirstaidwcprcandaed_20260619_fillable.pdf"
OUTPUT_PDF = "DEBUG_BASIC_FIRST_AID.pdf"

if not os.path.exists(INPUT_PDF):
    print(f"Error: {INPUT_PDF} not found.")
else:
    reader = PdfReader(INPUT_PDF)
    writer = PdfWriter()
    writer.append(reader)

    # Write the hidden field names into the text boxes
    debug_data = {field_name: field_name for field_name in reader.get_fields().keys()}
    
    for page in writer.pages:
        writer.update_page_form_field_values(page, debug_data)

    writer.root_object.update({NameObject("/AcroForm"): DictionaryObject()})
    writer.root_object["/AcroForm"][NameObject("/NeedAppearances")] = BooleanObject(True)

    with open(OUTPUT_PDF, "wb") as f:
        writer.write(f)
    print(f"Created {OUTPUT_PDF}. Open it and check Candidate 1's field names.")