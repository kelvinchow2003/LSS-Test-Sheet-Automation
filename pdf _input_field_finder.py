from pypdf import PdfReader

# REPLAACE WITH YOUR FILE NAME
pdf_path = "leadershipmastersheet_on_20250219_fillable.pdf" 
reader = PdfReader(pdf_path)
fields = reader.get_fields()

print("--- PDF FIELD NAMES ---")
for field_name, value in fields.items():
    print(f"Field Name: {field_name}")