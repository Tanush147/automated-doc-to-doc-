import pandas as pd
from docx import Document
import os

# Load data from the CSV file
df = pd.read_csv("Untitled Spreadsheet_Sheet1.csv")

# Load the template document
template_path = "INTERNSHIP OFFER LETTER 3.docx"

# Define output folder
output_folder = "offer_letters"
os.makedirs(output_folder, exist_ok=True)  # Create the folder if it doesn't exist

for index, row in df.iterrows():
    title = str(row['Title'])
    name = str(row['Name'])

    # Open the template for each student
    doc = Document(template_path)

    # Replace placeholders in paragraphs
    for para in doc.paragraphs:
        if '<<Title>>' in para.text or '<<Name>>' in para.text:
            para.text = para.text.replace('<<Title>>', title).replace('<<Name>>', name)

    # Replace placeholders in tables (if any)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if '<<Title>>' in cell.text or '<<Name>>' in cell.text:
                    cell.text = cell.text.replace('<<Title>>', title).replace('<<Name>>', name)

    # Save the personalized document in the offer_letters folder
    filename = f"Offer_Letter_{name.replace(' ', '_')}.docx"
    output_path = os.path.join(output_folder, filename)
    doc.save(output_path)

    print(f"Generated: {output_path}")

print("✅ All offer letters generated and saved in 'offer_letters' folder!")
