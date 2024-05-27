import requests
import PyPDF2
import pandas as pd
from io import BytesIO

# URL of the PDF file
pdf_url = "https://www.nrb.org.np/contents/uploads/2021/12/Key-Financial-Indicators-2078-Ashoj.pdf"

# Download the PDF file
response = requests.get(pdf_url)
pdf_data = BytesIO(response.content)

# Initialize PDF reader
pdf_reader = PyPDF2.PdfFileReader(pdf_data)

# Extract text from each page
text_data = []
for page_num in range(pdf_reader.numPages):
    page = pdf_reader.getPage(page_num)
    text_data.append(page.extract_text())


data = []
for page_text in text_data:
    lines = page_text.split('\n')
    for line in lines:
   
        columns = line.split()
        data.append(columns)

# Convert to DataFrame
df = pd.DataFrame(data)

# Save DataFrame to Excel
excel_file_path = "key_financial_indicators.xlsx"
df.to_excel(excel_file_path, index=False)

print(f"Data has been saved to {excel_file_path}")
