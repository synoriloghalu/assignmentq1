import fitz  # PyMuPDF
import pandas as pd


def extract_text_from_pdf(pdf_path, page_start, page_end):
    # Open the PDF
    doc = fitz.open(pdf_path)
    text_data = []

    # Adjust for zero-based indexing
    page_start -= 1
    page_end -= 1

    # Ensure page_start and page_end are within the document bounds
    page_start = max(0, page_start)
    page_end = min(len(doc) - 1, page_end)

    for page_num in range(page_start, page_end + 1):
        page = doc.load_page(page_num)
        text = page.get_text("dict")
        blocks = text["blocks"]

        # Group blocks by their x-coordinate (column)
        column_data = {}
        for b in blocks:
            if b["type"] == 0:  # block contains text
                x0, y0, x1, y1 = b["bbox"]
                column_key = round(x0)  # Group by rounded x-coordinate
                block_text = " ".join([line["spans"][0]["text"] for line in b["lines"]])
                column_data.setdefault(column_key, []).append((y0, block_text))

        # Sort each column by y coordinate and then combine
        for column in sorted(column_data.keys()):
            column_data[column].sort(key=lambda x: x[0])  # Sort by y-coordinate
            text_data.extend([text for _, text in column_data[column]])

    return text_data


def save_to_excel(text_data, output_file):
    # Convert to DataFrame
    df = pd.DataFrame(text_data, columns=["Text"])

    # Save to Excel
    df.to_excel(output_file, index=False)


# Path to PDF
pdf_path = "keppel-corporation-limited-annual-report-2018.pdf"
page_start = 12
page_end = 12
# Output Excel file
output_file = "output.xlsx"

# Extract text and save to Excel
text_data = extract_text_from_pdf(pdf_path, page_start, page_end)
save_to_excel(text_data, output_file)
