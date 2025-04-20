import os
import re
import pdfplumber
import pandas as pd


# Flexible section extractor using keywords
def extract_sections_by_keywords(text, section_keywords):
    lines = text.splitlines()
    sections = {}
    current_section = None
    buffer = []

    keyword_lookup = {}
    for sec, keywords in section_keywords.items():
        for keyword in keywords:
            keyword_lookup[keyword.lower()] = sec

    def match_section_header(line):
        for keyword in keyword_lookup:
            if re.search(rf"\b{re.escape(keyword)}\b", line, re.IGNORECASE):
                return keyword_lookup[keyword]
        return None

    for line in lines:
        matched_section = match_section_header(line.strip())
        if matched_section:
            if current_section and buffer:
                if current_section not in sections:
                    sections[current_section] = "\n".join(buffer).strip()
            current_section = matched_section
            buffer = [line]
        elif current_section:
            buffer.append(line)

    if current_section and buffer:
        if current_section not in sections:
            sections[current_section] = "\n".join(buffer).strip()

    return sections


# Extract fields from one PDF
def extract_fields_from_pdf(section_keywords, pdf_path):
    full_text = ""
    all_tables = []
    cas_number = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += "\n" + text
            tables = page.extract_tables()
            all_tables.extend(tables)

    # Extract sections using keywords
    sections = extract_sections_by_keywords(full_text, section_keywords)
    sec1 = sections.get("section_1", "")
    sec3 = sections.get("section_3", "")

    # --- Extract values from Section 1 ---
    product_name = re.search(r"(?i)Product name\s*[:\-]?\s*(.+)", sec1)
    product_code = re.search(r"(?i)(Product code|Product number)\s*[:\-]?\s*(.+)", sec1)
    manufacturer = re.search(r"(?i)(Company name of supplier|Manufacturer)\s*[:\-]?\s*(.+)", sec1)
    usage = re.search(r"(?i)Recommended use\s*[:\-]?\s*(.+)", sec1)
    revision_date = re.search(r"(?i)(Revision Date|Date of revision)\s*[:\-]?\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", sec1)

    # --- Extract CAS Number from Section 3 text ---
    match_cas = re.search(r"\b\d{2,7}-\d{2}-\d\b", sec3)
    if match_cas:
        cas_number = match_cas.group()
    else:
        # Try finding CAS from tables
        for table in all_tables:
            for row in table:
                for cell in row:
                    if cell:
                        cas_match = re.search(r"\b\d{2,7}-\d{2}-\d\b", str(cell))
                        if cas_match:
                            cas_number = cas_match.group()
                            break
                if cas_number:
                    break
            if cas_number:
                break

    return {
        "Product Name": product_name.group(1).strip() if product_name else None,
        "Product Number": product_code.group(2).strip() if product_code else None,
        "Manufacturer": manufacturer.group(2).strip() if manufacturer else None,
        "Usage": usage.group(1).strip() if usage else None,
        "Revision Date": revision_date.group(2).strip() if revision_date else None,
        "CAS Number": cas_number
    }


def main():
    pdf_folder = "Input_PDF"  # Replace with your actual folder name
    section_keywords = {
        "section_1": [
            "Identification", "Product Identification", "Section 1"
        ],
        "section_3": [
            "Composition", "Information on Ingredients", "Ingredients",
            "Substance", "Section 3", "Hazardous Ingredients"
        ]
    }

    data = []
    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            result = extract_fields_from_pdf(section_keywords, pdf_path)
            result["File Name"] = filename
            data.append(result)

    df = pd.DataFrame(data)
    df.to_excel("Extracted_Information.xlsx", index=False)
    print("✅ Done. Output written to 'Extracted_Information.xlsx'")


if __name__ == "__main__":
    main()
