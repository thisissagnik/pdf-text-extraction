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

    # Map all known keywords to section IDs
    keyword_lookup = {}
    for sec, keywords in section_keywords.items():
        for keyword in keywords:
            keyword_lookup[keyword.lower()] = sec

    def match_section_header(line):
        # for keyword in keyword_lookup:
        #     if re.search(rf"\b{re.escape(keyword)}\b", line, re.IGNORECASE):
        #         return keyword_lookup[keyword]
        # if (
        #     re.match(r"(Section\s+)?4(\.|:)?\s", line.strip(), re.IGNORECASE)
        #     or "First Aid Measures" in line
        # ):
        #     return "section_4"  # Next section marker
        # return None

        normalized = line.strip().lower()
        section_heading_match = re.match(r"^(section\s*)?(\d{1,2})[\.\: -]+(.+)?", normalized)
        if section_heading_match:
            for keyword in keyword_lookup:
                if keyword in normalized:
                    return keyword_lookup[keyword]
        if (
            re.match(r"(Section\s+)?4(\.|:)?\s", line.strip(), re.IGNORECASE)
            or "First Aid Measures" in line
        ):
            return "section_4"  # Next section marker


    for line in lines:
        matched_section = match_section_header(line.strip())
        if matched_section:
            if current_section and buffer:
                if current_section not in sections:
                    sections[current_section] = "\n".join(buffer).strip()
            current_section = keyword_lookup.get(
                matched_section.lower(), matched_section
            )
            buffer = [line]
        elif current_section:
            buffer.append(line)

    if current_section and buffer:
        if current_section not in sections:
            sections[current_section] = "\n".join(buffer).strip()

    return sections


# 🔍 New helper function for multiline field extraction
def extract_field_with_multiline_support(text, keyword_variants):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        clean_line = line.strip()
        for keyword in keyword_variants:
            match = re.search(rf"{re.escape(keyword)}\s*[:\-]\s*(.+)", clean_line, re.IGNORECASE)
            if match:
                return match.group(1).strip()
            if re.fullmatch(rf"{re.escape(keyword)}\s*[:\-]?", clean_line, re.IGNORECASE):
                for j in range(i + 1, len(lines)):
                    next_line = lines[j].strip()
                    if next_line:
                        return next_line
    return None


# Extract fields from one PDF
def extract_fields_from_pdf(section_keywords, pdf_path):
    full_text = ""
    all_tables = []
    cas_number_combined = None

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if text:
                full_text += "\n" + text
            tables = page.extract_tables()
            all_tables.append((i, tables))

    sections = extract_sections_by_keywords(full_text, section_keywords)
    sec1 = sections.get("section_1", "")
    sec3 = sections.get("section_3", "")
    sec16 = sections.get("section_16", "")

    product_name = extract_field_with_multiline_support(sec1, ["Product name", "Chemical name","Trade name","Identification of the substance","Identification"])
    product_code = extract_field_with_multiline_support(sec1, ["Product code", "Product number","Article number","cataloge number"])
    manufacturer = extract_field_with_multiline_support(sec1, ["Company name of supplier", "Manufacturer", "Company name","Company Identification","Supplier","Manufacturer/Supplier","Company"])
    usage = extract_field_with_multiline_support(sec1, ["Recommended use", "Intended use", "Use","Identified uses","Aplication of the substance"])
    revision_date = extract_field_with_multiline_support(sec1, ["Revision Date", "Date of revision","Revision"])

    # if revision_date is not found in section 1, check section 16
    if not revision_date:
        revision_date = extract_field_with_multiline_support(sec16, ["Revision Date", "Date of revision"])

    # if revision_date is not found in section 1 or section 16, check in header or footer of the PDF
    if not revision_date:
        revision_date = extract_field_with_multiline_support(full_text[:300], ["Revision Date", "Date of revision"]) # only check first 300 characters

    cas_entries = set()
    found_table = False

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_text = page.extract_text()
            if page_text and sec3 in page_text:
                tables = page.extract_tables()
                if tables:
                    found_table = True
                    for table in tables:
                        for row in table:
                            row_text = " | ".join([str(cell) for cell in row if cell])
                            cas_matches = re.findall(r"\b\d{2,7}-\d{2}-\d\b", row_text)
                            row_text = re.sub(r"\b\d{2,7}-\d{2}-\d\b", "", row_text)
                            conc_matches = re.findall(
                                r"(?:[<>]=?|~)?\s*\d+(?:\.\d+)?(?:\s*-\s*(?:[<>]=?|~)?\s*\d+(?:\.\d+)?\s*)?%?",
                                row_text
                            )
                            for idx, cas_number in enumerate(cas_matches):
                                concentration = conc_matches[idx] if idx < len(conc_matches) else None
                                if concentration:
                                    cas_entries.add(f"{cas_number} ({concentration})")
                                else:
                                    cas_entries.add(cas_number)

    if not found_table:
        cas_matches = re.findall(r"\b\d{2,7}-\d{2}-\d\b", sec3)
        conc_matches = re.findall(r"\d+(\.\d+)?\s*%|\d+\s*-\s*\d+\s*%", sec3)
        for i, cas in enumerate(cas_matches):
            conc = conc_matches[i].strip() if i < len(conc_matches) else None
            if conc:
                cas_entries.add(f"{cas} ({conc})")
            else:
                cas_entries.add(cas)

    cas_number_combined = ", ".join(sorted(cas_entries)) if cas_entries else None

    extracted_fields = {
        "Product Name": product_name,
        "Product Number": product_code,
        "Manufacturer": manufacturer,
        "Usage": usage,
        "Revision Date": revision_date,
        "CAS Numbers": cas_number_combined,
    }
    print(f"Extracted fields: {extracted_fields}")
    return extracted_fields


def main():
    pdf_folder = "Input_PDF"
    section_keywords = {
        "section_1": ["Identification", "Product Identification", "Section 1"],
        "section_3": [
            "Composition",
            "Information on Ingredients",
            "Ingredients",
            "Section 3",
            "Hazardous Ingredients",
        ],
        "section_16": ["Other information", "Section 16", "Additional Information"],
    }

    data = []
    test_file_list = [
        "A0178_US_EN.pdf",
        "A2237_US_EN.pdf",
        "CEHS_0003.pdf",
    ]

    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith(".pdf"):
            # commented out the test file list to process all files in the folder
            # if filename not in test_file_list:
            #     continue
            pdf_path = os.path.join(pdf_folder, filename)
            print(f"Processing for file: {filename}")
            result = extract_fields_from_pdf(section_keywords, pdf_path)
            result["File Name"] = filename
            data.append(result)

    df = pd.DataFrame(data)
    df.to_csv("Extracted_Information.csv", index=False)
    print("✅ Done. Output written to 'Extracted_Information.csv'")


if __name__ == "__main__":
    main()
