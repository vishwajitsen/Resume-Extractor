import os
import re
import fitz  # PyMuPDF
import spacy
import pandas as pd

# -------------------------
# CONFIGURATION
# -------------------------
RESUME_PATH = r"C:\Users\win11\OneDrive\Documents\MiSpy Documents\Resume_Extractor\Resume-Extractor\Vishwajit Sen CV.pdf"
OUTPUT_EXCEL = r"C:\Users\win11\OneDrive\Documents\MiSpy Documents\Resume_Extractor\Resume-Extractor\resume_extracted.xlsx"

# -------------------------
# PDF TEXT EXTRACTION
# -------------------------
def extract_text_from_pdf(pdf_path):
    """Extracts all text from a PDF using PyMuPDF."""
    text = ""
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            text += page.get_text("text") + "\n"
    return text

def clean_text(text):
    """Cleans up spacing in extracted text."""
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# -------------------------
# RESUME FIELD EXTRACTION
# -------------------------
def extract_resume_info(text):
    """Extracts key information from resume text using spaCy NER + regex."""
    info = {}
    nlp = spacy.load("en_core_web_sm")
    doc = nlp(text)

    # Regex extractions
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
    if email_match:
        info["Email"] = email_match.group()

    phone_match = re.search(r'(\+?\d[\d\s\-]{8,}\d)', text)
    if phone_match:
        info["Mobile Phone"] = phone_match.group().strip()

    links = re.findall(r'(https?://[^\s]+)', text)
    if links:
        info["Social Links"] = list(set(links))  # unique

    # NER-based extractions
    names = [ent.text for ent in doc.ents if ent.label_ == "PERSON"]
    if names:
        name_parts = names[0].split()
        if len(name_parts) >= 2:
            info["First Name"] = name_parts[0]
            info["Last Name"] = name_parts[-1]

    orgs = list(set(ent.text for ent in doc.ents if ent.label_ == "ORG"))
    if orgs:
        info["Employer / Organizations"] = orgs

    # Education detection using keywords + NER
    education_keywords = ["B.Tech", "M.Tech", "Bachelor", "Master", "PhD", 
                          "Diploma", "Degree", "University", "College"]
    education_lines = [
        sent.text.strip()
        for sent in doc.sents
        if any(kw.lower() in sent.text.lower() for kw in education_keywords)
    ]
    if education_lines:
        info["Education"] = list(set(education_lines))

    return info

# -------------------------
# SAVE TO EXCEL
# -------------------------
def save_to_excel(info, output_path):
    """Saves extracted info to an Excel file."""
    # Convert lists to comma-separated strings for Excel
    processed_info = {}
    for key, value in info.items():
        if isinstance(value, list):
            processed_info[key] = ", ".join(value)
        else:
            processed_info[key] = value

    df = pd.DataFrame(processed_info.items(), columns=["Field", "Value"])
    df.to_excel(output_path, index=False)
    print(f"\n✅ Data exported to Excel: {output_path}")

# -------------------------
# MAIN FUNCTION
# -------------------------
def main():
    if not os.path.exists(RESUME_PATH):
        print("❌ Resume file not found at:", RESUME_PATH)
        return

    # Extract and clean text
    raw_text = extract_text_from_pdf(RESUME_PATH)
    raw_text = clean_text(raw_text)

    # Extract structured info
    info = extract_resume_info(raw_text)

    # Display results in bullet points
    print("\n--- Extracted Resume Information ---\n")
    for key, value in info.items():
        if isinstance(value, list):
            print(f"• {key}:")
            for v in value:
                print(f"   - {v}")
        else:
            print(f"• {key}: {value}")

    # Save to Excel
    save_to_excel(info, OUTPUT_EXCEL)

if __name__ == "__main__":
    main()
