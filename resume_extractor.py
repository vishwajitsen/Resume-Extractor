import os
import re
from pathlib import Path
import fitz  # PyMuPDF
from openpyxl import Workbook

# ========== CONFIG ==========
PDF_PATH = r"C:\Users\win11\OneDrive\Documents\MiSpy Documents\Resume_Extractor\Resume-Extractor\Vishwajit Sen CV.pdf"
OUTPUT_XLSX = r"C:\Users\win11\OneDrive\Documents\MiSpy Documents\Resume_Extractor\Resume-Extractor\resume_extracted.xlsx"

# Social domains we will keep
SOCIAL_DOMAINS = [
    "linkedin.com", "github.com", "gitlab.com", "behance.net", "dribbble.com",
    "pinterest.com", "medium.com", "x.com", "twitter.com", "facebook.com",
    "instagram.com", "hashnode.com", "dev.to", "substack.com", "blogspot.com",
    "wordpress.com", "notion.site", "notion.so", "about.me", "me.linkedin.com"
]

# Words we never consider as a person-name token
NAME_TOKEN_BLOCKLIST = {
    "cv", "resume", "profile", "curriculum", "vitae", "summary",
    "data", "science", "scientist", "ai", "ml", "dl", "analytics",
    "engineer", "lead", "manager", "professional", "transforming",
    "insights", "consultant", "portfolio", "contact"
}

# ========== PDF TEXT HELPERS ==========
def extract_text_lines_from_pdf(pdf_path):
    """Extract text preserving line breaks from the PDF (page 1 first)."""
    text_lines = []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            page_text = page.get_text("text")
            lines = [ln.strip() for ln in page_text.splitlines() if ln.strip()]
            text_lines.extend(lines)
    return text_lines

def extract_full_text(pdf_path):
    with fitz.open(pdf_path) as doc:
        return "\n".join(page.get_text("text") for page in doc)

# ========== FIELD EXTRACTORS ==========
def extract_email(text):
    m = re.search(r'(?i)\b[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}\b', text)
    return m.group(0) if m else ""

def extract_phone(text):
    """
    Extract the most plausible phone number.
    - Accepts +91 formats, spaces/dashes/() allowed
    - Returns the candidate with the most digits (10–13)
    """
    candidates = re.findall(r'(\+?\d[\d\s\-\(\)]{8,}\d)', text)
    best = ""
    best_digits = 0
    for c in candidates:
        digits = re.sub(r'\D', '', c)
        if 10 <= len(digits) <= 13:
            # Prefer Indian-style numbers starting 6-9 or with country code 91
            score = len(digits)
            if digits.startswith("91") or digits[0] in "6789":
                score += 2
            if score > best_digits:
                best_digits = score
                best = c.strip()
    return best

def extract_social_links(text):
    """
    Extract URLs and keep only known social/work domains.
    Accepts with/without scheme.
    """
    urls = set()
    # with scheme
    for u in re.findall(r'(?i)\bhttps?://[^\s\)\]]+', text):
        urls.add(u.strip().rstrip('.,);]'))
    # without scheme (e.g., linkedin.com/in/...)
    for u in re.findall(r'(?i)\b(?:www\.)?[a-z0-9.-]+\.[a-z]{2,}[^\s,;)]+', text):
        if u.startswith("http"):
            continue
        urls.add(u.strip().rstrip('.,);]'))

    def is_social(u):
        lower = u.lower()
        return any(dom in lower for dom in SOCIAL_DOMAINS)

    # Normalize: ensure scheme
    normalized = []
    for u in urls:
        if is_social(u):
            if not u.lower().startswith(("http://", "https://")):
                u = "https://" + u.lstrip("/")
            normalized.append(u)
    # Stable order
    normalized.sort()
    return normalized

def tokens_look_like_name(tokens):
    """
    Check if tokens look like a human full name:
    - 2 to 4 tokens
    - all alphabetic
    - start with uppercase letter (allow all-caps last name)
    - none of the tokens are in blocklist
    """
    if not (2 <= len(tokens) <= 4):
        return False
    for t in tokens:
        t_clean = re.sub(r"[^A-Za-z]", "", t)
        if not t_clean:
            return False
        if t_clean.lower() in NAME_TOKEN_BLOCKLIST:
            return False
        # Allow all-caps (e.g., SEN) or Capitalized
        if not (t_clean.isupper() or (t_clean[0].isupper() and t_clean[1:].islower())):
            return False
    return True

def name_from_filename(pdf_path):
    """
    Try to parse name from the file name (e.g., 'Vishwajit Sen CV.pdf' -> Vishwajit Sen).
    """
    stem = Path(pdf_path).stem  # 'Vishwajit Sen CV'
    # Split on separators
    raw = re.split(r'[\s_\-]+', stem)
    # Drop non-name tokens and numerics
    cand = [w for w in raw if w and w.isalpha() and w.lower() not in NAME_TOKEN_BLOCKLIST]
    # Keep first 2-4 tokens if they look like a name
    for L in range(4, 1, -1):
        if len(cand) >= L and tokens_look_like_name(cand[:L]):
            return cand[:L]
    if len(cand) >= 2 and tokens_look_like_name(cand[:2]):
        return cand[:2]
    return []

def name_from_top_lines(lines, hints=None):
    """
    Scan the first N lines looking for a likely full name.
    Prefer lines containing hint substrings (from email/file name).
    """
    hints = hints or []
    FIRST_N = 15
    candidates = []

    for ln in lines[:FIRST_N]:
        # skip obvious non-name lines
        if any(x in ln.lower() for x in ["@", "http", "www.", "phone", "mobile", "contact", "linkedin", "github"]):
            continue
        # remove punctuation except spaces and hyphens
        cleaned = re.sub(r"[^\w\s\-]", " ", ln).strip()
        tokens = [t for t in re.split(r"\s+", cleaned) if t]
        if tokens_look_like_name(tokens):
            score = 0
            low = ln.lower()
            # hint boost (email user, filename pieces)
            for h in hints:
                if h and h in low:
                    score += 2
            # earlier lines get a small bonus
            score += max(0, FIRST_N - lines[:FIRST_N].index(ln))
            candidates.append((score, tokens))

    if not candidates:
        return []
    # highest score
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

def split_first_middle_last(tokens):
    if not tokens:
        return "", "", ""
    if len(tokens) == 1:
        return tokens[0], "", ""
    if len(tokens) == 2:
        return tokens[0], "", tokens[1]
    # 3 or 4 tokens: middle = tokens[1:-1]
    first = tokens[0]
    last = tokens[-1]
    middle = " ".join(tokens[1:-1])
    return first, middle, last

def extract_name(pdf_path, all_text, lines):
    # derive hints from email and filename to avoid picking taglines
    email = extract_email(all_text)
    email_user = email.split("@")[0] if email else ""
    hint_pieces = set()
    if email_user:
        # split on non-letters and keep plausible name bits
        for piece in re.split(r'[^a-zA-Z]+', email_user):
            if len(piece) >= 3:
                hint_pieces.add(piece.lower())
    # try filename
    file_name_tokens = name_from_filename(pdf_path)
    if file_name_tokens:
        return split_first_middle_last(file_name_tokens)
    # try top lines with hints
    top_tokens = name_from_top_lines(lines, hints=hint_pieces)
    if top_tokens:
        return split_first_middle_last(top_tokens)
    # fallback: empty
    return "", "", ""

# ========== SAVE EXCEL ==========
def save_to_excel(fields_dict, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resume"
    ws.append(["Field", "Value"])
    for key in ["First Name", "Middle Name", "Last Name", "Email", "Mobile Phone", "Social Links"]:
        val = fields_dict.get(key, "")
        if isinstance(val, list):
            val = ", ".join(val)
        ws.append([key, val])
    # basic column width
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 90
    wb.save(output_path)

# ========== MAIN ==========
def main():
    if not os.path.exists(PDF_PATH):
        print(f"❌ File not found: {PDF_PATH}")
        return

    lines = extract_text_lines_from_pdf(PDF_PATH)
    full_text = "\n".join(lines)

    email = extract_email(full_text)
    phone = extract_phone(full_text)
    socials = extract_social_links(full_text)
    first, middle, last = extract_name(PDF_PATH, full_text, lines)

    result = {
        "First Name": first,
        "Middle Name": middle,
        "Last Name": last,
        "Email": email,
        "Mobile Phone": phone,
        "Social Links": socials,
    }

    # Print neat bullet points
    print("\n--- Extracted Resume Information ---\n")
    print(f"• First Name: {result['First Name']}")
    print(f"• Middle Name: {result['Middle Name']}")
    print(f"• Last Name: {result['Last Name']}")
    print(f"• Email: {result['Email']}")
    print(f"• Mobile Phone: {result['Mobile Phone']}")
    if result["Social Links"]:
        print("• Social Links:")
        for u in result["Social Links"]:
            print(f"   - {u}")

    # Export to Excel
    save_to_excel(result, OUTPUT_XLSX)
    print(f"\n✅ Saved to: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
