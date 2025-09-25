import json
from docx import Document

SECTION_NAMES = ["SUMMARY", "SKILLS", "EXPERIENCE", "EDUCATION", "PROJECTS", "CERTIFICATIONS"]

def parse_resume(path):
    doc = Document(path)
    data = {s.lower(): [] for s in SECTION_NAMES}
    data["contact"] = []

    current = "contact"

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        # Check if this line is a section title
        if text.upper() in SECTION_NAMES:
            current = text.lower()
            continue
        data[current].append(text)

    # For summary, join into a single string
    data["summary"] = " ".join(data["summary"])

    # Skills → split by commas/semicolons
    if data["skills"]:
        skills_text = " ".join(data["skills"])
        data["skills"] = [s.strip() for s in skills_text.replace(";", ",").split(",") if s.strip()]

    return data

if __name__ == "__main__":
    infile = "Your_Resume.docx"   # change this to your resume file
    outfile = "resume.json"

    parsed = parse_resume(infile)

    with open(outfile, "w", encoding="utf-8") as f:
        json.dump(parsed, f, indent=2, ensure_ascii=False)

    print(f"✅ Resume parsed → {outfile}")
