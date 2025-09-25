# resume_parser_fixed_dropin.py
import re, json
from docx import Document

HEADINGS = {
    "FULL NAME": "full_name",
    "PERSONAL PROFILE": "profile",
    "CONTACT DETAILS": "contact",
    "SKILLS AND ABILITIES": "skills",
    "SOFTWARE/TOOLS": "tools",
    "INTERNSHIP EXPERIENCE": "experience",
    "ACADEMIC PROFILE": "education",
    "PROJECTS": "projects",
    "CERTIFICATION/COURSES": "certifications",
    "ACHIEVEMENTS": "achievements",
    "POSTION OF RESPONSIBILITY": "responsibility",
    "CO-CURRICULAR ACTIVITIES": "activities",
}

# helpers
def norm(s:str)->str: return (s or "").strip()
def is_heading(s:str)->str|None:
    s = norm(s).upper()
    return s if s in HEADINGS else None

BULLETS = "•-–—▪·◦* \t"
def clean_bullets(lines):
    out=[]
    for t in lines:
        t = norm(t).lstrip(BULLETS).strip()
        if t: out.append(t)
    return out

def split_items(line:str):
    if any(sep in line for sep in [",",";"]) or re.search(r"\s{2,}", line):
        return [x.strip() for x in re.split(r"[;,]|\s{2,}", line) if x.strip()]
    return [line.strip()] if line.strip() else []

# detectors
RE_EMAIL = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
RE_LINK  = re.compile(r"(?:https?://\S+|linkedin\.com/\S+|\b@[A-Za-z0-9_-]+)", re.I)
RE_PHONE = re.compile(r"(?:\+?\d{1,3}[\s-]?)?(?:\(?\d{2,4}\)?[\s-]?)?\d{3,5}[\s-]?\d{4}")
RE_DATE  = re.compile(r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|JULY|JUNE|AUG)[^-\n]*[-–][^-\n]*?(?:Present|Current|Now|'?\d{2,4})", re.I)
EDU_HINT = re.compile(r"\b(University|College|Institute|Academy|School)\b", re.I)
DEG_HINT = re.compile(r"\b(MBA|BBA|Bachelor'?s|Master'?s|B\.?Tech|M\.?Tech|PhD|Diploma)\b", re.I)

TOOL_TOKENS = {
    "microsoft office","office","excel","powerpoint","word",
    "g-suite","gsuite","google workspace","google sheets","google docs",
    "tableau","power bi","powerbi","sql","python"
}
def looks_tool(s:str)->bool:
    k = s.lower()
    return any(tok in k for tok in TOOL_TOKENS)

# column-wise from tables, then paragraphs outside
def table_columns(tbl):
    ncols = max(len(r.cells) for r in tbl.rows) if tbl.rows else 0
    cols = [[] for _ in range(ncols)]
    for row in tbl.rows:
        for ci, cell in enumerate(row.cells):
            for p in cell.paragraphs:
                t = norm(p.text)
                if t:
                    cols[ci].append(t)
    return cols

def bucket_column(lines):
    buckets, cur = {}, None
    for ln in lines:
        h = is_heading(ln)
        if h:
            cur = HEADINGS[h]
            buckets.setdefault(cur, [])
            continue
        if cur:
            buckets[cur].append(ln)
    return buckets

def parse_docx_by_columns(path:str)->dict:
    doc = Document(path)
    per_col = []

    # tables as columns
    for tbl in doc.tables:
        for ci, col_lines in enumerate(table_columns(tbl)):
            if not col_lines: continue
            b = bucket_column(col_lines)
            if not b: continue
            while len(per_col) <= ci:
                per_col.append({})
            for k,v in b.items():
                per_col[ci].setdefault(k, []).extend(v)

    # paragraphs outside tables → extra column
    outside_lines = [norm(p.text) for p in doc.paragraphs if norm(p.text)]
    out_b = bucket_column(outside_lines)
    if out_b:
        per_col.append(out_b)

    # final structure
    out = {
        "full_name": None,
        "profile": "",
        "contact": {"phone": None, "email": None, "location": None, "linkedin": None},
        "skills": [],
        "tools": [],
        "experience": [],
        "education": [],
        "projects": [],
        "certifications": [],
        "achievements": [],
        "responsibility": [],
        "activities": [],
    }

    # merge columns carefully
    for col in per_col:
        if not out["full_name"] and col.get("full_name"):
            out["full_name"] = " ".join(col["full_name"]).strip()
        if not out["profile"] and col.get("profile"):
            out["profile"] = " ".join(col["profile"]).strip()

        # CONTACT: fix location+linkedin on one line
        if col.get("contact"):
            lines = col["contact"]
            blob = " | ".join(lines)
            if not out["contact"]["email"]:
                m = RE_EMAIL.search(blob);  out["contact"]["email"] = m.group(0) if m else None
            if not out["contact"]["phone"]:
                m = RE_PHONE.search(blob);  out["contact"]["phone"] = m.group(0) if m else None
            if not out["contact"]["linkedin"]:
                m = RE_LINK.search(blob);   out["contact"]["linkedin"] = m.group(0) if m else None
            for ln in lines:
                low = ln.lower()
                if "current location" in low:
                    # split if LinkedIn text is stuck to it
                    before = ln.split("LinkedIn",1)[0]
                    loc = before.split(":",1)[-1].strip()
                    if loc: out["contact"]["location"] = loc

        # SKILLS & TOOLS (exclude edu-looking items)
        for key, target in [("skills","skills"), ("tools","tools")]:
            if col.get(key):
                for ln in clean_bullets(col[key]):
                    for item in split_items(ln):
                        if not item: continue
                        if EDU_HINT.search(item) or DEG_HINT.search(item):
                            # never allow edu lines into skills/tools
                            continue
                        (out[target]).append(item)

        # EXPERIENCE → structured jobs
        if col.get("experience"):
            lines = [ln for ln in col["experience"] if ln.strip()]
            # split blocks on empty line OR a new role line (heuristic for your template)
            blocks, cur = [], []
            for ln in lines + [""]:
                if not ln.strip():
                    if cur: blocks.append(cur); cur=[]
                    continue
                # heuristic: lines that look like a new role (capitalize words, shortish)
                if cur and (("|" in ln) or re.match(r"^[A-Z][A-Za-z\s]+$", ln) and len(ln) <= 40):
                    # if current already has a company|dates line earlier, start new block
                    if any("|" in x for x in cur[1:]):
                        blocks.append(cur); cur=[ln]; continue
                cur.append(ln)
            if cur: blocks.append(cur)

            for bl in blocks:
                title = bl[0] if bl else None
                company, dates, rest = None, None, bl[1:]
                # find the first 'Company | dates' line
                for i, l in enumerate(rest):
                    if "|" in l:
                        parts = [x.strip() for x in l.split("|",1)]
                        if len(parts)==2:
                            company, dates = parts
                            rest = rest[i+1:]
                            break
                bullets = clean_bullets(rest)
                out["experience"].append({"title": title, "company": company, "dates": dates, "bullets": bullets})

        # EDUCATION → (institution, degree, dates) in order
        if col.get("education"):
            edu_lines = [ln for ln in col["education"] if ln.strip()]
            pending = {"institution": None, "degree": None, "dates": None}
            def flush():
                if pending["institution"] or pending["degree"] or pending["dates"]:
                    out["education"].append(pending.copy())
                pending.update({"institution": None, "degree": None, "dates": None})

            for ln in edu_lines:
                if EDU_HINT.search(ln):
                    flush()
                    pending["institution"] = ln.strip()
                    continue
                if DEG_HINT.search(ln):
                    pending["degree"] = ln.strip()
                    continue
                m = RE_DATE.search(ln)
                if m:
                    pending["dates"] = m.group(0).strip()
                    continue
            flush()

        # PROJECTS: strip dates/tools from description
        if col.get("projects"):
            pr = col["projects"]
            title = pr[0]
            desc_lines = clean_bullets(pr[1:])
            cleaned=[]
            for s in desc_lines:
                if RE_DATE.search(s):
                    continue
                if looks_tool(s) or any(tok in s.lower() for tok in ["tableau","power bi","powerbi"]):
                    for it in split_items(s):
                        if it and it not in out["tools"]:
                            out["tools"].append(it)
                    continue
                cleaned.append(s)
            out["projects"].append({"title": title, "description": " ".join(cleaned)})

        # Simple lists
        if col.get("certifications"):
            for ln in clean_bullets(col["certifications"]):
                out["certifications"].extend(split_items(ln))
        if col.get("achievements"):
            out["achievements"].extend(clean_bullets(col["achievements"]))
        if col.get("activities"):
            out["activities"].extend(clean_bullets(col["activities"]))
        if col.get("responsibility"):
            rl = col["responsibility"]
            role = rl[0]
            bullets = clean_bullets(rl[1:])
            # merge wrapped bullet like "... team of" + "5 and ..."
            merged=[]
            for b in bullets:
                if merged and merged[-1].lower().endswith(" of"):
                    merged[-1] = merged[-1] + " " + b
                else:
                    merged.append(b)
            out["responsibility"].append({"role": role, "bullets": merged})

    # tidy
    out["skills"] = list(dict.fromkeys([s for s in out["skills"] if s]))
    out["tools"]  = list(dict.fromkeys([t for t in out["tools"] if t]))
    if "Microsoft Office G-Suite" in out["tools"]:
        out["tools"].remove("Microsoft Office G-Suite")
        for t in ["Microsoft Office","G-Suite"]:
            if t not in out["tools"]: out["tools"].append(t)

    # drop empty/partial education rows
    out["education"] = [e for e in out["education"] if e["institution"] or e["degree"] or e["dates"]]

    return out

if __name__ == "__main__":
    infile  = "Resume_Template.docx"   # your file
    outfile = "parsed_resume.json"
    data = parse_docx_by_columns(infile)
    with open(outfile, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print("✅ Parsed →", outfile)
