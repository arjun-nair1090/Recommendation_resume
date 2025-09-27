import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, Union

# -------------------- Headers & Helpers --------------------
SECTION_HEADERS = [
    "PERSONAL PROFILE",
    "CONTACT DETAILS",
    "SKILLS AND ABILITIES",
    "SOFTWARE/TOOLS",
    "CERTIFICATION/COURSES",
    "ACHIEVEMENTS",
    "INTERNSHIP EXPERIENCE",
    "ACADEMIC PROFILE",
    "PROJECTS",
    "POSITION OF RESPONSIBILITY",   # canonical
    "POSTION OF RESPONSIBILITY",    # common typo -> normalized
    "CO-CURRICULAR ACTIVITIES",
]

HEADER_CANON = {
    "PERSONAL PROFILE": "PERSONAL PROFILE",
    "CONTACT DETAILS": "CONTACT DETAILS",
    "SKILLS AND ABILITIES": "SKILLS AND ABILITIES",
    "SOFTWARE/TOOLS": "SOFTWARE/TOOLS",
    "CERTIFICATION/COURSES": "CERTIFICATION/COURSES",
    "ACHIEVEMENTS": "ACHIEVEMENTS",
    "INTERNSHIP EXPERIENCE": "INTERNSHIP EXPERIENCE",
    "ACADEMIC PROFILE": "ACADEMIC PROFILE",
    "PROJECTS": "PROJECTS",
    "POSITION OF RESPONSIBILITY": "POSITION OF RESPONSIBILITY",
    "POSTION OF RESPONSIBILITY": "POSITION OF RESPONSIBILITY",  # normalize typo
    "CO-CURRICULAR ACTIVITIES": "CO-CURRICULAR ACTIVITIES",
}

HEADER_PATTERN = re.compile(
    r"^\s*(" + "|".join(re.escape(h) for h in SECTION_HEADERS) + r")\s*:?\s*$",
    flags=re.IGNORECASE
)

EDU_HINTS = re.compile(
    r"\b(university|college|school|bachelor|master|mba|bba|b\.?tech|m\.?tech|gpa|cgpa|grade|"
    r"mumbai university|skilltech|semester|january|february|march|april|may|june|july|august|"
    r"september|october|november|december|’\d{2}|'\d{2}|20\d{2}|19\d{2})\b",
    flags=re.IGNORECASE
)
TOOLS_HINTS = re.compile(
    r"\b(microsoft|office|excel|word|powerpoint|power bi|tableau|g[ -]?suite|google workspace|"
    r"sql|python|r\b|jira|confluence|slack|notion|canva|photoshop|illustrator|figma)\b",
    flags=re.IGNORECASE
)
SOFT_SKILL_HINTS = re.compile(
    r"\b(communication|leadership|teamwork|collaboration|adaptability|flexibility|problem[- ]?solv|"
    r"conflict|negotiation|networking|relationship|time management|project management)\b",
    flags=re.IGNORECASE
)

def canon_header(h: str) -> str:
    h_up = h.strip().upper()
    return HEADER_CANON.get(h_up, h_up)

def is_header_line(line: str) -> Optional[str]:
    m = HEADER_PATTERN.match(line.strip())
    if not m:
        return None
    return canon_header(m.group(1))

def clean_line(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s

# -------------------- DOCX Primitives --------------------
try:
    from docx import Document
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
except ImportError as e:
    raise SystemExit(
        "Missing dependency: python-docx. Install with:\n  pip install python-docx"
    ) from e

def iter_block_items(parent) -> Iterable[Union[Paragraph, Table]]:
    """Yield paragraphs and tables in document order (works for Document and table cells)."""
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P

    body = getattr(parent.element, "body", None)
    # If we're inside a table cell
    if body is None and hasattr(parent, "_tc"):
        body = parent._tc

    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def para_lines(p: Paragraph) -> List[str]:
    t = clean_line(p.text)
    return [t] if t else []

def cell_lines(cell: _Cell) -> List[str]:
    lines: List[str] = []
    for item in iter_block_items(cell):
        if isinstance(item, Paragraph):
            lines.extend(para_lines(item))
        elif isinstance(item, Table):
            # Flatten nested tables row-wise
            lines.extend(flatten_table(item))
    return lines

def flatten_table(tbl: Table) -> List[str]:
    lines: List[str] = []
    for row in tbl.rows:
        row_cells = [clean_line(cell_text) for cell_text in (cell_lines(c) for c in row.cells)]
        # cell_lines returns list; flatten per cell then join
        row_flat: List[str] = []
        for per_cell in row_cells:
            if isinstance(per_cell, list):  # already lines
                joined = " ".join(per_cell).strip()
                if joined:
                    row_flat.append(joined)
            else:
                if per_cell:
                    row_flat.append(str(per_cell).strip())
        if row_flat:
            lines.append(" | ".join([x for x in row_flat if x]))
    return lines

# -------------------- Main Extraction --------------------
def docx_to_sections_live(doc_path: Union[str, Path]) -> Dict[str, List[str]]:
    """
    Single-pass: walk paragraphs and tables; detect headers as they appear;
    assign subsequent lines to the current section until the next header.
    Each 2-column table is processed row-by-row (left→right) and NOT merged
    with the next table, avoiding cross-table bleed.
    """
    doc = Document(str(doc_path))
    sections: Dict[str, List[str]] = defaultdict(list)
    current: Optional[str] = None
    unsectioned: List[str] = []

    def commit_line(line: str):
        nonlocal current
        h = is_header_line(line)
        if h:
            current = h
            # Don’t store the header itself
            return
        if current:
            sections[current].append(line)
        else:
            unsectioned.append(line)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            for ln in para_lines(block):
                commit_line(ln)

        elif isinstance(block, Table):
            # Heuristic: treat as two-column if all rows have 2 cells
            col_counts = {len(r.cells) for r in block.rows} if block.rows else set()
            is_two_col = (len(col_counts) == 1 and next(iter(col_counts), 0) == 2)
            if is_two_col:
                for row in block.rows:
                    # Process left then right so headers in each column are recognized in order
                    for cell in row.cells:
                        for ln in cell_lines(cell):
                            commit_line(ln)
            else:
                # Generic table: read row-wise
                for ln in flatten_table(block):
                    commit_line(ln)

    # Attach any unsectioned lines as FULL_TEXT (fallback)
    if unsectioned:
        sections["FULL_TEXT"] = unsectioned

    return sections

# -------------------- Repair / Normalization --------------------
def repair_sections(sections: Dict[str, List[str]]) -> Dict[str, List[str]]:
    """
    Fix common misfiled content based on simple, conservative heuristics:
      - Move education-ish lines out of SOFTWARE/TOOLS into ACADEMIC PROFILE.
      - Move soft-skill lines out of ACADEMIC PROFILE into SKILLS AND ABILITIES.
    """
    def move_matching(src_key: str, dst_key: str, pattern: re.Pattern):
        src = sections.get(src_key, [])
        keep, move = [], []
        for ln in src:
            if pattern.search(ln):
                move.append(ln)
            else:
                keep.append(ln)
        if move:
            sections[src_key] = keep
            sections[dst_key] = sections.get(dst_key, []) + move

    # 1) Tools should not contain schools/degrees/dates
    move_matching("SOFTWARE/TOOLS", "ACADEMIC PROFILE", EDU_HINTS)

    # 2) Academic Profile should not contain generic soft skills
    move_matching("ACADEMIC PROFILE", "SKILLS AND ABILITIES", SOFT_SKILL_HINTS)

    # 3) If SKILLS is still empty but ACADEMIC PROFILE looks 80% skills, move all
    skills = sections.get("SKILLS AND ABILITIES", [])
    acad = sections.get("ACADEMIC PROFILE", [])
    if not skills and acad:
        hits = sum(1 for ln in acad if SOFT_SKILL_HINTS.search(ln) and not EDU_HINTS.search(ln))
        if hits >= max(1, int(0.8 * len(acad))):
            sections["SKILLS AND ABILITIES"] = acad
            sections["ACADEMIC PROFILE"] = []

    # Strip empty sections
    for k in list(sections.keys()):
        sections[k] = [ln for ln in sections[k] if clean_line(ln)]
        if not sections[k]:
            # keep the key but as empty list (up to you). We’ll keep it for visibility.
            pass

    return sections

# -------------------- Public API --------------------
def docx_resume_to_json(docx_path: Union[str, Path]) -> Dict[str, str]:
    """
    Read a .docx resume and return a dict of {SECTION: "joined text"}.
    Always returns valid JSON-serializable data.
    """
    p = Path(docx_path)
    if not p.exists():
        raise FileNotFoundError(f"File not found: {p}")

    raw_sections = docx_to_sections_live(p)
    fixed_sections = repair_sections(raw_sections)

    # Normalize header keys to canonical form and join lines
    out: Dict[str, str] = {}
    for k, lines in fixed_sections.items():
        canon = canon_header(k)
        # Join with newlines (preserve bullet/line structure)
        out[canon] = "\n".join(lines).strip()

    return out

# -------------------- CLI --------------------
if __name__ == "__main__":
    import sys
    from pathlib import Path

    # Prefer CWD, then script folder
    candidates = [
        Path.cwd() / "Resume_Template.docx",
        Path(__file__).resolve().parent / "Resume_Template.docx",
    ]

    docx_path = next((p for p in candidates if p.exists()), None)
    if not docx_path:
        raise SystemExit(
            "Resume_Template.docx not found.\n"
            "Place it in the current working directory or next to this script."
        )

    out_path = docx_path.with_suffix(".json")

    try:
        data = docx_resume_to_json(docx_path)
    except Exception as e:
        # Make any failure obvious and actionable
        raise SystemExit(f"Failed to convert '{docx_path.name}': {e}") from e

    try:
        import json
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise SystemExit(f"Could not write JSON to '{out_path}': {e}") from e

    print(f"Saved JSON to: {out_path}")
