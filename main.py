#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Union

# ==================== Section Headers & Patterns ====================

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

# Hints used in the repair step
EDU_HINTS = re.compile(
    r"\b(university|college|school|bachelor|master|mba|bba|gpa|cgpa|grade|"
    r"semester|campus|institute|degree|diploma|"
    r"mumbai university|atlas skilltech|"
    r"january|february|march|april|may|june|july|august|september|october|november|december|"
    r"’\d{2}|'?\d{2}|20\d{2}|19\d{2})\b",
    flags=re.IGNORECASE
)

TOOLS_HINTS = re.compile(
    r"\b(microsoft(?: office)?|excel|word|powerpoint|power\s*bi|tableau|g[ -]?suite|google analytics|"
    r"sql|python|r\b|jira|confluence|notion|slack|figma|canva|photoshop|illustrator)\b",
    flags=re.IGNORECASE
)

SOFT_SKILL_HINTS = re.compile(
    r"\b(communication|leadership|teamwork|collaboration|adaptability|flexibility|problem[- ]?solv|"
    r"conflict|negotiation|networking|relationship|time management|project management)\b",
    flags=re.IGNORECASE
)

FANCY_APOST = re.compile(r"[\u2019\u2018]")  # ’ ‘
MONTH_WORDS = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)"
DATE_RANGE_RE = re.compile(rf"\b{MONTH_WORDS}\b.*\b\d{{2,4}}\b.*?-.*?\b{MONTH_WORDS}\b.*\b\d{{2,4}}\b", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")
DEGREE_HINTS = re.compile(
    r"\b(mba|master(?:'s)? of business administration|masters of business administration|"
    r"bba|b\.?ba|b\.?tech|m\.?tech|bsc|msc|ba|ma|"
    r"bachelor(?:'s)?|master(?:'s)?|degree|diploma)\b",
    re.IGNORECASE
)
GENERIC_NAME_LINES = re.compile(r"^(full\s*name|student|resume)$", re.IGNORECASE)

# ==================== Helpers ====================

def canon_header(h: str) -> str:
    return HEADER_CANON.get(h.strip().upper(), h.strip().upper())

def is_header_line(line: str) -> Optional[str]:
    m = HEADER_PATTERN.match((line or "").strip())
    if not m:
        return None
    return canon_header(m.group(1))

def clean_line(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def normalize_line(s: str) -> str:
    s = s or ""
    s = FANCY_APOST.sub("'", s)
    s = s.replace("DURATIONR", "DURATION")  # fix common typo
    s = re.sub(r"\s+", " ", s).strip()
    return s

def is_tools_only(line: str) -> bool:
    return bool(TOOLS_HINTS.search(line)) and not (EDU_HINTS.search(line) or DEGREE_HINTS.search(line))

def is_degreeish(line: str) -> bool:
    return bool(DEGREE_HINTS.search(line) or EDU_HINTS.search(line))

def is_dateish(line: str) -> bool:
    l = FANCY_APOST.sub("'", line or "")
    return bool(DATE_RANGE_RE.search(l) or YEAR_RE.search(l) or "DURATION" in l.upper())

def dedupe_preserve_order(lines: List[str]) -> List[str]:
    seen, out = set(), []
    for ln in lines:
        key = ln.lower()
        if key not in seen:
            seen.add(key)
            out.append(ln)
    return out

# ==================== DOCX Reading ====================

try:
    from docx import Document
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
except ImportError as e:
    raise SystemExit(
        "Missing dependency: python-docx.\nInstall with: pip install python-docx"
    ) from e

def iter_block_items(parent) -> Iterable[Union[Paragraph, Table]]:
    """Yield paragraphs and tables in document order (Document or table cell)."""
    if isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent.element.body

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def para_lines(p: Paragraph) -> List[str]:
    t = clean_line(p.text)
    return [t] if t else []

def cell_to_lines(cell: _Cell) -> List[str]:
    lines: List[str] = []
    for item in iter_block_items(cell):
        if isinstance(item, Paragraph):
            lines.extend(para_lines(item))
        elif isinstance(item, Table):
            lines.extend(table_iter_lines(item))  # nested table
    return lines

def table_iter_lines(tbl: Table) -> List[str]:
    """
    Return lines by reading the table row-wise and cell-wise.
    This preserves left→right order (good for 2-column resumes).
    """
    lines: List[str] = []
    for row in tbl.rows:
        for cell in row.cells:
            lines.extend(cell_to_lines(cell))
    return lines

# ==================== Core Extraction ====================

def docx_to_sections_live(doc_path: Union[str, Path]) -> Dict[str, List[str]]:
    """
    Single-pass parse: detect headers as they appear; assign subsequent lines
    to that section until the next header. Works across paragraphs and tables.
    """
    doc = Document(str(doc_path))
    sections: Dict[str, List[str]] = defaultdict(list)
    current: Optional[str] = None
    unsectioned: List[str] = []

    def commit_line(line: str):
        nonlocal current
        if not line:
            return
        h = is_header_line(line)
        if h:
            current = h  # don't store the header itself
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
            for ln in table_iter_lines(block):
                commit_line(ln)

    if unsectioned:
        sections["FULL_TEXT"] = unsectioned

    return sections

# ==================== Repair / Normalization ====================

def repair_sections(sections: Dict[str, List[str]]) -> Dict[str, List[str]]:
    """
    Strong repair:
      - Keep tools in SOFTWARE/TOOLS; move degrees/dates to ACADEMIC PROFILE.
      - PROJECTS: move tool-only lines to SOFTWARE/TOOLS; date-ish to ACADEMIC PROFILE.
      - Move soft-skill lines from ACADEMIC PROFILE to SKILLS AND ABILITIES.
      - Clean FULL_TEXT placeholders.
      - Normalize whitespace, dedupe, keep order.
    """
    # Normalize every line first
    for k in list(sections.keys()):
        sections[k] = [normalize_line(ln) for ln in sections.get(k, []) if normalize_line(ln)]

    # SOFTWARE/TOOLS: keep tools; move degree/date-ish to academic
    tools = sections.get("SOFTWARE/TOOLS", [])
    keep_tools, to_acad = [], []
    for ln in tools:
        if is_degreeish(ln) or is_dateish(ln):
            to_acad.append(ln)
        else:
            keep_tools.append(ln)
    sections["SOFTWARE/TOOLS"] = keep_tools

    # PROJECTS: move tool-only lines to tools; date-ish to academic (unless explicitly a "PROJECT" line)
    projs = sections.get("PROJECTS", [])
    proj_keep, proj_to_tools, proj_to_acad = [], [], []
    for ln in projs:
        if is_tools_only(ln):
            proj_to_tools.append(ln)
        elif is_dateish(ln) and "PROJECT" not in ln.upper():
            proj_to_acad.append(ln)
        else:
            proj_keep.append(ln)
    sections["PROJECTS"] = proj_keep

    # Consolidate moves
    acad = sections.get("ACADEMIC PROFILE", [])
    acad.extend(to_acad)
    acad.extend(proj_to_acad)
    sections["ACADEMIC PROFILE"] = acad

    sw = sections.get("SOFTWARE/TOOLS", [])
    sw.extend(proj_to_tools)
    sections["SOFTWARE/TOOLS"] = sw

    # ACADEMIC PROFILE: move soft skills to SKILLS AND ABILITIES
    acad = sections.get("ACADEMIC PROFILE", [])
    acad_keep, to_skills = [], []
    for ln in acad:
        if SOFT_SKILL_HINTS.search(ln) and not is_degreeish(ln):
            to_skills.append(ln)
        else:
            acad_keep.append(ln)
    sections["ACADEMIC PROFILE"] = acad_keep

    skills = sections.get("SKILLS AND ABILITIES", [])
    skills.extend(to_skills)
    sections["SKILLS AND ABILITIES"] = skills

    # FULL_TEXT: drop placeholder junk like "FULL NAME", "Student", "Resume"
    ft = [ln for ln in sections.get("FULL_TEXT", []) if not GENERIC_NAME_LINES.match(ln)]
    if ft:
        sections["FULL_TEXT"] = ft
    else:
        sections.pop("FULL_TEXT", None)

    # Deduplicate & strip empties (preserve keys, but contents cleaned)
    for k in list(sections.keys()):
        sections[k] = dedupe_preserve_order([ln for ln in sections[k] if ln.strip()])

    return sections

# ==================== Public API ====================

def docx_resume_to_json(docx_path: Union[str, Path]) -> Dict[str, str]:
    """
    Read a .docx resume and return a dict of {SECTION: "joined text"}.
    If no headers are found, you'll get {"FULL_TEXT": "..."}.
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
        out[canon] = "\n".join(lines).strip()
    return out

# ==================== Main (Auto-input: Resume_Template.docx) ====================

if __name__ == "__main__":
    # Locate Resume_Template.docx in CWD, else next to this script
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
        raise SystemExit(f"Failed to convert '{docx_path.name}': {e}") from e

    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise SystemExit(f"Could not write JSON to '{out_path}': {e}") from e

    print(f"Saved JSON to: {out_path}")
