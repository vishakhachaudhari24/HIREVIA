"""
=============================================================
AI-ASSISTED RESUME SCREENING SYSTEM - main.py
=============================================================
Fixes in this version:
  1. Skill dedup — removes any skill from "missing" that already
     appears in "matched" (AI inconsistency guard)
  2. Better name extraction — handles more resume formats,
     skips section headers like "Profile Summary",
     falls back to filename when name can't be found
  3. Word (.docx) support — extracts text from .docx files
     using python-docx alongside pdfplumber for PDFs
  4. Score-based criteria — anyone scoring 70%+ is marked
     as "Criteria Met" regardless of keyword filter result

Outputs:
  ranked_candidates.xlsx  — formatted Excel table
  ranked_resumes.zip      — renamed PDFs/DOCXs in ranked order

Uses: Groq API (llama-3.3-70b-versatile)
=============================================================
"""

import os
import sys
import zipfile
import re
import json
import time
import tempfile
import shutil
import pandas as pd
import pdfplumber
from dotenv import load_dotenv
from groq import Groq
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from scorer import calculate_final_score

# ── Try importing python-docx (optional — only needed for .docx resumes) ──
try:
    from docx import Document as DocxDocument
    DOCX_SUPPORTED = True
except ImportError:
    DOCX_SUPPORTED = False
    print("[INFO] python-docx not installed. .docx resumes will be skipped.")
    print("       Run: pip install python-docx  to enable Word support.\n")

load_dotenv()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

MAX_RESUME_CHARS    = 6000
OUTPUT_EXCEL        = "ranked_candidates.xlsx"
OUTPUT_ZIP          = "ranked_resumes.zip"
RANKED_RESUMES_DIR  = "ranked_resumes"
CRITERIA_MET_SCORE  = 70          # Score threshold for auto criteria-met

SUPPORTED_EXTS = [".pdf"] + ([".docx"] if DOCX_SUPPORTED else [])


# ══════════════════════════════════════════════════════════
# STEP 1 — TEXT EXTRACTION  (PDF + DOCX)
# ══════════════════════════════════════════════════════════

def extract_text_from_pdf(pdf_path: str) -> str:
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"  [WARNING] Could not read PDF '{pdf_path}': {e}")
    return text.strip()


def extract_text_from_docx(docx_path: str) -> str:
    """Extract plain text from a Word .docx file."""
    if not DOCX_SUPPORTED:
        return ""
    text = ""
    try:
        doc = DocxDocument(docx_path)
        for para in doc.paragraphs:
            if para.text.strip():
                text += para.text + "\n"
        # Also pull text from tables inside the document
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        text += cell.text + "\n"
    except Exception as e:
        print(f"  [WARNING] Could not read DOCX '{docx_path}': {e}")
    return text.strip()


def extract_text(file_path: str) -> str:
    """Route to the correct extractor based on file extension."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_path)
    elif ext == ".docx":
        return extract_text_from_docx(file_path)
    return ""


# ══════════════════════════════════════════════════════════
# STEP 2 — BIAS REMOVAL
# ══════════════════════════════════════════════════════════

def remove_bias_fields(text: str) -> str:
    text = re.sub(r'\S+@\S+\.\S+', '[EMAIL REMOVED]', text)
    text = re.sub(r'(\+?\d[\d\s\-\(\)]{7,}\d)', '[PHONE REMOVED]', text)
    text = re.sub(r'(date of birth|dob|d\.o\.b)[^\n]*', '[DOB REMOVED]', text, flags=re.IGNORECASE)
    text = re.sub(r'\bage\s*[:\-]?\s*\d{2}\b|\b\d{2}\s+years?\s+old\b', '[AGE REMOVED]', text, flags=re.IGNORECASE)
    text = re.sub(r'\d+\s+\w[\w\s]+(?:street|st|avenue|ave|road|rd|lane|ln|drive|dr|blvd|way)[^\n]*', '[ADDRESS REMOVED]', text, flags=re.IGNORECASE)
    text = re.sub(r'\b[A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2}\b', '[POSTCODE REMOVED]', text)
    text = re.sub(r'\b\d{5}(?:-\d{4})?\b', '[ZIP REMOVED]', text)
    text = re.sub(r'\b(he|she|his|her|him|himself|herself)\b', '[PRONOUN REMOVED]', text, flags=re.IGNORECASE)
    text = re.sub(r'(photo|photograph|image|picture)[^\n]*', '[PHOTO REFERENCE REMOVED]', text, flags=re.IGNORECASE)
    text = re.sub(r'(nationality|citizenship|citizen of|national of)[^\n]*', '[NATIONALITY REMOVED]', text, flags=re.IGNORECASE)
    lines = text.splitlines()
    text = "\n".join(lines[2:]) if len(lines) > 2 else text
    return text.strip()


# ══════════════════════════════════════════════════════════
# STEP 3 — SMART TRUNCATION
# ══════════════════════════════════════════════════════════

def smart_truncate(text: str, max_chars: int = MAX_RESUME_CHARS) -> str:
    if len(text) <= max_chars:
        return text
    priority_keywords = [
        "experience", "work history", "employment",
        "projects", "project",
        "skills", "technical skills", "core competencies"
    ]
    sections = {}
    lines = text.splitlines()
    current_section = "other"
    sections[current_section] = []
    for line in lines:
        for kw in priority_keywords:
            if re.search(rf'\b{kw}\b', line, re.IGNORECASE) and len(line) < 60:
                current_section = kw.lower()
                sections.setdefault(current_section, [])
                break
        sections.setdefault(current_section, []).append(line)
    priority_order = ["experience", "work history", "employment",
                      "projects", "project",
                      "skills", "technical skills", "core competencies"]
    result = []
    for key in priority_order:
        if key in sections:
            result.extend(sections[key])
    return "\n".join(result)[:max_chars]


# ══════════════════════════════════════════════════════════
# STEP 4 — HARD FILTER
# ══════════════════════════════════════════════════════════

def extract_hard_filter_criteria(job_description: str) -> dict:
    prompt = f"""
You are a job description parser.
Read this job description and extract ONLY hard requirements.
Return ONLY valid JSON — no explanation, no markdown fences.

JSON format:
{{
  "min_years_experience": <integer or null>,
  "required_skills": [<list of must-have skills or empty list>],
  "required_location": "<location string or null>",
  "work_authorization": "<authorization requirement string or null>"
}}

Job Description:
{job_description[:3000]}
"""
    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
            max_tokens=400,
        )
        raw = response.choices[0].message.content.strip()
        raw = re.sub(r"```(?:json)?", "", raw).strip("` \n")
        return json.loads(raw)
    except Exception as e:
        print(f"  [WARNING] Could not parse hard filter criteria: {e}")
        return {"min_years_experience": None, "required_skills": [],
                "required_location": None, "work_authorization": None}


def apply_hard_filter(resume_text: str, criteria: dict) -> tuple[bool, str]:
    flags = []
    min_yrs = criteria.get("min_years_experience")
    if min_yrs:
        year_matches = re.findall(r'(\d+)\+?\s+years?', resume_text, re.IGNORECASE)
        max_found = max((int(y) for y in year_matches), default=0)
        if max_found < min_yrs:
            flags.append(f"May not meet {min_yrs}+ years experience")
    req_skills = criteria.get("required_skills", [])
    missing_skills = [s for s in req_skills
                      if not re.search(rf'\b{re.escape(s)}\b', resume_text, re.IGNORECASE)]
    if missing_skills:
        flags.append(f"Skills not found: {', '.join(missing_skills)}")
    met = len(flags) == 0
    reason = "; ".join(flags) if flags else "All mandatory criteria appear met"
    return met, reason


# ══════════════════════════════════════════════════════════
# STEP 5 — AI EVALUATION  (skill dedup in prompt + post-process)
# ══════════════════════════════════════════════════════════

def evaluate_resume_with_ai(resume_text: str, job_description: str) -> dict:
    prompt = f"""
You are an expert, unbiased recruitment analyst.

Evaluate the candidate's resume against the job description.
Use ONLY information explicitly present in the resume.
Be precise and consistent — do NOT list the same skill in both
"key_skills_matched" AND "missing_or_weak_areas". If a skill is
found in the resume, it belongs ONLY in key_skills_matched.

Return ONLY a valid JSON object — no markdown fences, no extra text.

JSON schema:
{{
  "skill_match_score": <integer 0-100>,
  "experience_relevance_score": <integer 0-100>,
  "project_relevance_score": <integer 0-100>,
  "domain_fit_score": <integer 0-100>,
  "career_progression_score": <integer 0-100>,
  "relevant_experience_summary": "<3-4 sentences mentioning years of experience and 1-2 relevant projects if present>",
  "key_skills_matched": [<skills FOUND in resume, max 8>],
  "missing_or_weak_areas": [<skills/areas NOT found or weak in resume vs JD, max 5>],
  "ai_summary": "<3-4 line professional summary of candidate fit>"
}}

Scoring:
- skill_match_score: % of JD-required skills found in resume
- experience_relevance_score: relevance of work history to JD
- project_relevance_score: relevance of projects; 50 if none mentioned
- domain_fit_score: industry/domain alignment
- career_progression_score: evidence of growth and responsibility

=== JOB DESCRIPTION ===
{job_description[:2000]}

=== RESUME (anonymised) ===
{resume_text}
"""
    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=800,
        )
        raw = response.choices[0].message.content.strip()
        raw = re.sub(r"```(?:json)?", "", raw).strip("` \n")
        result = json.loads(raw)

        # Post-process: remove from "missing" anything already in "matched"
        matched_lower = [s.lower().strip() for s in result.get("key_skills_matched", [])]
        missing       = result.get("missing_or_weak_areas", [])
        result["missing_or_weak_areas"] = [
            m for m in missing
            if not any(word in m.lower() for word in matched_lower)
        ]

        return result

    except json.JSONDecodeError as e:
        print(f"  [ERROR] Invalid JSON from AI: {e}")
        return _empty_evaluation(f"JSON parse error: {e}")
    except Exception as e:
        print(f"  [ERROR] AI evaluation failed: {e}")
        return _empty_evaluation(str(e))


def _empty_evaluation(reason: str) -> dict:
    return {
        "skill_match_score": 0, "experience_relevance_score": 0,
        "project_relevance_score": 0, "domain_fit_score": 0,
        "career_progression_score": 0,
        "relevant_experience_summary": f"Evaluation failed: {reason}",
        "key_skills_matched": [], "missing_or_weak_areas": ["Could not evaluate"],
        "ai_summary": f"Evaluation failed: {reason}",
    }


# ══════════════════════════════════════════════════════════
# STEP 6 — CANDIDATE NAME EXTRACTION  (improved)
# ══════════════════════════════════════════════════════════

# Common resume section headers to skip — including multi-word ones
_SECTION_HEADERS = re.compile(
    r'^(resume|curriculum vitae|cv|profile|objective|summary|contact|'
    r'personal|information|details|about|introduction|overview|'
    r'profile summary|career summary|professional summary|career objective|'
    r'about me|personal details|personal information|executive summary)$',
    re.IGNORECASE
)

# Noise words that appear at the top of resumes but aren't names
_NOISE_WORDS = re.compile(
    r'\b(mr|ms|mrs|dr|prof|phone|email|mobile|address|linkedin|github|'
    r'portfolio|website|http|www)\b|[@,]',
    re.IGNORECASE
)


def extract_candidate_name(raw_text: str, filename: str = "") -> str:
    """
    Multi-strategy name extractor:
      1. Scan first 8 non-empty lines for a 2-4 word capitalised name
      2. Try to pull name from the filename itself
      3. Fall back to 'Unknown'
    """
    lines = [l.strip() for l in raw_text.splitlines() if l.strip()][:8]

    for line in lines:
        # Skip section headers (single or multi-word)
        if _SECTION_HEADERS.match(line.strip()):
            continue
        # Skip lines with noise keywords (emails, URLs, etc.)
        if _NOISE_WORDS.search(line):
            continue
        # Skip very long lines (unlikely to be just a name)
        if len(line) > 60:
            continue
        # Skip lines that are ALL CAPS sentences (likely headings, not names)
        if line.isupper() and len(line.split()) > 3:
            continue
        # Match 2–4 words where each word starts with a capital letter
        # Handles: "Rahul Sharma", "Mary-Jane Watson", "Riya A Kaurase"
        if re.match(r'^[A-Z][a-zA-Z\-]+([\s][A-Z][a-zA-Z\-]+){1,3}$', line):
            return line.strip()

    # Strategy 2: derive name from filename
    # e.g. "Riya_Kaurase_Resume.docx" → "Riya Kaurase"
    if filename:
        base = os.path.splitext(os.path.basename(filename))[0]
        # Remove common prefixes/suffixes
        base = re.sub(r'(resume|cv|candidate|applicant|intern|engineer|analyst)', '', base, flags=re.IGNORECASE)
        base = re.sub(r'[\(\)\d\-_\.]+', ' ', base).strip()
        parts = [p for p in base.split() if len(p) > 1 and re.match(r'^[A-Za-z]+$', p)]
        if 2 <= len(parts) <= 4:
            return ' '.join(p.capitalize() for p in parts)

    return "Unknown"


# ══════════════════════════════════════════════════════════
# STEP 7 — LOAD RESUMES  (PDF + DOCX)
# ══════════════════════════════════════════════════════════

def load_resume_paths(source_path: str) -> tuple[list[str], str]:
    temp_dir = ""

    if zipfile.is_zipfile(source_path):
        temp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(source_path, 'r') as zf:
            zf.extractall(temp_dir)
        folder = temp_dir
    elif os.path.isdir(source_path):
        folder = source_path
    else:
        print(f"[ERROR] '{source_path}' is not a valid folder or ZIP file.")
        sys.exit(1)

    resume_files = [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if os.path.splitext(f)[1].lower() in SUPPORTED_EXTS
    ]

    if not resume_files:
        print(f"[ERROR] No supported resume files found in '{source_path}'.")
        print(f"        Supported formats: {', '.join(SUPPORTED_EXTS)}")
        sys.exit(1)

    if not DOCX_SUPPORTED:
        skipped = [f for f in os.listdir(folder) if f.lower().endswith(".docx")]
        if skipped:
            print(f"  [INFO] Skipped {len(skipped)} .docx file(s). "
                  f"Run: pip install python-docx  to enable.\n")

    return sorted(resume_files), temp_dir


# ══════════════════════════════════════════════════════════
# STEP 8 — SAFE FILENAME HELPER
# ══════════════════════════════════════════════════════════

def safe_filename(name: str) -> str:
    name = re.sub(r'[^\w\s\-]', '', name)
    return name.strip().replace(' ', '_') or "Unknown"


# ══════════════════════════════════════════════════════════
# STEP 9 — EXCEL EXPORT
# ══════════════════════════════════════════════════════════

_HEADER_BG   = "1F3864"
_HEADER_FG   = "FFFFFF"
_TOP3_BG     = "E8F5E9"
_ALT_BG      = "F5F7FA"
_WHITE_BG    = "FFFFFF"
_BORDER_CLR  = "D0D7E3"
_MET_GREEN   = "2E7D32"
_NOT_MET_RED = "C62828"

_THIN   = Side(style="thin", color=_BORDER_CLR)
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_COLUMNS = [
    ("Rank",                   "Rank",                    10,  False),
    ("Resume ID",              "Resume ID",               12,  False),
    ("Candidate Name",         "Candidate Name",          24,  False),
    ("Final Match Score (%)",  "Final Match Score (%)",   20,  False),
    ("Mandatory Criteria Met", "Mandatory Criteria Met",  28,  True ),
    ("Key Skills Matched",     "Key Skills Matched",      36,  True ),
    ("Missing / Weak Areas",   "Missing / Weak Areas",    32,  True ),
    ("AI Summary",             "AI Summary",              54,  True ),
    ("Resume Link",            "resume_path",             30,  False),
]


def _score_fill(score: float) -> PatternFill:
    if score >= 70:
        return PatternFill("solid", fgColor="C8E6C9")
    if score >= 50:
        return PatternFill("solid", fgColor="FFF9C4")
    return PatternFill("solid", fgColor="FFCDD2")


def export_excel(df: pd.DataFrame, file_path_map: dict,
                 output_path: str, top_n: int = None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidate Rankings"

    if top_n:
        df = df.head(top_n).copy()

    ws.freeze_panes = "A2"

    h_font  = Font(name="Arial", bold=True, color=_HEADER_FG, size=11)
    h_fill  = PatternFill("solid", fgColor=_HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ci, (label, _, width, _) in enumerate(_COLUMNS, 1):
        cell = ws.cell(row=1, column=ci, value=label)
        cell.font = h_font; cell.fill = h_fill
        cell.alignment = h_align; cell.border = _BORDER
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[1].height = 30

    for ri, (_, row) in enumerate(df.iterrows(), 2):
        score    = row["Final Match Score (%)"]
        is_top3  = (ri - 2) < 3
        bg       = _TOP3_BG if is_top3 else (_ALT_BG if ri % 2 == 0 else _WHITE_BG)
        def_fill = PatternFill("solid", fgColor=bg)

        for ci, (_, key, _, wrap) in enumerate(_COLUMNS, 1):
            cell = ws.cell(row=ri, column=ci)
            cell.border    = _BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=wrap)

            if key == "Rank":
                cell.value     = f"Rank{(ri - 1):03d}"
                cell.font      = Font(name="Arial", bold=True, size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill      = def_fill

            elif key == "Final Match Score (%)":
                cell.value         = score
                cell.number_format = '0"%"'
                cell.font          = Font(name="Arial", bold=True, size=11)
                cell.alignment     = Alignment(horizontal="center", vertical="center")
                cell.fill          = _score_fill(score)

            elif key == "Mandatory Criteria Met":
                val = str(row.get(key, ""))
                met = val.strip().lower().startswith("yes")
                cell.value = val
                cell.font  = Font(name="Arial", size=10,
                                  color=_MET_GREEN if met else _NOT_MET_RED, bold=True)
                cell.fill  = def_fill

            elif key == "resume_path":
                resume_id = row.get("Resume ID", "")
                fpath     = file_path_map.get(resume_id, "")
                if fpath and os.path.exists(fpath):
                    abs_path = os.path.abspath(fpath)
                    uri = abs_path.replace("\\", "/")
                    if not uri.startswith("/"):
                        uri = "/" + uri
                    cell.hyperlink = f"file://{uri}"
                    cell.value     = "📄 Open Resume"
                    cell.font      = Font(name="Arial", color="1155CC",
                                         underline="single", size=10)
                else:
                    cell.value = "N/A"
                    cell.font  = Font(name="Arial", size=10)
                cell.fill      = def_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            else:
                cell.value = row.get(key, "")
                cell.font  = Font(name="Arial", size=10)
                cell.fill  = def_fill

        ws.row_dimensions[ri].height = 72

    last_col = get_column_letter(len(_COLUMNS))
    ws.auto_filter.ref = f"A1:{last_col}{len(df) + 1}"

    # ── Summary sheet ──────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "AI Resume Screening — Run Summary"
    ws2["A1"].font = Font(name="Arial", bold=True, size=14)

    criteria_met_count = int((df["Mandatory Criteria Met"].str.startswith("Yes")).sum())
    summary_rows = [
        ("Total Candidates",    len(df)),
        ("Criteria Met",        criteria_met_count),
        ("Criteria Not Met",    len(df) - criteria_met_count),
        ("Avg Match Score (%)", f"=AVERAGE('Candidate Rankings'!D2:D{len(df)+1})"),
        ("Top Score (%)",       f"=MAX('Candidate Rankings'!D2:D{len(df)+1})"),
        ("Low Score (%)",       f"=MIN('Candidate Rankings'!D2:D{len(df)+1})"),
        ("Top Candidate",       df.iloc[0]["Candidate Name"] if len(df) else "N/A"),
        ("Score Threshold for Criteria Met", f"{CRITERIA_MET_SCORE}%"),
    ]
    for r, (label, val) in enumerate(summary_rows, 3):
        ws2.cell(r, 1, label).font = Font(name="Arial", bold=True, size=10)
        ws2.cell(r, 2, val).font   = Font(name="Arial", size=10)
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 28

    wb.save(output_path)
    print(f"   ✅ Excel saved → {output_path}")


# ══════════════════════════════════════════════════════════
# STEP 10 — ZIP RANKED RESUMES
# ══════════════════════════════════════════════════════════

def build_ranked_zip(df: pd.DataFrame, file_path_map: dict,
                     resumes_dir: str, zip_path: str):
    os.makedirs(resumes_dir, exist_ok=True)
    seen_names = set()

    for rank, row in df.iterrows():
        display_rank  = rank + 1
        resume_id     = row["Resume ID"]
        score         = row["Final Match Score (%)"]
        name          = row["Candidate Name"]
        original_file = file_path_map.get(resume_id)

        if not original_file or not os.path.exists(original_file):
            print(f"  [WARNING] File not found for {resume_id}, skipping.")
            continue

        ext          = os.path.splitext(original_file)[1].lower()
        base_name    = safe_filename(name)
        new_filename = f"Rank{display_rank:03d}_{score}pct_{base_name}{ext}"

        # Collision guard
        if new_filename in seen_names:
            new_filename = f"Rank{display_rank:03d}_{score}pct_{base_name}_{resume_id}{ext}"
        seen_names.add(new_filename)

        shutil.copy2(original_file, os.path.join(resumes_dir, new_filename))

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in sorted(os.listdir(resumes_dir)):
            if os.path.splitext(fname)[1].lower() in SUPPORTED_EXTS:
                zf.write(os.path.join(resumes_dir, fname), arcname=fname)

    print(f"   ✅ ZIP saved  → {zip_path}  ({len(seen_names)} files)")


# ══════════════════════════════════════════════════════════
# MAIN PIPELINE
# ══════════════════════════════════════════════════════════

def run_pipeline(job_description_path: str, resumes_source: str,
                 top_n: int = None):

    with open(job_description_path, "r", encoding="utf-8") as f:
        job_description = f.read().strip()
    print(f"\n✅ Job description loaded ({len(job_description)} chars).")

    print("🔍 Parsing hard filter criteria from JD...")
    criteria = extract_hard_filter_criteria(job_description)
    print(f"   Criteria: {json.dumps(criteria, indent=2)}\n")

    resume_paths, temp_dir = load_resume_paths(resumes_source)
    total = len(resume_paths)
    print(f"📂 Found {total} resume(s) to process "
          f"({', '.join(SUPPORTED_EXTS)} supported).\n")

    results       = []
    file_path_map = {}

    for idx, file_path in enumerate(resume_paths, 1):
        resume_id = f"R{idx:03d}"
        filename  = os.path.basename(file_path)
        file_path_map[resume_id] = file_path
        print(f"[{idx}/{total}] Processing: {filename}")

        raw_text = extract_text(file_path)
        if not raw_text:
            print(f"  ⚠️  Empty or unreadable file — skipping.\n")
            results.append({
                "Resume ID":              resume_id,
                "Candidate Name":         "Unknown",
                "Mandatory Criteria Met": "N/A",
                "Final Match Score (%)":  0,
                "Key Skills Matched":     "",
                "Missing / Weak Areas":   "Unreadable file",
                "AI Summary":             "File could not be processed",
            })
            continue

        # Improved name extraction with filename fallback
        candidate_name = extract_candidate_name(raw_text, filename)
        clean_text     = remove_bias_fields(raw_text)
        truncated_text = smart_truncate(clean_text)

        criteria_met, filter_reason = apply_hard_filter(truncated_text, criteria)

        print(f"  🤖 Sending to Groq AI for evaluation...")
        evaluation  = evaluate_resume_with_ai(truncated_text, job_description)
        final_score = calculate_final_score(evaluation)

        # ── FIX 4: Score-based criteria override ──────────────────
        # If score >= threshold → always mark as Met
        # If score < threshold  → use keyword filter result
        if final_score >= CRITERIA_MET_SCORE:
            criteria_label = "Yes"
        elif criteria_met:
            criteria_label = "Yes"
        else:
            criteria_label = f"No — {filter_reason}"

        print(f"  ✅ Score: {final_score}% | "
              f"Name: {candidate_name} | "
              f"Criteria: {'Met' if criteria_label == 'Yes' else 'Not Met'}\n")

        results.append({
            "Resume ID":              resume_id,
            "Candidate Name":         candidate_name,
            "Mandatory Criteria Met": criteria_label,
            "Final Match Score (%)":  final_score,
            "Key Skills Matched":     ", ".join(evaluation.get("key_skills_matched", [])),
            "Missing / Weak Areas":   ", ".join(evaluation.get("missing_or_weak_areas", [])),
            "AI Summary":             evaluation.get("ai_summary", ""),
        })

        time.sleep(0.5)

    if temp_dir:
        shutil.rmtree(temp_dir, ignore_errors=True)

    df = pd.DataFrame(results)
    df = df.sort_values("Final Match Score (%)", ascending=False).reset_index(drop=True)

    print(f"\n📊 Generating Excel report...")
    export_excel(df, file_path_map, OUTPUT_EXCEL, top_n=top_n)

    print(f"📁 Building ranked ZIP...")
    build_ranked_zip(df, file_path_map, RANKED_RESUMES_DIR, OUTPUT_ZIP)

    top = df.iloc[0]
    print(f"\n🎉 Done!")
    print(f"   📊 Excel  → {OUTPUT_EXCEL}")
    print(f"   📦 ZIP    → {OUTPUT_ZIP}")
    print(f"   Processed {len(results)} resume(s).")
    print(f"   Top candidate: {top['Candidate Name']} ({top['Final Match Score (%)']}%)")
    print(f"\n💡 Send your client BOTH files.")


# ══════════════════════════════════════════════════════════
# CLI ENTRY POINT
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("\nUsage:")
        print("  python main.py <job_description.txt> <resumes_folder_or_zip> [top_n]")
        print("\nExamples:")
        print("  python main.py jd.txt ./resumes/")
        print("  python main.py jd.txt resumes.zip 20\n")
        sys.exit(1)

    jd_path      = sys.argv[1]
    resumes_path = sys.argv[2]
    top_n        = int(sys.argv[3]) if len(sys.argv) == 4 else None

    if not os.path.exists(jd_path):
        print(f"[ERROR] Job description file not found: '{jd_path}'")
        sys.exit(1)
    if not os.path.exists(resumes_path):
        print(f"[ERROR] Resumes source not found: '{resumes_path}'")
        sys.exit(1)

    run_pipeline(jd_path, resumes_path, top_n=top_n)