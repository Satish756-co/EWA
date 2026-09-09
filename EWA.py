#!/usr/bin/env python3
"""
SAP EWA (Early Watch Alert) Report Analyzer
Parses PDF / HTML / DOC / DOCX EWA reports and generates a formatted Excel
action item tracker using only content extracted from the report itself.

Usage:
    python ewa_analyzer.py --file "C:/path/to/EWA_Report.doc"
    python ewa_analyzer.py --file "C:/path/to/EWA_Report.pdf" --sid PRD
    python ewa_analyzer.py --file "C:/path/to/EWA_Report.docx" --output "C:/out/result.xlsx" --sid PRD
"""

import argparse
import os
import re
import sys
from datetime import datetime
from pathlib import Path


# ─────────────────────────────────────────────
# Dependency check
# ─────────────────────────────────────────────
def check_dependencies():
    missing = []
    for pkg, imp in [
        ("pdfplumber",    "pdfplumber"),
        ("beautifulsoup4","bs4"),
        ("openpyxl",      "openpyxl"),
        ("python-docx",   "docx"),
    ]:
        try:
            __import__(imp)
        except ImportError:
            missing.append(pkg)
    if missing:
        print("\n[ERROR] Missing packages. Install them with:")
        print(f"        pip install {' '.join(missing)}\n")
        sys.exit(1)

check_dependencies()

import pdfplumber
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import docx as python_docx


# ─────────────────────────────────────────────
# Classification helpers  (keyword → category / priority)
# These only CLASSIFY text extracted from the real report — they never
# substitute or invent content.
# ─────────────────────────────────────────────

CATEGORY_PATTERNS = {
    "Performance":           [r"response time", r"throughput", r"dialog.*time", r"workload", r"cpu.*utiliz"],
    "Database":              [r"database", r"tablespace", r"table growth", r"\bSQL\b", r"index",
                              r"db statistic", r"oracle", r"hana", r"db size"],
    "Security":              [r"security", r"authoriz", r"password", r"SAP_ALL", r"SAP\*",
                              r"\bDDIC\b", r"firefighter", r"emergency user"],
    "Memory / Buffers":      [r"memory", r"buffer", r"paging", r"swapping",
                              r"extended memory", r"heap", r"roll area", r"\bEM\b"],
    "ABAP Runtime":          [r"short dump", r"ABAP", r"runtime error", r"\bdump\b", r"ST22"],
    "Support Package":       [r"support package", r"\bSP\s+level\b", r"patch level",
                              r"kernel.*patch", r"upgrade"],
    "Background Processing": [r"background job", r"batch job", r"SM37",
                              r"job.*fail", r"job.*cancel", r"job.*abort"],
    "Update Processing":     [r"update.*error", r"update.*terminat", r"SM13", r"update.*fail"],
    "Enqueue / Locks":       [r"enqueue", r"\block\b", r"deadlock", r"SM12"],
    "Output Management":     [r"spool", r"output.*queue", r"print.*fail", r"SP01"],
    "Transport Management":  [r"transport", r"STMS", r"import queue", r"\bCTS\b"],
    "Availability":          [r"availability", r"uptime", r"downtime",
                              r"system.*restart", r"instance.*restart"],
    "Backup & Recovery":     [r"backup.*fail", r"no.*backup", r"backup.*miss",
                              r"restore", r"archive log"],
    "IDoc / Interface":      [r"IDoc", r"interface.*error", r"\bRFC\b.*error",
                              r"middleware", r"\bPI\b", r"\bPO\b"],
    "Workload":              [r"work process", r"dialog.*WP", r"background.*WP", r"RFC.*WP"],
    "System Configuration":  [r"profile parameter", r"instance profile",
                              r"configuration", r"parameter.*change"],
}

PRIORITY_PATTERNS = {
    "HIGH":   [r"\bRED\b", r"\bcritical\b", r"immediately", r"\bsevere\b", r"\burgent\b",
               r"must be corrected", r"action required", r"exceeded.*limit",
               r"not acceptable", r"serious.*issue", r"needs.*immediate"],
    "MEDIUM": [r"\bYELLOW\b", r"\bwarning\b", r"\battention\b", r"\bmonitor\b",
               r"suboptimal", r"should be", r"recommended", r"not optimal",
               r"could be improved", r"consider.*changing"],
    "LOW":    [r"\bGREEN\b", r"\binformation\b", r"\bnote\b", r"\bOK\b",
               r"no action required", r"for information", r"within.*limit",
               r"within.*range", r"satisfactory"],
}

# Headings that start a Recommendation sub-section inside an EWA section
REC_HEADING_RE = re.compile(
    r"^\s*(recommendation[s]?|action[s]?\s+required|suggested action[s]?|"
    r"maßnahme[n]?|proposed\s+action[s]?|next\s+step[s]?)[:\s]*$",
    re.IGNORECASE,
)
# Inline recommendation prefix on the same line  e.g. "Recommendation: ..."
REC_INLINE_RE = re.compile(
    r"^\s*(recommendation[s]?|action[s]?\s+required|suggested action[s]?|"
    r"maßnahme[n]?)[:\s]+(.+)",
    re.IGNORECASE | re.DOTALL,
)

HEADING_STYLES = frozenset([
    "heading 1", "heading 2", "heading 3", "heading 4", "heading 5",
    "überschrift 1", "überschrift 2", "überschrift 3",   # German Word
])


def _categorise(text):
    low = text.lower()
    for cat, pats in CATEGORY_PATTERNS.items():
        if any(re.search(p, low, re.IGNORECASE) for p in pats):
            return cat
    return "General"


def _prioritise(text):
    for pri, pats in PRIORITY_PATTERNS.items():
        if any(re.search(p, text, re.IGNORECASE) for p in pats):
            return pri
    return "MEDIUM"


def _rgb_priority(run):
    """Return HIGH/MEDIUM/LOW if a Word run has a recognisable status colour."""
    try:
        rgb = run.font.color.rgb        # e.g. "FF0000"
        if not rgb:
            return None
        r = int(str(rgb)[0:2], 16)
        g = int(str(rgb)[2:4], 16)
        b = int(str(rgb)[4:6], 16)
        if r > 180 and g < 100 and b < 100:   # red
            return "HIGH"
        if g > 150 and r < 100 and b < 100:   # green
            return "LOW"
        if r > 180 and g > 140 and b < 80:    # yellow / orange
            return "MEDIUM"
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────
# Parsers
# ─────────────────────────────────────────────

class EWADocParser:
    """
    Parses .docx (via python-docx) and .doc (via Word COM / LibreOffice / docx2txt).
    Produces:
        self.full_text  – plain text for regex fallbacks
        self.sections   – list of dicts {heading, description, recommendation,
                          priority, category} extracted from actual document structure
    """

    def __init__(self, filepath):
        self.filepath   = filepath
        self.full_text  = ""
        self.sections   = []

    def parse(self):
        ext = Path(self.filepath).suffix.lower()
        if ext == ".docx":
            self._parse_docx()
        else:
            if not (self._parse_doc_com() or self._parse_doc_docx2txt() or self._parse_doc_libreoffice()):
                self._abort()
        print(f"  [DOC] {len(self.full_text):,} chars, {len(self.sections)} sections extracted")
        return self.full_text

    # ── .docx ─────────────────────────────────────────────────────
    def _parse_docx(self):
        print(f"  [DOCX] Parsing: {self.filepath}")
        doc = python_docx.Document(self.filepath)
        self._build_sections(doc)

    def _build_sections(self, doc):
        """
        Walk every paragraph; group into sections by heading style.
        Within each section collect:
          - description  : body paragraphs BEFORE a Recommendation heading
          - recommendation : paragraphs AFTER a Recommendation heading
        Priority is taken from run colours first, then keyword matching.
        """
        text_parts = []
        sections   = []

        cur = None           # current section dict
        in_rec = False       # are we past the "Recommendation:" heading?

        def flush():
            if cur and (cur["_desc"] or cur["_rec"]):
                heading   = cur["heading"]
                desc_text = " ".join(cur["_desc"]).strip()
                rec_text  = " ".join(cur["_rec"]).strip()
                # If no explicit recommendation block, try pulling one from description body
                if not rec_text:
                    rec_text = _pull_rec_from_text(desc_text)
                full_desc = (heading + ".  " + desc_text).strip() if desc_text else heading
                sections.append({
                    "heading":        heading,
                    "description":    full_desc[:500],
                    "recommendation": rec_text[:500],
                    "priority":       cur["priority"],
                    "category":       cur["category"],
                })

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
            text_parts.append(text)

            style      = (para.style.name or "").lower()
            is_heading = any(style.startswith(h) for h in HEADING_STYLES)

            # Treat short all-bold lines as headings too
            if not is_heading and len(text) < 130:
                runs = [r for r in para.runs if r.text.strip()]
                if runs and all(r.bold for r in runs):
                    is_heading = True

            if is_heading:
                flush()
                cur    = {"heading": text, "_desc": [], "_rec": [],
                          "priority": "MEDIUM", "category": _categorise(text)}
                in_rec = False
                # Priority from heading colour
                for run in para.runs:
                    rp = _rgb_priority(run)
                    if rp:
                        cur["priority"] = rp
                        break
                # Priority from heading text keywords
                kp = _prioritise(text)
                if kp == "HIGH":
                    cur["priority"] = "HIGH"
                continue

            if cur is None:
                continue

            # ── Detect start of Recommendation block ──────────────
            if REC_HEADING_RE.match(text):
                in_rec = True
                continue

            inline = REC_INLINE_RE.match(text)
            if inline:
                in_rec = True
                cur["_rec"].append(inline.group(2).strip())
                continue

            # ── Accumulate colour-based priority from body runs ───
            for run in para.runs:
                rp = _rgb_priority(run)
                if rp == "HIGH":
                    cur["priority"] = "HIGH"
                    break
                if rp == "MEDIUM" and cur["priority"] != "HIGH":
                    cur["priority"] = "MEDIUM"

            # ── Priority from body text keywords ──────────────────
            body_pri = _prioritise(text)
            if body_pri == "HIGH" and cur["priority"] != "HIGH":
                cur["priority"] = "HIGH"

            if in_rec:
                cur["_rec"].append(text)
            else:
                cur["_desc"].append(text)

        flush()

        # Tables — add cell text to full_text only (tables rarely hold the
        # structured finding/recommendation pattern)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    ct = cell.text.strip()
                    if ct:
                        text_parts.append(ct)

        self.full_text = "\n".join(text_parts)
        self.sections  = sections

    # ── .doc via Windows COM ──────────────────────────────────────
    def _parse_doc_com(self):
        try:
            import win32com.client, pythoncom
        except ImportError:
            print("  [DOC] pywin32 not installed  →  trying docx2txt")
            print("        (install later: pip install pywin32)")
            return False

        abs_path = str(Path(self.filepath).resolve())

        # Remove internet zone mark → prevents Protected View
        try:
            import subprocess
            subprocess.run(
                ["powershell", "-Command", f'Unblock-File -LiteralPath "{abs_path}"'],
                capture_output=True, timeout=15,
            )
            print("  [DOC] Unblocked file (removed internet zone mark)")
        except Exception:
            pass

        try:
            print(f"  [DOC] Using Word COM: {abs_path}")
            pythoncom.CoInitialize()
            word             = win32com.client.Dispatch("Word.Application")
            word.Visible     = False
            word.DisplayAlerts = False

            doc = None
            try:
                doc = word.Documents.Open(abs_path, False, True, False)
            except Exception:
                print("  [DOC] Normal open landed in Protected View — promoting…")
                for pvw in word.ProtectedViewWindows:
                    if abs_path.lower() in str(pvw.Document.FullName).lower():
                        doc = pvw.Edit()
                        break

            if doc is None:
                raise RuntimeError("Could not obtain Word Document object")

            self.full_text = doc.Content.Text
            doc.Close(False)
            word.Quit()
            pythoncom.CoUninitialize()

            # Now re-parse as docx by saving to temp, to get structured sections
            try:
                import tempfile
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
                    tmp = tf.name
                pythoncom.CoInitialize()
                word2 = win32com.client.Dispatch("Word.Application")
                word2.Visible = False
                word2.DisplayAlerts = False
                d2 = word2.Documents.Open(abs_path, False, True, False)
                d2.SaveAs2(tmp, 16)   # 16 = wdFormatXMLDocument (.docx)
                d2.Close(False)
                word2.Quit()
                pythoncom.CoUninitialize()
                saved_path    = self.filepath
                self.filepath = tmp
                self._parse_docx()
                self.filepath = saved_path
                Path(tmp).unlink(missing_ok=True)
            except Exception as e:
                print(f"  [DOC] Structured re-parse skipped ({e}) — using plain text")

            return True

        except Exception as e:
            print(f"  [DOC] COM failed ({e})")
            try:
                word.Quit()
                pythoncom.CoUninitialize()
            except Exception:
                pass
            return False

    # ── .doc fallback: docx2txt ───────────────────────────────────
    def _parse_doc_docx2txt(self):
        try:
            import docx2txt
            print(f"  [DOC] Using docx2txt: {self.filepath}")
            self.full_text = docx2txt.process(self.filepath) or ""
            return len(self.full_text.strip()) > 50
        except ImportError:
            print("  [DOC] docx2txt not installed  →  pip install docx2txt")
        except Exception as e:
            print(f"  [DOC] docx2txt failed ({e})")
        return False

    # ── .doc fallback: LibreOffice ────────────────────────────────
    def _parse_doc_libreoffice(self):
        import subprocess, tempfile, shutil
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            return False
        try:
            print(f"  [DOC] LibreOffice convert: {self.filepath}")
            with tempfile.TemporaryDirectory() as td:
                subprocess.run(
                    [soffice, "--headless", "--convert-to", "docx", "--outdir", td, self.filepath],
                    check=True, capture_output=True,
                )
                converted = Path(td) / f"{Path(self.filepath).stem}.docx"
                if converted.exists():
                    saved = self.filepath
                    self.filepath = str(converted)
                    self._parse_docx()
                    self.filepath = saved
                    return True
        except Exception as e:
            print(f"  [DOC] LibreOffice failed ({e})")
        return False

    def _abort(self):
        print("\n" + "="*62)
        print("  [ERROR] Could not read the .doc file.  Options:")
        print("    pip install pywin32     (Word COM — recommended on Windows)")
        print("    pip install docx2txt   (no Word required)")
        print("    Open in Word → Save As .docx or .pdf and re-run")
        print("="*62 + "\n")
        sys.exit(1)

    def extract_date(self):
        return _extract_date(self.full_text[:4000])


# ─────────────────────────────────────────────
class EWAPDFParser:
    def __init__(self, filepath):
        self.filepath  = filepath
        self.full_text = ""
        self.sections  = []

    def parse(self):
        print(f"  [PDF] Parsing: {self.filepath}")
        with pdfplumber.open(self.filepath) as pdf:
            pages = len(pdf.pages)
            for page in pdf.pages:
                self.full_text += (page.extract_text() or "") + "\n"
        print(f"  [PDF] {pages} pages, {len(self.full_text):,} chars")
        return self.full_text

    def extract_date(self):
        return _extract_date(self.full_text[:4000])


# ─────────────────────────────────────────────
class EWAHTMLParser:
    def __init__(self, filepath):
        self.filepath  = filepath
        self.full_text = ""
        self.sections  = []

    def parse(self):
        print(f"  [HTML] Parsing: {self.filepath}")
        with open(self.filepath, "r", encoding="utf-8", errors="ignore") as fh:
            soup = BeautifulSoup(fh.read(), "html.parser")
        self.full_text = soup.get_text(separator="\n")
        print(f"  [HTML] {len(self.full_text):,} chars")
        return self.full_text

    def extract_date(self):
        return _extract_date(self.full_text[:4000])


# ─────────────────────────────────────────────
def _extract_date(text):
    for p in [
        r"Report\s+Date[:\s]+(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"Generated\s+on[:\s]+(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"EarlyWatch\s+Alert.*?(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"Date[:\s]+(\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4})",
        r"(\d{1,2}\.\d{1,2}\.\d{4})",
        r"(\d{4}-\d{2}-\d{2})",
        r"(\w+\s+\d{1,2},?\s+\d{4})",
    ]:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return datetime.now().strftime("%d.%m.%Y")


def _pull_rec_from_text(text):
    """
    Try to extract a recommendation sentence from body text.
    Returns the actual sentence found, or empty string.
    """
    for pat in [
        r"recommend[ation]*s?[:\s]+(.{20,400}?)(?:\.|$)",
        r"action[s]?\s+required[:\s]+(.{20,400}?)(?:\.|$)",
        r"should\s+(.{20,300}?)(?:\.|$)",
        r"please\s+(.{20,300}?)(?:\.|$)",
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return re.sub(r"\s+", " ", m.group(1)).strip()
    return ""


# ─────────────────────────────────────────────
# Finding Extractor — uses ONLY real report content
# ─────────────────────────────────────────────

def extract_findings(text, sections=None):
    """
    Build the findings list exclusively from content in the real document.
    No pre-written descriptions or recommendations exist in this function.

    Priority 1 : structured sections (Word paragraph styles + colours)
    Priority 2 : RED / YELLOW / GREEN rating lines
    Priority 3 : Recommendation / Action Required text blocks
    Priority 4 : Numbered section headings with body text
    """
    findings  = []
    seen_keys = set()

    def add(f):
        key = f["description"][:50].lower().strip()
        if key and key not in seen_keys and len(f["description"].strip()) > 15:
            seen_keys.add(key)
            findings.append(f)

    # ── 1. Structured Word sections ───────────────────────────────
    if sections:
        for s in sections:
            pri = s.get("priority", "MEDIUM")
            # Skip sections that are clearly just GREEN / informational with no action
            if pri == "LOW" and not s.get("recommendation"):
                continue
            add({
                "category":       s.get("category") or _categorise(s.get("heading", "")),
                "priority":       pri,
                "description":    s.get("description", s.get("heading", ""))[:400],
                "recommendation": s.get("recommendation", "")[:400],
                "source":         "Document Section",
            })
        print(f"  Structured sections → {len(findings)} findings")

    # ── 2. Inline RED / YELLOW / GREEN rating lines ───────────────
    rating_re = re.compile(
        r"(RED|YELLOW|GREEN)\s*[:\-–]\s*(.{20,350}?)(?=\n\n|RED\s*[:\-–]|YELLOW\s*[:\-–]|GREEN\s*[:\-–]|\Z)",
        re.IGNORECASE | re.DOTALL,
    )
    for m in rating_re.finditer(text):
        rating   = m.group(1).upper()
        content  = re.sub(r"\s+", " ", m.group(2)).strip()
        priority = {"RED": "HIGH", "YELLOW": "MEDIUM", "GREEN": "LOW"}[rating]
        if priority == "LOW":
            continue
        after = text[m.end(): m.end() + 600]
        rec   = _pull_rec_from_text(content + " " + after)
        add({
            "category":       _categorise(content),
            "priority":       priority,
            "description":    content[:350],
            "recommendation": rec,
            "source":         f"Rating:{rating}",
        })

    # ── 3. Recommendation / Action Required blocks ────────────────
    rec_re = re.compile(
        r"(?:Recommendation[s]?|Suggested Action[s]?|Action[s]?\s+Required)"
        r"[:\s]+(.{30,600}?)(?=\n\n|\n[A-Z][A-Z]|\Z)",
        re.IGNORECASE | re.DOTALL,
    )
    for m in rec_re.finditer(text):
        rec_text    = re.sub(r"\s+", " ", m.group(1)).strip()
        ctx_start   = max(0, m.start() - 500)
        context     = text[ctx_start: m.start()]
        priority    = _prioritise(context + rec_text)
        category    = _categorise(context + rec_text)
        desc_lines  = [l.strip() for l in context.split("\n") if len(l.strip()) > 20]
        description = desc_lines[-1][:350] if desc_lines else context[-180:].strip()
        add({
            "category":       category,
            "priority":       priority,
            "description":    description,
            "recommendation": rec_text[:400],
            "source":         "Recommendation block",
        })

    # ── 4. Numbered section headings (e.g. "2.3 Database Activity") ──
    section_re = re.compile(
        r"^(\d+[\.\d]*\s+[A-Z][^\n]{10,90})\n((?:.{15,}\n?){1,10})",
        re.MULTILINE,
    )
    for m in section_re.finditer(text):
        heading  = m.group(1).strip()
        body     = re.sub(r"\s+", " ", m.group(2)).strip()
        priority = _prioritise(heading + " " + body)
        if priority == "LOW":
            continue
        rec = _pull_rec_from_text(body)
        add({
            "category":       _categorise(heading + " " + body),
            "priority":       priority,
            "description":    (heading + ".  " + body[:250]).strip(),
            "recommendation": rec,
            "source":         "Section heading",
        })

    print(f"  Total unique findings: {len(findings)}")
    return findings


# ─────────────────────────────────────────────
# Excel Generator
# ─────────────────────────────────────────────

PALETTE = {
    "header_bg":  "1F4E79",
    "header_fg":  "FFFFFF",
    "high_bg":    "FFCCCC",
    "medium_bg":  "FFEB99",
    "low_bg":     "CCFFCC",
    "high_txt":   "CC0000",
    "medium_txt": "7F4F00",
    "low_txt":    "1A6B1A",
    "alt_row":    "EEF4FB",
    "open_bg":    "FFE699",
    "section_bg": "2E75B6",
    "border":     "B8CCE4",
}

def _fill(h):
    return PatternFill(start_color=h, end_color=h, fill_type="solid")

def _border():
    t = Side(style="thin", color=PALETTE["border"])
    return Border(left=t, right=t, top=t, bottom=t)

def _font(bold=False, size=10, color="000000", name="Calibri"):
    return Font(name=name, bold=bold, size=size, color=color)

def _align(h="left", v="top", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def build_excel(findings, report_date, source_file, output_path, sid=""):
    wb = openpyxl.Workbook()

    # ── Sheet 1 : Action Items ────────────────────────────────────
    ws = wb.active
    ws.title = "Action Items"

    COLUMNS = [
        ("No.",                        5),
        ("EWA Report Date",           14),
        ("System / SID",              12),
        ("Priority",                  10),
        ("Type / Category",           22),
        ("Finding / Description",     50),
        ("SAP Recommendation",        50),
        ("Current Status",            13),
        ("Analysis & Action Remark",  42),
        ("Assigned To",               16),
        ("Target Date",               14),
        ("Completed Date",            14),
        ("Remarks",                   22),
    ]

    for col, (label, width) in enumerate(COLUMNS, 1):
        c = ws.cell(row=1, column=col, value=label)
        c.fill      = _fill(PALETTE["header_bg"])
        c.font      = _font(bold=True, size=11, color=PALETTE["header_fg"])
        c.alignment = _align("center", "center")
        c.border    = _border()
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 32

    pri_order  = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}
    sorted_f   = sorted(findings, key=lambda x: pri_order.get(x.get("priority", "MEDIUM"), 1))

    status_dv  = DataValidation(
        type="list",
        formula1='"Open,In Progress,Closed,N/A,Deferred"',
        allow_blank=True,
    )
    ws.add_data_validation(status_dv)

    for idx, f in enumerate(sorted_f, 1):
        row    = idx + 1
        pri    = f.get("priority", "MEDIUM")
        row_bg = PALETTE["alt_row"] if idx % 2 == 0 else "FFFFFF"

        row_data = [
            idx, report_date, sid or "", pri,
            f.get("category", "General"),
            f.get("description", ""),
            f.get("recommendation", ""),
            "Open", "", "", "", "", "",
        ]

        for col, val in enumerate(row_data, 1):
            c = ws.cell(row=row, column=col, value=val)
            c.border = _border()
            c.font   = _font(size=10)
            if col in (1, 2, 3):
                c.alignment = _align("center", "center", False)
                c.fill      = _fill(row_bg)
            elif col == 4:
                bg  = {"HIGH": PALETTE["high_bg"],  "MEDIUM": PALETTE["medium_bg"],
                       "LOW":  PALETTE["low_bg"]}.get(pri, "FFFFFF")
                txt = {"HIGH": PALETTE["high_txt"],  "MEDIUM": PALETTE["medium_txt"],
                       "LOW":  PALETTE["low_txt"]}.get(pri, "000000")
                c.fill      = _fill(bg)
                c.font      = _font(bold=True, size=10, color=txt)
                c.alignment = _align("center", "center", False)
            elif col == 8:
                c.fill      = _fill(PALETTE["open_bg"])
                c.font      = _font(bold=True, size=10)
                c.alignment = _align("center", "center", False)
                status_dv.add(c)
            elif col == 9:
                c.fill      = _fill("FFFFF0")
                c.alignment = _align("left", "top", True)
            else:
                c.fill      = _fill(row_bg)
                c.alignment = _align("left", "top", True)

        ws.row_dimensions[row].height = 50

    ws.auto_filter.ref  = f"A1:{get_column_letter(len(COLUMNS))}{len(sorted_f)+1}"
    ws.freeze_panes     = "A2"
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1

    # ── Sheet 2 : Summary ─────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:G1")
    t           = ws2["A1"]
    t.value     = "SAP EWA Report — Analysis Summary"
    t.fill      = _fill(PALETTE["header_bg"])
    t.font      = _font(bold=True, size=15, color="FFFFFF")
    t.alignment = _align("center", "center", False)
    ws2.row_dimensions[1].height = 40

    for i, (label, val) in enumerate([
        ("EWA Report Date", report_date),
        ("Source File",     os.path.basename(source_file)),
        ("System / SID",    sid or "—"),
        ("Analysis Run At", datetime.now().strftime("%Y-%m-%d  %H:%M")),
        ("Total Findings",  len(findings)),
    ], 3):
        ws2.cell(row=i, column=1, value=label).font = _font(bold=True, size=11)
        ws2.cell(row=i, column=2, value=val).font   = _font(size=11)

    _section_hdr(ws2, 9, "PRIORITY BREAKDOWN")
    for col, h in enumerate(["Priority", "Count", "% of Total"], 1):
        c = ws2.cell(row=10, column=col, value=h)
        c.fill = _fill("BDD7EE"); c.font = _font(bold=True, size=10)
        c.border = _border(); c.alignment = _align("center", "center", False)

    counts = {"HIGH": 0, "MEDIUM": 0, "LOW": 0}
    for f in findings:
        counts[f.get("priority", "MEDIUM")] = counts.get(f.get("priority", "MEDIUM"), 0) + 1
    total = len(findings) or 1
    for i, (p, cnt) in enumerate(counts.items(), 11):
        ws2.cell(row=i, column=1, value=p).fill = _fill(
            {"HIGH": PALETTE["high_bg"], "MEDIUM": PALETTE["medium_bg"], "LOW": PALETTE["low_bg"]}[p])
        ws2.cell(row=i, column=2, value=cnt)
        ws2.cell(row=i, column=3, value=f"{cnt/total*100:.1f}%")
        for col in range(1, 4):
            ws2.cell(row=i, column=col).border    = _border()
            ws2.cell(row=i, column=col).alignment = _align("center", "center", False)
            ws2.cell(row=i, column=col).font      = _font(size=10)

    cats = {}
    for f in findings:
        cats[f.get("category", "General")] = cats.get(f.get("category", "General"), 0) + 1
    _section_hdr(ws2, 16, "CATEGORY BREAKDOWN")
    for col, h in enumerate(["Category", "Count", "% of Total"], 1):
        c = ws2.cell(row=17, column=col, value=h)
        c.fill = _fill("BDD7EE"); c.font = _font(bold=True, size=10)
        c.border = _border(); c.alignment = _align("center", "center", False)
    for i, (cat, cnt) in enumerate(sorted(cats.items(), key=lambda x: -x[1]), 18):
        bg = PALETTE["alt_row"] if i % 2 == 0 else "FFFFFF"
        for col, v in enumerate([cat, cnt, f"{cnt/total*100:.1f}%"], 1):
            c = ws2.cell(row=i, column=col, value=v)
            c.fill = _fill(bg); c.border = _border()
            c.alignment = _align("center", "center", False); c.font = _font(size=10)

    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 14

    # ── Sheet 3 : EWA Checklist ───────────────────────────────────
    ws3 = wb.create_sheet("EWA Checklist")
    chk_hdrs   = ["#","Check Area","Check Item","Expected Threshold",
                   "Actual Finding","Status","Action Required"]
    chk_widths = [4, 22, 36, 28, 32, 12, 36]
    for col, (h, w) in enumerate(zip(chk_hdrs, chk_widths), 1):
        c = ws3.cell(row=1, column=col, value=h)
        c.fill = _fill(PALETTE["header_bg"]); c.font = _font(bold=True, size=11, color="FFFFFF")
        c.alignment = _align("center", "center"); c.border = _border()
        ws3.column_dimensions[get_column_letter(col)].width = w
    ws3.row_dimensions[1].height = 30

    checklist = [
        ("Security","SAP* default password changed","Changed","","Check",""),
        ("Security","DDIC default password changed","Changed","","Check",""),
        ("Security","EARLYWATCH password changed","Changed","","Check",""),
        ("Security","No user assigned SAP_ALL in PRD","0 users","","Check",""),
        ("Security","Password policy active","Minimum complexity","","Check",""),
        ("Security","Inactive users locked/deleted","0 inactive > 90d","","Check",""),
        ("Performance","Dialog avg response time","< 1,000 ms","","Check",""),
        ("Performance","DB avg response time","< 40 ms","","Check",""),
        ("Performance","CPU utilization (avg)","< 70%","","Check",""),
        ("Performance","Memory paging / swapping","None","","Check",""),
        ("Database","Tablespace utilization","< 85%","","Check",""),
        ("Database","DB statistics current","Updated < 7 days","","Check",""),
        ("Database","Daily DB backup successful","Green","","Check",""),
        ("Database","No missing or invalid indexes","0","","Check",""),
        ("Memory","Extended memory not exhausted","No exhaustion","","Check",""),
        ("Memory","Buffer hit ratios (table, NTAB, prog)","> 95%","","Check",""),
        ("ABAP","Short dump count (last 7 days)","< 100","","Check",""),
        ("Updates","Update terminations in SM13","0","","Check",""),
        ("Batch Jobs","Failed critical background jobs","0","","Check",""),
        ("Availability","System availability","> 99.5%","","Check",""),
        ("Sp. Package","SP stack level","Within 2 SP","","Check",""),
        ("Sp. Package","Kernel patch level","< 6 months old","","Check",""),
        ("Interfaces","IDoc error rate","< 1%","","Check",""),
        ("Interfaces","RFC connection errors","0 critical","","Check",""),
    ]
    for i, row_data in enumerate(checklist, 2):
        bg = PALETTE["alt_row"] if i % 2 == 0 else "FFFFFF"
        ws3.cell(row=i, column=1, value=i-1)
        for col, val in enumerate(row_data, 2):
            c = ws3.cell(row=i, column=col, value=val)
            c.fill = _fill(bg); c.border = _border()
            c.font = _font(size=10); c.alignment = _align("left", "center", False)
        ws3.row_dimensions[i].height = 18
    ws3.auto_filter.ref = f"A1:{get_column_letter(len(chk_hdrs))}1"
    ws3.freeze_panes = "A2"

    wb.save(output_path)
    return output_path


def _section_hdr(ws, row, title):
    ws.merge_cells(f"A{row}:C{row}")
    c = ws.cell(row=row, column=1, value=title)
    c.fill = _fill(PALETTE["section_bg"]); c.font = _font(bold=True, size=11, color="FFFFFF")
    c.alignment = _align("center", "center", False)
    ws.row_dimensions[row].height = 22


# ─────────────────────────────────────────────
# Orchestrator
# ─────────────────────────────────────────────

def analyze(filepath, output_path=None, sid=""):
    filepath = Path(filepath)
    if not filepath.exists():
        print(f"\n[ERROR] File not found: {filepath}")
        sys.exit(1)

    ext = filepath.suffix.lower()
    if ext == ".pdf":
        parser = EWAPDFParser(str(filepath))
    elif ext in (".html", ".htm"):
        parser = EWAHTMLParser(str(filepath))
    elif ext in (".doc", ".docx"):
        parser = EWADocParser(str(filepath))
    else:
        print(f"\n[ERROR] Unsupported type '{ext}'.  Supported: .pdf .html .htm .doc .docx")
        sys.exit(1)

    print("\n" + "="*62)
    print("  SAP EWA Report Analyzer")
    print("="*62)
    print(f"  File : {filepath}")

    text        = parser.parse()
    report_date = parser.extract_date()
    sections    = getattr(parser, "sections", [])
    print(f"  Date : {report_date}")
    if sid:
        print(f"  SID  : {sid}")

    findings = extract_findings(text, sections)

    if not findings:
        print("\n  [WARN] No findings detected — creating blank 20-row template.")
        findings = [{"category": "", "priority": "MEDIUM",
                     "description": "", "recommendation": "", "source": "Manual"}] * 20

    findings.sort(key=lambda x: {"HIGH": 0, "MEDIUM": 1, "LOW": 2}.get(x.get("priority", "MEDIUM"), 1))

    if not output_path:
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        output_path = filepath.parent / f"{filepath.stem}_EWA_Analysis_{ts}.xlsx"

    print(f"\n  Generating Excel…")
    saved  = build_excel(findings, report_date, str(filepath), str(output_path), sid)
    high   = sum(1 for f in findings if f.get("priority") == "HIGH")
    medium = sum(1 for f in findings if f.get("priority") == "MEDIUM")
    low    = sum(1 for f in findings if f.get("priority") == "LOW")

    print(f"\n{'='*62}")
    print(f"  ANALYSIS COMPLETE")
    print(f"  Total Findings : {len(findings)}")
    print(f"  HIGH  (Red)    : {high}")
    print(f"  MEDIUM (Yellow): {medium}")
    print(f"  LOW   (Green)  : {low}")
    print(f"  Output Excel   : {saved}")
    print(f"{'='*62}\n")
    return saved


# ─────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(
        description="SAP EWA Analyzer — extracts real report content → Excel action tracker",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python ewa_analyzer.py --file "C:/EWA/report.doc"
  python ewa_analyzer.py --file "C:/EWA/report.docx"  --sid PRD
  python ewa_analyzer.py --file "C:/EWA/report.pdf"   --output "C:/out/EWA_PRD.xlsx" --sid PRD
        """,
    )
    ap.add_argument("--file",   required=True)
    ap.add_argument("--output", default=None)
    ap.add_argument("--sid",    default="")
    args = ap.parse_args()
    analyze(args.file, args.output, args.sid)


if __name__ == "__main__":
    main()
