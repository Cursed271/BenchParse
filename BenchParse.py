# ----- License -------------------------------------------------- # 

#  BenchParse - BenchParse extracts CIS Benchmark recommendations from 
#  PDF documents and creates structured Excel workbooks containing 
#  control details, assessment fields, and summaries.
#  Copyright (c) 2026 - CursedSec (Operated by Cursed271). All rights reserved.

#  This software is an proprietary intellectual property developed for
#  penetration testing, threat modeling, and security research. It   
#  is licensed under the CURSEDSEC OWNERSHIP EDICT:
#
#  🚫 PROHIBITION WARNING 🚫
#  Redistribution, re-uploading, and unauthorized modification are strictly forbidden 
#  under the COE. Use is granted ONLY under the limited terms defined in the official 
#  LICENSE file (COE), which must be included in all copies.

#  DISCLAIMER:
#  This tool is intended for **educational or ethical testing** purposes only.
#  Unauthorized or malicious use of this software against systems without 
#  proper authorization is strictly prohibited and may violate laws and regulations.
#  The author assumes no liability for misuse or damage caused by this tool.

#  🔗 LICENSE: CURSEDSEC OWNERSHIP EDICT (COE)
#  🔗 Repository: https://github.com/Cursed271
#  🔗 Author: Steven Pereira (@Cursed271)

# ----- Libraries ------------------------------------------------ #

import os
import re
import sys
import typer
import tempfile
import pdfplumber
from pathlib import Path
from rich.rule import Rule
from rich.text import Text
from rich.panel import Panel
from openpyxl import Workbook
from rich.prompt import Prompt
from rich.console import Console
import typer.rich_utils as typer_rich
from typing import Dict, List, Optional
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn

# ----- Global Declaration --------------------------------------- #

console = Console()
CIS_ID_START_REGEX = re.compile(r"^(\d+(?:\.\d+)+)(?:\s+(.*))?$", re.DOTALL)
ASSESS_TAG_REGEX = re.compile(r"(?:\(\s*(Automated|Manual)\s*\)|(?<!\w)(Automated|Manual)(?!\w))", re.IGNORECASE)
LEVEL_TAG_REGEX = re.compile(r"(?:\(\s*(L[12]|Level\s*[12])\s*\)|(?<!\w)(L[12]|Level\s*[12])(?!\w))", re.IGNORECASE)

app = typer.Typer(
    rich_markup_mode = "rich",
    add_completion = False,
    context_settings = {"help_option_names": ["-h", "--help"]}
)

HEADER_FOOTER_PATTERNS = [
    re.compile(r"^Page\s+\d+\s*$", re.IGNORECASE),
    re.compile(r"^CIS\s+.*?Benchmark\s*$", re.IGNORECASE),
    re.compile(r"^https?://.*?cisecurity\.org", re.IGNORECASE),
]

SECTION_HEADERS = [
    "Profile Applicability:",
    "Description:",
    "Rationale:",
    "Impact:",
    "Audit:",
    "Remediation:",
    "Default Value:",
    "References:",
    "CIS Controls:"
]

# ----- PDF Path Validation -------------------------------------- #

def pdf_clean_path(value: str) -> Path:
    cleaned = value.strip()
    if len(cleaned) >= 2 and cleaned[0] == cleaned[-1] and cleaned[0] in ("'", '"'):
        cleaned = cleaned[1:-1].strip()
    if not cleaned:
        raise ValueError("A file path is required.")
    return Path(cleaned).expanduser().resolve()

def pdf_validate(value: str) -> Path:
    try:
        path = pdf_clean_path(value)
        if not path.is_file():
            raise ValueError("The input must be an existing PDF file.")
        if path.suffix.lower() != ".pdf":
            raise ValueError("The input file must use the .pdf extension.")
        if not os.access(path, os.R_OK):
            raise ValueError("The input PDF is not readable.")
        return path
    except (ValueError, OSError, RuntimeError) as exc:
        raise typer.BadParameter(str(exc), param_hint = "--input / -i") from exc

def pdf_output(value: str, input_path: Path) -> Path:
    try:
        path = pdf_clean_path(value)
        if path == input_path:
            raise ValueError("The output cannot replace the input PDF.")
        if path.suffix.lower() != ".xlsx":
            raise ValueError("The output file must use the .xlsx extension.")
        if path.exists() and not path.is_file():
            raise ValueError("The output path is not a regular file.")
        return path
    except (ValueError, OSError, RuntimeError) as exc:
        raise typer.BadParameter(str(exc), param_hint = "--output / -o") from exc

# ----- PDF Extraction ------------------------------------------- #

def pdf_clean_lines(page) -> List[str]:
    text = page.extract_text(layout = True)
    if not text:
        return []
    cleaned = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if any(pattern.match(stripped) for pattern in HEADER_FOOTER_PATTERNS):
            continue
        cleaned.append(stripped)
    return cleaned

def pdf_parse(pdf_path: str) -> List[Dict[str, str]]:
    recommendations = []
    current_rec = None
    current_section = None
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        console.print(f"[#BB9AF7][~] Opened PDF with {total_pages} pages. Analyzing layout...")
        with Progress(
            SpinnerColumn(),
            TextColumn("[progress.description]{task.description}"),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TimeElapsedColumn(),
            console = console
        ) as progress:
            task = progress.add_task("[#C6ECE3]Parsing CIS recommendations...", total = total_pages)
            for page in pdf.pages:
                lines = pdf_clean_lines(page)
                line_idx = 0
                while line_idx < len(lines):
                    line = lines[line_idx]
                    if "Appendix: Summary Table" in line or "Appendix: Change Log" in line:
                        break
                    id_match = CIS_ID_START_REGEX.match(line)
                    is_real_control = False
                    if id_match:
                        lookahead_limit = min(line_idx + 8, len(lines))
                        for k in range(line_idx + 1, lookahead_limit):
                            if lines[k].strip().startswith("Profile Applicability:"):
                                is_real_control = True
                                break
                            if CIS_ID_START_REGEX.match(lines[k].strip()):
                                break
                    if is_real_control:
                        if current_rec:
                            recommendations.append(current_rec)
                        control_id = id_match.group(1).strip()
                        initial_title = (id_match.group(2) or "").strip()
                        title_lines = [initial_title] if initial_title else []
                        peek_idx = line_idx + 1
                        while peek_idx < len(lines):
                            peek_line = lines[peek_idx].strip()
                            if any(peek_line.startswith(section) for section in SECTION_HEADERS):
                                break
                            title_lines.append(peek_line)
                            peek_idx += 1
                        raw_full_title = " ".join(title_lines).strip()
                        # Preserve the original defaults when title tags are absent.
                        level = "Level 1"
                        level_match = LEVEL_TAG_REGEX.search(raw_full_title)
                        if level_match:
                            value = (level_match.group(1) or level_match.group(2) or "").upper()
                            level = "Level 2" if "2" in value else "Level 1"
                            raw_full_title = raw_full_title[:level_match.start()] + raw_full_title[level_match.end():]
                        assessment = "Automated"
                        assessment_match = ASSESS_TAG_REGEX.search(raw_full_title)
                        if assessment_match:
                            assessment = (assessment_match.group(1) or assessment_match.group(2) or "").title()
                            raw_full_title = raw_full_title[:assessment_match.start()] + raw_full_title[assessment_match.end():]
                        clean_title = re.sub(r"\s+", " ", raw_full_title).strip(" -:–()")
                        current_rec = {
                            "ID": control_id,
                            "Title": clean_title,
                            "Level": level,
                            "Assessment": assessment,
                            "Profile Applicability": "",
                            "Description": "",
                            "Rationale": "",
                            "Impact": "",
                            "Audit": "",
                            "Remediation": "",
                            "Default Value": "",
                        }
                        current_section = None
                        line_idx = peek_idx
                        continue
                    if not current_rec:
                        line_idx += 1
                        continue
                    matched_section = None
                    for section in SECTION_HEADERS:
                        if line.startswith(section):
                            matched_section = section.rstrip(":")
                            remainder = line[len(section):].strip()
                            current_section = matched_section
                            if remainder:
                                if current_rec.get(current_section):
                                    current_rec[current_section] += "\n" + remainder
                                else:
                                    current_rec[current_section] = remainder
                            break
                    if matched_section:
                        line_idx += 1
                        continue
                    if current_section and current_section in current_rec:
                        if current_rec[current_section]:
                            current_rec[current_section] += "\n" + line
                        else:
                            current_rec[current_section] = line
                    line_idx += 1
                progress.update(task, advance = 1)
    if current_rec:
        recommendations.append(current_rec)
    seen = set()
    deduped = []
    for recommendation in recommendations:
        key = (recommendation["ID"], recommendation["Title"])
        if key not in seen:
            seen.add(key)
            deduped.append(recommendation)
    return deduped

# ----- Excel Report --------------------------------------------- #

def save_cis_to_excel(recommendations: List[Dict[str, str]], output_path: Path, overwrite: bool = False) -> None:
    output_path = Path(output_path)
    if output_path.exists() and not overwrite:
        raise FileExistsError("Output already exists. Use --overwrite to replace it.")
    wb = Workbook()
    ws = wb.active
    ws.title = "CIS Compliance Matrix"
    ws.views.sheetView[0].showGridLines = True
    navy_header_fill = PatternFill(start_color = "1E293B", end_color = "1E293B", fill_type = "solid")
    header_font = Font(name = "Segoe UI", size = 11, bold = True, color = "FFFFFF")
    data_font = Font(name = "Segoe UI", size = 10, color = "0F172A")
    title_font = Font(name = "Segoe UI", size = 10, bold = True, color = "0F172A")
    code_font = Font(name = "Consolas", size = 9, color = "0F172A")
    id_font = Font(name = "Segoe UI", size = 10, bold = True, color = "0284C7")
    thin_border = Border(
        left = Side(style = "thin", color = "E2E8F0"),
        right = Side(style = "thin", color = "E2E8F0"),
        top = Side(style = "thin", color = "E2E8F0"),
        bottom = Side(style = "thin", color = "E2E8F0")
    )
    columns = [
        ("CIS ID", 12, "center", id_font),
        ("Control Title", 45, "left", title_font),
        ("Level", 12, "center", data_font),
        ("Assessment", 14, "center", data_font),
        ("Status", 16, "center", Font(name = "Segoe UI", size = 10, bold = True)),
        ("Owner / Team", 16, "left", data_font),
        ("Description", 45, "left", data_font),
        ("Rationale", 40, "left", data_font),
        ("Impact", 35, "left", data_font),
        ("Audit Procedure", 55, "left", code_font),
        ("Remediation Steps", 55, "left", code_font),
        ("Default Value", 30, "left", data_font),
        ("Auditor Notes & Artifacts", 30, "left", data_font),
    ]
    ws.append([column[0] for column in columns])
    for col_idx in range(1, len(columns) + 1):
        cell = ws.cell(row = 1, column = col_idx)
        cell.fill = navy_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal = "center", vertical = "center", wrap_text = True)
    ws.row_dimensions[1].height = 28
    status_validation = DataValidation(
        type = "list",
        formula1 = '"Compliant,Non-Compliant,Exception,In Progress,Not Applicable"',
        allow_blank = True
    )
    status_validation.error = "Select a valid compliance status from the dropdown."
    status_validation.errorTitle = "Invalid Status"
    status_validation.showErrorMessage = True
    status_validation.errorStyle = "stop"
    ws.add_data_validation(status_validation)
    for row_idx, recommendation in enumerate(recommendations, start = 2):
        row_data = [
            recommendation.get("ID", ""),
            recommendation.get("Title", ""),
            recommendation.get("Level", ""),
            recommendation.get("Assessment", ""),
            "In Progress",
            "",
            recommendation.get("Description", "").strip(),
            recommendation.get("Rationale", "").strip(),
            recommendation.get("Impact", "").strip(),
            recommendation.get("Audit", "").strip(),
            recommendation.get("Remediation", "").strip(),
            recommendation.get("Default Value", "").strip(),
            ""
        ]
        row_data = [ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value for value in row_data]
        ws.append(row_data)
        for col_idx, column in enumerate(columns, start = 1):
            cell = ws.cell(row = row_idx, column = col_idx)
            # Extracted PDF content is text, not an Excel formula.
            if isinstance(cell.value, str):
                cell.data_type = "s"
            cell.font = column[3]
            cell.border = thin_border
            cell.alignment = Alignment(horizontal = column[2], vertical = "top", wrap_text = True)
        status_validation.add(f"E{row_idx}")
    for col_idx, column in enumerate(columns, start = 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = column[1]
    ws.freeze_panes = "C2"
    last_col_letter = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col_letter}{len(recommendations) + 1}"
    ws_dash = wb.create_sheet(title = "Executive Summary", index = 0)
    ws_dash.views.sheetView[0].showGridLines = True
    ws_dash["B2"] = "CIS Benchmark Compliance Scorecard"
    ws_dash["B2"].font = Font(name = "Segoe UI", size = 16, bold = True, color = "1E293B")
    ws_dash["B4"] = "Total CIS Controls Parsed:"
    ws_dash["C4"] = len(recommendations)
    ws_dash["B4"].font = Font(name = "Segoe UI", size = 11, bold = True)
    ws_dash["C4"].font = Font(name = "Segoe UI", size = 11, bold = True, color = "0284C7")
    temporary_path = None
    try:
        output_path.parent.mkdir(parents = True, exist_ok = True)
        with tempfile.NamedTemporaryFile(prefix = ".benchparse-", suffix = ".xlsx", dir = output_path.parent, delete = False) as temporary:
            temporary_path = Path(temporary.name)
        wb.save(temporary_path)
        if output_path.exists() and not overwrite:
            raise FileExistsError("Output already exists. Use --overwrite to replace it.")
        os.replace(temporary_path, output_path)
    finally:
        wb.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok = True)

# ----- Orchestration -------------------------------------------- #

def orchestrate(input_path: Path, output_path: Path, overwrite: bool = False) -> None:
    if output_path.exists() and not overwrite:
        raise typer.BadParameter("Output already exists. Use --overwrite or choose another path.", param_hint = "--output / -o")
    try:
        recommendations = pdf_parse(str(input_path))
    except Exception as exc:
        console.print(f"[bold #FF6B6B][!]Unable to parse the PDF: {exc}")
        raise typer.Exit(code = 1) from exc
    if not recommendations:
        console.print(f"[bold #FF6B6B][!] No recommendations found. The PDF may be scanned or use an unsupported layout.")
        raise typer.Exit(code = 1)
    console.print(f"[#2ECC71][+] Parsed {len(recommendations)} distinct CIS recommendations.")
    try:
        save_cis_to_excel(recommendations, output_path, overwrite = overwrite)
    except Exception as exc:
        console.print(f"[bold #FF6B6B][!] Unable to save the Excel report: {exc}")
        raise typer.Exit(code = 1) from exc
    console.print(f"[#2ECC71][+] Conversion complete! Spreadsheet saved to {output_path}")

# ----- ASCII Banner --------------------------------------------- #

def ascii() -> int:
    art_content = rf"""
oooooooooo.                                  oooo        ooooooooo.
`888'   `Y8b                                 `888        `888   `Y88.
 888     888  .ooooo.  ooo. .oo.    .ooooo.   888 .oo.    888   .d88'  .oooo.   oooo d8b  .oooo.o  .ooooo.
 888oooo888' d88' `88b `888P"Y88b  d88' `"Y8  888P"Y88b   888ooo88P'  `P  )88b  `888""8P d88(  "8 d88' `88b
 888    `88b 888ooo888  888   888  888        888   888   888          .oP"888   888     `"Y88b.  888ooo888
 888    .88P 888    .o  888   888  888   .o8  888   888   888         d8(  888   888     o.  )88b 888    .o
o888bood8P'  `Y8bod8P' o888o o888o `Y8bod8P' o888o o888o o888o        `Y888""8o d888b    8""888P' `Y8bod8P'
"""
    art_color = Text(art_content, style = "#C6ECE3")
    art_panel = Panel(art_color, border_style = "#C6ECE3", expand = False)
    console.print(art_panel)
    banner_width = console.measure(art_panel).maximum
    console.print(Rule(style = "#C6ECE3 blink"), width = banner_width)
    console.print(f"  [~] Created by [bold red blink]Cursed271[/bold red blink]")
    console.print(f"  [~] BenchParse - From benchmark pages to audit progress.")
    console.print(Rule(style = "#C6ECE3 blink"), width = banner_width)
    console.print()
    return banner_width

# ----- Argument Parser ------------------------------------------ #

@app.command()
def argu(
    input_pdf: Optional[str] = typer.Option(None, "--input", "-i", help = "Enter the path to the CIS Benchmark PDF"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help = "Enter the Excel output path; defaults to the input filename with .xlsx"),
    overwrite: bool = typer.Option(False, "--overwrite", help = "Allow replacement of an existing Excel report"),
) -> None:
    interactive = input_pdf is None
    if interactive:
        input_pdf = Prompt.ask(f"[#C6ECE3][?] Enter the path to the CIS Benchmark PDF")
    input_path = pdf_validate(input_pdf)
    default_output = input_path.with_suffix(".xlsx")
    if interactive and output is None:
        output = Prompt.ask(f"[#C6ECE3][?] Enter the Excel output path", default = str(default_output))
    output_path = pdf_output(output if output is not None else str(default_output), input_path)
    console.print(f"[#2ECC71][+] BenchParse initiated conversion of {input_path}")
    orchestrate(input_path = input_path, output_path = output_path, overwrite = overwrite)

# ----- Main Function -------------------------------------------- #

if __name__ == "__main__":
    os.system("cls" if os.name == "nt" else "clear")
    banner_width = ascii()
    typer_rich.MAX_WIDTH = banner_width
    app()

# ----- End ------------------------------------------------------ #