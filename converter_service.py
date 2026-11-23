from __future__ import annotations

import sys
from pathlib import Path

# Garantir carregamento das bibliotecas do ambiente virtual local
_venv_site = Path(__file__).resolve().parent / ".venv" / "Lib" / "site-packages"
if _venv_site.exists() and str(_venv_site) not in sys.path:
    sys.path.insert(0, str(_venv_site))

import csv
import io
import os
import re
import zipfile
from typing import List, Optional, Dict, Any, Tuple, Union

# Safe lazy imports for optional libraries
try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    PdfReader, PdfWriter = None, None

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
except ImportError:
    docx = None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.cell import WriteOnlyCell
except ImportError:
    openpyxl = None
    ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")
    WriteOnlyCell = None

# Reportlab imports para geração de PDF
try:
    from reportlab.lib.pagesizes import letter, A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image as RLImage
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
except ImportError:
    SimpleDocTemplate = None


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def html_escape(text: str) -> str:
    """Escapa caracteres HTML para uso seguro no ReportLab Paragraph."""
    if text is None:
        return ""
    return (
        str(text)
        .replace('&', '&amp;')
        .replace('<', '&lt;')
        .replace('>', '&gt;')
        .replace('"', '&quot;')
        .replace("'", '&#39;')
    )


def clean_inline_text(text: Any) -> str:
    """Remove quebras de linha espúrias e múltiplos espaços de textos e células."""
    if text is None:
        return ""
    s = str(text)
    s = ILLEGAL_CHARACTERS_RE.sub('', s)
    s = re.sub(r'[\r\n\t\v\f]+', ' ', s)
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip()


def extract_page_table_clean(page) -> List[List[str]]:
    """
    Extrai dados tabulares ou estruturados de uma página de PDF com máxima precisão:
    1. Detecta tabelas explícitas (com bordas). Se a tabela explícita cobrir os dados da página, usa-a.
    2. Para tabelas implícitas (sem bordas completas, listas públicas, concursos, faturamentos),
       agrupa palavras por linha vertical (Y) e agrupa colunas horizontais (X) com tolerância inteligente.
    3. Normaliza e remove qualquer quebra de linha interna dentro de células.
    """
    # 1. Tabelas explícitas com linhas/bordas
    table_rows = []
    try:
        tables = page.extract_tables()
        if tables:
            for t in tables:
                for row in t:
                    cleaned_row = [clean_inline_text(c) for c in row]
                    if any(cleaned_row):
                        table_rows.append(cleaned_row)
    except Exception:
        table_rows = []

    max_table_cols = max((len([c for c in r if c]) for r in table_rows), default=0)
    raw_text = page.extract_text() or ""
    text_lines_count = len([l for l in raw_text.split("\n") if l.strip()])

    # Se a tabela explícita for consistente e cobrir a maior parte do conteúdo da página, usa direto
    if max_table_cols > 1 and len(table_rows) >= 6 and len(table_rows) >= (text_lines_count * 0.45):
        return table_rows

    # 2. Extração posicional por palavras
    try:
        words = page.extract_words(x_tolerance=2, y_tolerance=2, keep_blank_chars=False)
    except Exception:
        words = []

    if not words:
        return table_rows

    # Agrupa palavras por linha horizontal (coordenada top com tolerância de 3.5pt)
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    lines: List[List[Dict[str, Any]]] = []
    curr_line: List[Dict[str, Any]] = []
    curr_top = -1.0

    for w in sorted_words:
        if curr_top < 0:
            curr_top = w['top']
            curr_line.append(w)
        elif abs(w['top'] - curr_top) <= 3.5:
            curr_line.append(w)
            curr_top = (curr_top * len(curr_line) + w['top']) / (len(curr_line) + 1)
        else:
            if curr_line:
                curr_line.sort(key=lambda item: item['x0'])
                lines.append(curr_line)
            curr_line = [w]
            curr_top = w['top']

    if curr_line:
        curr_line.sort(key=lambda item: item['x0'])
        lines.append(curr_line)

    # Identificar linhas com características tabulares (gaps horizontais >= 8.0pt)
    multi_col_lines = []
    for l_words in lines:
        gaps = 0
        for idx in range(len(l_words) - 1):
            if (l_words[idx + 1]['x0'] - l_words[idx]['x1']) >= 8.0:
                gaps += 1
        if gaps >= 1:
            multi_col_lines.append(l_words)

    # Clusterização de âncoras de colunas baseada em x0
    col_starts: List[float] = []
    if len(multi_col_lines) >= 2:
        raw_starts = []
        for l_words in multi_col_lines:
            raw_starts.append(l_words[0]['x0'])
            for idx in range(len(l_words) - 1):
                if (l_words[idx + 1]['x0'] - l_words[idx]['x1']) >= 8.0:
                    raw_starts.append(l_words[idx + 1]['x0'])
        
        raw_starts.sort()
        clusters = []
        for x in raw_starts:
            if not clusters or (x - clusters[-1][-1]) > 16.0:
                clusters.append([x])
            else:
                clusters[-1].append(x)
        
        min_occ = max(2, int(len(multi_col_lines) * 0.12))
        raw_col_starts = sorted([sum(c) / len(c) for c in clusters if len(c) >= min_occ])

        # Mescla âncoras excessivamente próximas (< 35pt) para não fragmentar células
        merged_starts: List[float] = []
        for s in raw_col_starts:
            if not merged_starts:
                merged_starts.append(s)
            elif (s - merged_starts[-1]) >= 35.0:
                merged_starts.append(s)
        col_starts = merged_starts

    output_rows: List[List[str]] = []
    if len(col_starts) >= 2:
        for l_words in lines:
            if not l_words:
                continue
            row_cells = [""] * len(col_starts)
            for w in l_words:
                col_idx = 0
                for c_i in range(len(col_starts) - 1, -1, -1):
                    if w['x0'] >= (col_starts[c_i] - 12.0):
                        col_idx = c_i
                        break
                
                if row_cells[col_idx]:
                    row_cells[col_idx] += " " + w['text']
                else:
                    row_cells[col_idx] = w['text']

            row_cells = [clean_inline_text(c) for c in row_cells]
            if any(row_cells):
                output_rows.append(row_cells)
    else:
        for l_words in lines:
            cols = []
            curr_words = []
            last_x1 = -1.0
            for w in l_words:
                if last_x1 >= 0 and (w['x0'] - last_x1) >= 12.0:
                    if curr_words:
                        cols.append(clean_inline_text(" ".join(curr_words)))
                        curr_words = []
                curr_words.append(w['text'])
                last_x1 = w['x1']
            if curr_words:
                cols.append(clean_inline_text(" ".join(curr_words)))
            if cols:
                output_rows.append(cols)

    # Se a extração posicional gerou mais dados ou se table_rows foi fragmentada/vazia, prefere output_rows
    if len(output_rows) >= len(table_rows):
        return output_rows
    return table_rows or output_rows


# =========================================================================
# FUNÇÕES AUXILIARES DE INTERVALO DE PÁGINAS
# =========================================================================

def parse_page_ranges(range_str: str, total_pages: int) -> List[int]:
    """Converte string de intervalos (ex: '1-3, 5, 8-10', '1 a 5', '1 ate 5') em lista de índices 0-indexed."""
    if not range_str or not str(range_str).strip():
        return list(range(total_pages))
    
    normalized = re.sub(r'\s+(?:a|ate|até)\s+', '-', str(range_str), flags=re.IGNORECASE)
    pages = set()
    parts = normalized.split(",")
    for part in parts:
        part = part.strip()
        if "-" in part:
            try:
                start_s, end_s = part.split("-", 1)
                start = max(1, int(start_s.strip()))
                end = min(total_pages, int(end_s.strip()))
                if start <= end:
                    for p in range(start, end + 1):
                        pages.add(p - 1)
                else:
                    for p in range(end, start + 1):
                        pages.add(p - 1)
            except ValueError:
                continue
        elif part.isdigit():
            p = int(part)
            if 1 <= p <= total_pages:
                pages.add(p - 1)
    
    res = sorted(list(pages))
    return res if res else list(range(total_pages))


def resolve_page_indices(
    total_pages: int, 
    options: Optional[Union[Dict[str, Any], str]] = None, 
    custom_range: Optional[str] = None
) -> List[int]:
    """Resolve a lista de índices 0-indexed de páginas selecionadas baseado em options ou range direto."""
    if total_pages <= 0:
        return []
    
    if isinstance(options, str):
        custom_range = options
        options = {}
    else:
        options = options or {}

    range_mode = options.get("pageRangeMode") or options.get("splitMode") or options.get("mode") or "all"
    start = options.get("pageStart") or options.get("page_start")
    end = options.get("pageEnd") or options.get("page_end")
    range_str = custom_range or options.get("pageRange") or options.get("page_range") or options.get("range") or options.get("splitRange") or ""
    
    if range_mode == "all" and not custom_range and not start and not end and not range_str:
        return list(range(total_pages))
    
    # Verificar se veio page_start / page_end direto
    if start is not None and str(start).strip():
        try:
            s_val = max(1, int(str(start).strip()))
            if end is not None and str(end).strip():
                e_val = min(total_pages, int(str(end).strip()))
            else:
                e_val = total_pages
            if s_val <= e_val:
                return list(range(s_val - 1, e_val))
            else:
                return list(range(e_val - 1, s_val))
        except (ValueError, TypeError):
            pass

    range_str = custom_range or options.get("pageRange") or options.get("page_range") or options.get("range") or options.get("splitRange") or ""
    if range_str and str(range_str).strip():
        return parse_page_ranges(str(range_str).strip(), total_pages)
    
    return list(range(total_pages))


# =========================================================================
# 1. CONVERTER DE PDF
# =========================================================================

def convert_pdf_to_word(
    input_path: str | Path, 
    output_path: str | Path,
    options: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> str:
    """
    Converte PDF para documento Word (.docx) preservando estrutura,
    reconstruindo parágrafos contínuos sem quebras de linha artificiais,
    e convertendo listas/tabelas em tabelas formatadas do Word.
    Suporta seleção de páginas específicas através de options.
    """
    import pdfplumber

    doc = docx.Document()
    
    # Configurar margens normais de 2,5 cm (~1 polegada)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    with pdfplumber.open(str(input_path)) as pdf:
        total_pages = len(pdf.pages)
        selected_indices = resolve_page_indices(total_pages, options)
        pages_to_process = [pdf.pages[i] for i in selected_indices if 0 <= i < total_pages]
        if not pages_to_process:
            pages_to_process = pdf.pages

        total_selected = len(pages_to_process)
        has_written_page = False
        for page_idx, page in enumerate(pages_to_process, start=1):
            if progress_callback:
                progress_callback(page_idx, total_selected, f"Processando página {page_idx} de {total_selected} no Word...")
            if has_written_page:
                doc.add_page_break()


            rows = extract_page_table_clean(page)
            if not rows:
                continue

            has_written_page = True
            # Identificar se a página é composta majoritariamente por tabela
            tabular_rows = [r for r in rows if len(r) > 1]
            is_tabular_page = len(tabular_rows) >= 3 and (len(tabular_rows) / len(rows)) >= 0.4

            if is_tabular_page:
                # Cria tabela formatada no Word para os dados tabulares
                max_cols = max(len(r) for r in rows)
                padded_rows = [r + [""] * (max_cols - len(r)) for r in rows]

                doc_table = doc.add_table(rows=len(padded_rows), cols=max_cols)
                doc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                doc_table.style = 'Table Grid'

                for r_idx, row_data in enumerate(padded_rows):
                    for c_idx, cell_value in enumerate(row_data):
                        cell = doc_table.cell(r_idx, c_idx)
                        cell.text = clean_inline_text(cell_value)
                        if r_idx == 0:
                            for p in cell.paragraphs:
                                for run in p.runs:
                                    run.font.bold = True
                            try:
                                shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>')
                                cell._tc.get_or_add_tcPr().append(shading)
                            except Exception:
                                pass
                doc.add_paragraph()
            else:
                # Página com fluxo de texto / parágrafos contínuos
                # Unifica linhas consecutivas que formam o mesmo parágrafo
                paragraphs = []
                current_p = []

                for r in rows:
                    line_text = clean_inline_text("   ".join(r))
                    if not line_text:
                        if current_p:
                            paragraphs.append(" ".join(current_p))
                            current_p = []
                        continue

                    # Se for título evidente
                    is_title = len(line_text) < 70 and (
                        line_text.isupper() or 
                        line_text.endswith(":") or 
                        line_text.startswith(("EDITAL", "CAPÍTULO", "SEÇÃO", "ANEXO", "1)", "2)", "3)"))
                    )

                    if is_title:
                        if current_p:
                            paragraphs.append(" ".join(current_p))
                            current_p = []
                        paragraphs.append(line_text)
                    else:
                        # Se a linha anterior terminava com hífen (separação silábica)
                        if current_p and current_p[-1].endswith("-"):
                            current_p[-1] = current_p[-1][:-1] + line_text
                        else:
                            current_p.append(line_text)

                        # Se a linha termina com ponto final ou dois pontos, conclui o parágrafo
                        if line_text.endswith((".", "!", "?", ":")):
                            paragraphs.append(" ".join(current_p))
                            current_p = []

                if current_p:
                    paragraphs.append(" ".join(current_p))

                for p_text in paragraphs:
                    clean = clean_inline_text(p_text)
                    if not clean:
                        continue
                    heading_prefixes = ('EDITAL', 'CAPÍTULO', 'SEÇÃO', 'ANEXO', '1)', '2)', '3)')
                    if len(clean) < 70 and (clean.isupper() or clean.startswith(heading_prefixes)):
                        h = doc.add_heading(clean, level=2)
                        h.paragraph_format.space_before = Pt(8)
                        h.paragraph_format.space_after = Pt(4)
                    else:
                        p = doc.add_paragraph(clean)
                        p.paragraph_format.space_after = Pt(4)

    doc.save(str(output_path))
    return str(output_path)


# =========================================================================
# PADRÕES E REGRAS DE DETECÇÃO DE RUÍDO E ENTIDADES ÚTEIS PARA ENRIQUECIMENTO
# =========================================================================

RE_CPF = re.compile(r'\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b')
RE_CPF_MASKED = re.compile(r'(?:^|(?<=[\s\(\[\:\,\;]))(?:\d{3}\.?[*xX]{3}\.?[*xX]{3}[-.]?(?:\d{2}|[*xX]{2})|[*xX]{3}\.?\d{3}\.?\d{3}[-.]?(?:\d{2}|[*xX]{2})|\d{3}\.?[*xX\.\-]{5,11}(?:\d{2}|[*xX]{2})|\d{3}[*xX]{3,8}\d{0,2})(?=$|[\s\)\]\,\;\.\-])')
RE_RG = re.compile(r'\b\d{1,2}\.?\d{3}\.?\d{3}[-.]?[0-9xXkK]\b')
RE_CNPJ = re.compile(r'\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b')
RE_PHONE = re.compile(r'(?:\+?55\s?)?(?:\(?([1-9]{2})\)?\s?)?(?:(9\d{4}|\d{4})[-\s]?(\d{4}))\b')
RE_EMAIL = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b')
RE_CEP = re.compile(r'\b\d{5}-?\d{3}\b')
RE_DATE = re.compile(r'\b\d{2}/\d{2}/\d{4}\b')


VALID_BRAZILIAN_DDDS = {
    "11", "12", "13", "14", "15", "16", "17", "18", "19",
    "21", "22", "24", "27", "28",
    "31", "32", "33", "34", "35", "37", "38",
    "41", "42", "43", "44", "45", "46", "47", "48", "49",
    "51", "53", "54", "55",
    "61", "62", "63", "64", "65", "66", "67", "68", "69",
    "71", "73", "74", "75", "77", "79",
    "81", "82", "83", "84", "85", "86", "87", "88", "89",
    "91", "92", "93", "94", "95", "96", "97", "98", "99"
}


def normalize_phone_value(value: Any, ddd: str = "") -> str:
    """
    Normaliza qualquer telefone brasileiro (celular ou fixo), tratando com precisão:
    - Prefixos de tipo de contato (ex: 'p:', 'm:', 'c:', 't:', 'w:', 'tel:', 'cel:').
    - O DDI do Brasil (+55, 0055, +), distinguindo-o com precisão do DDD 55 (Santa Maria / RS).
    - Textos explicativos (ex: 'Tel:', 'WhatsApp:', 'Cel:').
    - Inserção do 9º dígito em celulares de 8 dígitos.
    - Associação de DDD fornecido quando o número for local (8 ou 9 dígitos).
    """
    if value is None or value == "":
        return ""
    val_str = str(value).strip()
    if not val_str:
        return ""

    # Remove prefixos de contato como 'p:', 'p:+55', 'tel:', 'cel:', 'm:', 'c:', 'w:', 'whats:', etc.
    val_str = re.sub(r"^\s*(?:(?:p|m|c|t|w|f|v|tel|telefone|cel|celular|whats|whatsapp|zap|fone|phone|mobile|cell|contato)\s*[:.-]?\s*)+", "", val_str, flags=re.IGNORECASE).strip()
    if not val_str:
        return ""

    clean_ddd = re.sub(r"\D", "", str(ddd or ""))
    if clean_ddd.startswith("0") and len(clean_ddd) == 3:
        clean_ddd = clean_ddd[1:]
    clean_ddd = clean_ddd[:2]

    # 1. Se a string começa explicitamente com DDI internacional (+55, + 55, 0055)
    if re.search(r"^\s*(?:\+|00)\s*55\b", val_str):
        remainder = re.sub(r"^\s*(?:\+|00)\s*55[\s.-]*", "", val_str)
        rem_digits = re.sub(r"\D", "", remainder)
        if rem_digits.startswith("0") and len(rem_digits) in (11, 12):
            rem_digits = rem_digits[1:]

        if len(rem_digits) == 11 and rem_digits[2] == "9":
            return rem_digits
        elif len(rem_digits) == 10:
            if rem_digits[:2] in VALID_BRAZILIAN_DDDS and rem_digits[2] in "6789":
                return f"{rem_digits[:2]}9{rem_digits[2:]}"
            elif len(clean_ddd) == 2 and rem_digits[0] in "6789":
                return f"{clean_ddd}9{rem_digits}"
            return rem_digits
        elif len(rem_digits) == 9:
            if len(clean_ddd) == 2:
                return f"{clean_ddd}{rem_digits}"
            return f"55{rem_digits}"
        elif len(rem_digits) == 8:
            if rem_digits[0] in "6789":
                rem_digits = f"9{rem_digits}"
            if len(clean_ddd) == 2:
                return f"{clean_ddd}{rem_digits}"
            return f"55{rem_digits}"

    digits = re.sub(r"\D", "", val_str)
    if not digits:
        return val_str

    if digits.startswith("0055"):
        digits = digits[4:]
    elif digits.startswith("00"):
        digits = digits[2:]

    # 2. 13 ou 14 dígitos iniciando com DDI 55 (ex: 5555999998888 ou 5582996506901)
    if digits.startswith("55") and len(digits) in (13, 14):
        digits = digits[2:]
        if digits.startswith("0") and len(digits) == 12:
            digits = digits[1:]
        if len(digits) == 11 and digits[2] == "9":
            return digits

    # 3. 12 dígitos (ex: 55 + DDD + 8 dígitos móvel OU 0 + DDD + 9 dígitos móvel)
    if len(digits) == 12:
        if digits.startswith("0"):
            d11 = digits[1:]
            if len(d11) == 11 and d11[2] == "9":
                return d11
            elif len(d11) == 10 and d11[2] in "6789":
                return f"{d11[:2]}9{d11[2:]}"
        elif digits.startswith("55") and digits[2:4] in VALID_BRAZILIAN_DDDS:
            extracted_ddd = digits[2:4]
            local8 = digits[4:]
            if local8[0] in "6789":
                return f"{extracted_ddd}9{local8}"
            else:
                return f"{extracted_ddd}{local8}"
        elif digits[:2] in VALID_BRAZILIAN_DDDS and digits[2] == "0":
            return f"{digits[:2]}{digits[3:]}"

    # 4. 11 dígitos
    if len(digits) == 11:
        if digits.startswith("0"):
            d10 = digits[1:]
            if d10[:2] in VALID_BRAZILIAN_DDDS and d10[2] in "6789":
                return f"{d10[:2]}9{d10[2:]}"
            return d10
        
        # Número com DDD brasileiro válido (inclusive DDD 55!)
        if digits[:2] in VALID_BRAZILIAN_DDDS and digits[2] == "9":
            return digits
        
        if digits.startswith("55") and digits[2:4] in VALID_BRAZILIAN_DDDS:
            extracted_ddd = digits[2:4]
            local7 = digits[4:]
            return f"{extracted_ddd}9{local7}" if len(local7) == 7 else digits

        return digits

    # 5. 10 dígitos (DDD + 8 dígitos)
    if len(digits) == 10:
        if digits.startswith("0"):
            local9 = digits[1:]
            if len(clean_ddd) == 2:
                return f"{clean_ddd}{local9}"
            return local9
        
        extracted_ddd = digits[:2]
        local8 = digits[2:]
        if extracted_ddd in VALID_BRAZILIAN_DDDS:
            if local8[0] in "6789":
                return f"{extracted_ddd}9{local8}"
            else:
                return f"{extracted_ddd}{local8}"
        elif len(clean_ddd) == 2:
            if digits[0] == "9":
                return f"{clean_ddd}{digits[1:]}"

    # 6. 9 dígitos (móvel local)
    if len(digits) == 9:
        if len(clean_ddd) == 2:
            return f"{clean_ddd}{digits}"
        return digits

    # 7. 8 dígitos (local 8 dígitos)
    if len(digits) == 8:
        if digits[0] in "6789":
            digits = f"9{digits}"
        if len(clean_ddd) == 2:
            digits = f"{clean_ddd}{digits}"
        return digits

    return digits


PHONE_ANNOTATION_WORDS = {
    "p", "m", "c", "t", "w", "f", "v",
    "tel", "telefone", "telefones", "cel", "celular", "celulares",
    "whats", "whatsapp", "zap", "fone", "fones", "contato", "contatos",
    "phone", "mobile", "cell", "work", "home", "voice", "fax",
    "fixo", "fixos", "ramal", "comercial", "residencial", "principal",
    "casa", "trabalho", "recado", "recados", "pessoal", "movel", "moveis",
    "wpp", "tel1", "tel2", "cel1", "cel2"
}


def is_annotation_only_text(text: str) -> bool:
    if not text:
        return True
    s = str(text).strip().lower()
    s = re.sub(r'^\s*(?:\+|00)?\s*55\b', '', s)
    s = re.sub(r'[\(\)\s:.\-\/\+\\]+', '', s).strip()
    if not s or s.isdigit():
        return True
    return s in PHONE_ANNOTATION_WORDS


def extract_phone_and_name_from_combined_text(val: Any) -> Tuple[Optional[str], Optional[str]]:
    """
    Detecta se um texto contém simultaneamente um Telefone (10 ou 11 dígitos) e um Nome,
    suportando prefixos como 'p:+55', '+55', '+', 'Tel:', 'Cel:', 'WhatsApp:', 'Zap:', 'Fone:', 'Contato:', etc.
    Exemplos:
      '82996506901 AMANDA KESIA CUNHA ROCHA' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
      '(82) 99650-6901 AMANDA KESIA CUNHA ROCHA' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
      '+55 82 99650-6901 AMANDA KESIA CUNHA ROCHA' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
      'p:+55 82 99650-6901 AMANDA KESIA' -> ('82996506901', 'AMANDA KESIA')
      'Tel: +55 (82) 99650-6901 AMANDA KESIA' -> ('82996506901', 'AMANDA KESIA')
      'AMANDA KESIA CUNHA ROCHA (Tel: +55 82 99650-6901)' -> ('82996506901', 'AMANDA KESIA CUNHA ROCHA')
    """
    if not val:
        return None, None
    s = str(val).strip()
    s_clean = re.sub(r'[\r\n\t]+', ' ', s).strip()
    s_single_space = re.sub(r'\s{2,}', ' ', s_clean)

    prefix_pattern = r'(?:(?:p|m|c|t|w|f|v|tel|telefone|cel|celular|whats|whatsapp|zap|fone|phone|mobile|cell|contato)\s*[:.-]?\s*)?'
    phone_pattern = r'(?:(?:\+|00)?55[\s.-]?)?(?:\(?0?([1-9]{2})\)?[\s.-]?)?(?:9[\s.-]?)?(?:(9\d{4}|\d{4})[\s.-]?(\d{4}))'

    # 1. Telefone no início (com ou sem prefixos)
    m_start = re.match(
        rf'^{prefix_pattern}{phone_pattern}\s*[-–—/:]?\s*([A-Za-zÀ-ÿ].*)$',
        s_single_space,
        flags=re.IGNORECASE
    )
    if m_start:
        name_part = m_start.group(4).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part):
            return None, None
        phone_raw = s_single_space[:len(s_single_space) - len(name_part)]
        digits = normalize_phone_value(phone_raw)
        if len(digits) in (10, 11) and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    # 2. Telefone no final (com ou sem prefixos ou parênteses)
    m_end = re.match(
        rf'^([A-Za-zÀ-ÿ].*?)\s*[-–—/:]?\s*\(?{prefix_pattern}{phone_pattern}\)?$',
        s_single_space,
        flags=re.IGNORECASE
    )
    if m_end:
        name_part = m_end.group(1).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part):
            return None, None
        phone_raw = s_single_space[len(name_part):]
        digits = normalize_phone_value(phone_raw)
        if len(digits) in (10, 11) and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    # 3. Telefone numérico puro no início (ex: '82996506901 AMANDA KESIA' ou '+5582996506901 AMANDA' ou 'p:+5582996506901 AMANDA')
    m_digits_start = re.match(r'^(?:(?:p|m|c|t|w|f|v|tel|cel|whats|fone|phone)\s*[:.-]?\s*)?(?:(?:\+|00)?55)?([1-9]{2}9?\d{8})\s+([A-Za-zÀ-ÿ].*)$', s_single_space, flags=re.IGNORECASE)
    if m_digits_start:
        digits = m_digits_start.group(1)
        if len(digits) == 10 and digits[2] in ("6", "7", "8", "9"):
            digits = f"{digits[:2]}9{digits[2:]}"
        name_part = m_digits_start.group(2).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part):
            return None, None
        if len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    return None, None


def extract_cpf_and_name_from_combined_text(val: Any) -> Tuple[Optional[str], Optional[str]]:
    """
    Detecta se um texto contém simultaneamente um CPF e um Nome real de pessoa.
    Exemplos:
      '123.456.789-00 AMANDA KESIA CUNHA ROCHA' -> ('12345678900', 'AMANDA KESIA CUNHA ROCHA')
      '12345678900 AMANDA KESIA CUNHA ROCHA' -> ('12345678900', 'AMANDA KESIA CUNHA ROCHA')
      'AMANDA KESIA CUNHA ROCHA 123.456.789-00' -> ('12345678900', 'AMANDA KESIA CUNHA ROCHA')
    """
    if not val:
        return None, None
    s = str(val).strip()
    s_clean = re.sub(r'[\r\n\t]+', ' ', s).strip()
    s_single_space = re.sub(r'\s{2,}', ' ', s_clean)

    # CPF mascarado no final
    m_mask_end = re.search(r'\s+((?:\d{3}\.?[*xX]{3}\.?[*xX]{3}[-.]?(?:\d{2}|[*xX]{2})|[*xX]{3}\.?\d{3}\.?\d{3}[-.]?(?:\d{2}|[*xX]{2})|\d{3}\.?[*xX\.\-]{5,11}(?:\d{2}|[*xX]{2})|\d{3}[*xX]{3,8}\d{0,2}))\s*$', s_single_space)
    if m_mask_end:
        doc_part = m_mask_end.group(1).strip()
        name_part = s_single_space[:m_mask_end.start()].strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if len(name_part) >= 2:
            return doc_part, name_part

    # CPF mascarado no início
    m_mask_start = re.match(r'^((?:\d{3}\.?[*xX]{3}\.?[*xX]{3}[-.]?(?:\d{2}|[*xX]{2})|[*xX]{3}\.?\d{3}\.?\d{3}[-.]?(?:\d{2}|[*xX]{2})|\d{3}\.?[*xX\.\-]{5,11}(?:\d{2}|[*xX]{2})|\d{3}[*xX]{3,8}\d{0,2}))\s*[-–—/:]?\s*([A-Za-zÀ-ÿ].*)$', s_single_space)
    if m_mask_start:
        doc_part = m_mask_start.group(1).strip()
        name_part = m_mask_start.group(2).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if len(name_part) >= 2:
            return doc_part, name_part

    # CPF no início
    m_cpf_start = re.match(r'^(\d{3}\.?\d{3}\.?\d{3}-?\d{2})\s*[-–—/:]?\s*([A-Za-zÀ-ÿ].*)$', s_single_space)
    if m_cpf_start:
        name_part = m_cpf_start.group(2).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part) or '+' in name_part or ':' in name_part:
            return None, None
        cpf_raw = m_cpf_start.group(1)
        digits = re.sub(r'\D', '', cpf_raw)
        if len(digits) == 11 and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    # CPF no final
    m_cpf_end = re.match(r'^([A-Za-zÀ-ÿ].*?)\s*[-–—/:]?\s*(\d{3}\.?\d{3}\.?\d{3}-?\d{2})$', s_single_space)
    if m_cpf_end:
        name_part = m_cpf_end.group(1).strip()
        name_part = re.sub(r'^[-–—/:\(\)\s]+|[-–—/:\(\)\s]+$', '', name_part)
        if is_annotation_only_text(name_part) or '+' in name_part or ':' in name_part:
            return None, None
        cpf_raw = m_cpf_end.group(2)
        digits = re.sub(r'\D', '', cpf_raw)
        if len(digits) == 11 and len(name_part) >= 3 and not name_part.isdigit():
            return digits, name_part

    return None, None

UF_LIST = {
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA',
    'MT', 'MS', 'MG', 'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN',
    'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
}

NOISE_PHRASES = [
    "CONSIDERANDO", "DISPOE SOBRE", "DISPÕE SOBRE", "RESOLVE", "DECRETA", "PORTARIA", "EDITAL", "LEI",
    "DIARIO OFICIAL", "DIÁRIO OFICIAL", "SECRETARIA", "SECRETARIO", "SECRETÁRIO", "PREFEITURA", "GOVERNO",
    "MINISTERIO", "MINISTÉRIO", "REPUBLICA", "REPÚBLICA", "PODER EXECUTIVO", "PODER JUDICIARIO",
    "PODER JUDICIÁRIO", "PODER LEGISLATIVO", "TRIBUNAL", "GABINETE", "DEPARTAMENTO", "COMISSAO",
    "COMISSÃO", "TERMOS E CONDICOES", "TERMOS E CONDIÇÕES", "TERMOS DE USO", "POLITICA DE PRIVACIDADE",
    "POLÍTICA DE PRIVACIDADE", "DIREITOS RESERVADOS", "COPYRIGHT", "PAGINA", "PÁGINA", "AUTENTICIDADE",
    "CHAVE DE ACESSO", "ASSINADO DIGITALMENTE", "DOCUMENTO ASSINADO", "HOMOLOGA", "HOMOLOGACAO",
    "HOMOLOGAÇÃO", "CONVOCA", "CONVOCACAO", "CONVOCAÇÃO", "ESTE EDITAL", "FICAM CONVOCADOS",
    "TOTAL DE REGISTROS", "RELATORIO", "RELATÓRIO", "EMISSAO", "EMISSÃO", "INFORMACAO", "INFORMAÇÃO",
    "DISPOSICOES", "DISPOSIÇÕES", "CAPITULO", "CAPÍTULO", "SECAO", "SEÇÃO", "ARTIGO", "ART.",
    "PARAGRAFO", "PARÁGRAFO", "ATRIBUICOES", "ATRIBUIÇÕES", "LEGISLACAO", "LEGISLAÇÃO", "DOCUMENTOS",
    "CONTRARIO", "CONTRÁRIO", "APRESENTACAO", "APRESENTAÇÃO", "NO USO DE", "DISPOSTO NA", "VIGENTE"
]

HEADER_SYNONYMS = {
    "nome": ["nome", "nome completo", "nome do cliente", "nome do candidato", "nome do servidor", "candidato", "cliente", "razao social", "razao", "titular", "contato", "favorecido", "servidor", "beneficiario", "beneficiário", "aluno", "paciente", "funcionario", "funcionário", "proponente"],
    "cpf": ["cpf", "cpf/cnpj", "documento", "doc", "cpf_cnpj", "num_cpf", "nr_cpf", "rg", "documento_identidade"],
    "cnpj": ["cnpj", "num_cnpj", "nr_cnpj"],
    "telefone": ["telefone", "celular", "fone", "tel", "cel", "contato_tel", "whatsapp", "telefone_1", "tel_1", "telefone nome", "nome telefone", "tel nome", "nome tel", "celular nome", "nome celular"],
    "telefone_2": ["telefone_2", "celular_2", "fone_2", "tel_2", "cel_2", "outro_telefone", "recado"],
    "email": ["email", "e-mail", "correio eletronico", "email_contato"],
    "logradouro": ["logradouro", "endereco", "endereço", "rua", "av", "avenida", "logr"],
    "numero": ["numero", "número", "num", "nº", "nr"],
    "complemento": ["complemento", "compl", "apto", "bloco"],
    "bairro": ["bairro", "bair"],
    "cidade": ["cidade", "municipio", "município", "localidade"],
    "uf": ["uf", "estado", "est"],
    "cep": ["cep", "codigo_postal"],
    "data_nascimento": ["data_nascimento", "nascimento", "dt_nasc", "dt_nascimento", "data de nascimento"],
    "matricula": ["matricula", "matrícula", "inscricao", "inscrição", "codigo", "código", "id", "registro", "protocolo", "ordem", "item", "seq", "sequencia", "sequência", "posicao", "posição"],
    "cargo": ["cargo", "profissao", "profissão", "funcao", "função", "ocupacao", "ocupação", "especialidade", "disciplina", "area", "área", "area de atuacao", "área de atuação", "setor", "departamento", "divisao", "divisão", "lotacao", "lotação", "polo", "unidade", "vinculo", "vínculo"],
    "status": ["status", "situacao", "situação", "resultado", "parecer", "fase", "etapa", "condicao", "condição", "classificacao", "classificação", "ranking", "nota", "pontos", "pontuacao", "pontuação", "opcao", "opção", "vaga", "curso", "turno", "modalidade", "categoria", "grupo", "tipo", "observacao", "observação", "obs"],
    "valor": ["valor", "total", "saldo", "vlr", "quantia", "renda", "salario", "salário", "vencimento", "remuneracao", "remuneração"]
}


COMMON_NON_NAME_WORDS = {
    'AS', 'OS', 'UM', 'UMA', 'UNS', 'UMAS', 'ESTE', 'ESTA', 'ESTES', 'ESTAS', 'ESSE', 'ESSA',
    'AQUELE', 'AQUELA', 'QUE', 'QUAL', 'QUAIS', 'ONDE', 'COMO', 'QUANDO', 'MUITO', 'MAIS',
    'INSTRUCOES', 'CONTIDAS', 'ALI', 'TERMOS', 'CONDICOES', 'DIAS', 'PRAZO', 'DATA', 'HORA',
    'LOCAL', 'SALA', 'BLOCO', 'PREDIO', 'ANDAR', 'MES', 'ANO', 'LEI', 'DECRETO', 'PORTARIA',
    'EDITAL', 'ARTIGO', 'INCISO', 'PARAGRAFO', 'CAPUT', 'ITEM', 'ALINEA', 'ANEXO', 'TABELA',
    'TOTAL', 'SUBTOTAL', 'GERAL', 'FINAL', 'PARCIAL', 'RESULTADO', 'HOMOLOGACAO', 'CONVOCACAO',
    'PRESENTE', 'PUBLICACAO', 'SECRETARIA', 'PREFEITURA', 'GOVERNO', 'ESTADO', 'MUNICIPAL'
}


def looks_like_person_name(text: str) -> bool:
    """Verifica se um texto tem estrutura de nome próprio de pessoa (ex: 'AMANDA KESIA CUNHA ROCHA')."""
    clean = clean_inline_text(text)
    if not clean or clean.endswith(('.', ';', ':', '!', '?', ',')):
        return False
    words = clean.split()
    if len(words) < 2 or len(words) > 7:
        return False
    if re.search(r'\d', clean) or clean.islower():
        return False
    header_terms = {
        "NOME", "COMPLETO", "CANDIDATO", "CLIENTE", "TITULAR", "CONTATO", "SERVIDOR", 
        "CARGO", "FUNCAO", "AREA", "SETOR", "STATUS", "SITUACAO", "ORDEM", "INSCRICAO",
        "MATRICULA", "CPF", "TELEFONE", "CELULAR", "EMAIL", "ENDERECO", "CIDADE", "ESTADO",
        "DOCUMENTO", "DOCUMENTOS", "EDITAL", "PREFEITURA", "SECRETARIA"
    }
    norm_words = [normalize_text_unaccented(w) for w in words]
    if any(w in header_terms or w in COMMON_NON_NAME_WORDS for w in norm_words):
        return False
    for w in words:
        w_clean = re.sub(r"['\-\.]", "", w)
        if len(w_clean) < 2 or not w_clean.isalpha():
            return False
    return True



def normalize_text_unaccented(text: Any) -> str:
    """Normaliza texto removendo acentos e convertendo para maiúsculo para matching exato."""
    if text is None:
        return ""
    import unicodedata
    nfkd = unicodedata.normalize('NFKD', str(text))
    return ''.join([c for c in nfkd if not unicodedata.combining(c)]).upper().strip()


def is_noise_text(text: str) -> bool:
    """Verifica se um texto isolado é ruído administrativo, legal ou informativo sem valor de entidade."""
    if not text:
        return True
    
    t_clean = text.strip()
    norm_t = normalize_text_unaccented(t_clean)

    # Página x de y, Pág. x, etc.
    if re.match(r'^(PAGINA|PAG\.?|FOLHA)\s*\d+(\s*(DE|/)\s*\d+)?$', norm_t):
        return True

    # Data/Hora isolada
    if re.match(r'^\d{2}/\d{2}/\d{4}(\s+\d{2}:\d{2}(:\d{2})?)?$', t_clean):
        return True

    # Traços ou separadores
    if re.match(r'^[-=_*#~.\s]{3,}$', t_clean):
        return True

    # Frases de aviso ou preâmbulo institucional
    for phrase in NOISE_PHRASES:
        norm_p = normalize_text_unaccented(phrase)
        if norm_p and norm_p in norm_t:
            if not RE_CPF.search(t_clean) and not RE_CPF_MASKED.search(t_clean) and not RE_EMAIL.search(t_clean) and not RE_PHONE.search(t_clean):
                return True

    # Frase longa explicativa com pontuação final e sem entidades
    if len(t_clean) > 55 and t_clean.endswith(('.', ':', ';', '!', '?')):
        if not RE_CPF.search(t_clean) and not RE_CPF_MASKED.search(t_clean) and not RE_EMAIL.search(t_clean) and not RE_PHONE.search(t_clean):
            return True

    return False


def is_valid_entity_record(row: List[str], detected_headers: Optional[List[str]] = None) -> bool:
    """Valida se uma linha representa um registro de dados úteis para enriquecimento."""
    non_empty = [clean_inline_text(c) for c in row if clean_inline_text(c)]
    if not non_empty:
        return False
    
    full_text = " ".join(non_empty)
    norm_full = normalize_text_unaccented(full_text)

    # Subtítulos, legendas de classificação ou avisos que não são registros
    norm_lower = norm_full.lower()
    if norm_lower in ['ordem alfabetica', 'classificacao geral', 'ampla concorrencia', 'vagas reservadas', 'deferimentos', 'indeferimentos', 'inscricoes deferidas', 'inscricoes indeferidas']:
        return False
    
    # Detecção de entidades fortes (CPF, CPF Mascarado, CNPJ, Email, Telefone ou RG)
    has_strong_entity = bool(
        RE_CPF.search(full_text) or 
        RE_CPF_MASKED.search(full_text) or 
        RE_CNPJ.search(full_text) or 
        RE_EMAIL.search(full_text) or 
        RE_PHONE.search(full_text) or
        RE_RG.search(full_text)
    )

    # Detecção se contém nome próprio de pessoa em alguma célula
    has_person_name = any(looks_like_person_name(c) for c in non_empty)
    
    # Se contém qualquer palavra-chave de ruído institucional/legal em linha de célula única e sem entidade forte
    if len(non_empty) == 1 and not has_strong_entity:
        for kw in NOISE_PHRASES:
            norm_kw = normalize_text_unaccented(kw)
            if norm_kw and norm_kw in norm_full:
                return False
        # Linha de texto corrido narrativo
        if full_text.endswith(('.', ':', ';', '!', '?')) or full_text.islower():
            return False

    # Linha de cabeçalho provável não deve ser tratada como linha de registro
    if is_probable_header_row(row):
        return False

    # Se possui entidade forte ou nome de pessoa estruturado, é dado válido
    if has_strong_entity or has_person_name:
        return True

    # Validação contextual contra cabeçalhos conhecidos se disponíveis
    if detected_headers and not has_strong_entity:
        cpf_idx = next((i for i, h in enumerate(detected_headers) if h == 'CPF'), None)
        phone_idx = next((i for i, h in enumerate(detected_headers) if h.startswith('Telefone')), None)
        email_idx = next((i for i, h in enumerate(detected_headers) if h == 'Email'), None)
        
        if cpf_idx is not None and cpf_idx < len(row) and row[cpf_idx]:
            if not (RE_CPF.search(str(row[cpf_idx])) or RE_CPF_MASKED.search(str(row[cpf_idx]))):
                return False
        if phone_idx is not None and phone_idx < len(row) and row[phone_idx]:
            if not RE_PHONE.search(str(row[phone_idx])):
                return False
        if email_idx is not None and email_idx < len(row) and row[email_idx]:
            if not RE_EMAIL.search(str(row[email_idx])):
                return False

    # Linha com múltiplas colunas estruturadas
    if len(non_empty) >= 2:
        return True

    # Linha com apenas 1 célula contendo nome de pessoa
    if len(non_empty) == 1 and looks_like_person_name(non_empty[0]):
        return True
        
    return False


def is_noise_row(row: List[str]) -> bool:
    """Verifica se uma linha inteira extraída da tabela é mero ruído."""
    return not is_valid_entity_record(row)


def normalize_header_name(val: str) -> str:
    """Normaliza nomes de colunas para os padrões aceitos pelo enriquecedor com ordem de prioridade."""
    v = clean_inline_text(val).lower()
    v_norm = re.sub(r'[^a-z0-9/ ]', '', v).strip()
    
    # 1. Prioridade: Telefones (evita que 'telefone de contato' caia em 'nome')
    if any(t in v_norm for t in ['telefone', 'celular', 'fone', 'tel ', ' cel ']) or v_norm.endswith(' tel') or v_norm.endswith(' cel') or v_norm == 'tel' or v_norm == 'cel':
        return 'Telefone_2' if ('2' in v_norm or 'segundo' in v_norm or 'outro' in v_norm) else 'Telefone'

    # 2. Prioridade: Documentos
    if 'cpf' in v_norm:
        return 'CPF'
    if 'cnpj' in v_norm:
        return 'CNPJ'
    if 'rg' in v_norm:
        return 'RG'

    # 3. Prioridade: Email
    if 'email' in v_norm or 'e-mail' in v_norm:
        return 'Email'

    # 4. Prioridade: Nome e Favorecido
    if any(t in v_norm for t in ['nome', 'cliente', 'titular', 'candidato', 'favorecido', 'servidor', 'beneficiario', 'beneficiário', 'razao social', 'razao']):
        return 'Nome'

    # 5. Prioridade: Ordem / Inscrição / Matrícula
    if any(t in v_norm for t in ['ordem', 'item', 'seq', 'posicao', 'posição']):
        return 'Ordem'
    if any(t in v_norm for t in ['matricula', 'matrícula', 'inscricao', 'inscrição', 'registro', 'codigo', 'código', 'protocolo']):
        return 'Inscricao'

    # 6. Prioridade: Área / Cargo / Função
    if any(t in v_norm for t in ['area', 'área', 'setor', 'departamento', 'lotacao', 'lotação', 'polo', 'unidade', 'especialidade', 'disciplina']):
        return 'Area'
    if any(t in v_norm for t in ['cargo', 'profissao', 'profissão', 'funcao', 'função', 'ocupacao', 'ocupação']):
        return 'Cargo'

    # 7. Prioridade: Status / Situação / Pontuação / Opção
    if any(t in v_norm for t in ['status', 'situacao', 'situação', 'resultado', 'parecer', 'fase', 'etapa', 'condicao', 'condição']):
        return 'Status'
    if any(t in v_norm for t in ['pontos', 'pontuacao', 'pontuação', 'nota', 'classificacao', 'classificação', 'ranking']):
        return 'Pontuacao'
    if any(t in v_norm for t in ['opcao', 'opção', 'vaga', 'curso', 'turno', 'modalidade', 'categoria', 'grupo', 'tipo']):
        return 'Opcao'

    # 8. Prioridade: Endereço e Localização
    if 'logradouro' in v_norm or 'endereco' in v_norm or 'rua' in v_norm or 'avenida' in v_norm:
        return 'Logradouro'
    if 'numero' in v_norm or 'número' in v_norm or v_norm in ['num', 'nº', 'nr']:
        return 'Numero'
    if 'complemento' in v_norm or 'apto' in v_norm or 'bloco' in v_norm:
        return 'Complemento'
    if 'bairro' in v_norm:
        return 'Bairro'
    if 'cidade' in v_norm or 'municipio' in v_norm:
        return 'Cidade'
    if v_norm in ['uf', 'estado', 'est']:
        return 'UF'
    if 'cep' in v_norm:
        return 'CEP'

    # 9. Outros campos padrão
    if 'nasc' in v_norm:
        return 'Data_Nascimento'
    if any(t in v_norm for t in ['valor', 'total', 'saldo', 'renda', 'salario', 'salário', 'vencimento', 'remuneracao', 'remuneração']):
        return 'Valor'

    return clean_inline_text(val).title()


def is_probable_header_row(row: List[str]) -> bool:
    """Verifica se a linha é claramente um cabeçalho de tabela."""
    non_empty = [clean_inline_text(c) for c in row if clean_inline_text(c)]
    if not non_empty or len(non_empty) < 2:
        return False

    full_t = " ".join(non_empty)
    # Se contém CPF formatado, CPF mascarado, e-mail ou número de telefone explícito de 10-11 dígitos, é dado real e não cabeçalho
    if RE_CPF.search(full_t) or RE_CPF_MASKED.search(full_t) or RE_EMAIL.search(full_t) or re.search(r'\b[1-9]{2}9\d{7,8}\b', full_t):
        return False

    # Se alguma célula é claramente um nome de pessoa (ex: 'AMANDA KESIA CUNHA ROCHA'), é dado real
    if any(looks_like_person_name(c) for c in non_empty):
        return False

    norm_full = normalize_text_unaccented(full_t).lower()
    # Avisos de ordenação ou subtítulos que não são cabeçalhos de coluna
    if norm_full in ['ordem alfabetica', 'classificacao geral', 'ampla concorrencia', 'vagas reservadas', 'deferimentos', 'indeferimentos']:
        return False

    matches = 0
    for cell in non_empty:
        c_clean = normalize_text_unaccented(cell).lower()
        c_clean = re.sub(r'[^a-z0-9/ ]', '', c_clean).strip()
        for canon, synonyms in HEADER_SYNONYMS.items():
            found = False
            for s in synonyms:
                norm_s = normalize_text_unaccented(s).lower()
                if c_clean == norm_s or c_clean.startswith(norm_s + " ") or c_clean.endswith(" " + norm_s):
                    matches += 1
                    found = True
                    break
            if found:
                break
    
    # Se pelo menos 2 células são termos de cabeçalho ou >= 40% das células
    return matches >= 2 or (matches >= 1 and len(non_empty) == 2 and any(canon in norm_full for canon in ['nome', 'cpf', 'cargo', 'documento', 'telefone', 'email']))


def is_probable_footer_row(row: List[str]) -> bool:
    """Detecta se uma linha é claramente um rodapé de página ou carimbo de paginação."""
    non_empty = [clean_inline_text(c) for c in row if clean_inline_text(c)]
    if not non_empty:
        return False
    full_text = " ".join(non_empty)
    text_norm = normalize_text_unaccented(full_text).lower()

    # Se contém CPF, email ou telefone explícito de 10-11 dígitos, é dado real
    if RE_CPF.search(full_text) or RE_EMAIL.search(full_text) or re.search(r'\b[1-9]{2}9\d{7,8}\b', full_text):
        return False

    # 1. Padrão "Página X de Y" / "Pág. X/Y" / "Page X of Y"
    if re.search(r'\b(pagina|pag|page|p)\.?\s*\d+\s*(de|/|of)\s*\d+\b', text_norm):
        return True

    # 2. Número de página isolado (ex: "1", "- 2 -", "[ 3 ]")
    if re.match(r'^\s*[-–—\[\(]?\s*\d{1,4}\s*[-–—\]\)]?\s*$', full_text):
        return True

    # 3. Data de emissão ou carimbo de sistema no rodapé (ex: "Emitido em 25/08/2026", "Relatório gerado em...")
    if re.search(r'\b(emitido|gerado|impresso|processado|relatorio emitido)\s+(em|as|por|no dia)\b', text_norm):
        return True

    # 4. URLs governamentais ou links de rodapé isolados (ex: "www.prefeitura.sp.gov.br")
    if re.match(r'^(?:https?://)?(?:www\.)?[a-z0-9.-]+\.(?:gov|org|edu|leg)\.br\s*$', text_norm):
        return True

    return False


def infer_column_headers(rows: List[List[str]], num_cols: int) -> List[str]:
    """Infere os nomes mais prováveis para cada coluna com base no tipo de dado predominante."""
    col_scores: List[Dict[str, int]] = [{"cpf": 0, "phone": 0, "email": 0, "name": 0, "uf": 0, "cep": 0, "date": 0, "num": 0} for _ in range(num_cols)]

    sample_rows = rows[:100]
    total_samples = len(sample_rows)
    if total_samples == 0:
        return [f"Coluna_{i+1}" for i in range(num_cols)]

    for r in sample_rows:
        for c_idx in range(min(len(r), num_cols)):
            val = r[c_idx].strip()
            if not val:
                continue

            digits = re.sub(r'\D', '', val)

            # Telefone explícito com parênteses, hífen de 4 dígitos ou prefixo internacional
            if any(c in val for c in ['(', ')', '+55']) or (len(digits) in (10, 11) and '-' in val and '.' not in val):
                col_scores[c_idx]["phone"] += 1
            # CPF explícito com máscara padrão
            elif re.search(r'\b\d{3}\.\d{3}\.\d{3}-\d{2}\b', val):
                col_scores[c_idx]["cpf"] += 1
            # CNPJ com máscara padrão
            elif re.search(r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b', val):
                col_scores[c_idx]["cpf"] += 1
            # 11 dígitos numéricos puros (diferenciação entre Celular e CPF)
            elif len(digits) == 11 and digits.isdigit():
                if digits[2] == '9' and int(digits[:2]) in range(11, 99):
                    col_scores[c_idx]["phone"] += 1
                else:
                    col_scores[c_idx]["cpf"] += 1
            elif len(digits) == 10 and int(digits[:2]) in range(11, 99):
                col_scores[c_idx]["phone"] += 1
            elif RE_EMAIL.search(val):
                col_scores[c_idx]["email"] += 1
            elif RE_CEP.search(val) or (len(digits) == 8 and digits.isdigit() and "-" in val):
                col_scores[c_idx]["cep"] += 1
            elif val.upper() in UF_LIST:
                col_scores[c_idx]["uf"] += 1
            elif RE_DATE.search(val):
                col_scores[c_idx]["date"] += 1
            elif re.match(r'^[A-Za-zÀ-ÿ]{2,}(\s+[A-Za-zÀ-ÿ]{1,}){1,6}$', val):
                col_scores[c_idx]["name"] += 1
            elif val.isdigit():
                col_scores[c_idx]["num"] += 1

    headers = []
    used_headers = set()

    for idx, scores in enumerate(col_scores):
        best_type = max(scores, key=scores.get)
        best_count = scores[best_type]
        threshold = max(2, int(total_samples * 0.20))

        if best_count >= threshold:
            if best_type == "cpf" and "CPF" not in used_headers:
                h = "CPF"
            elif best_type == "phone":
                h = "Telefone" if "Telefone" not in used_headers else "Telefone_2"
            elif best_type == "email" and "Email" not in used_headers:
                h = "Email"
            elif best_type == "name" and "Nome" not in used_headers:
                h = "Nome"
            elif best_type == "uf" and "UF" not in used_headers:
                h = "UF"
            elif best_type == "cep" and "CEP" not in used_headers:
                h = "CEP"
            elif best_type == "date" and "Data" not in used_headers:
                h = "Data"
            elif best_type == "num" and idx == 0 and "Inscricao" not in used_headers:
                h = "Inscricao"
            else:
                h = f"Coluna_{idx+1}"
        else:
            h = f"Coluna_{idx+1}"

        used_headers.add(h)
        headers.append(h)

    return headers


def filter_and_structure_table_rows(
    pages_data: List[Tuple[int, List[List[str]]]], 
    smart_enrich: bool = True
) -> Tuple[List[str], List[List[str]]]:
    """
    Filtra ruídos e estrutura os dados tabulares do PDF de forma limpa e inteligente.
    Garante que cabeçalhos e títulos repetidos em cada página NÃO se repitam nas linhas de dados.
    Retorna (headers, rows).
    """
    if not pages_data:
        return ["Dados"], [["Nenhum conteúdo extraído do PDF."]]

    if not smart_enrich:
        all_raw_rows: List[List[str]] = []
        max_cols = 1
        first_page_top_sigs = set()

        for p_idx, (_, p_rows) in enumerate(pages_data):
            for r_idx, r in enumerate(p_rows):
                cleaned = [clean_inline_text(c) for c in r]
                if not any(cleaned):
                    continue
                non_empty = [c for c in cleaned if c]
                sig = "|".join(normalize_text_unaccented(c) for c in non_empty)
                
                # Na primeira página, registra os cabeçalhos e títulos do topo
                if p_idx == 0:
                    if r_idx <= 3:
                        first_page_top_sigs.add(sig)
                    all_raw_rows.append(cleaned)
                    if len(cleaned) > max_cols:
                        max_cols = len(cleaned)
                else:
                    # Em páginas seguintes (2, 3...), se a linha for repetição do topo ou cabeçalho provável, ignora
                    if (r_idx <= 4 and sig in first_page_top_sigs) or (r_idx <= 3 and is_probable_header_row(cleaned)):
                        continue
                    if len(cleaned) > max_cols:
                        max_cols = len(cleaned)
                    all_raw_rows.append(cleaned)
        
        if not all_raw_rows:
            return ["Dados"], [["Nenhum dado encontrado."]]
            
        padded = [r + [""] * (max_cols - len(r)) for r in all_raw_rows]
        headers = [f"Coluna_{i+1}" for i in range(max_cols)]
        return headers, padded

    candidate_rows: List[List[str]] = []
    detected_headers: Optional[List[str]] = None
    known_header_signatures: set[str] = set()
    known_header_tokens: set[str] = set()
    repeated_page_banners: set[str] = set()
    repeated_page_footers: set[str] = set()

    # 1. Identificar títulos/banners no topo e rodapés na base de múltiplas páginas
    if len(pages_data) > 1:
        from collections import Counter
        top_lines_per_page = []
        bottom_lines_per_page = []
        for _, p_rows in pages_data:
            page_tops = []
            page_bottoms = []
            for r in p_rows[:4]:
                cl = [clean_inline_text(c) for c in r if clean_inline_text(c)]
                if cl:
                    sig = " ".join(normalize_text_unaccented(c) for c in cl)
                    page_tops.append(sig)
            if len(p_rows) >= 6:
                for r in p_rows[-3:]:
                    cl = [clean_inline_text(c) for c in r if clean_inline_text(c)]
                    if cl:
                        sig = " ".join(normalize_text_unaccented(c) for c in cl)
                        page_bottoms.append(sig)
            top_lines_per_page.append(set(page_tops))
            bottom_lines_per_page.append(set(page_bottoms))
        
        all_top_sigs = [sig for p_set in top_lines_per_page for sig in p_set]
        top_counts = Counter(all_top_sigs)
        for sig, count in top_counts.items():
            if count >= 2:
                repeated_page_banners.add(sig)

        all_bottom_sigs = [sig for p_set in bottom_lines_per_page for sig in p_set]
        bottom_counts = Counter(all_bottom_sigs)
        for sig, count in bottom_counts.items():
            if count >= 2:
                repeated_page_footers.add(sig)

    # 2. Processar linhas página a página
    for p_num, p_rows in pages_data:
        total_p_rows = len(p_rows)
        for r_idx, r in enumerate(p_rows):
            cleaned = [clean_inline_text(c) for c in r]
            if not any(cleaned):
                continue
            
            non_empty_cells = [c for c in cleaned if c]
            row_sig = "|".join(normalize_text_unaccented(c) for c in non_empty_cells)
            row_text_sig = " ".join(normalize_text_unaccented(c) for c in non_empty_cells)

            # 1. Detecta se é linha de cabeçalho da tabela (prioridade máxima)
            if is_probable_header_row(cleaned):
                known_header_signatures.add(row_sig)
                for c in non_empty_cells:
                    known_header_tokens.add(normalize_text_unaccented(c))
                norm_candidates = [normalize_header_name(c) for c in non_empty_cells]
                
                def header_quality(hdrs):
                    score = len(hdrs)
                    if any('Nome' in h for h in hdrs): score += 5
                    if any('CPF' in h or 'Documento' in h or 'RG' in h for h in hdrs): score += 5
                    if any('Cargo' in h for h in hdrs): score += 3
                    if any('Telefone' in h for h in hdrs): score += 4
                    return score

                if detected_headers is None or header_quality(norm_candidates) > header_quality(detected_headers):
                    detected_headers = norm_candidates
                continue

            # 2. Detecta e descarta rodapés explícitos (ex: "Página X de Y", carimbos de emissão) ou rodapés repetidos na base
            is_at_bottom = r_idx >= max(0, total_p_rows - 3)
            if is_probable_footer_row(cleaned) or (is_at_bottom and row_text_sig in repeated_page_footers):
                continue

            # 3. Ignora assinatura exata de cabeçalho já conhecido
            if row_sig in known_header_signatures:
                continue

            # 4. Ignora banners/títulos de cabeçalho de página repetidos (documentos, secretarias, editais)
            if row_text_sig in repeated_page_banners:
                continue

            # Se já detectamos cabeçalhos, verifica se esta linha em páginas 2+ é repetição de cabeçalho
            if detected_headers and len(non_empty_cells) >= 2:
                norm_cells = [normalize_text_unaccented(c) for c in non_empty_cells]
                matches_header = sum(1 for c in norm_cells if c in known_header_tokens or any(c == normalize_text_unaccented(h) for h in detected_headers))
                if matches_header >= max(2, int(len(norm_cells) * 0.4)):
                    full_str = " ".join(non_empty_cells)
                    if not (RE_CPF.search(full_str) or RE_CPF_MASKED.search(full_str) or re.search(r'\b[1-9]{2}9\d{7,8}\b', full_str) or RE_EMAIL.search(full_str)):
                        known_header_signatures.add(row_sig)
                        continue

            # Filtra ruídos administrativos e textos soltos
            if not is_valid_entity_record(cleaned, detected_headers=detected_headers):
                continue

            candidate_rows.append(cleaned)

    if not candidate_rows:
        # Fallback garantido: nunca devolver vazio ou com mensagem de erro quando há dados no PDF
        all_raw_fallback: List[List[str]] = []
        max_c = 1
        for _, p_rows in pages_data:
            for r in p_rows:
                cl = [clean_inline_text(c) for c in r if clean_inline_text(c)]
                if cl:
                    all_raw_fallback.append(cl)
                    if len(cl) > max_c:
                        max_c = len(cl)
        if all_raw_fallback:
            padded_fb = [r + [""] * (max_c - len(r)) for r in all_raw_fallback]
            headers_fb = [f"Coluna_{i+1}" for i in range(max_c)]
            return headers_fb, padded_fb
        return ["Dados"], [["Nenhum conteúdo extraído do PDF."]]

    # 3. Limpeza final de qualquer cabeçalho que possa ter entrado em candidate_rows
    if detected_headers:
        norm_det_hdrs = [normalize_text_unaccented(h) for h in detected_headers]
        filtered_candidates = []
        for r in candidate_rows:
            r_norm = [normalize_text_unaccented(c) for c in r if c]
            if r_norm == norm_det_hdrs:
                continue
            if len(r_norm) >= 2 and all(c in known_header_tokens for c in r_norm):
                continue
            filtered_candidates.append(r)
        candidate_rows = filtered_candidates or candidate_rows

    # 4. Detecção e separação de colunas com Telefone+Nome ou CPF+Nome combinados
    sample_cand = candidate_rows[:100]
    num_c = max(len(r) for r in sample_cand)
    col_split_plans = {}

    for c_idx in range(num_c):
        phone_name_cnt = 0
        cpf_name_cnt = 0
        non_empty_cnt = 0

        for r in sample_cand:
            if c_idx < len(r) and r[c_idx].strip():
                non_empty_cnt += 1
                p_d, n_p = extract_phone_and_name_from_combined_text(r[c_idx])
                if p_d and n_p:
                    phone_name_cnt += 1
                else:
                    c_d, cn_p = extract_cpf_and_name_from_combined_text(r[c_idx])
                    if c_d and cn_p:
                        cpf_name_cnt += 1

        if non_empty_cnt > 0:
            min_matches = 1 if len(sample_cand) == 1 else 2
            if phone_name_cnt >= max(min_matches, int(non_empty_cnt * 0.30)):
                col_split_plans[c_idx] = "phone_name"
            elif cpf_name_cnt >= max(min_matches, int(non_empty_cnt * 0.30)):
                col_split_plans[c_idx] = "cpf_name"

    if col_split_plans:
        new_candidates = []
        for r in candidate_rows:
            new_r = []
            for c_idx, val in enumerate(r):
                if c_idx in col_split_plans:
                    plan = col_split_plans[c_idx]
                    if plan == "phone_name":
                        p_d, n_p = extract_phone_and_name_from_combined_text(val)
                        if p_d and n_p:
                            new_r.extend([n_p, p_d])
                        else:
                            new_r.extend([val, ""])
                    elif plan == "cpf_name":
                        c_d, cn_p = extract_cpf_and_name_from_combined_text(val)
                        if c_d and cn_p:
                            new_r.extend([c_d, cn_p])
                        else:
                            new_r.extend(["", val])
                else:
                    new_r.append(val)
            new_candidates.append(new_r)
        candidate_rows = new_candidates

        if detected_headers:
            new_headers = []
            for c_idx, h in enumerate(detected_headers):
                if c_idx in col_split_plans:
                    plan = col_split_plans[c_idx]
                    if plan == "phone_name":
                        new_headers.extend(["Nome", "Telefone"])
                    elif plan == "cpf_name":
                        new_headers.extend(["CPF", "Nome"])
                else:
                    new_headers.append(h)
            detected_headers = new_headers

    max_cols = max(len(r) for r in candidate_rows)

    if detected_headers:
        if len(detected_headers) < max_cols:
            detected_headers += [f"Coluna_{i+1}" for i in range(len(detected_headers), max_cols)]
        headers = detected_headers[:max_cols]
    else:
        headers = infer_column_headers(candidate_rows, max_cols)

    final_rows: List[List[str]] = []
    for r in candidate_rows:
        row_padded = r + [""] * (max_cols - len(r))
        final_rows.append(row_padded[:max_cols])

    return headers, final_rows


def convert_pdf_to_excel_streaming(
    input_path: str | Path, 
    output_path: str | Path, 
    options: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> str:
    """
    Pipeline streaming de alta escala para conversao de PDFs extensos (50 a 5.000+ paginas).
    Utiliza amostragem estatistica inicial (Smart Sampling) para perfilamento de schema e regras
    e gravacao direta no disco via openpyxl.Workbook(write_only=True), mantendo o consumo de memoria
    RAM constante e estavel (< 300 MB) independentemente do volume de paginas.
    """
    import gc
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.cell import WriteOnlyCell

    options = options or {}
    excel_mode = options.get("excelMode", "smart_enrich")
    smart_enrich = (excel_mode != "full")

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    font_data = Font(name="Calibri", size=11)

    def is_numeric_val(val: str) -> bool:
        v = val.replace(".", "").replace(",", "").replace("-", "").replace("+", "").replace("R$", "").replace("%", "").strip()
        return v.isdigit() and len(v) > 0

    if progress_callback:
        progress_callback(0, 100, "Iniciando leitura e analise estatistica do documento...")

    import pdfplumber
    with pdfplumber.open(str(input_path)) as pdf:
        total_pages = len(pdf.pages)
        selected_indices = resolve_page_indices(total_pages, options)
        total_selected = len(selected_indices) if selected_indices else total_pages
        if not selected_indices:
            selected_indices = list(range(total_pages))

        # 1. Fase de Amostragem Inteligente (Smart Sampling)
        if total_selected <= 24:
            sample_indices = list(selected_indices)
        else:
            first_part = selected_indices[:12]
            mid_idx = total_selected // 2
            mid_part = selected_indices[max(0, mid_idx - 3):min(total_selected, mid_idx + 3)]
            end_part = selected_indices[-6:]
            sample_indices = sorted(list(set(first_part + mid_part + end_part)))

        sample_pages_data: List[Tuple[int, List[List[str]]]] = []
        for s_idx in sample_indices:
            if 0 <= s_idx < total_pages:
                try:
                    p = pdf.pages[s_idx]
                    p_rows = extract_page_table_clean(p)
                    p.flush_cache()
                    if p_rows:
                        sample_pages_data.append((s_idx + 1, p_rows))
                except Exception:
                    pass
        try:
            pdf.flush_cache()
        except Exception:
            pass

        # Fallback se a leitura da amostra falhou no pdfplumber
        if not sample_pages_data:
            try:
                from pypdf import PdfReader
                reader = PdfReader(str(input_path))
                for s_idx in sample_indices:
                    if 0 <= s_idx < len(reader.pages):
                        txt = reader.pages[s_idx].extract_text() or ""
                        lines = [l.strip() for l in txt.split("\n") if l.strip()]
                        fb_rows = []
                        for line in lines:
                            cols = [clean_inline_text(c) for c in re.split(r'\t+|\s{2,}', line) if c.strip()]
                            if cols:
                                fb_rows.append(cols)
                        if fb_rows:
                            sample_pages_data.append((s_idx + 1, fb_rows))
            except Exception:
                pass

        if not sample_pages_data:
            sample_pages_data = extract_text_from_scanned_pdf(input_path)

        if not sample_pages_data:
            wb_empty = openpyxl.Workbook()
            ws_e = wb_empty.active
            ws_e.title = "Dados"
            ws_e.cell(row=1, column=1, value="Nenhum conteudo extraido do PDF.")
            wb_empty.save(str(output_path))
            return str(output_path)

        # 2. Perfilamento de Cabecalhos e Regras da Amostra
        sample_headers, sample_structured = filter_and_structure_table_rows(
            sample_pages_data, 
            smart_enrich=smart_enrich
        )

        if smart_enrich and (not sample_structured or sample_structured == [["Nenhum conteudo extraido do PDF."]] or sample_structured == [["Nenhum dado encontrado."]]):
            smart_enrich = False
            sample_headers, sample_structured = filter_and_structure_table_rows(
                sample_pages_data, 
                smart_enrich=False
            )

        headers = sample_headers
        max_cols = len(headers)

        # Identificar assinaturas de banners de topo e rodapes repetidos
        repeated_page_banners = set()
        repeated_page_footers = set()
        if len(sample_pages_data) > 1:
            from collections import Counter
            top_lines_per_page = []
            bottom_lines_per_page = []
            for _, p_rows in sample_pages_data:
                page_tops = []
                page_bottoms = []
                for r in p_rows[:4]:
                    cl = [clean_inline_text(c) for c in r if clean_inline_text(c)]
                    if cl:
                        sig = " ".join(normalize_text_unaccented(c) for c in cl)
                        page_tops.append(sig)
                if len(p_rows) >= 6:
                    for r in p_rows[-3:]:
                        cl = [clean_inline_text(c) for c in r if clean_inline_text(c)]
                        if cl:
                            sig = " ".join(normalize_text_unaccented(c) for c in cl)
                            page_bottoms.append(sig)
                top_lines_per_page.append(set(page_tops))
                bottom_lines_per_page.append(set(page_bottoms))

            for sig, count in Counter([s for p_s in top_lines_per_page for s in p_s]).items():
                if count >= 2:
                    repeated_page_banners.add(sig)
            for sig, count in Counter([s for p_s in bottom_lines_per_page for s in p_s]).items():
                if count >= 2:
                    repeated_page_footers.add(sig)

        known_header_signatures = set()
        known_header_tokens = set()
        for h in headers:
            known_header_tokens.add(normalize_text_unaccented(h))

        # Detectar planos de divisao de colunas combinadas (telefone+nome ou cpf+nome)
        col_split_plans = {}
        sample_cand = sample_structured[:100]
        for c_idx in range(max_cols):
            phone_name_cnt = 0
            cpf_name_cnt = 0
            non_empty_cnt = 0
            for r in sample_cand:
                if c_idx < len(r) and r[c_idx].strip():
                    non_empty_cnt += 1
                    p_d, n_p = extract_phone_and_name_from_combined_text(r[c_idx])
                    if p_d and n_p:
                        phone_name_cnt += 1
                    else:
                        c_d, cn_p = extract_cpf_and_name_from_combined_text(r[c_idx])
                        if c_d and cn_p:
                            cpf_name_cnt += 1
            if non_empty_cnt > 0:
                min_matches = 1 if len(sample_cand) == 1 else 2
                if phone_name_cnt >= max(min_matches, int(non_empty_cnt * 0.30)):
                    col_split_plans[c_idx] = "phone_name"
                elif cpf_name_cnt >= max(min_matches, int(non_empty_cnt * 0.30)):
                    col_split_plans[c_idx] = "cpf_name"

        # Calcular larguras proporcionais das colunas
        col_widths: Dict[str, float] = {}
        for c_idx in range(max_cols):
            max_len = len(str(headers[c_idx] or ""))
            for r in sample_structured[:300]:
                if c_idx < len(r):
                    max_len = max(max_len, len(str(r[c_idx] or "")))
            col_letter = get_column_letter(c_idx + 1)
            col_widths[col_letter] = max(min(max_len + 4, 65), 14)

        # 3. Inicializar Planilha em Modo Streaming (write_only=True)
        wb = openpyxl.Workbook(write_only=True)
        main_sheet_title = "Dados" if smart_enrich else "Tabela Completa"
        ws = wb.create_sheet(title=main_sheet_title)
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        def create_header_cells(sheet_obj):
            h_cells = []
            for h_text in headers:
                c = WriteOnlyCell(sheet_obj, value=clean_inline_text(h_text))
                c.font = header_font
                c.fill = header_fill
                c.alignment = align_center
                c.border = border_thin
                h_cells.append(c)
            return h_cells

        ws.append(create_header_cells(ws))

        # 4. Stream de Gravacao Pagina a Pagina
        total_written_rows = 0
        current_sheet_rows = 0
        sheet_part = 1
        first_page_top_sigs = set()

        for p_i, idx in enumerate(selected_indices, start=1):
            page_rows = []
            if 0 <= idx < total_pages:
                try:
                    p = pdf.pages[idx]
                    page_rows = extract_page_table_clean(p)
                    p.flush_cache()
                except Exception:
                    page_rows = []

            if page_rows:
                total_p_rows = len(page_rows)
                for r_idx, r in enumerate(page_rows):
                    cleaned = [clean_inline_text(c) for c in r]
                    if not any(cleaned):
                        continue
                    non_empty_cells = [c for c in cleaned if c]
                    row_sig = "|".join(normalize_text_unaccented(c) for c in non_empty_cells)
                    row_text_sig = " ".join(normalize_text_unaccented(c) for c in non_empty_cells)

                    if not smart_enrich:
                        if p_i == 1:
                            if r_idx <= 3:
                                first_page_top_sigs.add(row_sig)
                        else:
                            if (r_idx <= 4 and row_sig in first_page_top_sigs) or (r_idx <= 3 and is_probable_header_row(cleaned)):
                                continue
                    else:
                        # 1. Cabecalho da tabela
                        if is_probable_header_row(cleaned):
                            known_header_signatures.add(row_sig)
                            for c in non_empty_cells:
                                known_header_tokens.add(normalize_text_unaccented(c))
                            continue

                        # 2. Rodapes
                        is_at_bottom = r_idx >= max(0, total_p_rows - 3)
                        if is_probable_footer_row(cleaned) or (is_at_bottom and row_text_sig in repeated_page_footers):
                            continue

                        # 3. Assinatura exata de cabecalho
                        if row_sig in known_header_signatures:
                            continue

                        # 4. Banners repetidos
                        if row_text_sig in repeated_page_banners:
                            continue

                        # 5. Cabecalhos repetidos no topo de paginas 2+
                        if len(non_empty_cells) >= 2:
                            norm_cells = [normalize_text_unaccented(c) for c in non_empty_cells]
                            matches_header = sum(1 for c in norm_cells if c in known_header_tokens or any(c == normalize_text_unaccented(h) for h in headers))
                            if matches_header >= max(2, int(len(norm_cells) * 0.4)):
                                full_str = " ".join(non_empty_cells)
                                if not (RE_CPF.search(full_str) or RE_CPF_MASKED.search(full_str) or re.search(r'\b[1-9]{2}9\d{7,8}\b', full_str) or RE_EMAIL.search(full_str)):
                                    known_header_signatures.add(row_sig)
                                    continue

                        # 6. Validacao de entidade valida (ignora ruidos e textos soltos)
                        if not is_valid_entity_record(cleaned, detected_headers=headers):
                            continue

                        # 7. Divisao de colunas combinadas
                        if col_split_plans:
                            new_r = []
                            for col_i, val in enumerate(cleaned):
                                if col_i in col_split_plans:
                                    plan = col_split_plans[col_i]
                                    if plan == "phone_name":
                                        p_d, n_p = extract_phone_and_name_from_combined_text(val)
                                        if p_d and n_p:
                                            new_r.extend([n_p, p_d])
                                        else:
                                            new_r.extend([val, ""])
                                    elif plan == "cpf_name":
                                        c_d, cn_p = extract_cpf_and_name_from_combined_text(val)
                                        if c_d and cn_p:
                                            new_r.extend([c_d, cn_p])
                                        else:
                                            new_r.extend(["", val])
                                else:
                                    new_r.append(val)
                            cleaned = new_r

                    # Ajustar colunas ao tamanho de headers
                    padded = cleaned + [""] * max(0, max_cols - len(cleaned))
                    row_data = padded[:max_cols]

                    # Particionar em nova aba se atingir limite maximo do Excel (1 milhao de linhas)
                    if current_sheet_rows >= 1000000:
                        sheet_part += 1
                        ws = wb.create_sheet(title=f"Dados_Parte_{sheet_part}")
                        for col_letter, width in col_widths.items():
                            ws.column_dimensions[col_letter].width = width
                        ws.append(create_header_cells(ws))
                        current_sheet_rows = 0

                    total_written_rows += 1
                    current_sheet_rows += 1
                    is_even = (current_sheet_rows % 2 == 0)

                    row_cells = []
                    for val in row_data:
                        val_clean = clean_inline_text(val)
                        c = WriteOnlyCell(ws, value=val_clean)
                        c.font = font_data
                        c.border = border_thin
                        if is_numeric_val(val_clean):
                            c.alignment = align_right
                        elif len(val_clean) <= 4:
                            c.alignment = align_center
                        else:
                            c.alignment = align_left

                        if is_even:
                            c.fill = zebra_fill
                        row_cells.append(c)

                    ws.append(row_cells)

            # Limpeza periodica de memoria e atualizacao de progresso
            if p_i % 20 == 0 or p_i == total_selected:
                try:
                    pdf.flush_cache()
                except Exception:
                    pass
                gc.collect()

                if progress_callback:
                    progress_callback(
                        p_i, 
                        total_selected, 
                        f"Extraindo dados da página {p_i} de {total_selected} ({total_written_rows} registros extraídos)..."
                    )

        if total_written_rows == 0:
            ws.append([WriteOnlyCell(ws, value="Nenhum conteúdo extraído do PDF.")])

    if progress_callback:
        progress_callback(total_selected, total_selected, "Salvando planilha final...")

    wb.save(str(output_path))
    wb.close()
    return str(output_path)


def convert_pdf_to_excel(
    input_path: str | Path, 
    output_path: str | Path, 
    options: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> str:
    """
    Extrai tabelas e dados de PDF para planilha Excel (.xlsx) profissional,
    sem quebras de linha indesejadas, com alinhamento perfeito e sem truncamento.
    No modo 'smart_enrich' (padrão), elimina ruídos, disclaimers e textos soltos,
    estruturando apenas as entidades e dados relevantes prontos para enriquecimento direto.
    Para documentos com mais de 30 páginas, ativa automaticamente o pipeline streaming
    de alta performance e consumo constante de memória (suporta até 5.000+ páginas).
    """
    options = options or {}

    # 0. Verificar quantidade de páginas para direcionar ao pipeline streaming de alta escala
    try:
        import pdfplumber
        with pdfplumber.open(str(input_path)) as _p_check:
            _tot_p = len(_p_check.pages)
            _sel_idx = resolve_page_indices(_tot_p, options)
            _tot_sel = len(_sel_idx) if _sel_idx else _tot_p
    except Exception:
        _tot_sel = 1

    if _tot_sel > 30:
        return convert_pdf_to_excel_streaming(
            input_path=input_path,
            output_path=output_path,
            options=options,
            progress_callback=progress_callback
        )

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    excel_mode = options.get("excelMode", "smart_enrich")
    smart_enrich = (excel_mode != "full")

    wb = openpyxl.Workbook()
    if wb.active:
        wb.remove(wb.active)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    def is_numeric(val: str) -> bool:
        """Verifica se string representa valor numérico."""
        v = val.replace(".", "").replace(",", "").replace("-", "").replace("+", "").replace("R$", "").replace("%", "").strip()
        return v.isdigit() and len(v) > 0

    pages_data: List[Tuple[int, List[List[str]]]] = []

    if progress_callback:
        progress_callback(0, 100, "Iniciando leitura do documento...")

    # 1. Extração estruturada inteligente via pdfplumber
    try:
        import pdfplumber
        with pdfplumber.open(str(input_path)) as pdf:
            total_pages = len(pdf.pages)
            selected_indices = resolve_page_indices(total_pages, options)
            total_selected = len(selected_indices) if selected_indices else total_pages
            for i, idx in enumerate(selected_indices, start=1):
                if 0 <= idx < total_pages:
                    if progress_callback:
                        progress_callback(i, total_selected, f"Extraindo dados da página {i} de {total_selected}...")
                    page = pdf.pages[idx]
                    page_rows = extract_page_table_clean(page)
                    try:
                        page.flush_cache()
                    except Exception:
                        pass
                    if page_rows:
                        pages_data.append((idx + 1, page_rows))
            try:
                pdf.flush_cache()
            except Exception:
                pass
    except Exception:
        pass

    # 2. Fallbacks de leitura caso nada tenha sido extraído
    if not pages_data:
        try:
            from pypdf import PdfReader
            reader = PdfReader(str(input_path))
            total_pages = len(reader.pages)
            selected_indices = resolve_page_indices(total_pages, options)
            total_selected = len(selected_indices) if selected_indices else total_pages
            for i, idx in enumerate(selected_indices, start=1):
                if 0 <= idx < total_pages:
                    if progress_callback:
                        progress_callback(i, total_selected, f"Lendo texto da página {i} de {total_selected}...")
                    page = reader.pages[idx]
                    text = page.extract_text() or ""
                    lines = [l.strip() for l in text.split("\n") if l.strip()]
                    page_rows = []
                    for line in lines:
                        cols = [clean_inline_text(c) for c in re.split(r'\t+|\s{2,}', line) if c.strip()]
                        if cols:
                            page_rows.append(cols)
                    if page_rows:
                        pages_data.append((idx + 1, page_rows))
        except Exception:
            pass

    if not pages_data:
        if progress_callback:
            progress_callback(1, 1, "Aplicando OCR em documento escaneado...")
        pages_data = extract_text_from_scanned_pdf(input_path)

    # Se ainda estiver vazio
    if not pages_data:
        ws = wb.create_sheet(title="Dados")
        ws.cell(row=1, column=1, value="Nenhum conteúdo extraído do PDF.")
        wb.save(str(output_path))
        return str(output_path)

    if progress_callback:
        tot = total_selected if 'total_selected' in locals() else 1
        progress_callback(tot, tot, "Formatando tabelas e colunas da planilha...")

    # 3. Filtragem inteligente e estruturação de cabeçalhos e registros
    headers, structured_rows = filter_and_structure_table_rows(pages_data, smart_enrich=smart_enrich)


    def write_formatted_sheet(ws_target, sheet_headers: List[str], sheet_rows: List[List[str]]):
        # Escreve o cabeçalho na Linha 1
        for c_idx, h_text in enumerate(sheet_headers, start=1):
            cell = ws_target.cell(row=1, column=c_idx, value=clean_inline_text(h_text))
            cell.border = border_thin
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Escreve as linhas de dados a partir da Linha 2
        for r_idx, row in enumerate(sheet_rows, start=2):
            for c_idx, val in enumerate(row, start=1):
                val_clean = clean_inline_text(val)
                cell = ws_target.cell(row=r_idx, column=c_idx, value=val_clean)
                cell.border = border_thin

                if is_numeric(val_clean):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif len(val_clean) <= 4:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

                if r_idx % 2 == 0:
                    cell.fill = zebra_fill

        try:
            ws_target.auto_filter.ref = ws_target.dimensions
        except Exception:
            pass

        # Auto-ajuste proporcional da largura das colunas
        for col in ws_target.columns:
            max_len = 0
            for c in col:
                val_len = len(str(c.value or ''))
                if val_len > max_len:
                    max_len = val_len
            col_letter = get_column_letter(col[0].column)
            ws_target.column_dimensions[col_letter].width = max(min(max_len + 4, 65), 14)

    # 4. Criação da Aba Principal Consolidada (compatível direto com /api/upload)
    main_sheet_title = "Dados" if smart_enrich else "Tabela Completa"
    ws_main = wb.create_sheet(title=main_sheet_title)
    write_formatted_sheet(ws_main, headers, structured_rows)

    wb.save(str(output_path))
    return str(output_path)


def convert_word_to_excel(
    input_path: str | Path, 
    output_path: str | Path, 
    options: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> str:
    """
    Converte documento do Word (.docx) para planilha Excel (.xlsx) profissional,
    extraindo todas as tabelas e dados estruturados, tratando parágrafos tabulares
    e formatando com cabeçalhos elegantes prontos para enriquecimento direto.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import docx

    options = options or {}
    excel_mode = options.get("excelMode", "smart_enrich")
    smart_enrich = (excel_mode != "full")

    if progress_callback:
        progress_callback(0, 100, "Iniciando leitura do documento Word...")

    doc = docx.Document(str(input_path))

    wb = openpyxl.Workbook()
    if wb.active:
        wb.remove(wb.active)

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    def is_numeric(val: str) -> bool:
        v = val.replace(".", "").replace(",", "").replace("-", "").replace("+", "").replace("R$", "").replace("%", "").strip()
        return v.isdigit() and len(v) > 0

    pages_data: List[Tuple[int, List[List[str]]]] = []

    # 1. Extração de Tabelas do Word
    tables = doc.tables
    total_tables = len(tables)

    if total_tables > 0:
        for t_idx, table in enumerate(tables, start=1):
            if progress_callback:
                progress_callback(t_idx, total_tables, f"Extraindo dados da tabela {t_idx} de {total_tables}...")

            table_rows: List[List[str]] = []
            for row in table.rows:
                row_vals = []
                for cell in row.cells:
                    c_text = " ".join([p.text.strip() for p in cell.paragraphs if p.text.strip()])
                    row_vals.append(clean_inline_text(c_text))

                if any(v.strip() for v in row_vals):
                    table_rows.append(row_vals)

            if table_rows:
                pages_data.append((t_idx, table_rows))

    # 2. Se não houver tabelas ou houver apenas registros em parágrafos
    if not pages_data or len(pages_data[0][1]) < 3:
        if progress_callback:
            progress_callback(50, 100, "Analisando estrutura de parágrafos e listas...")

        para_rows: List[List[str]] = []
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            cols = [clean_inline_text(c) for c in re.split(r'\t+|\s{2,}', text) if c.strip()]
            if cols:
                para_rows.append(cols)

        if para_rows:
            pages_data.append((len(pages_data) + 1, para_rows))

    if not pages_data:
        ws = wb.create_sheet(title="Dados")
        ws.cell(row=1, column=1, value="Nenhum conteúdo ou tabela extraída do documento Word.")
        wb.save(str(output_path))
        return str(output_path)

    if progress_callback:
        progress_callback(90, 100, "Formatando tabelas e colunas da planilha...")

    # 3. Filtragem inteligente e estruturação
    headers, structured_rows = filter_and_structure_table_rows(pages_data, smart_enrich=smart_enrich)

    def write_formatted_sheet(ws_target, sheet_headers: List[str], sheet_rows: List[List[str]]):
        for c_idx, h_text in enumerate(sheet_headers, start=1):
            cell = ws_target.cell(row=1, column=c_idx, value=clean_inline_text(h_text))
            cell.border = border_thin
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r_idx, row in enumerate(sheet_rows, start=2):
            for c_idx, val in enumerate(row, start=1):
                val_clean = clean_inline_text(val)
                cell = ws_target.cell(row=r_idx, column=c_idx, value=val_clean)
                cell.border = border_thin

                if is_numeric(val_clean):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif len(val_clean) <= 4:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

                if r_idx % 2 == 0:
                    cell.fill = zebra_fill

        try:
            ws_target.auto_filter.ref = ws_target.dimensions
        except Exception:
            pass

        for col in ws_target.columns:
            max_len = 0
            for c in col:
                val_len = len(str(c.value or ''))
                if val_len > max_len:
                    max_len = val_len
            col_letter = get_column_letter(col[0].column)
            ws_target.column_dimensions[col_letter].width = max(min(max_len + 4, 65), 14)

    # 4. Criação da Aba Principal
    main_sheet_title = "Dados" if smart_enrich else "Tabela Completa"
    ws_main = wb.create_sheet(title=main_sheet_title)
    write_formatted_sheet(ws_main, headers, structured_rows)

    if progress_callback:
        progress_callback(100, 100, "Gravando arquivo Excel...")

    wb.save(str(output_path))
    return str(output_path)




def extract_text_from_scanned_pdf(input_path: str | Path) -> List[Tuple[int, List[List[str]]]]:
    """Extrai texto e tabelas de PDF escaneado (imagens) usando OCR quando disponível."""
    pages_data: List[Tuple[int, List[List[str]]]] = []
    try:
        import pytesseract
        import pypdfium2 as pdfium
        from PIL import Image

        doc = pdfium.PdfDocument(str(input_path))
        for page_idx in range(len(doc)):
            pil_img = doc[page_idx].render(scale=2.0).to_pil()
            text = pytesseract.image_to_string(pil_img, lang="por+eng") or ""
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            page_rows = []
            for line in lines:
                cols = [clean_inline_text(c) for c in re.split(r'\t+|\s{2,}', line) if c.strip()]
                if cols:
                    page_rows.append(cols)
            if page_rows:
                pages_data.append((page_idx + 1, page_rows))
            pil_img.close()
        doc.close()
    except Exception:
        pass
    return pages_data


def convert_pdf_to_images(
    input_path: str | Path, 
    output_dir: str | Path, 
    img_format: str = "png", 
    scale: float = 2.0,
    options: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> List[str]:
    """Renderiza páginas selecionadas do PDF em imagens de alta definição usando pypdfium2."""
    out_dir = ensure_dir(Path(output_dir))
    pdf = pdfium.PdfDocument(str(input_path))
    image_paths: List[str] = []
    
    stem = Path(input_path).stem
    ext = img_format.lower()
    if ext not in ["png", "jpg", "jpeg", "webp"]:
        ext = "png"

    total_pages = len(pdf)
    selected_indices = resolve_page_indices(total_pages, options)
    total_selected = len(selected_indices) if selected_indices else total_pages

    for i, page_idx in enumerate(selected_indices, start=1):
        if 0 <= page_idx < total_pages:
            if progress_callback:
                progress_callback(i, total_selected, f"Renderizando imagem da página {i} de {total_selected}...")
            page = pdf[page_idx]
            image = page.render(scale=scale).to_pil()
            
            img_filename = f"{stem}_pagina_{page_idx + 1:03d}.{ext}"
            img_path = out_dir / img_filename
            
            if ext in ["jpg", "jpeg"]:
                if image.mode != "RGB":
                    image = image.convert("RGB")
                image.save(str(img_path), "JPEG", quality=95)
            elif ext == "webp":
                image.save(str(img_path), "WEBP", quality=95)
            else:
                image.save(str(img_path), "PNG")

            image_paths.append(str(img_path))
            page.close()

    pdf.close()
    return image_paths


def convert_pdf_to_text(
    input_path: str | Path, 
    output_path: str | Path,
    options: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> str:
    """Extrai texto das páginas selecionadas do PDF com fluxo contínuo de parágrafos sem quebras de linha duras."""
    text_content = []
    
    try:
        import pdfplumber
        with pdfplumber.open(str(input_path)) as pdf:
            total_pages = len(pdf.pages)
            selected_indices = resolve_page_indices(total_pages, options)
            total_selected = len(selected_indices) if selected_indices else total_pages
            for i, idx in enumerate(selected_indices, start=1):
                if 0 <= idx < total_pages:
                    if progress_callback:
                        progress_callback(i, total_selected, f"Extraindo texto da página {i} de {total_selected}...")
                    page = pdf.pages[idx]
                    raw_text = page.extract_text(layout=False) or ""
                    if not raw_text.strip():
                        continue

                    raw_lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
                    paragraphs = []
                    curr_p = []

                    for line in raw_lines:
                        if curr_p and curr_p[-1].endswith("-"):
                            curr_p[-1] = curr_p[-1][:-1] + line
                        else:
                            curr_p.append(line)

                        if line.endswith((".", "!", "?", ":")):
                            paragraphs.append(clean_inline_text(" ".join(curr_p)))
                            curr_p = []

                    if curr_p:
                        paragraphs.append(clean_inline_text(" ".join(curr_p)))

                    page_clean = "\n\n".join(paragraphs).strip()
                    if page_clean:
                        text_content.append(f"--- [Página {idx + 1}] ---\n{page_clean}\n")
    except Exception:
        pass

    if not text_content:
        try:
            from pypdf import PdfReader
            reader = PdfReader(str(input_path))
            total_pages = len(reader.pages)
            selected_indices = resolve_page_indices(total_pages, options)
            total_selected = len(selected_indices) if selected_indices else total_pages
            for i, idx in enumerate(selected_indices, start=1):
                if 0 <= idx < total_pages:
                    if progress_callback:
                        progress_callback(i, total_selected, f"Lendo texto da página {i} de {total_selected}...")
                    page = reader.pages[idx]
                    page_text = page.extract_text() or ""
                    clean_text = clean_inline_text(page_text)
                    if clean_text:
                        text_content.append(f"--- [Página {idx + 1}] ---\n{clean_text}\n")
        except Exception:
            pass

    if not text_content:
        text_content.append("--- [Documento] ---\nNenhum texto detectado nas páginas do PDF.")

    full_text = "\n".join(text_content)

    with open(str(output_path), "w", encoding="utf-8") as f:
        f.write(full_text)
    return str(output_path)


# =========================================================================
# 2. CONVERTER PARA PDF
# =========================================================================

def convert_word_to_pdf(input_path: str | Path, output_path: str | Path) -> str:
    """Converte documento DOCX para PDF com ReportLab preservando formatação e larguras."""
    doc = docx.Document(str(input_path))
    
    page_width, page_height = A4
    margin = 36  # ~0.5 inch
    usable_width = page_width - (2 * margin)

    pdf_doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=margin,
        leftMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )
    
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle(
        'DocxNormal',
        parent=styles['Normal'],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#1e293b")
    )
    h1_style = ParagraphStyle(
        'DocxH1',
        parent=styles['Heading1'],
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#1e293b"),
        spaceBefore=8,
        spaceAfter=6
    )
    h2_style = ParagraphStyle(
        'DocxH2',
        parent=styles['Heading2'],
        fontSize=13,
        leading=17,
        textColor=colors.HexColor("#334155"),
        spaceBefore=6,
        spaceAfter=4
    )

    story = []

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        style_name = p.style.name.lower() if p.style and p.style.name else ""
        if 'heading 1' in style_name or 'título 1' in style_name:
            story.append(Paragraph(html_escape(text), h1_style))
        elif 'heading 2' in style_name or 'título 2' in style_name:
            story.append(Paragraph(html_escape(text), h2_style))
        else:
            story.append(Paragraph(html_escape(text), normal_style))
            story.append(Spacer(1, 4))

    for table in doc.tables:
        raw_rows = []
        for row in table.rows:
            row_data = [html_escape(cell.text.strip()) for cell in row.cells]
            raw_rows.append(row_data)

        if not raw_rows:
            continue

        num_cols = max(len(r) for r in raw_rows)
        if num_cols == 0:
            continue

        # Calcula larguras de colunas proporcionais ao conteúdo
        col_max_lens = [max((len(r[c]) if c < len(r) else 1) for r in raw_rows) for c in range(num_cols)]
        total_lens = max(sum(col_max_lens), 1)
        col_widths = [(max(l, 3) / total_lens) * usable_width for l in col_max_lens]

        # Formata células com Paragraph para quebra de linha suave
        formatted_table = []
        for r_idx, row in enumerate(raw_rows):
            formatted_row = []
            for c_idx in range(num_cols):
                val = row[c_idx] if c_idx < len(row) else ""
                st = normal_style
                formatted_row.append(Paragraph(val, st))
            formatted_table.append(formatted_row)

        rl_table = Table(formatted_table, colWidths=col_widths, repeatRows=1)
        rl_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0066EB')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F8FAFC'), colors.white]),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ]))
        story.append(Spacer(1, 8))
        story.append(rl_table)
        story.append(Spacer(1, 8))

    if not story:
        story.append(Paragraph("Documento vazio.", normal_style))

    pdf_doc.build(story)
    return str(output_path)


def convert_excel_to_pdf(input_path: str | Path, output_path: str | Path) -> str:
    """
    Converte planilhas (.xlsx, .xls, .csv) em relatório PDF com
    orientação dinâmica (Paisagem/Retrato), larguras proporcionais e sem corte de colunas.
    """
    input_p = Path(input_path)
    all_rows: List[List[str]] = []

    # 1. Leitura do arquivo original
    if input_p.suffix.lower() in [".csv", ".txt"]:
        with open(str(input_p), "r", encoding="utf-8", errors="ignore") as f:
            sample = f.read(4096)
            delimiter = ";" if ";" in sample else ("," if "," in sample else "\t")
            f.seek(0)
            reader = csv.reader(f, delimiter=delimiter)
            for row in reader:
                if any(str(c).strip() for c in row):
                    all_rows.append([str(c).strip() for c in row])
    else:
        wb = openpyxl.load_workbook(str(input_p), data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() for cell in row):
                all_rows.append([str(cell if cell is not None else '').strip() for cell in row])

    if not all_rows:
        all_rows = [["Nenhum dado encontrado na planilha."]]

    num_cols = max(len(r) for r in all_rows)
    
    # 2. Escolha dinâmica da orientação da página:
    is_landscape = num_cols > 5
    page_size = landscape(A4) if is_landscape else A4
    page_width, page_height = page_size
    margin = 24  # Margem compacta e elegante
    usable_width = page_width - (2 * margin)

    pdf_doc = SimpleDocTemplate(
        str(output_path),
        pagesize=page_size,
        rightMargin=margin,
        leftMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'ExcelTitle',
        parent=styles['Heading2'],
        fontSize=13,
        leading=16,
        textColor=colors.HexColor('#0066EB'),
        spaceAfter=10
    )
    
    font_size = 7.5 if num_cols > 8 else (8.5 if num_cols > 5 else 9.5)
    leading_size = font_size + 2.5

    cell_style = ParagraphStyle(
        'ExcelCell',
        parent=styles['Normal'],
        fontSize=font_size,
        leading=leading_size,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_LEFT
    )
    cell_style_num = ParagraphStyle(
        'ExcelCellNum',
        parent=styles['Normal'],
        fontSize=font_size,
        leading=leading_size,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_RIGHT
    )
    cell_style_center = ParagraphStyle(
        'ExcelCellCenter',
        parent=styles['Normal'],
        fontSize=font_size,
        leading=leading_size,
        textColor=colors.HexColor('#1E293B'),
        alignment=TA_CENTER
    )
    header_style = ParagraphStyle(
        'ExcelHeader',
        parent=styles['Normal'],
        fontSize=font_size,
        leading=leading_size,
        textColor=colors.white,
        fontName='Helvetica-Bold',
        alignment=TA_CENTER
    )

    story = [
        Paragraph(f"<b>Relatório: {html_escape(input_p.stem)}</b>", title_style),
        Spacer(1, 4)
    ]

    # 3. Cálculo de largura proporcional exata para cada coluna
    col_max_lens = []
    for c_idx in range(num_cols):
        max_l = 4
        for r in all_rows:
            val = r[c_idx] if c_idx < len(r) else ""
            if len(val) > max_l:
                max_l = len(val)
        col_max_lens.append(min(max_l, 45))  # Limita peso de colunas gigantes

    total_len_weights = sum(col_max_lens)
    col_widths = [(w / total_len_weights) * usable_width for w in col_max_lens]

    # 4. Formatação de cada célula com Paragraph para quebra de linha elegante
    formatted_table = []
    for r_idx, row in enumerate(all_rows):
        formatted_row = []
        is_header = (r_idx == 0)
        for c_idx in range(num_cols):
            val = row[c_idx] if c_idx < len(row) else ""
            val_escaped = html_escape(val)
            
            if is_header:
                st = header_style
            else:
                clean_v = val.replace(".", "").replace(",", "").replace("-", "").replace("R$", "").replace("%", "").strip()
                if clean_v.isdigit():
                    st = cell_style_num
                elif len(val) <= 4:
                    st = cell_style_center
                else:
                    st = cell_style

            formatted_row.append(Paragraph(val_escaped, st))
        formatted_table.append(formatted_row)

    # 5. Estilo profissional de tabela
    rl_table = Table(formatted_table, colWidths=col_widths, repeatRows=1)
    rl_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0066EB')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#F8FAFC'), colors.white]),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(rl_table)

    pdf_doc.build(story)
    return str(output_path)


def convert_images_to_pdf(image_paths: List[str | Path], output_path: str | Path) -> str:
    """Combina uma ou múltiplas imagens (.jpg, .png, .webp, etc.) em um único PDF."""
    if not image_paths:
        raise ValueError("Nenhuma imagem fornecida para conversão.")

    opened_images: List[Image.Image] = []
    try:
        for p in image_paths:
            img = Image.open(str(p))
            if img.mode in ("RGBA", "LA", "P"):
                rgb_img = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                rgb_img.paste(img, mask=img.split()[-1] if img.mode == "RGBA" else None)
                opened_images.append(rgb_img)
            elif img.mode != "RGB":
                opened_images.append(img.convert("RGB"))
            else:
                opened_images.append(img)

        if opened_images:
            first_img = opened_images[0]
            rest_imgs = opened_images[1:] if len(opened_images) > 1 else []
            first_img.save(str(output_path), "PDF", resolution=100.0, save_all=True, append_images=rest_imgs)
    finally:
        for img in opened_images:
            try:
                img.close()
            except Exception:
                pass

    return str(output_path)


# =========================================================================
# 3. ORGANIZAR E OTIMIZAR PDF (MERGE, SPLIT, COMPRESS, ROTATE, LOCK)
# =========================================================================

def merge_pdfs(pdf_paths: List[str | Path], output_path: str | Path) -> str:
    """Junta múltiplos arquivos PDF em um único documento ordenado."""
    if not pdf_paths:
        raise ValueError("Nenhum arquivo PDF fornecido para junção.")

    writer = PdfWriter()
    for pdf_file in pdf_paths:
        reader = PdfReader(str(pdf_file))
        for page in reader.pages:
            writer.add_page(page)

    with open(str(output_path), "wb") as f_out:
        writer.write(f_out)
    return str(output_path)


def split_pdf(
    input_path: str | Path, 
    output_dir: str | Path, 
    pages_mode: str = "all", 
    custom_range: Optional[str] = None,
    options: Optional[Dict[str, Any]] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> List[str]:
    """Divide um PDF em páginas separadas ou em intervalos específicos."""
    out_dir = ensure_dir(Path(output_dir))
    reader = PdfReader(str(input_path))
    total_pages = len(reader.pages)
    created_files: List[str] = []

    opts = dict(options or {})
    if pages_mode:
        opts["pageRangeMode"] = pages_mode
    if custom_range:
        opts["pageRange"] = custom_range

    selected_indices = resolve_page_indices(total_pages, opts, custom_range)
    total_selected = len(selected_indices) if selected_indices else total_pages

    if (pages_mode == "range" or opts.get("pageRangeMode") == "range") and len(selected_indices) < total_pages:
        writer = PdfWriter()
        for i, idx in enumerate(selected_indices, start=1):
            if 0 <= idx < total_pages:
                if progress_callback:
                    progress_callback(i, total_selected, f"Extraindo página {idx + 1} ({i} de {total_selected})...")
                writer.add_page(reader.pages[idx])

        clean_range = str(custom_range or opts.get("pageRange") or f"{selected_indices[0]+1}a{selected_indices[-1]+1}").replace(',', '_').replace('-', 'a').replace(' ', '')
        out_file = out_dir / f"extraido_paginas_{clean_range}.pdf"
        with open(str(out_file), "wb") as f_out:
            writer.write(f_out)
        created_files.append(str(out_file))

    else:
        stem = Path(input_path).stem
        for i, idx in enumerate(selected_indices, start=1):
            if 0 <= idx < total_pages:
                if progress_callback:
                    progress_callback(i, total_selected, f"Dividindo página {idx + 1} ({i} de {total_selected})...")
                writer = PdfWriter()
                writer.add_page(reader.pages[idx])
                out_file = out_dir / f"{stem}_pagina_{idx + 1:03d}.pdf"
                with open(str(out_file), "wb") as f_out:
                    writer.write(f_out)
                created_files.append(str(out_file))

    return created_files



def compress_pdf(input_path: str | Path, output_path: str | Path, quality: str = "medium") -> str:
    """Otimiza e reduz o tamanho do PDF comprimindo fluxos de dados."""
    reader = PdfReader(str(input_path))
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    for page in writer.pages:
        try:
            page.compress_content_streams()
        except Exception:
            pass

    try:
        writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
    except Exception:
        pass

    with open(str(output_path), "wb") as f_out:
        writer.write(f_out)
    return str(output_path)


def rotate_pdf(
    input_path: str | Path, 
    output_path: str | Path, 
    angle: int = 90,
    options: Optional[Dict[str, Any]] = None
) -> str:
    """Gira todas as páginas ou páginas selecionadas do PDF no ângulo especificado (90, 180, 270 graus)."""
    reader = PdfReader(str(input_path))
    writer = PdfWriter()
    total_pages = len(reader.pages)
    selected_indices = resolve_page_indices(total_pages, options)
    selected_set = set(selected_indices)

    for idx, page in enumerate(reader.pages):
        if idx in selected_set:
            page.rotate(angle)
        writer.add_page(page)

    with open(str(output_path), "wb") as f_out:
        writer.write(f_out)
    return str(output_path)


def protect_pdf(input_path: str | Path, output_path: str | Path, password: str) -> str:
    """Adiciona proteção por senha a um documento PDF."""
    if not password:
        raise ValueError("A senha não pode estar em branco.")

    reader = PdfReader(str(input_path))
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(user_password=password, owner_password=password, use_128bit=True)

    with open(str(output_path), "wb") as f_out:
        writer.write(f_out)
    return str(output_path)


def unlock_pdf(input_path: str | Path, output_path: str | Path, password: str) -> str:
    """Remove a senha de um PDF protegido."""
    reader = PdfReader(str(input_path))
    if reader.is_encrypted:
        success = reader.decrypt(password)
        if not success:
            raise ValueError("Senha incorreta. Não foi possível desbloquear o PDF.")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with open(str(output_path), "wb") as f_out:
        writer.write(f_out)
    return str(output_path)


def create_zip_from_files(files: List[str | Path], zip_output_path: str | Path) -> str:
    """Empacota uma lista de arquivos em um arquivo ZIP."""
    with zipfile.ZipFile(str(zip_output_path), "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            p = Path(f)
            zf.write(str(p), arcname=p.name)
    return str(zip_output_path)
