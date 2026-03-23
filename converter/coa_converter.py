#!/usr/bin/env python3
"""
COA PDF to Template Converter
从供应商COA PDF中提取检测数据，填充到Key In Nutrition的模板中（支持 XLSX 和 DOCX）。

使用方式:
    python coa_converter.py <pdf_path> <template_path> [output_path]

依赖:
    pip install pdfplumber openpyxl PyMuPDF python-docx
"""

import sys
import os
import re
import logging
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field

import pdfplumber

# ============ 日志配置 ============
logging.basicConfig(
    level=logging.INFO,
    format='[COA-Converter] %(levelname)s: %(message)s'
)
logger = logging.getLogger(__name__)


# ============ 文本空格修复 ============

def fix_spacing(text: str) -> str:
    """
    修复PDF提取文本中缺失的空格。
    仅做安全的、不会破坏已有单词的修复。
    """
    if not text:
        return ""

    # 移除 · 前缀（分组标题常见）
    text = re.sub(r'^[·•]\s*', '', text)

    # 在小写字母+大写字母之间插入空格: "ProductName" → "Product Name"
    text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)

    # 在小写字母+数字之间插入空格: "pass80" → "pass 80"
    text = re.sub(r'([a-z])(\d)', r'\1 \2', text)

    # 在数字+小写字母之间插入空格: "80mesh" → "80 mesh"（但保留 "10ppm", "5cfu" 等单位前缀）
    text = re.sub(r'(\d)(mesh|pass)', r'\1 \2', text, flags=re.IGNORECASE)

    # 方法代码中连字符后的字母+数字: "OMA991" → "OMA 991"
    text = re.sub(r'(OMA|RI)(\d)', r'\1 \2', text)

    # 在逗号后添加空格，但不对千分位数字（逗号后跟3位数字）
    text = re.sub(r',(?!\d{3}(?:\D|$))(\S)', r', \1', text)

    # 修复 "N.W.:25" → "N.W.: 25"，但不拆ratio如 "10:1"
    text = re.sub(r'(?<!\d):(\d)', r': \1', text)

    # 压缩多余空格
    text = re.sub(r'\s+', ' ', text).strip()

    return text


def fix_row_spacing(row: list) -> list:
    """对表格行的所有单元格应用空格修复"""
    return [fix_spacing(cell) if cell else "" for cell in row]


# ============ 数据结构 ============

@dataclass
class COAData:
    """COA文档提取结果"""
    header: Dict[str, str] = field(default_factory=dict)
    assay: Dict[str, str] = field(default_factory=dict)
    analytical_items: List[Dict[str, str]] = field(default_factory=list)
    microbiology_items: List[Dict[str, str]] = field(default_factory=list)
    packing_storage: str = ""
    unmapped_items: List[Dict[str, str]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# ============ 字段映射配置 ============

# PDF头部字段 → 标准化键名
HEADER_FIELD_ALIASES = {
    "product name": "product_name",
    "product": "product_name",
    "botanical latin name": "botanical_name",
    "botanical name": "botanical_name",
    "botanical source": "botanical_name",
    "plant part": "plant_part",
    "part used": "plant_part",
    "batch number": "batch_number",
    "batch no": "batch_number",
    "batch no.": "batch_number",
    "lot number": "batch_number",
    "lot no": "batch_number",
    "lot no.": "batch_number",
    "country of origin": "country",
    "origin": "country",
    "quantity": "quantity",
    "pack size": "quantity",
    "manufacture date": "mfg_date",
    "manufacturing date": "mfg_date",
    "date of manufacture": "mfg_date",
    "mfg date": "mfg_date",
    "production date": "mfg_date",
    "expire date": "exp_date",
    "expiration date": "exp_date",
    "expiry date": "exp_date",
    "date of expiry": "exp_date",
    "exp date": "exp_date",
    "issue date": "issue_date",
    "solvent": "solvent",
}

# 注意：HEADER_TO_CELL 已移除，改用 template_detector 自动检测布局

# PDF检测项名称 → 标准化键（用于匹配XLSX模板行）
# 注意：包含空格修复前后的各种变体
ITEM_NAME_NORMALIZE = {
    # Assay相关
    "assay": "assay",
    "ratio": "assay",
    # PDF特有项（XLSX无对应行，识别后归入unmapped）
    "extract ratio": "extract_ratio",
    "identification": "identification",
    # Analytical Data
    "appearance": "appearance",
    "description": "appearance",
    "color": "color",
    "odor": "odor",
    "taste": "taste",
    "loss on drying": "lod",
    "moisture": "lod",
    "residue on ignition": "ash",
    "ash": "ash",
    "ash content": "ash",
    "particle size": "sieve",
    "sieve analysis": "sieve",
    "sieve test": "sieve",
    "sieve test (passes through)": "sieve",
    "mesh": "sieve",
    "40 mesh": "sieve_40",
    "80 mesh": "sieve",
    "mesh size": "sieve",
    "bulk density": "bulk_density",
    "heavy metals": "heavy_metals",
    "heavy metal": "heavy_metals",
    "total heavy metals": "heavy_metals",
    "total heavy metal": "heavy_metals",
    "arsenic (as)": "as",
    "arsenic": "as",
    "as": "as",
    "lead (pb)": "pb",
    "lead": "pb",
    "pb": "pb",
    "cadmium (cd)": "cd",
    "cadmium(cd)": "cd",
    "cadmium": "cd",
    "cd": "cd",
    "mercury (hg)": "hg",
    "mercury(hg)": "hg",
    "mercury": "hg",
    "hg": "hg",
    "pesticides residue": "pesticides",
    "pesticide residue": "pesticides",
    "pesticide residues": "pesticides",
    "pesticides": "pesticides",
    "foreign matter": "foreign_matter",
    "elemental impurities": "elemental_impurities",
    # Microbiology
    "total plate count": "tpc",
    "total aerobic count": "tpc",
    "total aerobic microbial count": "tpc",
    "total bacterial count": "tpc",
    "total yeast & mold": "yeast_mold",
    "total yeast and mold": "yeast_mold",
    "total yeasts and molds count": "yeast_mold",
    "total yeasts and molds": "yeast_mold",
    "yeast & mold": "yeast_mold",
    "yeast and mold": "yeast_mold",
    "total fungal count": "yeast_mold",
    "fungal count": "yeast_mold",
    "escherchia coli": "e_coli",
    "escherichia coli": "e_coli",
    "e. coli": "e_coli",
    "e. coli.": "e_coli",
    "e.coli": "e_coli",
    "salmonella species": "salmonella",
    "salmonella": "salmonella",
    "staphylococcus aurreus": "s_aureus",
    "staphylococcus aureus": "s_aureus",
    "staphylococcus": "s_aureus",
    "s. aureus": "s_aureus",
    "pseudomonas aeruginosa": "pseudomonas",
    "bile-tolerant gram-negative bacteria count": "bile_tolerant",
    "coliform": "coliform",
}

# PDF特有的检测项（XLSX模板中没有对应行）
PDF_ONLY_ITEMS = {"extract_ratio", "identification", "bulk_density", "pesticides", "coliform",
                   "foreign_matter", "elemental_impurities", "pseudomonas", "bile_tolerant",
                   "sieve_40"}

# 注意：XLSX_ROW_KEYS 已移除，改用 template_detector 自动检测行号

# 分组标题行识别模式（支持有空格和无空格版本）
GROUP_PATTERNS = {
    "chemical": re.compile(r'chemical\s*physical\s*control', re.IGNORECASE),
    "physical": re.compile(r'^Physical$', re.IGNORECASE),
    "others": re.compile(r'^Others$', re.IGNORECASE),
    "microbiology": re.compile(r'microbiol', re.IGNORECASE),
    "additional": re.compile(r'additional\s*information', re.IGNORECASE),
    "packing": re.compile(r'pack\w*\s*(?:&|and)?\s*stor', re.IGNORECASE),
    "shelf": re.compile(r'shelf\s*l\s*i\s*f\s*e', re.IGNORECASE),
}

# "Additional Information" section中可提取的key-value项
ADDITIONAL_INFO_KEYS = {
    "country of origin": "country",
    "storage condition": "storage_condition",
}


# ============ 日期格式转换 ============

MONTH_MAP = {
    'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
    'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
    'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12',
    'january': '01', 'february': '02', 'march': '03',
    'april': '04', 'june': '06', 'july': '07',
    'august': '08', 'september': '09', 'october': '10',
    'november': '11', 'december': '12',
}


def convert_date(date_str: str) -> str:
    """
    将各种日期格式转换为 YYYY.MM.DD
    支持: "Feb.03, 2026", "2026-02-03", "02/03/2026", "Feb 03, 2026", "2026.02.03"
    """
    if not date_str or not date_str.strip():
        return ""

    date_str = date_str.strip()
    logger.debug(f'[日期转换] 输入: {date_str}')

    # 已经是目标格式 YYYY.MM.DD
    m = re.match(r'^(\d{4})\.(\d{1,2})\.(\d{1,2})$', date_str)
    if m:
        return date_str

    # 格式: "Feb.03, 2026" 或 "Feb 03, 2026" 或 "February 03, 2026"
    m = re.match(r'([A-Za-z]+)\.?\s*(\d{1,2}),?\s*(\d{4})', date_str)
    if m:
        month_str = m.group(1).lower()
        day = m.group(2).zfill(2)
        year = m.group(3)
        month = MONTH_MAP.get(month_str, '')
        if month:
            result = f"{year}.{month}.{day}"
            logger.debug(f'[日期转换] 结果: {result}')
            return result

    # 格式: "03 Feb 2026" 或 "03 February 2026"
    m = re.match(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', date_str)
    if m:
        day = m.group(1).zfill(2)
        month_str = m.group(2).lower()
        year = m.group(3)
        month = MONTH_MAP.get(month_str, '')
        if month:
            return f"{year}.{month}.{day}"

    # 格式: "Feb-2028" 或 "Mar-2025" (Month-Year only, no day)
    m = re.match(r'^([A-Za-z]+)-(\d{4})$', date_str)
    if m:
        month_str = m.group(1).lower()
        year = m.group(2)
        month = MONTH_MAP.get(month_str, '')
        if month:
            result = f"{year}.{month}"
            logger.debug(f'[日期转换] Month-Year格式结果: {result}')
            return result

    # 格式: "2026-02-03"
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', date_str)
    if m:
        return f"{m.group(1)}.{m.group(2).zfill(2)}.{m.group(3).zfill(2)}"

    # 格式: "02/03/2026" (MM/DD/YYYY)
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', date_str)
    if m:
        return f"{m.group(3)}.{m.group(1).zfill(2)}.{m.group(2).zfill(2)}"

    logger.warning(f'[日期转换] 无法识别日期格式: "{date_str}"，保持原样')
    return date_str


# ============ 数据格式转换 ============

def _build_nospace_lookup():
    """构建去空格的反向查找表，用于模糊匹配"""
    lookup = {}
    for key, value in ITEM_NAME_NORMALIZE.items():
        # 去除所有空格、标点的纯字母数字版本
        nospace = re.sub(r'[\s\.\(\)&]', '', key.lower())
        lookup[nospace] = value
    return lookup

_NOSPACE_LOOKUP = _build_nospace_lookup()


def normalize_item_name(name: str) -> str:
    """将检测项名称标准化为键值，支持模糊匹配"""
    if not name:
        return ""
    cleaned = name.strip().lower()
    # 移除多余空格
    cleaned = re.sub(r'\s+', ' ', cleaned)

    # 1. 精确匹配
    if cleaned in ITEM_NAME_NORMALIZE:
        return ITEM_NAME_NORMALIZE[cleaned]

    # 2. 去除括号内容后匹配
    no_paren = re.sub(r'\s*\(.*?\)\s*', '', cleaned).strip()
    if no_paren in ITEM_NAME_NORMALIZE:
        return ITEM_NAME_NORMALIZE[no_paren]

    # 3. 去空格模糊匹配（核心：解决PDF提取的空格缺失问题）
    nospace = re.sub(r'[\s\.\(\)&]', '', cleaned)
    if nospace in _NOSPACE_LOOKUP:
        return _NOSPACE_LOOKUP[nospace]

    return ""


def is_group_header(item_text: str) -> Optional[str]:
    """检查文本是否是分组标题行，返回分组类型或None"""
    if not item_text:
        return None
    # 同时对去空格版本进行匹配
    text_nospace = re.sub(r'\s+', '', item_text.lower())
    for group_type, pattern in GROUP_PATTERNS.items():
        if pattern.search(item_text) or pattern.search(text_nospace):
            return group_type
    return None


# ============ PDF数据提取 ============

def extract_from_pdf(pdf_path: str) -> COAData:
    """
    从COA PDF中提取所有结构化数据
    """
    logger.info(f'[提取] 开始处理PDF: {pdf_path}')
    coa = COAData()

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                logger.info(f'[提取] 处理第 {page_num + 1} 页')

                # 查找表格
                tables = page.find_tables(table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 5,
                    "join_tolerance": 5,
                })

                if not tables:
                    logger.info('[提取] 有线表格未找到，尝试文本策略')
                    tables = page.find_tables(table_settings={
                        "vertical_strategy": "text",
                        "horizontal_strategy": "text",
                        "snap_tolerance": 5,
                        "min_words_vertical": 2,
                        "min_words_horizontal": 1,
                    })

                if not tables:
                    logger.warning(f'[提取] 第 {page_num + 1} 页未检测到表格')
                    continue

                for table in tables:
                    rows = table.extract()
                    _parse_table_rows(rows, coa)

    except Exception as e:
        logger.error(f'[提取] pdfplumber提取失败: {e}，尝试PyMuPDF降级')
        _fallback_pymupdf(pdf_path, coa)

    # 检查表格提取是否有效——如果头部或数据全为空，降级到words策略
    has_data = bool(coa.header) and (bool(coa.analytical_items) or bool(coa.microbiology_items) or bool(coa.assay))
    if not has_data:
        logger.info('[提取] 表格提取数据不足，降级到基于word位置的策略')
        coa_words = COAData()
        try:
            _extract_by_words(pdf_path, coa_words)
            # 用words策略的结果替换
            if coa_words.header or coa_words.analytical_items or coa_words.microbiology_items:
                coa = coa_words
        except Exception as e:
            logger.error(f'[提取-Words] words策略失败: {e}')

    logger.info(f'[提取] 完成 - 头部字段: {len(coa.header)}, '
                f'分析项: {len(coa.analytical_items)}, '
                f'微生物项: {len(coa.microbiology_items)}, '
                f'未映射项: {len(coa.unmapped_items)}')
    return coa


def _parse_table_rows(rows: List, coa: COAData):
    """解析表格行数据，分类到COAData各字段"""
    current_section = "general"  # general / analytical / microbiology / footer
    found_data_header = False

    for row in rows:
        if row is None:
            continue

        # 清理单元格并修复空格
        cells = fix_row_spacing(row)
        non_empty = [c for c in cells if c]

        if not non_empty:
            continue

        # 检测表格表头行（Item | Specification | Result | Test Methods | Parameters）
        first_lower = cells[0].lower() if cells[0] else ""
        if any(kw in first_lower for kw in ['item', 'items of analysis', 'parameter']):
            found_data_header = True
            continue

        # 优先检查 packing/shelf 分组标题（可能出现在无数据表头的独立表格中，如第2页）
        group = is_group_header(cells[0])
        if group in ("packing", "shelf"):
            if group == "packing":
                current_section = "footer"
                packing_text = " ".join(c for c in cells[1:] if c)
                if packing_text:
                    coa.packing_storage = packing_text
            continue

        # 解析键值对行（头部信息区域）
        if not found_data_header:
            _parse_header_row(cells, coa)
            continue

        # 检查分组标题
        # group already computed above; re-check for non-packing/shelf types
        if not group:
            group = is_group_header(cells[0])
        if group in ("chemical", "physical", "others"):
            current_section = "analytical"
            continue
        elif group == "microbiology":
            current_section = "microbiology"
            continue
        elif group == "additional":
            current_section = "additional"
            continue
        elif group == "packing":
            current_section = "footer"
            # 提取包装信息
            packing_text = " ".join(c for c in cells[1:] if c)
            if packing_text:
                coa.packing_storage = packing_text
            continue
        elif group == "shelf":
            # Shelf Life行也属于footer
            continue

        # 检测Assay行
        item_key = normalize_item_name(cells[0])
        if item_key == "assay":
            coa.assay = _make_item_dict(cells)
            current_section = "analytical"  # Assay后面通常是分析数据
            continue

        # 普通数据行 - 至少要有Item和一个其他值
        if len(non_empty) < 2 and not item_key:
            # 可能是分组标题或独立文本行
            if cells[0] and not is_group_header(cells[0]):
                # 可能是另一种分组标题格式
                text_lower = cells[0].lower()
                if 'analytical' in text_lower or 'physical' in text_lower:
                    current_section = "analytical"
                elif 'micro' in text_lower:
                    current_section = "microbiology"
                elif 'additional' in text_lower or 'packing' in text_lower:
                    current_section = "footer"
            continue

        item_dict = _make_item_dict(cells)

        if current_section == "footer":
            # Footer区域的内容追加到packing_storage
            if cells[0] and any(c for c in cells[1:] if c):
                text = " ".join(c for c in cells if c)
                if coa.packing_storage:
                    coa.packing_storage += " " + text
                else:
                    coa.packing_storage = text
            continue

        if not item_key:
            coa.unmapped_items.append(item_dict)
            coa.warnings.append(f'未识别的检测项: "{cells[0]}"')
            continue

        # "extract_ratio" 特殊处理：将其作为 Assay/Ratio 数据
        if item_key == "extract_ratio" and not coa.assay:
            coa.assay = _make_item_dict(cells)
            logger.info(f'[提取] Extract Ratio 映射为 Assay: {coa.assay}')
            continue

        # PDF特有项（XLSX模板中没有对应行），记录但不报错
        if item_key in PDF_ONLY_ITEMS:
            coa.unmapped_items.append(item_dict)
            logger.info(f'[提取] PDF特有项(无XLSX映射): {cells[0]} = {item_dict.get("result", "")}')
            continue

        # 分配到对应的section
        if item_key in ("tpc", "yeast_mold", "e_coli", "salmonella", "s_aureus"):
            coa.microbiology_items.append(item_dict)
            current_section = "microbiology"
        else:
            coa.analytical_items.append(item_dict)


def _parse_header_row(cells: List[str], coa: COAData):
    """解析头部键值对行"""
    # COA头部通常是 [key, value, key, value] 四列格式
    # 按照成对的方式解析: (cells[0],cells[1]), (cells[2],cells[3])
    pairs = []
    i = 0
    while i < len(cells) - 1:
        if cells[i] and cells[i + 1]:
            pairs.append((cells[i], cells[i + 1]))
            i += 2
        elif cells[i]:
            i += 1
        else:
            i += 1

    for key_text, value_text in pairs:
        key_lower = key_text.lower().strip()
        key_nospace = re.sub(r'\s+', '', key_lower)
        matched_key = None
        for alias, std_key in HEADER_FIELD_ALIASES.items():
            alias_nospace = re.sub(r'\s+', '', alias)
            if alias in key_lower or key_lower in alias or \
               alias_nospace in key_nospace or key_nospace in alias_nospace:
                matched_key = std_key
                break

        if matched_key:
            coa.header[matched_key] = value_text.strip()
            logger.debug(f'[提取-头部] {matched_key} = {value_text.strip()}')


def _make_item_dict(cells: List[str]) -> Dict[str, str]:
    """将单元格列表转为标准化的字典"""
    result = {
        "item": cells[0] if len(cells) > 0 else "",
        "specification": "",
        "result": "",
        "method": "",
    }
    # 处理不同列数的情况（有些PDF是3列，有些是4列）
    if len(cells) == 4:
        result["specification"] = cells[1]
        result["result"] = cells[2]
        result["method"] = cells[3]
    elif len(cells) == 3:
        result["specification"] = cells[1]
        result["result"] = cells[2]
    elif len(cells) >= 5:
        # 可能有合并列的情况，取第2-4列
        result["specification"] = cells[1]
        result["result"] = cells[2]
        result["method"] = cells[3]
    return result


def _extract_by_words(pdf_path: str, coa: COAData):
    """
    基于word位置的文本提取策略（适用于无线表格PDF）。
    通过分析words的x/y坐标自动检测列边界并重建表格结构。
    """
    logger.info('[提取-Words] 使用基于word位置的提取策略')

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            if not words:
                continue

            # 按y坐标分组
            y_groups = {}
            for w in words:
                y_key = round(w['top'])
                # 合并相近y坐标（±3px视为同一行）
                merged = False
                for existing_y in list(y_groups.keys()):
                    if abs(y_key - existing_y) <= 3:
                        y_groups[existing_y].append(w)
                        merged = True
                        break
                if not merged:
                    y_groups[y_key] = [w]

            # 找到数据表头行（必须同时包含 SPECIFICATION 和 RESULT/METHOD）
            header_y = None
            data_start_y = None
            for y in sorted(y_groups.keys()):
                row_text = ' '.join(w['text'] for w in sorted(y_groups[y], key=lambda x: x['x0']))
                row_upper = row_text.upper()
                # 必须包含 SPECIFICATION 或 (RESULT + METHOD)，排除 "CERTIFICATE OF ANALYSIS"
                has_spec = 'SPECIFICATION' in row_upper
                has_result = 'RESULT' in row_upper
                has_method = 'METHOD' in row_upper
                has_analysis = 'ANALYSIS' in row_upper or 'ITEMS' in row_upper or 'PARAMETER' in row_upper
                if (has_spec or (has_result and has_method)) and (has_analysis or has_spec):
                    header_y = y
                    continue
                if header_y and data_start_y is None:
                    data_start_y = y
                    break

            if not header_y:
                logger.warning(f'[提取-Words] 第 {page_num + 1} 页未找到表格表头')
                continue

            # 从表头行检测列边界（x坐标）
            # 表头如 "Parameters  Specifications  Result  Reference"
            # 注意 "TEST METHOD" 是两个word但属于同一列，用较大间距(>60px)区分列
            header_words = sorted(y_groups[header_y], key=lambda x: x['x0'])
            header_col_centers = []
            for w in header_words:
                x0 = w['x0']
                if not header_col_centers or x0 - header_col_centers[-1] > 60:
                    header_col_centers.append(x0)

            # 确保恰好4列；如果超过4列，合并最后几列
            if len(header_col_centers) > 4:
                header_col_centers = header_col_centers[:4]

            # 用数据行的实际x坐标来校准列边界
            # 表头位置可能居中显示，与实际数据列不对齐
            # 策略：使用相邻表头列的中点作为列分隔边界，并用数据行最小x作为第一列起始
            data_x0_samples = []
            for y in sorted(y_groups.keys()):
                if y <= header_y:
                    continue
                row_w = sorted(y_groups[y], key=lambda x: x['x0'])
                if row_w:
                    data_x0_samples.append(row_w[0]['x0'])

            if data_x0_samples and len(header_col_centers) >= 2:
                min_data_x = min(data_x0_samples)
                # 如果数据起始位置比表头第一列靠左较多，说明表头居中
                if header_col_centers[0] - min_data_x > 20:
                    # 使用中点分隔策略：
                    # 第一列起始 = 数据行最小x
                    # 后续列的分隔线 = 相邻表头列中点
                    col_starts = [min_data_x]
                    for ci in range(1, len(header_col_centers)):
                        midpoint = (header_col_centers[ci - 1] + header_col_centers[ci]) / 2
                        col_starts.append(midpoint)
                    logger.info(f'[提取-Words] 使用中点分隔策略校准列边界: {[f"{x:.1f}" for x in col_starts]}')
                else:
                    col_starts = header_col_centers
            else:
                col_starts = header_col_centers

            logger.info(f'[提取-Words] 检测到 {len(col_starts)} 列，x边界: {col_starts}')

            # 用全页文本解析头部（比word位置更可靠）
            full_text = page.extract_text()
            if full_text:
                _parse_header_from_text(full_text, coa)

            # 处理数据行（表头之后）
            found_data_header = True
            current_section = "general"

            for y in sorted(y_groups.keys()):
                if y <= header_y:
                    continue

                row_words = sorted(y_groups[y], key=lambda x: x['x0'])

                # 根据列边界将words分配到各列
                cells = _words_to_cells(row_words, col_starts)
                cells = fix_row_spacing(cells)
                non_empty = [c for c in cells if c]

                if not non_empty:
                    continue

                # 检查分组标题
                group = is_group_header(cells[0])
                if group in ("chemical", "physical", "others"):
                    current_section = "analytical"
                    continue
                elif group == "microbiology":
                    current_section = "microbiology"
                    continue
                elif group == "additional":
                    current_section = "additional"
                    continue
                elif group == "packing":
                    current_section = "footer"
                    full_text_line = ' '.join(w['text'] for w in row_words)
                    packing_parts = re.split(r'(?:Packing\s*&?\s*Storage)\s*', full_text_line, flags=re.IGNORECASE)
                    if len(packing_parts) > 1:
                        coa.packing_storage = packing_parts[1].strip()
                    continue
                elif group == "shelf":
                    continue

                # "Additional Information" section: 提取key-value对
                if current_section == "additional":
                    full_text_line = ' '.join(w['text'] for w in row_words)
                    # 跳过备注、免责等非结构化文本行
                    fl_lower = full_text_line.lower()
                    if fl_lower.startswith('*') or 'disclaimer' in fl_lower or 'remarks' in fl_lower or 'analysed' in fl_lower or 'approved' in fl_lower:
                        continue
                    # 尝试匹配已知的Additional Info key
                    matched_additional = False
                    for ai_pattern, ai_key in ADDITIONAL_INFO_KEYS.items():
                        if ai_pattern in fl_lower:
                            # 提取value（key之后的文本）
                            idx = fl_lower.index(ai_pattern) + len(ai_pattern)
                            ai_value = full_text_line[idx:].strip()
                            if ai_key == "country" and ai_value:
                                coa.header["country"] = ai_value
                                logger.info(f'[提取-Words-附加] country = {ai_value}')
                            elif ai_key == "storage_condition" and ai_value:
                                if not coa.packing_storage:
                                    coa.packing_storage = ai_value
                                else:
                                    coa.packing_storage += " " + ai_value
                                logger.info(f'[提取-Words-附加] storage = {ai_value}')
                            matched_additional = True
                            break
                    if not matched_additional:
                        coa.unmapped_items.append(_make_item_dict(cells))
                    continue

                # 继续拼接packing/storage的后续行
                if current_section == "footer":
                    full_text_line = ' '.join(w['text'] for w in row_words)
                    if coa.packing_storage:
                        coa.packing_storage += " " + full_text_line
                    else:
                        coa.packing_storage = full_text_line
                    continue

                # Assay / 数据行
                item_key = normalize_item_name(cells[0])
                if item_key == "assay":
                    coa.assay = _make_item_dict(cells)
                    current_section = "analytical"
                    continue

                if len(non_empty) < 2 and not item_key:
                    text_lower = cells[0].lower()
                    if 'analytical' in text_lower or 'physical' in text_lower:
                        current_section = "analytical"
                    elif 'micro' in text_lower:
                        current_section = "microbiology"
                    elif 'additional' in text_lower:
                        current_section = "additional"
                    continue

                item_dict = _make_item_dict(cells)

                if not item_key:
                    coa.unmapped_items.append(item_dict)
                    coa.warnings.append(f'未识别的检测项: "{cells[0]}"')
                    continue

                # "extract_ratio" 特殊处理：将其作为 Assay/Ratio 数据
                if item_key == "extract_ratio" and not coa.assay:
                    coa.assay = _make_item_dict(cells)
                    logger.info(f'[提取-Words] Extract Ratio 映射为 Assay: {coa.assay}')
                    continue

                if item_key in PDF_ONLY_ITEMS:
                    coa.unmapped_items.append(item_dict)
                    logger.info(f'[提取-Words] PDF特有项: {cells[0]} = {item_dict.get("result", "")}')
                    continue

                if item_key in ("tpc", "yeast_mold", "e_coli", "salmonella", "s_aureus"):
                    coa.microbiology_items.append(item_dict)
                    current_section = "microbiology"
                else:
                    coa.analytical_items.append(item_dict)


def _parse_header_from_text(full_text: str, coa: COAData):
    """从全页文本中用正则提取头部字段。
    直接匹配已知的key pattern及其后续值。
    """
    lines = full_text.split('\n')

    # 合并所有头部行为一个文本（从第一个已知key到ANALYSIS/SPECIFICATION之前）
    header_text = ""
    in_header = False
    for line in lines:
        line_upper = line.strip().upper()
        # 检测头部开始
        if any(kw in line_upper for kw in ['PRODUCT NAME', 'PRODUCT:', 'BATCH NUMBER', 'BATCH NO']):
            in_header = True
        # 检测头部结束
        if 'SPECIFICATION' in line_upper or 'ITEMS OF ANALYSIS' in line_upper:
            break
        if in_header:
            header_text += " " + line.strip()

    if not header_text:
        return

    logger.info(f'[提取-Text-头部] 头部文本: {header_text[:200]}...')

    # 定义提取模式：每个key pattern后面跟着值（直到下一个已知key）
    # 按在文本中出现的顺序定义所有可能的key
    # 包含所有可能出现在头部区域的key（含非映射key作为值边界标记）
    key_patterns = [
        (r'Product\s*Name\s*', 'product_name'),
        (r'Botanical\s*(?:Latin\s*)?(?:Name|Source)\s*', 'botanical_name'),
        (r'Batch\s*(?:Number|No\.?)\s*', 'batch_number'),
        (r'Lot\s*(?:Number|No\.?)\s*', 'batch_number'),
        (r'Part\s*Used\s*', 'plant_part'),
        (r'Plant\s*Part\s*', 'plant_part'),
        (r'Pack\s*Size\s*', 'quantity'),
        (r'Quantity\s*', 'quantity'),
        (r'Country\s*of\s*Origin\s*', 'country'),
        (r'Origin\s*', 'country'),
        (r'Date\s*of\s*Analysis\s*', '_date_of_analysis'),
        (r'Date\s*of\s*Manufacturing\s*', 'mfg_date'),
        (r'(?:Manufacture|Manufacturing|Production|Mfg\.?)\s*Date\s*', 'mfg_date'),
        (r'Date\s*of\s*Manufacture\s*', 'mfg_date'),
        (r'(?:Expire|Expiration|Expiry|Exp\.?)\s*Date\s*', 'exp_date'),
        (r'Date\s*of\s*Expiry\s*', 'exp_date'),
        (r'Issue\s*Date\s*', 'issue_date'),
        (r'Solvent\s*', 'solvent'),
        # 非映射key（仅作为值边界标记，用 _ 前缀区分）
        (r'T\.?\s*R\.?\s*No\.?\s*', '_tr_no'),
        (r'Category\s*', '_category'),
        (r'GMO\s*Status\s*', '_gmo_status'),
        (r'Carrier\s*', '_carrier'),
    ]

    # 找到所有key在文本中的位置
    matches = []
    for pattern, std_key in key_patterns:
        for m in re.finditer(pattern, header_text, re.IGNORECASE):
            matches.append((m.start(), m.end(), std_key))

    # 按位置排序
    matches.sort(key=lambda x: x[0])

    # 提取每个key的值（从key结束到下一个key开始）
    for i, (start, end, std_key) in enumerate(matches):
        # 跳过非映射标记key（以 _ 开头）
        if std_key.startswith('_'):
            continue

        # 值的结束位置 = 下一个key的开始位置，或文本末尾
        val_end = matches[i + 1][0] if i + 1 < len(matches) else len(header_text)
        value = header_text[end:val_end].strip()

        # 清理值末尾的多余内容
        value = value.strip()

        if value:
            coa.header[std_key] = value
            logger.info(f'[提取-Text-头部] {std_key} = {value}')


def _parse_header_zone(header_rows: list, coa: COAData):
    """基于x坐标区域检测的头部解析。
    自动识别4区域(key1/val1/key2/val2)的x边界。
    """
    # 收集所有header words的x0坐标
    all_x0 = []
    for row_words in header_rows:
        for w in row_words:
            all_x0.append(round(w['x0']))

    # 用聚类找出主要的x起始位置
    # COA头部通常是4区域布局，区域间距约60-100px
    # 同区域内word间距一般<50px
    all_x0.sort()
    clusters = []
    for x in all_x0:
        if not clusters or x - clusters[-1][-1] > 50:
            clusters.append([x])
        else:
            clusters[-1].append(x)

    # 如果检测到超过4个区域，合并最接近的相邻区域直到剩4个
    while len(clusters) > 4:
        min_gap = float('inf')
        merge_idx = 0
        for i in range(len(clusters) - 1):
            gap = min(clusters[i + 1]) - max(clusters[i])
            if gap < min_gap:
                min_gap = gap
                merge_idx = i
        clusters[merge_idx] = clusters[merge_idx] + clusters[merge_idx + 1]
        del clusters[merge_idx + 1]

    zone_starts = [min(c) for c in clusters]
    logger.info(f'[提取-Words-头部] 检测到 {len(zone_starts)} 个区域，x边界: {zone_starts}')

    if len(zone_starts) < 2:
        # 无法分区，退回逐行解析
        for row_words in header_rows:
            _parse_header_from_words(row_words, coa)
        return

    # 对每行按zone分配words
    for row_words in header_rows:
        zones = [[] for _ in zone_starts]
        for w in row_words:
            x0 = w['x0']
            best_zone = 0
            for zi, zs in enumerate(zone_starts):
                if x0 >= zs:
                    best_zone = zi
            zones[best_zone].append(w['text'])

        zone_texts = [' '.join(words) for words in zones]

        # 按 key-value 配对解析
        i = 0
        while i < len(zone_texts) - 1:
            key_text = zone_texts[i].strip()
            val_text = zone_texts[i + 1].strip()
            if not key_text:
                i += 1
                continue

            key_lower = key_text.lower()
            key_nospace = re.sub(r'\s+', '', key_lower)
            matched_key = None
            for alias, std_key in HEADER_FIELD_ALIASES.items():
                alias_nospace = re.sub(r'\s+', '', alias)
                if alias in key_lower or key_lower in alias or \
                   alias_nospace in key_nospace or key_nospace in alias_nospace:
                    matched_key = std_key
                    break

            if matched_key and val_text:
                coa.header[matched_key] = val_text
                logger.info(f'[提取-Words-头部] {matched_key} = {val_text}')
                i += 2
            else:
                i += 1


def _parse_header_from_words(row_words: list, coa: COAData):
    """从word列表解析头部键值对。
    COA头部通常是4段布局: [key1] [value1] [key2] [value2]
    通过检测x坐标的大间距(>40px)来分段。
    """
    if not row_words:
        return

    # 按x坐标间距分段（间距>40px视为新段）
    segments = []
    current_seg = [row_words[0]]
    for i in range(1, len(row_words)):
        prev_end = current_seg[-1]['x0'] + len(current_seg[-1]['text']) * 7
        gap = row_words[i]['x0'] - prev_end
        if gap > 30:
            segments.append(' '.join(w['text'] for w in current_seg))
            current_seg = [row_words[i]]
        else:
            current_seg.append(row_words[i])
    segments.append(' '.join(w['text'] for w in current_seg))

    # 每两个segment组成一对 key-value
    i = 0
    while i < len(segments) - 1:
        key_text = segments[i]
        val_text = segments[i + 1]

        key_lower = key_text.lower().strip()
        key_nospace = re.sub(r'\s+', '', key_lower)
        matched_key = None
        for alias, std_key in HEADER_FIELD_ALIASES.items():
            alias_nospace = re.sub(r'\s+', '', alias)
            if alias in key_lower or key_lower in alias or \
               alias_nospace in key_nospace or key_nospace in alias_nospace:
                matched_key = std_key
                break

        if matched_key:
            coa.header[matched_key] = val_text.strip()
            logger.info(f'[提取-Words-头部] {matched_key} = {val_text.strip()}')
            i += 2
        else:
            i += 1


def _words_to_cells(row_words: list, col_starts: list) -> list:
    """根据列边界将words分配到各列，返回4列cells"""
    if len(col_starts) < 2:
        # 列边界不足，退回全文
        return [' '.join(w['text'] for w in row_words), "", "", ""]

    # 每个word根据x0分配到最近的列
    # 使用"最后一个 <= x0 的列边界"策略，无额外容差
    col_contents = [[] for _ in col_starts]
    for w in row_words:
        x0 = w['x0']
        best_col = 0
        for ci, cs in enumerate(col_starts):
            if x0 >= cs:
                best_col = ci
        col_contents[best_col].append(w['text'])

    cells = [' '.join(words) for words in col_contents]

    # 确保返回4列
    while len(cells) < 4:
        cells.append("")

    # 如果超过4列，合并前几列为item，保留后3列
    if len(cells) > 4:
        # item = cells[0], spec = cells[1..n-2], result = cells[-2], method = cells[-1]
        cells = [cells[0], ' '.join(cells[1:-2]), cells[-2], cells[-1]]

    return cells[:4]


def _fallback_pymupdf(pdf_path: str, coa: COAData):
    """PyMuPDF降级提取方案"""
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            logger.error('[降级] PyMuPDF未安装，无法执行降级提取')
            return

    logger.info('[降级] 使用PyMuPDF进行表格提取')
    doc = pymupdf.open(pdf_path)
    for page in doc:
        tabs = page.find_tables()
        for tab in tabs:
            rows = []
            for row in tab.extract():
                rows.append(row)
            _parse_table_rows(rows, coa)
    doc.close()


# ============ 模板填充（委托给 xlsx_filler / docx_filler）============

def fill_template(coa: COAData, template_path: str, output_path: str):
    """
    根据模板格式自动检测布局并填充数据。
    支持 .xlsx 和 .docx 模板。
    返回检测到的 TemplateLayout 供后续验证使用。
    """
    from template_detector import detect_template_layout
    layout = detect_template_layout(template_path)

    if layout.format == "xlsx":
        from xlsx_filler import fill_xlsx
        fill_xlsx(coa, layout, template_path, output_path)
    elif layout.format == "docx":
        from docx_filler import fill_docx
        fill_docx(coa, layout, template_path, output_path)
    else:
        raise ValueError(f'不支持的模板格式: {layout.format}')

    return layout


# ============ 数据验证 ============

def validate_coa(coa: COAData) -> List[str]:
    """
    验证提取数据的完整性和合理性
    返回警告信息列表
    """
    warnings = list(coa.warnings)

    # 检查必填头部字段
    required_headers = ["product_name", "batch_number", "mfg_date", "exp_date"]
    for key in required_headers:
        if key not in coa.header or not coa.header[key]:
            warnings.append(f'必填头部字段缺失: {key}')

    # 检查日期合理性
    mfg = coa.header.get("mfg_date", "")
    exp = coa.header.get("exp_date", "")
    if mfg and exp:
        mfg_conv = convert_date(mfg)
        exp_conv = convert_date(exp)
        if mfg_conv > exp_conv:
            warnings.append(f'日期异常: 生产日期({mfg_conv}) > 过期日期({exp_conv})')

    # 检查Assay是否存在
    if not coa.assay:
        warnings.append('Assay数据缺失')

    # 检查关键检测项
    all_items = coa.analytical_items + coa.microbiology_items
    all_keys = set()
    for item in all_items:
        key = normalize_item_name(item["item"])
        if key:
            all_keys.add(key)

    critical_keys = {"lod", "heavy_metals", "tpc", "e_coli", "salmonella"}
    missing_critical = critical_keys - all_keys
    if missing_critical:
        warnings.append(f'缺少关键检测项: {", ".join(missing_critical)}')

    # 报告未映射项
    if coa.unmapped_items:
        names = [item["item"] for item in coa.unmapped_items]
        warnings.append(f'未映射的检测项({len(names)}): {", ".join(names)}')

    return warnings


# ============ 主函数 ============

def convert_coa(pdf_path: str, template_path: str, output_path: Optional[str] = None) -> str:
    """
    COA转换主函数

    Args:
        pdf_path: PDF文件路径
        template_path: 模板文件路径（.xlsx 或 .docx）
        output_path: 输出文件路径（可选，默认自动生成）

    Returns:
        输出文件路径
    """
    # 验证输入文件
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f'PDF文件不存在: {pdf_path}')
    if not os.path.exists(template_path):
        raise FileNotFoundError(f'模板文件不存在: {template_path}')

    # 检测模板格式
    template_ext = os.path.splitext(template_path)[1].lower()
    if template_ext not in ('.xlsx', '.docx'):
        raise ValueError(f'不支持的模板格式: {template_ext} (仅支持 .xlsx 和 .docx)')

    # 生成默认输出路径
    if not output_path:
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_dir = os.path.dirname(pdf_path)
        output_path = os.path.join(output_dir, f'{pdf_name}{template_ext}')

    logger.info('=' * 60)
    logger.info(f'COA转换开始')
    logger.info(f'  PDF: {pdf_path}')
    logger.info(f'  模板: {template_path}')
    logger.info(f'  输出: {output_path}')
    logger.info('=' * 60)

    # Step 1: 提取PDF数据
    coa = extract_from_pdf(pdf_path)

    # Step 2: 数据验证
    warnings = validate_coa(coa)
    if warnings:
        logger.warning(f'数据验证发现 {len(warnings)} 个警告:')
        for w in warnings:
            logger.warning(f'  - {w}')

    # Step 3: 检测布局并填充模板
    layout = fill_template(coa, template_path, output_path)

    # Step 4: 输出验证
    verification = None
    if layout.format == "xlsx" and layout.template_type in ('coa_assay', 'coa_powder', 'coa_ratio'):
        from xlsx_filler import verify_xlsx_output
        verification = verify_xlsx_output(coa, layout, output_path, template_path=template_path)
        if verification["failed"] > 0:
            logger.warning(f'[验证] 发现 {verification["failed"]} 个字段不一致，正确率: {verification["accuracy"]:.1%}')
        else:
            logger.info(f'[验证] 全部通过，正确率: {verification["accuracy"]:.1%}')
    elif layout.format == "docx":
        from docx_filler import verify_docx_output
        verification = verify_docx_output(coa, layout, output_path)
        if verification["failed"] > 0:
            logger.warning(f'[验证] 发现 {verification["failed"]} 个字段不一致，正确率: {verification["accuracy"]:.1%}')
        else:
            logger.info(f'[验证] 全部通过，正确率: {verification["accuracy"]:.1%}')

    logger.info('=' * 60)
    logger.info(f'COA转换完成: {output_path}')
    if warnings:
        logger.info(f'共 {len(warnings)} 个警告，请检查输出文件')
    if verification and verification["failed"] > 0:
        logger.warning(f'输出验证: {verification["failed"]}/{verification["total"]} 个字段不一致')
    logger.info('=' * 60)

    return output_path


def main():
    if len(sys.argv) < 3:
        print("用法: python coa_converter.py <pdf_path> <template_path> [output_path]")
        print("示例: python coa_converter.py coa.pdf template.xlsx output.xlsx")
        print("       python coa_converter.py coa.pdf template.docx output.docx")
        sys.exit(1)

    pdf_path = sys.argv[1]
    template_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else None

    try:
        result = convert_coa(pdf_path, template_path, output_path)
        print(f"\n转换成功: {result}")
    except Exception as e:
        logger.error(f'转换失败: {e}')
        sys.exit(1)


if __name__ == "__main__":
    main()
