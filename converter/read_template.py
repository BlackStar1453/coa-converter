#!/usr/bin/env python3
"""
read_template.py — 读取模板结构并输出 JSON
AI-First 架构的第一步：让 AI 了解模板布局，用于生成 template_mapping。

用法:
    python3 read_template.py <template_path>
    # 输出 JSON 到 stdout
"""

import sys
import json
import logging

from template_detector import (
    detect_template_layout, TemplateLayout, _col_letter
)

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)


def layout_to_json(layout: TemplateLayout) -> dict:
    """将 TemplateLayout 序列化为 JSON-serializable dict"""
    result = {
        "format": layout.format,
        "template_type": layout.template_type,
    }

    # 头部字段
    header_fields = {}
    for key, mapping in layout.header_fields.items():
        entry = {}
        if mapping.label_pos and mapping.label_pos.row:
            entry["label_cell"] = f"{_col_letter(mapping.label_pos.col)}{mapping.label_pos.row}"
        if mapping.value_pos:
            if mapping.value_pos.cell_ref:
                entry["value_cell"] = mapping.value_pos.cell_ref
            elif mapping.value_pos.row:
                entry["value_cell"] = f"{_col_letter(mapping.value_pos.col)}{mapping.value_pos.row}"
        header_fields[key] = entry
    result["header_fields"] = header_fields

    # 数据表头行
    result["data_header_row"] = layout.data_header_row

    # 数据列
    data_columns = {}
    for key, col_num in layout.data_columns.items():
        data_columns[key] = _col_letter(col_num)
    result["data_columns"] = data_columns

    # 数据行标签（item_key → row + 原始标签文本）
    row_labels = []
    for item_key, row_num in sorted(layout.table_rows.items(), key=lambda x: x[1]):
        row_labels.append({
            "row": row_num,
            "key": item_key,
        })
    result["row_labels"] = row_labels

    # Packing & Storage 位置
    if layout.packing_position:
        pos = layout.packing_position
        result["packing_cell"] = pos.cell_ref or f"{_col_letter(pos.col)}{pos.row}"
    else:
        result["packing_cell"] = None

    if layout.storage_position:
        pos = layout.storage_position
        result["storage_cell"] = pos.cell_ref or f"{_col_letter(pos.col)}{pos.row}"
    else:
        result["storage_cell"] = None

    # DOCX 特有字段
    if layout.format == "docx":
        result["docx_tables"] = layout.docx_tables
        result["product_name_positions"] = [
            {"para_index": p.para_index} for p in layout.product_name_positions
        ]

    return result


def main():
    if len(sys.argv) < 2:
        print("用法: python3 read_template.py <template_path>", file=sys.stderr)
        sys.exit(1)

    template_path = sys.argv[1]
    layout = detect_template_layout(template_path)
    result = layout_to_json(layout)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
