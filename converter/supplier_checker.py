#!/usr/bin/env python3
"""
Supplier Checker
检查 PDF 是否来自已验证的供应商，判断是否需要 AI 验证。
"""

import json
import os
import logging
from datetime import datetime

import pdfplumber

logger = logging.getLogger(__name__)

REGISTRY_PATH = os.path.join(os.path.dirname(__file__), "supplier_registry.json")


def load_registry() -> dict:
    """加载供应商注册表"""
    if not os.path.exists(REGISTRY_PATH):
        return {"suppliers": []}
    with open(REGISTRY_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_registry(registry: dict):
    """保存供应商注册表"""
    with open(REGISTRY_PATH, 'w', encoding='utf-8') as f:
        json.dump(registry, f, ensure_ascii=False, indent=2)


def extract_pdf_text_sample(pdf_path: str, max_chars: int = 2000) -> str:
    """提取 PDF 前几页的文本用于签名匹配"""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:2]:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
                if len(text) >= max_chars:
                    break
    except Exception as e:
        logger.error(f'[供应商检查] PDF读取失败: {e}')
    return text[:max_chars]


def check_supplier(pdf_path: str) -> dict:
    """
    检查 PDF 是否来自已验证的供应商。

    Returns:
        {
            "known": True/False,
            "supplier_id": "lipond" or None,
            "supplier_name": "Lipond" or None,
            "needs_ai_verification": True/False,
            "message": "说明信息"
        }
    """
    registry = load_registry()
    text_sample = extract_pdf_text_sample(pdf_path)
    text_lower = text_sample.lower()

    pdf_name = os.path.basename(pdf_path)

    for supplier in registry.get("suppliers", []):
        signatures = supplier.get("signatures", [])
        matched = sum(1 for sig in signatures if sig.lower() in text_lower)

        if matched >= 1:
            # 检查这个具体 PDF 是否已验证过
            verified_pdfs = supplier.get("verified_pdfs", [])
            already_verified = pdf_name in verified_pdfs

            return {
                "known": True,
                "supplier_id": supplier["id"],
                "supplier_name": supplier["name"],
                "needs_ai_verification": False,
                "already_verified_pdf": already_verified,
                "message": f"已验证供应商: {supplier['name']} "
                           f"(签名匹配 {matched}/{len(signatures)}, "
                           f"历史正确率: {supplier.get('accuracy', 'N/A')})"
            }

    return {
        "known": False,
        "supplier_id": None,
        "supplier_name": None,
        "needs_ai_verification": True,
        "already_verified_pdf": False,
        "message": f"未知供应商 — 需要 AI 验证首次转换结果"
    }


def register_supplier(supplier_id: str, supplier_name: str, signatures: list,
                       pdf_format: str, pdf_name: str, accuracy: str,
                       notes: str = ""):
    """将新供应商添加到注册表"""
    registry = load_registry()

    # 检查是否已存在
    for supplier in registry.get("suppliers", []):
        if supplier["id"] == supplier_id:
            # 已存在，追加 PDF 记录
            if pdf_name not in supplier.get("verified_pdfs", []):
                supplier.setdefault("verified_pdfs", []).append(pdf_name)
            supplier["accuracy"] = accuracy
            supplier["verified_date"] = datetime.now().strftime("%Y-%m-%d")
            if notes:
                supplier["notes"] = notes
            save_registry(registry)
            logger.info(f'[供应商注册] 更新已有供应商: {supplier_name}, 新增PDF: {pdf_name}')
            return

    # 新增供应商
    new_entry = {
        "id": supplier_id,
        "name": supplier_name,
        "signatures": signatures,
        "format": pdf_format,
        "verified_date": datetime.now().strftime("%Y-%m-%d"),
        "verified_pdfs": [pdf_name],
        "accuracy": accuracy,
        "notes": notes,
    }
    registry.setdefault("suppliers", []).append(new_entry)
    save_registry(registry)
    logger.info(f'[供应商注册] 新增供应商: {supplier_name}')


# ============ CLI ============

def main():
    import sys
    if len(sys.argv) < 2:
        print("用法: python supplier_checker.py <pdf_path>")
        print("检查 PDF 是否来自已验证的供应商")
        sys.exit(1)

    logging.basicConfig(level=logging.INFO, format='[供应商检查] %(message)s')
    pdf_path = sys.argv[1]

    if not os.path.exists(pdf_path):
        print(f"文件不存在: {pdf_path}")
        sys.exit(1)

    result = check_supplier(pdf_path)
    print(f"\n{'='*50}")
    print(f"PDF: {os.path.basename(pdf_path)}")
    print(f"状态: {'✅ 已知供应商' if result['known'] else '⚠️ 未知供应商'}")
    print(f"供应商: {result.get('supplier_name') or '未识别'}")
    print(f"需要AI验证: {'是' if result['needs_ai_verification'] else '否'}")
    print(f"说明: {result['message']}")
    print(f"{'='*50}")

    # 退出码: 0=已知供应商(可直接转换), 1=未知供应商(需AI验证)
    sys.exit(0 if result['known'] else 1)


if __name__ == "__main__":
    main()
