# XLSX 模板详细布局规则

## 1. Key In COA - Assay

**适用场景**：含 Assay（含量测定）行的标准 COA，如 "5% Withanolides"

### 布局

**公司信息区（不可修改）：**
- `C1`: Key In Nutrition 公司信息（电话/地址/网站/邮箱）

**头部字段（Row 4-7）：**

| 标签位置 | 值位置 | 值来源 | 格式要求 |
|---------|--------|--------|----------|
| `A4` Product Name | `C4` | PDF: Product Name | 原文 |
| `A5` Botanical Source | `C5` | PDF: Botanical Latin Name / Botanical Source | 原文 |
| `E5` Part Used | `F5` | PDF: Plant Part / Part Used | 原文 |
| `A6` Batch Number | `C6` | PDF: Batch Number / Lot No. | 原文 |
| `E6` Country of Origin | `F6` | PDF: Country of Origin | 原文 |
| `A7` Manufacturing Date | `C7` | PDF: Mfg/Production Date | **YYYY.MM.DD**（仅 Month-Year 则 YYYY.MM） |
| `E7` Expiration Date | `F7` | PDF: Exp/Expiry Date | **YYYY.MM.DD**（同上） |

**数据表头（Row 9，不可修改）：**
- `A9`=Items of Analysis, `C9`=Specification, `E9`=Result, `F9`=Test Method

**数据行：**

| 行号 | A列（标签） | C列（Spec） | E列（Result） | F列（Method） |
|------|------------|-------------|---------------|---------------|
| 10 | **Assay** | PDF Assay 规格 | PDF Assay 结果 | PDF Assay 方法 |
| 11 | Analytical Data | — section header，不填 — | | |
| 12 | Appearance | PDF | PDF | PDF |
| 13 | Odor | PDF | PDF | PDF |
| 14 | Taste | PDF | PDF | PDF |
| 15 | Loss on drying | PDF | PDF | PDF |
| 16 | Ash | PDF | PDF | PDF |
| 17 | Sieve Analysis | PDF | PDF | PDF |
| 18 | Heavy metals | PDF | PDF | PDF |
| 19 | Pb | PDF | PDF | PDF |
| 20 | As | PDF | PDF | PDF |
| 21 | Cd | PDF | PDF | PDF |
| 22 | Hg | PDF | PDF | PDF |
| 23 | Microbiological Test | — section header — | | |
| 24 | Total Plate count | PDF | PDF | PDF |
| 25 | Yeast & Mold | PDF | PDF | PDF |
| 26 | E. Coli. | PDF | PDF | PDF |
| 27 | Salmonella | PDF | PDF | PDF |
| 28 | Staphylococcus aureus | PDF | PDF | PDF |
| 30 | Additional information | — section header — | | |
| 31 | Packing and Storage | `C31` = Packing 信息 | | |
| 32 | | `C32` = Storage 信息 | | |

---

## 2. Key In COA - Ratio

**适用场景**：含 Ratio（提取比例）行的提取物 COA，如 "10:1"

与 Assay 模板**完全相同**，仅 Row 10 不同：

| 行号 | A列 | 说明 |
|------|-----|------|
| 10 | **Ratio** | C10=比例值（如 "10:1"），E10=结果，F10=方法 |

头部（Row 4-7）和其余数据行（Row 11-32）与 Assay 模板一致。

---

## 3. Key In COA - Powder

**适用场景**：无 Assay/Ratio 行的粉末类 COA

**注意**：没有 Assay/Ratio 行，所有数据行号比 Assay/Ratio 模板**少 1 行**。

头部（Row 4-7）与 Assay 模板一致。

**数据行：**

| 行号 | A列（标签） | C列（Spec） | E列（Result） | F列（Method） |
|------|------------|-------------|---------------|---------------|
| 10 | Analytical Data | — section header — | | |
| 11 | Appearance | PDF | PDF | PDF |
| 12 | Odor | PDF | PDF | PDF |
| 13 | Taste | PDF | PDF | PDF |
| 14 | Loss on drying | PDF | PDF | PDF |
| 15 | Ash | PDF | PDF | PDF |
| 16 | Sieve Analysis | PDF | PDF | PDF |
| 17 | Heavy metals | PDF | PDF | PDF |
| 18 | Pb | PDF | PDF | PDF |
| 19 | As | PDF | PDF | PDF |
| 20 | Cd | PDF | PDF | PDF |
| 21 | Hg | PDF | PDF | PDF |
| 22 | Microbiological Test | — section header — | | |
| 23 | Total Plate count | PDF | PDF | PDF |
| 24 | Yeast & Mold | PDF | PDF | PDF |
| 25 | E. Coli. | PDF | PDF | PDF |
| 26 | Salmonella | PDF | PDF | PDF |
| 27 | Staphylococcus aureus | PDF | PDF | PDF |
| 29 | Additional information | — section header — | | |
| 30 | Packing and Storage | `C30` = Packing 信息 | | |
| 31 | | `C31` = Storage 信息 | | |

---

## 4. Allergen -

**适用场景**：过敏原声明文档，仅需替换产品名

### 布局

| 位置 | 内容 | 填充规则 |
|------|------|----------|
| `B1` | Key In Nutrition 公司信息 | 不可修改 |
| `B2` | "Allergen Statement" | 不可修改 |
| `B3` | "Product Name: " | **追加** PDF 产品名（如 → "Product Name: Rhodiola Rosea P.E."） |

**过敏原数据表（Row 4-18，不可修改）：**

| 行号 | B列（Items） | C列 | D列 | E列 | F列 | G列 |
|------|-------------|-----|-----|-----|-----|-----|
| 4 | *表头* | Contain | Absent | Same production line | Same production facility | Comments |
| 5 | Cereals containing gluten... | | √ | | | |
| 6 | Crustaceans | | √ | | | |
| 7 | Eggs | | √ | | | |
| 8 | Fish | | √ | | | |
| 9 | Peanuts | | √ | | | |
| 10 | Soybeans | | √ | | | |
| 11 | Milk (including lactose) | | √ | | | |
| 12 | Nuts | | √ | | | |
| 13 | Celery | | √ | | | |
| 14 | Mustard | | √ | | | |
| 15 | Sesame seeds | | √ | | | |
| 16 | Sulphur dioxide and sulphites | | √ | | | |
| 17 | Lupin | | √ | | | |
| 18 | Mollusks | | √ | | | |

**填充规则**：仅替换 `B3` 中的产品名。14 项过敏原数据默认全部 "Absent（√）"，保持模板默认值不修改。

---

## 5. Flow Chart

**适用场景**：生产流程图，仅需替换标题中的产品名

### 布局

| 位置 | 内容 | 填充规则 |
|------|------|----------|
| `C1` | Key In Nutrition 公司信息 | 不可修改 |
| `A2` | 标题（如 "Cordyceps Powder Flow Chart"） | **替换为** "{产品名} Flow Chart" |

**填充规则**：仅替换 `A2` 标题中的产品名。流程图内容保持模板默认值不修改。
