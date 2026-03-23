---
name: coa-to-template
description: >
  COA to template filler. MUST USE for any of these scenarios:
  (1) User mentions "COA" or "Certificate of Analysis" or "分析证书" or "检测报告" or "检测数据" with any template/模板/filling task.
  (2) User wants to fill Assay, Ratio, Powder, Allergen, Flow Chart, Composition Statement, CS, Nutrition Info, or Safety Data Sheet templates from PDF data.
  (3) User asks to extract data from a lab/supplier PDF into XLSX or DOCX templates.
  Keywords that signal this skill: COA, 分析证书, 检测数据, Assay模板, Ratio template, 过敏原声明, 流程图, 成分声明, 营养信息, 安全数据表, "fill template from PDF", "PDF转模板".
  NEVER use for: generic PDF reading/merging/converting, script writing, web development, or image conversion.
---

# COA PDF to Template Converter

> **平台适配**：本文档中 Steps 2-6 的命令示例以 bash 编写。在 Windows 上，Claude 应自动将 `python3` 替换为 `python`，将 Unix 路径（`~/`、`/tmp/`）替换为 Windows 等价路径（`$HOME\`、`$env:TEMP\`），将 bash 语法替换为 PowerShell 等价命令。

将供应商COA (Certificate of Analysis) PDF文件中的检测数据提取并填充到Key In Nutrition的模板中。支持 XLSX 和 DOCX 两种模板格式。

## 项目根目录检测（必须首先执行）

所有路径均基于项目根目录 `$PROJECT_ROOT`。在执行任何步骤前，必须先定位项目根目录：

```bash
# 自动检测项目根目录：从当前工作目录向上查找包含 converter/ 和 templates/ 的目录
PROJECT_ROOT="$(pwd)"
while [ "$PROJECT_ROOT" != "/" ]; do
  if [ -d "$PROJECT_ROOT/converter" ] && [ -d "$PROJECT_ROOT/templates" ] && [ -f "$PROJECT_ROOT/app.py" ]; then
    break
  fi
  PROJECT_ROOT="$(dirname "$PROJECT_ROOT")"
done

if [ "$PROJECT_ROOT" = "/" ]; then
  echo "错误：无法找到 coa-converter-web 项目根目录。请在项目目录内运行。"
  exit 1
fi

echo "项目根目录: $PROJECT_ROOT"
```

> **重要**：后续所有命令中的路径均使用 `$PROJECT_ROOT` 作为基准，不依赖任何固定安装位置。

## 使用方式

```
/coa-to-template <pdf_file_path> <template_path> [output_path]
```

## 目录结构

| 目录 | 用途 |
|------|------|
| `$PROJECT_ROOT/input/` | 存放待转换的 PDF 源文件 |
| `$PROJECT_ROOT/output/` | 存放转换后的输出文件 |
| `$PROJECT_ROOT/templates/` | 存放所有模板文件（XLSX + DOCX） |

## 参数说明

- `pdf_file_path`: COA PDF文件路径（必需，建议放在 `$PROJECT_ROOT/input/`）
- `template_path`: 模板文件路径，支持 `.xlsx` 或 `.docx`（必需，模板位于 `$PROJECT_ROOT/templates/`）
- `output_path`: 输出文件路径（可选，默认输出到 `$PROJECT_ROOT/output/`）

## 首次使用引导（自动执行）

首次运行时，skill 会自动完成环境初始化。**用户只需提供模板文件即可开始使用。**

### 自动完成的步骤

以下步骤在 Step 1 中自动执行，无需用户手动操作：

1. **定位项目根目录**：自动检测 `$PROJECT_ROOT`
2. **创建工作目录**：自动创建 `input/` 和 `output/` 子目录
3. **创建 Python 虚拟环境**：在 `$PROJECT_ROOT/.venv` 创建并安装依赖
4. **检测模板文件**：检查 `templates/` 目录是否有模板文件

### 用户需要手动完成的步骤

**将模板文件放到 `templates/` 目录**：

模板文件是 Key In Nutrition 公司内部标准模板，需由用户自行获取并放置到以下位置：

```
$PROJECT_ROOT/templates/
```

### 关于模板选择

**模板必须由用户在调用时明确指定**，skill 不会自动选择模板。

- 用户调用时需提供 `<template_path>` 参数
- **如果用户未指定模板**，skill 应**停止执行**并提醒用户：列出 `$PROJECT_ROOT/templates/` 目录中当前可用的模板文件，让用户选择一个
- **如果 `templates/` 目录为空**，skill 应**停止执行**并提示用户先将模板文件放入该目录

## 执行步骤

### Step 1. 环境初始化（自动检测项目 + 安装依赖）

**macOS / Linux：**
```bash
# 1a. 检测项目根目录（见上方"项目根目录检测"）
# 1b. 创建工作目录
mkdir -p "$PROJECT_ROOT/input" "$PROJECT_ROOT/output"

# 1c. 创建/激活 Python 虚拟环境
VENV_DIR="$PROJECT_ROOT/.venv"
if [ ! -d "$VENV_DIR" ]; then
  python3 -m venv "$VENV_DIR"
  source "$VENV_DIR/bin/activate"
  pip install -r "$PROJECT_ROOT/requirements.txt" -q
else
  source "$VENV_DIR/bin/activate"
fi

# 1d. 检查模板目录是否有文件
TEMPLATE_COUNT=$(ls "$PROJECT_ROOT/templates/"*.xlsx "$PROJECT_ROOT/templates/"*.docx 2>/dev/null | wc -l)
echo "模板文件数量: $TEMPLATE_COUNT"
```

**Windows (PowerShell)：**
```powershell
# 1a. 检测项目根目录
$ProjectRoot = (Get-Location).Path
while ($ProjectRoot -ne [System.IO.Path]::GetPathRoot($ProjectRoot)) {
  if ((Test-Path "$ProjectRoot\converter") -and (Test-Path "$ProjectRoot\templates") -and (Test-Path "$ProjectRoot\app.py")) { break }
  $ProjectRoot = Split-Path $ProjectRoot -Parent
}

# 1b. 创建工作目录
New-Item -ItemType Directory -Path "$ProjectRoot\input","$ProjectRoot\output" -Force | Out-Null

# 1c. 创建/激活 Python 虚拟环境
$VenvDir = "$ProjectRoot\.venv"
if (-not (Test-Path $VenvDir)) {
  python -m venv $VenvDir
  & "$VenvDir\Scripts\Activate.ps1"
  pip install -r "$ProjectRoot\requirements.txt" -q
} else {
  & "$VenvDir\Scripts\Activate.ps1"
}

# 1d. 检查模板目录是否有文件
$TEMPLATE_COUNT = (Get-ChildItem "$ProjectRoot\templates\*" -Include *.xlsx,*.docx -ErrorAction SilentlyContinue).Count
Write-Host "模板文件数量: $TEMPLATE_COUNT"
```

> **平台检测**：Claude 应根据当前操作系统自动选择对应的命令。

**如果 `TEMPLATE_COUNT` 为 0（templates 目录为空）**，停止执行并向用户显示以下提示：

```
⚠️ 首次使用：模板目录为空

请将模板文件（.xlsx / .docx）放入以下目录后重新运行：
  $PROJECT_ROOT/templates/

工作目录已自动创建：
  📂 $PROJECT_ROOT/templates/  ← 放置模板文件
  📂 $PROJECT_ROOT/input/      ← 放置待转换的 PDF 文件（批量处理时使用）
  📂 $PROJECT_ROOT/output/     ← 转换后的文件将输出到此目录
```

### Step 2. 参数检查与模板扫描
```bash
ARGS=($ARGUMENTS)
PDF_PATH="${ARGS[0]}"
TEMPLATE_PATH="${ARGS[1]}"
OUTPUT_PATH="${ARGS[2]:-}"
```

**如果用户未指定 `TEMPLATE_PATH`**，执行以下扫描并停止：

```bash
# 扫描 templates 目录中所有可用模板
echo "可用模板列表："
ls -1 "$PROJECT_ROOT/templates/"*.{xlsx,docx} 2>/dev/null | while read f; do
  echo "  $(basename "$f")"
done
```

停止执行，向用户展示扫描结果并提醒选择一个模板后重新运行。

**用户已指定模板后，继续执行：**
```bash
if [ -z "$OUTPUT_PATH" ]; then
  PDF_BASE=$(basename "$PDF_PATH" .pdf)
  TEMPLATE_EXT="${TEMPLATE_PATH##*.}"
  OUTPUT_PATH="$PROJECT_ROOT/output/${PDF_BASE}.${TEMPLATE_EXT}"
fi

# 供应商检查
python3 "$PROJECT_ROOT/converter/supplier_checker.py" "$PDF_PATH"
SUPPLIER_CHECK=$?
```

供应商检查结果决定后续流程：
- **退出码 0**（已知供应商）→ 转换后仍需执行验证（Step 4）
- **退出码 1**（未知供应商）→ 转换后执行验证（Step 4），验证通过后注册供应商（Step 5）

### Step 3. 运行转换
```bash
python3 "$PROJECT_ROOT/converter/coa_converter.py" "$PDF_PATH" "$TEMPLATE_PATH" "$OUTPUT_PATH"
```

### Step 4. AI 驱动验证与修复

转换完成后，程序会自动执行内嵌验证（`verify_xlsx_output`）。**但无论内嵌验证结果如何，都必须执行以下 AI 验证步骤。**

#### 4.1 读取三方文件

AI 必须同时读取以下三份文件进行三方对比：

1. **PDF 源文件**（ground truth）：使用 Read 工具直接读取 PDF 获取文本数据
2. **模板文件**（结构权威）：使用 Python 脚本读取模板 XLSX/DOCX 的所有非空单元格
3. **输出文件**（待验证对象）：使用 Python 脚本读取输出文件的所有非空单元格

读取 XLSX 文件的脚本（模板和输出都用同一脚本）：
```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('FILE_PATH', data_only=True)
ws = wb.active
for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=10, values_only=False):
    cells = []
    for cell in row:
        if cell.value is not None:
            cells.append(f'{cell.coordinate}={cell.value}')
    if cells:
        print(f'Row {row[0].row}: {\"  |  \".join(cells)}')
wb.close()
"
```

> **重要**：模板文件是结构和布局的最终权威。references/ 目录中的模板规则仅作为快速参考，如果规则文档与模板文件内容不一致，**以模板文件为准**。

#### 4.2 三方对比验证

对照**模板文件**和**规则文档**（见 `references/`），逐项检查输出文件：

1. **结构合规性**：输出文件的布局是否与模板文件一致
2. **数据准确性**：每个填充单元格的值是否与 PDF 源文件中的对应数据一致
3. **格式正确性**：日期格式是否为 `YYYY.MM.DD`（或 `YYYY.MM`），数值是否保留 PDF 原文格式
4. **完整性**：PDF 中有的数据是否都已填入输出文件**对应单元格**
5. **无残留默认值**：对比模板文件和输出文件，检查是否有模板示例数据被错误保留

> 详细的模板布局规则见：`references/xlsx-templates.md` 和 `references/docx-templates.md`
> 详细的验证检查清单见：`references/validation-checklist.md`

#### 4.3 修复错误

如果发现任何不一致：

1. **定位根因**：判断是提取阶段、填充阶段、还是模板检测阶段的问题
2. **直接修复代码**：修改对应的 `converter/coa_converter.py`、`converter/xlsx_filler.py` 或 `converter/template_detector.py`
3. **重新运行转换**：
   ```bash
   python3 "$PROJECT_ROOT/converter/coa_converter.py" "$PDF_PATH" "$TEMPLATE_PATH" "$OUTPUT_PATH"
   ```
4. **重新验证**：重复 4.1-4.2（三方对比），确认修复生效
5. **最多迭代 3 次**：如果 3 次迭代仍无法修复，停止并报告需要人工介入

### Step 5. 注册新供应商（AI 验证通过后）

验证通过后，使用以下命令将供应商注册到记录中：
```bash
cd "$PROJECT_ROOT/converter" && python3 -c "
from supplier_checker import register_supplier
register_supplier(
    supplier_id='<供应商ID，小写英文>',
    supplier_name='<供应商名称>',
    signatures=['<签名1: 公司名/地址/网站等PDF中出现的唯一标识>'],
    pdf_format='<lined_table 或 text_based>',
    pdf_name='<PDF文件名>',
    accuracy='<正确率，如100%>',
    notes='<格式特征备注>'
)
"
```

注册后，同供应商的后续 PDF 将自动识别为已知供应商，跳过 AI 验证。

### Step 6. 报告转换结果
- 列出成功填充的字段数量
- 列出自动验证结果（总字段数、通过数、失败数、正确率）
- 列出任何警告或未映射的项目
- 如有 AI 验证修复，列出修复轮次和最终结果
- 提供输出文件路径
- 标注供应商状态（已知/新注册）

---

## 模板规则总览

以下是当前已记录的 10 个模板（5 XLSX + 5 DOCX）：

| # | 模板文件 | 格式 | 类别 | 填充方式 |
|---|---------|------|------|----------|
| 1 | Key In COA - Assay | XLSX | COA 检测报告 | 头部 + 数据行全填充 |
| 2 | Key In COA - Ratio | XLSX | COA 检测报告 | 头部 + 数据行全填充 |
| 3 | Key In COA - Powder | XLSX | COA 检测报告 | 头部 + 数据行全填充 |
| 4 | Allergen - | XLSX | 过敏原声明 | 仅替换产品名 |
| 5 | Flow Chart | XLSX | 生产流程图 | 仅替换标题产品名 |
| 6 | Composition Statement - Powder & Ratio | DOCX | 成分声明 | 替换产品名 + 选择对应比例页 |
| 7 | Composition Statement - Standardized Material | DOCX | 成分声明 | 替换产品名 + 选择含量范围页 |
| 8 | CS - | DOCX | 综合声明 | 替换产品名 + 原产国 |
| 9 | Nutrition info - | DOCX | 营养信息 | 替换产品名（营养数据需手动） |
| 10 | Safety Data Sheet - | DOCX | 安全数据表 | 替换产品名 + 物料名 |

> 详细布局规则见 `references/xlsx-templates.md` 和 `references/docx-templates.md`

### 新增模板的自动发现

当 `templates/` 目录中出现上表未记录的新模板文件时，AI 应：

1. **扫描发现**：在 Step 2 扫描模板时，对比目录中的文件与上表已记录的文件名，识别新增模板
2. **读取结构**：使用 Python 脚本读取新模板的所有非空单元格，分析其布局和字段结构
3. **追加规则**：参照已有模板规则的格式，将新模板的布局规则追加到 references/ 对应文件
4. **更新总览表**：将新模板添加到上方总览表中

---

## 自动检测

- 模板布局自动检测（无需硬编码行号/列号）
- 模板类型自动识别（COA/Allergen/FlowChart/CS/SDS等）
- 公司信息（Key-In Nutrition logo、联系方式）通过模板复制自动保留

## 错误处理

- PDF格式无法识别时自动降级到PyMuPDF
- 未识别的检测项会被记录为警告但不影响其他数据
- 缺失的必填字段会生成警告

## 关于"未映射项"警告的说明

转换过程中出现的"未映射的检测项"警告**不是错误**，而是正常行为。原因如下：

- **模板是数据边界的最终权威**：模板定义了哪些检测项需要填入。如果模板中没有某个检测项的行，说明该项**不需要填入**，不属于标准报告范围。
- **警告仅为信息提示**：警告的目的是让用户知道 PDF 中存在哪些额外数据未被填入模板，以便在需要时人工参考。
- **Assay 缺失（Powder 模板）**：Powder 模板本身没有 Assay/Ratio 行，因此粉末类 COA 报告 "Assay数据缺失" 是预期行为。
- **验证判定标准**：AI 验证时，应将未映射项视为"已知的模板范围外数据"，而非验证失败。
