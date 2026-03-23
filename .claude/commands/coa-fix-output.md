---
name: coa-fix-output
description: >
  Fix COA conversion output errors. Use when:
  (1) User reports errors in a converted XLSX/DOCX output file (wrong values, missing data, misplaced fields, formatting issues).
  (2) User says "输出有误", "结果不对", "转换错误", "fix output", "修复输出", "数据不对".
  (3) User pastes error descriptions about COA conversion results.
  This skill reads the source PDF, template, output, and converter code to diagnose and fix both the output file AND the underlying conversion rules.
  NEVER use for: initial conversion (use /coa-to-template instead), generic file editing, or non-COA tasks.
---

# COA 输出修复工具

用户在转换后的输出文件中发现错误时，使用此 command 进行诊断和修复。

**核心原则**：先修复输出文件满足用户需求，再修复转换代码防止问题复发。

## 项目根目录检测（必须首先执行）

```bash
PROJECT_ROOT="$(pwd)"
while [ "$PROJECT_ROOT" != "/" ]; do
  if [ -d "$PROJECT_ROOT/converter" ] && [ -d "$PROJECT_ROOT/templates" ] && [ -f "$PROJECT_ROOT/app.py" ]; then
    break
  fi
  PROJECT_ROOT="$(dirname "$PROJECT_ROOT")"
done

if [ "$PROJECT_ROOT" = "/" ]; then
  echo "错误：无法找到 coa-converter-web 项目根目录。"
  exit 1
fi
```

## 使用方式

```
/coa-fix-output <用户描述的错误>
```

示例：
```
/coa-fix-output Lot No. 填到了 Mfg. Date 的位置，Expiry Date 是空的
/coa-fix-output Assay 的 Result 列显示的是 Specification 的值
/coa-fix-output 微生物检测数据全部缺失
```

## 执行步骤

### Step 1. 定位相关文件

找到最近一次转换涉及的文件。按以下顺序查找：

```bash
# 1a. 查找最近的输出文件
ls -lt "$PROJECT_ROOT/output/"*.{xlsx,docx} 2>/dev/null | head -5

# 1b. 查找最近的输入 PDF（可能已被清理，检查 input/ 和 job 记录）
ls -lt "$PROJECT_ROOT/input/"*.pdf 2>/dev/null | head -5
```

**如果文件不明确**，向用户确认：
- 哪个 PDF 文件是源文件？
- 哪个输出文件有问题？
- 使用的是哪个模板？

### Step 2. 三方读取（PDF + 模板 + 输出）

必须同时读取三份文件进行对比：

#### 2.1 读取 PDF 源文件（ground truth）

使用 Read 工具直接读取 PDF 获取文本数据。这是数据的最终权威来源。

#### 2.2 读取模板文件（结构权威）

```bash
# XLSX 模板
python3 -c "
import sys; sys.path.insert(0, '$PROJECT_ROOT/converter')
import openpyxl
wb = openpyxl.load_workbook('TEMPLATE_PATH', data_only=True)
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

#### 2.3 读取输出文件（待修复对象）

使用与 2.2 相同的脚本读取输出文件，替换路径即可。

### Step 3. 诊断错误

对照用户描述的错误，进行以下分析：

1. **定位错误单元格**：找到输出文件中用户指出的错误位置
2. **对比 PDF 原始数据**：确认 PDF 中该字段的正确值是什么
3. **对比模板结构**：确认该单元格在模板中的预期用途
4. **分类错误类型**：
   - **数据错位**：正确的值填到了错误的位置（如 Lot No. 填到 Mfg. Date）
   - **数据缺失**：PDF 中有数据但输出文件中为空
   - **数据错误**：填入的值与 PDF 原始值不一致
   - **格式错误**：日期格式、数值格式不正确
   - **残留数据**：模板中的示例/默认值未被替换

### Step 4. 修复输出文件（立即满足用户需求）

直接修改输出文件，使其符合 PDF 源文件数据和模板结构要求：

```bash
# XLSX 修复示例
python3 -c "
import sys; sys.path.insert(0, '$PROJECT_ROOT/converter')
import openpyxl
wb = openpyxl.load_workbook('OUTPUT_PATH')
ws = wb.active

# 修复具体单元格（根据诊断结果）
ws['C4'] = '正确的值'  # 示例
ws['E15'] = '正确的值'  # 示例

wb.save('OUTPUT_PATH')
print('输出文件已修复')
"
```

修复后重新读取输出文件，向用户展示修复结果，确认所有报告的错误已解决。

### Step 5. 溯源根因（分析转换代码）

定位导致错误的转换代码。按错误类型检查对应模块：

| 错误类型 | 优先检查的文件 | 关注点 |
|----------|---------------|--------|
| 数据错位 | `converter/xlsx_filler.py` 或 `converter/docx_filler.py` | 单元格坐标映射 |
| 数据缺失 | `converter/coa_converter.py` | PDF 解析正则、字段提取逻辑 |
| 数据错误 | `converter/coa_converter.py` | 数据清洗、格式转换 |
| 格式错误 | `converter/xlsx_filler.py` | 日期/数值格式化 |
| 模板检测错误 | `converter/template_detector.py` | 行号/列号检测、模板类型判断 |

**读取转换代码**：

```bash
# 根据错误类型读取相关文件
cat "$PROJECT_ROOT/converter/coa_converter.py"      # PDF 解析
cat "$PROJECT_ROOT/converter/xlsx_filler.py"         # XLSX 填充
cat "$PROJECT_ROOT/converter/docx_filler.py"         # DOCX 填充
cat "$PROJECT_ROOT/converter/template_detector.py"   # 模板检测
```

### Step 6. 修复转换规则（防止复发）

修改对应的转换代码，确保同类 PDF 后续转换不再出现相同错误。

**修复原则**：
- **最小修改**：只改必要的代码，避免引入新问题
- **向后兼容**：修复不能破坏其他模板类型的正常转换
- **添加日志**：在修复位置添加 `logger.info()` 方便后续追踪

**修复后验证**：

```bash
# 使用修复后的代码重新转换同一个 PDF
source "$PROJECT_ROOT/.venv/bin/activate"
python3 "$PROJECT_ROOT/converter/coa_converter.py" "PDF_PATH" "TEMPLATE_PATH" "OUTPUT_PATH_NEW"
```

重新读取新输出文件，确认：
1. 用户报告的错误已消失
2. 其他字段仍然正确
3. 无新的回归错误

### Step 7. 报告修复结果

向用户汇报：

1. **错误诊断**：简要说明错误原因
2. **输出修复**：列出修改的单元格和值
3. **代码修复**：说明修改了哪个文件的哪个逻辑
4. **验证结果**：确认重新转换后输出正确
5. **影响范围**：说明此修复是否影响其他模板类型

格式示例：
```
## 修复报告

### 错误原因
Lot No. 和 Mfg. Date 的正则匹配顺序错误，导致 Lot No. 的值被赋给了 Mfg. Date。

### 输出文件修复
- C4: "" → "LOT2024001"（Lot No.）
- C5: "LOT2024001" → "2024.03.15"（Mfg. Date）

### 代码修复
- `converter/coa_converter.py:156` — 调整 `_extract_header_fields()` 中 Lot No. 和 Mfg. Date 的匹配优先级

### 验证
重新转换验证通过，所有字段正确。此修复仅影响头部字段提取逻辑，不影响数据行填充。
```

## 注意事项

- **先修输出，后修代码**：用户的即时需求是获得正确的输出文件，代码修复是防止复发
- **不要猜测**：必须读取 PDF 原始数据确认正确值，不要根据上下文推测
- **保留模板格式**：修复输出文件时，保留模板原有的字体、颜色、边框等格式
- **多错误情况**：如果用户报告多个错误，逐一处理，每个错误都要完成 Step 3-6 的完整流程
