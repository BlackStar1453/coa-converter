# DOCX 模板详细布局规则

## 6. Composition Statement - Powder & Ratio

**适用场景**：粉末和提取物的成分声明，包含 6 个比例版本的页面

### 文档结构（6 个独立页面）

| 页面 | 标题 | 比例 | 成分表内容 |
|------|------|------|-----------|
| 第 1 页 | Composition Statement 4:1 | 4:1 | 88% {产品}, 12% Maltodextrin |
| 第 2 页 | Composition Statement 10:1 | 10:1 | 90% {产品}, 10% Maltodextrin |
| 第 3 页 | Composition Statement 20:1 | 20:1 | 93% {产品}, 7% Maltodextrin |
| 第 4 页 | Composition Statement 50:1 | 50:1 | 96% {产品}, 4% Maltodextrin |
| 第 5 页 | Composition Statement 100:1 | 100:1 | 98% {产品}, 2% Maltodextrin |
| 第 6 页 | Composition Statement No Maltodextrin Used | 无载体 | 100% {产品} |

**每页结构：**
- 正文：`We, Key-In Nutrition, hereby certify that the composition of our product, {产品名}, is as follows:`
- 表格（3 行 × 2 列）：
  - Row 0: `Product name` | `Content value`
  - Row 1: `{产品名}` | `{百分比} {产品名}\n{百分比} Maltodextrin`
  - Row 2: `Conclusion: Comply with the specifications of` | 同左

**填充规则**：
1. 根据 PDF 中的提取比例（Ratio）选择对应页面
2. 在选定页面中替换所有 `{产品名}` 占位符
3. 其余页面保持不变（模板包含所有比例，使用时选择对应页面）

---

## 7. Composition Statement - Standardized Material

**适用场景**：标准化材料的成分声明，按活性成分含量分 3 个范围

### 文档结构（3 个独立页面）

| 页面 | 标题 | 含量范围 | 成分表内容 |
|------|------|---------|-----------|
| 第 1 页 | Standardized Material < 50% | < 50% | 90% {产品}, 10% Maltodextrin |
| 第 2 页 | Standardized Material 50% - 80% | 50%-80% | 92% {产品}, 8% Maltodextrin |
| 第 3 页 | Standardized Material 80% - 99% | 80%-99% | 100% {产品} |

**每页结构**：与 Powder & Ratio 模板相同（正文 + 3行表格）。

**填充规则**：
1. 根据 PDF 中 Assay 的含量百分比选择对应范围的页面
2. 替换 `{产品名}` 占位符

---

## 8. CS -（Combined Statement）

**适用场景**：综合声明文档，声明产品的多项合规属性

### 文档结构（纯段落，无表格）

```
Combined Statement

To whom it may concern,
We hereby certify that our product {产品名} is:
- Non – GMO
- Non – ETO
- Non – Irradiation
- Country of Origin- {原产国}
- Melamine Free
- Gluten Free
- Cruelty Free
- Yeast Free
- Residual Solvents - Comply
- Pesticides Residual - Comply
- BSE/TSE Free
- GRAS – Yes
- Vegan/Vegetarian – Yes
- Sewer sludge Free
- Prop 65 - Complies
- No artificial flavors/colors
- WADA - Complies
- Kosher/Halal - Complies

Sincerely,
Key-In Nutrition
```

**填充规则**：
1. 替换第 3 段（P3）中的产品名（"our product {产品名} is:"）
2. 替换 "Country of Origin- {原产国}"（P7）中的原产国
3. 其余声明项保持默认值不修改

---

## 9. Nutrition info -

**适用场景**：营养信息声明

### 文档结构

**正文段落**：
- P1: "Nutrition Statement"
- P4: `We, Key-In Nutrition, hereby certify that Nutrition info of our product, {产品名}, is as follow:`

**营养数据表（1 个表格，12 行 × 2 列）：**

| 行号 | Item（Col 0） | Result Per 100g（Col 1） |
|------|--------------|------------------------|
| 0 | Item（表头） | Result Per 100g（表头） |
| 1 | Calories | {值} kcal |
| 2 | Protein | {值} g |
| 3 | Fat | {值} g |
| 4 | Total carbohydrate | {值} g |
| 5 | Potassium | {值} mg |
| 6 | Dietary Fiber | {值} g |
| 7 | Sodium | {值} mg |
| 8 | Calcium | {值} mg |
| 9 | Iron | {值} mg |
| 10 | Vitamin A | {值} mcg |
| 11 | Vitamin C | {值} mcg |

**填充规则**：
1. 替换正文中的 `{产品名}`
2. **营养数据通过网络搜索获取并自动填充**（COA PDF 通常不包含营养数据）：
   - 使用 WebSearch 搜索产品的营养成分数据（per 100g）
   - 优先使用 USDA FoodData Central 等权威数据源
   - 至少交叉验证 2 个数据源确保可靠性
   - 将查询到的数据填充到表格对应行的 Col 1
   - 无法查询到的项目保持模板默认值，并在报告中标注
   - 详细流程见 SKILL.md Step 3.5

---

## 10. Safety Data Sheet -（SDS）

**适用场景**：安全数据表

### 文档结构（17 个表格对应 SDS 各章节）

**Table 0 — 产品标识（需填充）：**

| 行号 | 标签（Col 0） | 值（Col 1） | 填充规则 |
|------|-------------|-----------|----------|
| 0 | Product Name: | {产品名} | **替换为 PDF 产品名** |
| 1 | Material Name: | {物料名/拉丁学名} | **替换为 PDF Botanical Name** |
| 2 | REACH No.: | 保持默认 | 不修改 |

**Table 1 — 企业信息（不可修改）：**
- Business Name: Key-In Nutrition Inc
- Business Address: 2031 S Lynx Pl, Ontario, CA, 91761
- Emergency telephone: 909 – 830 – 9376

**Table 2-16 — SDS 各章节（不可修改）：**
- Table 2: Composition / Nanoforms
- Table 3: Hazard Identification (GHS Classification)
- Table 4: First Aid Measures
- Table 5: Fire Fighting Measures
- Table 6: Accidental Release Measures
- Table 7: Handling and Storage
- Table 8: Exposure Controls / Personal Protection
- Table 9: Physical and Chemical Properties (18 rows)
- Table 10: Stability and Reactivity
- Table 11: Stability Conditions
- Table 12: Toxicological Information
- Table 13: Ecological Information
- Table 14: Disposal Considerations
- Table 15: Transport Information
- Table 16: Regulatory Information

**填充规则**：
1. 仅替换 Table 0 中的 Product Name（Row 0, Col 1）和 Material Name（Row 1, Col 1）
2. 所有其余表格和内容保持模板默认值不修改
