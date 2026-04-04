# COA 模板验证检查清单

AI 验证 COA 模板（Assay/Ratio/Powder）时必须逐项检查：

## 1. 头部字段（Row 4-7）
- [ ] C4 产品名 = PDF 中的 Product Name
- [ ] C5 植物来源 = PDF 中的 Botanical Source/Latin Name
- [ ] F5 使用部位 = PDF 中的 Part Used/Plant Part
- [ ] C6 批次号 = PDF 中的 Batch Number/Lot No.
- [ ] F6 原产国 = PDF 中的 Country of Origin
- [ ] C7 生产日期 = PDF 日期转为 YYYY.MM.DD（或 YYYY.MM）
- [ ] F7 过期日期 = PDF 日期转为 YYYY.MM.DD（或 YYYY.MM）
- [ ] 以上字段均不是模板示例值（不应出现 "Ashwagandha"、"Camellia sinensis" 等）

## 2. Assay/Ratio 行（仅 Assay 和 Ratio 模板）
- [ ] Row 10 的 Specification、Result、Method 与 PDF 一致

## 3. 检测数据行
- [ ] 每个检测项的 Specification（C列）= PDF 中的规格值
- [ ] 每个检测项的 Result（E列）= PDF 中的检测结果
- [ ] 每个检测项的 Method（F列）= PDF 中的检测方法
- [ ] PDF 中没有的检测项，对应单元格应为空（不应保留模板默认值）

## 4. Packing & Storage
- [ ] Packing 单元格 = PDF 中的包装信息（如有）
- [ ] Storage 单元格 = PDF 中的储存条件（如有）
- [ ] 如果 PDF 中没有包装/储存信息，这些单元格应为空

## 5. 不可修改区域
- [ ] C1 公司信息保持不变
- [ ] A3 "Certificate of Analysis" 保持不变
- [ ] Row 9 数据表头保持不变
- [ ] Section header 行（Analytical Data、Microbiological Test、Additional information）保持不变

## 6. 未映射项处理（信息项，非验证失败）
- [ ] 确认所有"未映射的检测项"警告均为**模板中无对应行**的项目（非填充遗漏）
- [ ] 这些项属于模板范围外数据，不计入验证失败
- [ ] 在最终报告中列出未映射项供用户参考，但明确标注为"模板范围外"

## 7. DOCX 模板验证（Composition/CS/Nutrition/SDS）
- [ ] 产品名已替换到所有出现位置（不应残留模板示例名称如 "Ashwagandha"、"Pumpkin Seed"）
- [ ] **CS 模板**：原产国已替换到正确位置
- [ ] **SDS 模板**：物料名已替换到 Table 0 的对应单元格
- [ ] **Composition 模板**：已根据 PDF 中的比例或含量范围选择正确的页面
- [ ] **Nutrition 模板**：所有营养数据已通过 Step 3.5 网络搜索填充，不应保留模板占位符值
- [ ] 文档格式（字体、字号、对齐）与模板一致，未被修改
