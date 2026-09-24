---
version: "1.0.0"
name: excel-sheet-inspector
description: "Excel 只读勘查工具。列出工作表、行列规模、原始网格与逐列统计；并可提取 BOM 报表的物料行、状态分布、零价行与两版差异。"
author: ggq-eng
---

# excel-sheet-inspector

Excel 只读勘查工具。列出工作表、行列规模、原始网格与逐列统计；并可提取 BOM 报表的物料行、状态分布、零价行与两版差异。

## 能力

- **结构勘查**：工作表清单、行列规模、逐列非空统计与类型分布
- **BOM 提取**：物料行、状态分布、零价行、同名相邻主/替对、两版差异
- **只读保证**：不写入、不改名、不改格式

## 使用方式

```bash
python inspect_sheets.py <文件.xlsx>
python extract_bom.py <版本A.xlsx> <版本B.xlsx>
```

## 关键规则

- **只读**，绝不修改原文件
- 遇到合并单元格、多级表头时先打印原始网格再解析，不臆测表头
- 差异输出按「新增 / 删除 / 数量变化 / 价格变化」分类

---
*excel-sheet-inspector by ggq-eng*
