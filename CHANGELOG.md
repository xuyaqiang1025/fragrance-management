# 变更日志 / CHANGELOG

本文件记录 `香料管理系统` 各版本的提交与主要改动。

---

## v3.3 — 多格式导入、删除持久化加固、GC-MS/智能分析/统计模块优化

**提交：** `136e42f`
**Tag：** `v3.3`
**日期：** 2026-08-31
**推送：** `https://github.com/xuyaqiang1025/fragrance-management.git`

### 问题1 — 多格式导入
- `import_gcms_compounds`（及原料、配方导入）改用 `read_table_any` + `SUPPORTED_FILE_FILTER`，
  支持 xlsx / xlsm / xls / csv / tsv / txt / HTML / XML 等格式的自动识别与读取。
- 文件选择框只列出 xlsx/xlsm/xls/csv/tsv/txt；XML 与 HTML 需经「所有文件(*)」选中，
  但 `read_table_any` 均可正确解析。此处 XML 指 **Excel 2003 SpreadsheetML**。
- 该 XML 路径曾有一个真实 bug，已在 `v3.3.1`（提交 `4f912f8`）中修复，详见下文。

### 问题2 — 删除配方后重新载入又出现
- `delete_formula` 开头重新 `session.query(Formula).filter_by(id=...).first()`，
  防御分页加载器关闭 session 导致的 detached instance 问题，确保真正删除并落库。
- `_init_sample_data` 强化「一次性播种」：
  - 无论新旧库都写入 `sample_formulas_seeded` 标记；
  - 仅在库不存在且 `Formula` 为空时播种；
  - 按已有 `number` 去重，避免重复生成示例配方。
- 注：该 bug 此前表现为「旧版 exe（2026-08-29 打包）重新载入后又生成已删配方」，根因为
  用户运行的是早于 v3.1/v3.2 修复的打包产物；本次随 `v3.3` 重新打包后彻底解决。

### 问题3 — GC-MS / 智能分析 / 数据统计优化
- **GC-MS 分析对话框（重要更正）**：
  经核查，真正的 `GCMSAnalysisFuncDialog` 定义在 `src/main.py`（内联类，约 4561–5215 行），
  由 `show_gcms_analysis_func` 直接实例化，**功能完整且正确**：
    - `load_analyses` 以 `analysis.id` 存储（无 number/analysis_id 混淆）；
    - matplotlib 后端使用 `backend_qtagg`（PyQt6 兼容）；
    - 已实现 特征物质筛查 / 差异物质分析（含差异倍数与 log2FC）/ 热图 / PCA 降维 / 图片与数据导出。
  原 `ui/gcms_analysis_func_dialog.py` 是**未被任何代码引用的孤儿模块**（全项目 grep 无引用、打包 spec 未包含），
  其上的改动不影响运行产物。该孤儿模块已在后续提交中移除，避免误导。
  结论：GC-MS 模块在运行产物中已是优化后的完整实现，本次无需额外修复。
- 数据统计新增第 4 类 **「GC-MS化合物频次统计」**（跨样品按化合物汇总 Top15 + 占比）。
- `ai_analysis_module.py` 新增 `FormulaAnalyzer._ing_ids(f)`：
  优先从 `f.ingredients` 取原料 id；缺失时回退解析 `f.content`（名称:百分比:用量）按名查找，
  使相似度 / 重合度分析在无 `ingredients` 关联时仍可用。

### 其他
- `.gitignore` 增加 `.workbuddy/` 排除助手工作区。
- 移除孤儿模块 `ui/gcms_analysis_func_dialog.py`（未被引用的死代码）。

---

## v3.3.1 — Excel 2003 XML 导入修复

**提交：** `4f912f8`
**日期：** 2026-09-01

### 问题
`read_table_any` 会先把文件头转小写，Excel 2003 XML（SpreadsheetML）中的
`<Table>` 因此变成 `<table>`，命中了 HTML 分支的 `'<table' in head` 判断，
导致 XML 被误送进 `pd.read_html`，报「**HTML文件中未找到数据表格**」；
而排在它后面的 `<?xml` 分支永远执行不到 —— XML 导入实际是失效的。

### 修复
- 将 `<?xml` 判定**提到 HTML 判定之前**：文件开头的 `<?xml` 声明是无歧义信号。

### 验证
- 实测 `read_table_any`：**CSV / TSV / TXT / XLSX / HTML / XML(Excel 2003) 6 种格式全部通过**。
- 说明：此处 XML 特指 **Excel 2003 SpreadsheetML**
  （命名空间 `urn:schemas-microsoft-com:office:spreadsheet`）。

### 打包
- 该修复已随重新打包的 `release/香料管理系统.exe`（409,831,160 字节）发布。
- 旧 exe 备份为 `release/香料管理系统_v3.2_20260829_旧版.exe`。

---

## v3.2 — 全量导入格式支持、数据持久化加固、GC-MS/智能分析/统计模块优化
**提交：** `3b39e49` | **Tag：** `v3.2-baseline`
基线提交，含多格式导入框架与示例数据重复播种修复。

## v3.1 — 多格式导入支持与示例数据重复播种修复
**提交：** `1648330`
引入 `table_io.read_table_any` 多格式读取与一次性播种标记。

## v3.0 — 香料管理系统（含AI调香专家模块）
**提交：** `ed078d2`
初始完整版本，含 AI 调香专家、数据模型与基础界面。
