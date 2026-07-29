# v1.1.0

- 新增版本化 channel x variant profile 契约。`legacy-v1/main-800` 保持 v1.0 横向表兼容；输出和报告会记录 profile 与版本。
- 已冻结天猫型竖表字段：`SKU` 到 `文件名称`，`变量01` 至 `变量07` 分别到利益点1、利益点2、预估到手价、价格1、价格2、规格、卖点，`图片目录路径` 到产品；`变量08/09` 不读取也不校验。
- 未知 profile、未声明列、字段 schema 不符会以 `E_PROFILE_UNSUPPORTED` 或 `E_PROFILE_SCHEMA_MISMATCH` 拒绝，不再按全局变量名猜测。
- 批处理在导出前检查 profile 目标尺寸；v1.1 Variables profile 还会检查 PSD Variables 名称、类型和应用绑定，并为每一行创建临时 Data Set 后导出。`E_*` 行不会输出 JPG 或 PSD，`W_*` 保留输出并写入报告。`@` 文本/`!` 智能对象仅保留给 `legacy-v1`。
- 结果报告增加 `profile_id`、`profile_version`、`severity`、`variable_binding_status` 和 `未导出`。
- 京东自营第一版维持 `legacy-v1/main-800`；天猫型 profile 已按 750/800 两个独立规格登记。仓库没有对应实机 PSD，待审批 profile 不能用于生产导出，也不声称已验证。
- 天猫官旗改为按运营 Sheet 选择规格：`现货-800` 对应 `main-800`，`现货-750` 对应 `main-750`。每次只生成当前 Sheet，不再自动派发两套任务；设计师可在当前 Sheet 内选择全部商品或部分商品。未知 Sheet 记录为 `E_PROFILE_SHEET_MISMATCH` 并阻断本次任务。

限制：本版本未在 Windows Photoshop 上实机验证 Photoshop Variables；Mac 环境仅运行 Python 契约测试。
