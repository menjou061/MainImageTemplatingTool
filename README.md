# 电商主图套版工具

面向电商设计师的 Windows + Adobe Photoshop 批量套版工具。选择商品 Excel 和 PSD 模板后，工具会自动完成数据检查、素材替换，并同时导出 JPG 成品和可编辑 PSD。

> 当前版本：**v1.0.0**。新手安装与使用教程见[飞书文档](https://pcnfsebpg8ae.feishu.cn/wiki/HE7JwtImKiaRhLkVNFwcakrmnxg)。

## 下载与安装

1. 打开 [Releases](https://github.com/menjou061/MainImageTemplatingTool/releases/latest)。
2. 下载 `MainImageTemplatingTool_Windows_v1.0.0.zip`。
3. 把 ZIP 完整解压到本地文件夹，不要直接在压缩包里运行。
4. 确认 Windows 已安装并登录 Adobe Photoshop。
5. 双击 `开始套版.cmd`。

工具已内置 Python 运行环境，正常使用不需要另外安装 Python。

## 使用流程

1. 选择商品信息表格（`.xlsx`）。
2. 选择 PSD 模板。
3. 确认数据工作表和要生成的商品。
4. 点击“开始生成”。

首次使用时，成品默认保存到桌面的 `电商主图套版成品`。每次任务会自动创建独立文件夹，不会覆盖以前的结果：

```text
套版成品_日期_时间_工作表名/
├─ JPG成品/      可直接交付的图片
├─ PSD源文件/    可继续编辑的 PSD
└─ 任务记录/     异常清单、生成结果和任务日志
```

异常商品会自动跳过，原因记录在 `任务记录/异常记录.csv` 中，不会反复弹出长错误信息。

## 使用要求

- Windows 10 x64 或更高版本
- Windows 版 Adobe Photoshop，已完成安装和授权
- 商品素材路径能在当前 Windows 主机上直接访问
- 安装包必须完整解压，不能只复制启动文件

## 项目结构

- `L0_Windows命令行版/`：Windows 交付程序及内置运行环境
- `tests/`：Windows 界面回归与截图脚本
- `M1_Windows验证包/`：早期 Windows/Photoshop 最小验证工具
- `电商主图批量套版工具_PRD_v*.docx`：产品需求文档历史版本

## 版本记录

| 版本 | 说明 |
|---|---|
| **v1.0.0** | 首个设计师团队交付版：四步套版流程、商品勾选、桌面默认保存、JPG/PSD/任务记录分类、异常自动跳过与任务日志、Windows 界面回归验证 |
