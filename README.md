# 查房表更新工具

用于将病历系统导出的 Excel 列表合并到昨日查房表模板中，并按维护的病区、医疗组、床位范围生成更新后的查房表。

## v1.0

- 工作台式界面。
- 支持拖拽或选择两份 Excel。
- 支持维护病区、医疗组和床位。
- 支持设置最小行高，导出时按内容自动加高到刚好够显示。
- 床号包含 `zr` 的行会排在普通床号后面。

## 隐私说明

仓库不包含实际查房表、患者数据或本地配置文件。`*.xls`、`*.xlsx`、`config.ini` 已被 `.gitignore` 排除。

## 打包

Win7 兼容版使用 Python 3.8 打包：

```powershell
py -3.8 -m PyInstaller --noconfirm --clean --onefile --windowed --collect-all tkinterdnd2 "查房表更新工具.py"
```
