# SeaTable Excel 生成器

这是一个用于从 SeaTable 生成 Excel 文件的工具，支持菜单选择不同的配置文件。

## 功能特性

- 从 SeaTable 获取数据并生成 Excel 文件
- 支持多个配置文件，每个配置文件可以有不同的 SeaTable API 配置
- 自动格式化日期和数字
- 智能处理百分比数值（0-1小数自动转换为0-100百分比）
- 使用 SUBTOTAL(109,...) 函数计算合计（只统计可见行）
- 支持文件合并功能
- 自动设置 Excel 样式和格式

## 安装依赖

```bash
pip install openpyxl python-dotenv seatable-api
```

## 使用方法

1. 配置 `.env` 文件，设置 SeaTable 服务器地址和 API Token
2. 运行程序: `python main-pro.py`
3. 选择配置文件
4. 选择要生成的文件或操作

## 构建独立可执行文件

### 本地构建

```bash
# 构建适用于当前平台的可执行文件
python build_standalone.py
```

### GitHub Actions自动构建

该项目配置了 GitHub Actions 工作流，会在推送到 `main` 或 `master` 分支时自动构建适用于 Linux、Windows 和 macOS 的可执行文件。

构建的可执行文件可以通过 GitHub Actions 的 Artifacts 下载，或在发布版本时自动上传到 Release 页面。