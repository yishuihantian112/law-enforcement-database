# 行政执法事项数据库管理工具

一个用于管理行政执法事项清单的Python工具，支持从Excel、PDF、Word文档中批量提取执法事项数据，自动识别执法机关和行政区划，支持增量添加和数据导出。

## 功能特性

- **多格式支持**：支持从Excel (.xlsx)、PDF (.pdf)、Word (.docx) 文档中提取执法事项
- **自动识别**：自动从文件内容或文件路径中识别执法机关和行政区划
- **增量添加**：基于文件哈希值避免重复导入
- **数据库存储**：使用SQLite数据库存储，支持快速查询和统计
- **数据导出**：支持导出为Excel或CSV格式
- **网站抓取**：支持从政府网站批量下载执法清单文件（框架功能）
- **批量处理**：支持单文件或整个文件夹批量导入

## 安装

### 环境要求
- Python 3.7+
- pip

### 安装依赖

```bash
pip install pandas pdfplumber python-docx beautifulsoup4 openpyxl
```

或使用以下命令（如果使用py启动器）：

```bash
py -m pip install pandas pdfplumber python-docx beautifulsoup4 openpyxl
```

## 使用方法

### 基本命令

```bash
# 查看帮助
python law_enforcement_db.py --help

# 添加单个文件
python law_enforcement_db.py --add 文件.xlsx

# 添加多个文件
python law_enforcement_db.py --add 文件1.xlsx 文件2.pdf 文件3.docx

# 添加整个文件夹
python law_enforcement_db.py --add 文件夹路径

# 指定行政区划和执法机关
python law_enforcement_db.py --add 文件.xlsx --region 浙江省 --authority 市场监管部门

# 导出数据到Excel
python law_enforcement_db.py --export output.xlsx

# 导出数据到CSV
python law_enforcement_db.py --export output.csv

# 查看最近30条记录
python law_enforcement_db.py --list

# 查看统计信息
python law_enforcement_db.py --stats

# 从网站抓取文件
python law_enforcement_db.py --fetch https://example.com --fetch-dir downloaded
```

### 命令行参数说明

| 参数 | 简写 | 说明 |
|------|------|------|
| `--add` | `-a` | 添加文件或文件夹，支持.xlsx/.pdf/.docx格式 |
| `--export` | `-e` | 导出到Excel或CSV文件 |
| `--list` | `-l` | 列出最近30条记录 |
| `--stats` | `-s` | 显示数据库统计信息 |
| `--db` | - | 指定数据库文件路径（默认：law_enforcement.db） |
| `--region` | - | 指定行政区划（如"浙江省"） |
| `--authority` | - | 指定执法机关（如"应急管理部门"） |
| `--fetch` | - | 从政府网站URL抓取清单文件 |
| `--fetch-dir` | - | 抓取下载的目录（默认：downloaded） |

## 支持的文件格式

### Excel格式 (.xlsx)
- 自动检测表头行位置
- 支持多种列名变体
- 支持HTML格式化的列名

### PDF格式 (.pdf)
- 自动提取表格数据
- 支持表格文本提取
- 无表格时支持文本规则提取

### Word格式 (.docx)
- 支持表格数据提取
- 无表格时支持文本规则提取

## 数据结构

### 数据库表结构

```sql
CREATE TABLE enforcement_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    serial_number TEXT,              -- 序号
    item_name TEXT NOT NULL,         -- 事项名称
    power_type TEXT,                 -- 职权类型
    legal_basis TEXT,                -- 实施依据
    legal_subject TEXT,              -- 法定实施主体
    first_responsibility_level TEXT, -- 第一责任层级建议
    enforcement_authority TEXT,      -- 执法机关（自动识别）
    administrative_region TEXT,      -- 行政区划（自动识别）
    source_file TEXT,                -- 来源文件
    file_hash TEXT,                  -- 文件哈希（防重复）
    import_time TIMESTAMP            -- 导入时间
);
```

### 预期列名

Excel文件应包含以下列（支持变体）：

| 标准列名 | 支持变体 |
|----------|----------|
| 序号 | 序号、serial_number、sn |
| 事项名称 | 事项名称、item_name |
| 职权类型 | 职权类型、power_type |
| 实施依据 | 实施依据、legal_basis |
| 法定实施主体 | 法定实施主体、法  定<br>实施主体、legal_subject |
| 第一责任层级建议 | 第一责任层级建议、责任层级、first_responsibility_level |

## 使用示例

### 示例1：导入单个Excel文件

```bash
python law_enforcement_db.py --add /path/to/enforcement_list.xlsx
```

输出：
```
2024-01-01 10:00:00 - INFO - 从Excel解析到 150 条事项
2024-01-01 10:00:01 - INFO - 成功添加 150 条记录，来源: /path/to/enforcement_list.xlsx
2024-01-01 10:00:01 - INFO - 共添加 150 条记录
```

### 示例2：批量导入文件夹

```bash
python law_enforcement_db.py --add /path/to/documents --region 浙江省
```

### 示例3：导出合并数据

```bash
python law_enforcement_db.py --export /path/to/merged_list.xlsx
```

### 示例4：查看数据统计

```bash
python law_enforcement_db.py --stats
```

输出：
```
数据库统计:
  总事项数: 3429
  执法机关数: 5
  涉及行政区划数: 8
```

## 行政区划识别

工具内置了中国主要省市级别的行政区划数据库，会自动从以下位置识别：

1. 文件路径中的行政区划名称
2. 文档内容中的行政区划名称

支持的部分行政区划：
- **省级**：北京市、天津市、上海市、重庆市、河北省、山西省、辽宁省、吉林省、黑龙江省、江苏省、浙江省、安徽省、福建省、江西省、山东省、河南省、湖北省、湖南省、广东省、海南省、四川省、贵州省、云南省、陕西省、甘肃省、青海省等
- **市级**：石家庄市、唐山市、太原市、沈阳市、大连市、长春市、哈尔滨市、南京市、苏州市、杭州市、宁波市、合肥市、福州市、厦门市、南昌市、济南市、青岛市、郑州市、武汉市、长沙市、广州市、深圳市、南宁市、海口市、成都市、贵阳市、昆明市、西安市、兰州市、西宁市、银川市、乌鲁木齐市等

## 常见问题

### Q: 如何处理中文乱码？

A: CSV导出使用`utf-8-sig`编码，Excel可直接打开无乱码。如仍有问题，建议使用.xlsx格式导出。

### Q: 文件重复导入怎么办？

A: 工具使用SHA256哈希值检测重复文件，已导入的文件会自动跳过。

### Q: 支持哪些Excel版本？

A: 支持.xlsx格式（Excel 2007+），不支持旧版.xls格式。如有.xls文件，请先用Excel另存为.xlsx格式。

### Q: PDF提取不完整怎么办？

A: 扫描PDF需要OCR支持。建议使用PDF中的表格格式，或转换为Excel后再导入。

### Q: 如何自定义行政区划列表？

A: 编辑脚本中的`CHINA_REGIONS`字典，添加您需要的省市区名称。

## 项目结构

```
.
├── law_enforcement_db.py    # 主程序文件
├── law_enforcement.db       # SQLite数据库文件（运行后生成）
├── downloaded/              # 网站抓取下载目录（自动创建）
└── README.md                # 本说明文件
```

## 技术栈

- **Python** 3.7+
- **pandas** - 数据处理
- **pdfplumber** - PDF解析
- **python-docx** - Word文档解析
- **beautifulsoup4** - HTML解析
- **openpyxl** - Excel读写
- **SQLite** - 数据存储

## 许可证

本项目仅供学习和研究使用。

## 贡献

欢迎提交问题和改进建议。

## 联系方式

如有问题或建议，请通过GitHub Issues联系。

---

**最后更新**：2026年6月