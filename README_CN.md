# Scholar-City：学术期刊数据流水线

一套用于从 Web of Science 构建大规模文献计量数据集、追踪期刊主编信息、并建立作者-主编名称对照表的工具集，服务于科学经济学领域的实证研究。

## 概述

本项目提供三个集成模块，构成从原始期刊元数据到结构化分析数据的完整流水线：

| 模块 | 功能 |
|------|------|
| **journal-eic-search** | 解析期刊分级目录 PDF，搜索各期刊历任主编，记录其机构、城市和任期信息 |
| **wos-paper-crawler** | 爬取 Web of Science 中目标期刊的全部论文（1995 年至今），提取元数据，清洗城市/国家/机构字段 |
| **editor-name-crosswalk** | 构建主编规范名到论文数据库中所有名称变体的映射表，用于识别主编参与署名的论文 |

整套流水线设计用于主编更替效应研究、同城发表溢价分析和学术网络分析。所有输出均为 Stata 友好的长格式 Excel。

## 环境要求

- **Python 3.9+**，需安装：`openpyxl`、`selenium`、`pandas`
- **Microsoft Edge** 浏览器 + [Edge WebDriver](https://developer.microsoft.com/zh-cn/microsoft-edge/tools/webdriver/)（WoS 爬虫需要）
- 可访问 Web of Science 的校园网络（爬虫需要）
- **Exa** API（EIC 模块中用于网络搜索，通过 opencode agent 调用）

```bash
pip install openpyxl selenium pandas
```

## 安装

```bash
git clone https://github.com/YOUR_USERNAME/scholar-city.git
cd scholar-city
```

将技能目录复制到 opencode skills 文件夹，或作为独立 Python 脚本使用：

```bash
cp -r journal-eic-search wos-paper-crawler editor-name-crosswalk ~/.config/opencode/skills/
```

也可直接运行脚本：

```bash
python journal-eic-search/scripts/parse_pdf.py input.pdf --output journals.csv
python wos-paper-crawler/scripts/crawl_wos.py
python editor-name-crosswalk/scripts/build_crosswalk.py --eic-xlsx 期刊目录.xlsx --paper-xlsx WoS_Papers_All.xlsx --output 主编名称对照.xlsx
```

## 工作流

### 第一步：构建期刊目录

解析期刊分级 PDF，创建期刊元数据表。

```bash
python journal-eic-search/scripts/parse_pdf.py "期刊目录.pdf" --output "journals_temp.csv"
python journal-eic-search/scripts/write_excel.py "期刊目录.xlsx" --add-journals "journals_temp.csv"
```

### 第二步：查询主编信息

使用 `journal-eic-search` 技能，查找每本期刊的历任主编、所属机构、所在城市和任期起止年份。结果写入 `期刊目录.xlsx` 的 Sheet2（EIC_Long 长表格式）。

### 第三步：爬取 Web of Science

编辑 `wos-paper-crawler/scripts/crawl_wos.py`，设置目标期刊、输出路径和停止时间：

```python
OUT = r"D:\Scholar-city\WoS_JournalName.csv"
STOP_TIME = datetime(2026, 6, 1, 22, 0, 0)
JOURNAL = "American Economic Review"
START_YEAR = 1995
```

然后运行：

```bash
python wos-paper-crawler/scripts/crawl_wos.py
```

爬虫支持断点续传、空页自动重启浏览器、按时停止等功能。

### 第四步：后处理爬虫输出

```bash
python wos-paper-crawler/scripts/postprocess.py --input WoS_JournalName.csv --output WoS_JournalName_long.xlsx
```

自动完成以下清洗：
- 国家名标准化（WoS 名称 → ISO 3166-1 alpha-3）
- 地址中提取城市（处理邮政编码、缩写、多校区）
- 机构名清洗（去除 WoS 编号前缀）
- 机构缺城市时推断城市
- 转换成长格式 Excel（Papers + Authors 两表）

### 第五步：构建主编名称对照表

```bash
python editor-name-crosswalk/scripts/build_crosswalk.py \
    --eic-xlsx 期刊目录.xlsx \
    --paper-xlsx WoS_Papers_All.xlsx \
    --output 主编名称对照.xlsx
```

将每位主编的规范名映射到论文数据库中出现的所有名称变体（如 "Acemoglu, Daron" ↔ "Acemoglu, D"），用于识别主编参与署名的论文。

## 输出格式

### WoS_Papers_All.xlsx

**Papers 表：**
| 列名 | 内容 |
|------|------|
| paper_id | WoS 唯一标识符 |
| journal_name | 期刊名 |
| article_title | 论文标题 |
| published_date | 发表年月（YYYY/MM） |
| citations_wos_core | WoS Core 被引次数 |
| citations_all_db | 全库被引次数 |
| cited_references | 参考文献数量 |
| document_type | 文献类型（Article、Review 等） |

**Authors 表：**
| 列名 | 内容 |
|------|------|
| paper_id | 关联 Papers 表 |
| author_seq | 作者顺序（1, 2, ...） |
| author_name | "姓氏, 名字" 格式 |
| institution | 主要隶属机构 |
| city | 提取的城市 |
| country | ISO 3166-1 alpha-3 国家代码 |
| inst_canonical | 规范化机构名 |

### 期刊目录.xlsx

**Sheet1（期刊元数据）：** 每刊一行，含 ISSN、分级、信息来源等。

**Sheet2（EIC_Long）：** 每位主编任期一行：

| 列 | 内容 |
|----|------|
| 期刊名 | 期刊名称 |
| 姓名 | 主编姓名 |
| 单位 | 主要机构 |
| 单位所在城市 | 机构城市 |
| 上任时间 | 开始年份 |
| 卸任时间 | 结束年份 |

### 主编名称对照.xlsx

| 列 | 内容 |
|----|------|
| 主编规范名 | EIC 标准名 |
| 数据库中匹配到的名称变体 | 管道分隔的所有匹配变体 |
| 匹配数量 | 匹配到的变体数 |
| 匹配状态 | 已匹配 / 未找到匹配 |

## 文件结构

```
scholar-city/
├── README.md
├── README_CN.md
├── journal-eic-search/
│   ├── SKILL.md
│   └── scripts/
│       ├── parse_pdf.py          # 解析期刊目录 PDF
│       └── write_excel.py        # 写入期刊元数据到 xlsx
├── wos-paper-crawler/
│   ├── SKILL.md
│   └── scripts/
│       ├── crawl_wos.py          # Edge + Selenium WoS 爬虫
│       └── postprocess.py        # 城市/国家清洗 + 长格式导出
└── editor-name-crosswalk/
    ├── SKILL.md
    └── scripts/
        └── build_crosswalk.py    # 主编名称变体匹配
```

## 技术说明

### WoS 爬虫

- 使用 **Microsoft Edge**（非 Chrome），适配作者所在机构的校园网络环境。Edge 驱动路径可在 `crawl_wos.py` 中配置。
- 爬虫运行于 `webofscience.clarivate.cn`（中国 WoS 门户），可适配任何 WoS 端点。
- 实现功能：渐进式页面滚动（每页 50 条）、自动补全点击、断点续传、空页自动重启浏览器、按时停止。
- 原始输出使用 `utf-8-sig` 编码，Windows 环境下兼容 GBK。

### 城市/国家清洗

后处理器处理以下已知数据质量问题：
- 加拿大邮政编码（FSA → 城市映射）
- 城市名中嵌入的欧洲邮政编码（"CH-8006 Zurich" → "Zurich"）
- 已知大学从机构名推断城市
- 城市字段误填国家名（"U Arab Emirates" → ARE / Abu Dhabi）
- 机构名规范化（WoS 原始 "1 \nHarvard Univ" → "Harvard Univ"）

### 主编名称匹配

对照表生成器生成名称变体（仅姓氏、仅首字母、完整名）并通过严格的姓氏 + 首字母匹配规则与论文作者数据库比对，最大程度减少误匹配。

## 许可证

MIT

## 引用

如果在研究中使用本工具集，请引用：

```
[Your Name]. (2026). Scholar-City: Academic Journal Data Pipeline.
GitHub 仓库：https://github.com/YOUR_USERNAME/scholar-city
```
