# 豆瓣图书 Top250 爬虫

一个用于爬取豆瓣图书 Top250 数据的 Python 脚本，支持导出 Excel、存入 SQLite 数据库以及生成中文词云。

## 功能特性

- **多页爬取**：自动遍历豆瓣图书 Top250 的全部 10 页（每页 25 本）
- **反爬策略**：每次请求随机 User-Agent，请求间自带 2 秒延迟
- **数据导出**：
  - 导出为 Excel（`data/douban_book_top250.xlsx`）
  - 存入 SQLite 数据库（`db/douban_books.db`）
- **数据分析**：
  - 查询评分最高的前 10 本书
  - 统计出现频率最高的前 10 个出版社
- **可视化**：基于书名生成加权词云图（评分越高的书字号越大），跨平台字体适配

## 爬取字段

| 字段 | 说明 | 类型 |
|------|------|------|
| `book_name` | 书名 | TEXT |
| `author` | 作者 | TEXT |
| `publisher` | 出版社 | TEXT |
| `publish_date` | 出版日期 | TEXT |
| `price` | 定价 | TEXT |
| `rating` | 评分 (0~10) | REAL |
| `comment_count` | 评价人数 | INTEGER |
| `recommendation` | 推荐语 | TEXT |
| `detail_url` | 图书详情页链接 | TEXT (UNIQUE) |

## 环境要求

- Python 3.8+
- 依赖包见下方安装说明

## 安装与运行

### 1. 安装依赖

```bash
pip install requests beautifulsoup4 fake-useragent pandas openpyxl wordcloud matplotlib
```

> **注意**：`fake-useragent` 首次运行需要联网下载浏览器 UA 列表，请确保网络可达。

### 2. 运行脚本

```bash
python Douban_Book_Crawl.py
```

运行完成后会在项目目录下生成三个文件夹：

```
data/       ← douban_book_top250.xlsx（Excel 文件）
db/         ← douban_books.db（SQLite 数据库）
output/     ← book_wordcloud.png（词云图）
```

## 项目结构

```
crawl_demo/
├── Douban_Book_Crawl.py    # 主脚本
├── README.md               # 项目说明
├── data/                   # Excel 输出目录
├── db/                     # SQLite 数据库目录
└── output/                 # 词云图输出目录
```

## 核心流程

```
main()
 ├─ ensure_dirs()              → 创建输出目录
 ├─ crawl_all_pages()          → 爬取 10 页图书数据
 │    └─ get_html()            → 随机 UA + 延时请求
 │    └─ extract_book_info()   → BeautifulSoup 解析
 ├─ save_to_excel()            → 写入 Excel
 ├─ insert_books_to_db()       → 写入 SQLite（INSERT OR IGNORE 去重）
 ├─ db_query_operations()      → 查询评分 Top10 / 出版社统计
 └─ generate_wordcloud()       → 生成中文词云图
```

## 常见问题

### Q: 词云图中文显示为方块？

A: 缺少中文字体。脚本会自动检测系统字体：
- **macOS**：使用 `STHeiti Light`
- **Windows**：使用 `simhei.ttf`
- **Linux**：优先查找 `wqy-zenhei`

如果均未找到，词云生成会跳过但不会报错。也可以手动指定字体路径。

### Q: 爬取过程中出现大量 418 或 403 错误？

A: 豆瓣有反爬机制，建议：
- 增大 `REQUEST_DELAY` 的值（默认 2 秒）
- 检查网络是否正常，`fake-useragent` 无法联网会导致 UA 固定

### Q: 如何只爬取特定页数？

A: 修改 `main()` 中的 `crawl_all_pages()` 调用参数，例如：

```python
all_books = crawl_all_pages(total_pages=5)  # 只爬前 5 页
```
