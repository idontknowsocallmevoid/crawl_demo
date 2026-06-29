# 导包
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import re
import pandas as pd
import os
import sqlite3
from sqlite3 import Error
import time
import platform
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import matplotlib

# 反爬虫：每次请求随机 UA，fallback 用于联网失败时兜底
ua = UserAgent(
    fallback='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36'
)

# 先给出爬取网址和各存储路径
BASE_URL = 'https://book.douban.com/top250'
DATA_DIR = 'data'
DB_DIR = 'db'
OUTPUT_DIR = 'output'
EXCEL_PATH = os.path.join(DATA_DIR, 'douban_book_top250.xlsx')
DB_PATH = os.path.join(DB_DIR, 'douban_books.db')
WORD_CLOUD_PATH = os.path.join(OUTPUT_DIR, 'book_wordcloud.png')

# 请求间隔（秒），避免频繁请求触发反爬
REQUEST_DELAY = 2


def ensure_dirs():
    """确保输出目录存在"""
    for d in (DATA_DIR, DB_DIR, OUTPUT_DIR):
        os.makedirs(d, exist_ok=True)


# 获取豆瓣网页源码
def get_html(url):
    """发起 GET 请求，自动轮换 UA，返回 HTML 文本"""
    headers = {
        "User-Agent": ua.random,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Connection": "keep-alive",
    }
    try:
        response = requests.get(url, timeout=20, headers=headers)
        response.raise_for_status()
        response.encoding = 'utf-8'
        return response.text
    except requests.RequestException as e:
        print(f"  [ERROR] 获取网页失败: {e}")
        return ""
    except Exception as e:
        print(f"  [ERROR] 未知错误: {e}")
        return ""


def clean_text(text):
    if not text:
        return ""
    return re.sub(r'\s+', ' ', text.strip())


# 获取单页图书信息
def extract_book_info(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('tr', attrs={'class': 'item'})
    book_list = []
    for idx, item in enumerate(items):
        try:
            book_info = {}

            # 获取每本书的详细 url
            nbg_tag = item.find('a', attrs={'class': 'nbg'})
            if nbg_tag and nbg_tag.get('href'):
                book_info['detail_url'] = nbg_tag['href']
            else:
                # 部分条目可能没有 nbg 链接，尝试从其他 <a> 提取
                alt_link = item.find('a')
                book_info['detail_url'] = alt_link['href'] if alt_link else ""

            # 获取书名
            title_tag = item.find('div', attrs={'class': 'pl2'}).find('a')
            book_info['book_name'] = clean_text(title_tag['title']) if title_tag else ""

            # 获取作者、出版社、出版时间、定价
            pl_tag = item.find('p', attrs={'class': 'pl'})
            if pl_tag:
                pl_text = clean_text(pl_tag.get_text())
                pl_parts = pl_text.split('/')
                if len(pl_parts) >= 4:
                    book_info['author'] = clean_text(pl_parts[0])
                    book_info['publisher'] = clean_text(pl_parts[1])
                    book_info['publish_date'] = clean_text(pl_parts[2])
                    book_info['price'] = clean_text(pl_parts[3])
                elif len(pl_parts) == 3:
                    # 有些书没有出版社信息，只有 作者/出版时间/定价
                    book_info['author'] = clean_text(pl_parts[0])
                    book_info['publisher'] = ""
                    book_info['publish_date'] = clean_text(pl_parts[1])
                    book_info['price'] = clean_text(pl_parts[2])
                else:
                    book_info['author'] = clean_text(pl_parts[0]) if pl_parts else ""
                    book_info['publisher'] = ""
                    book_info['publish_date'] = ""
                    book_info['price'] = ""
            else:
                book_info['author'] = ""
                book_info['publisher'] = ""
                book_info['publish_date'] = ""
                book_info['price'] = ""

            # 获取评分
            rating_tag = item.find('span', attrs={'class': 'rating_nums'})
            book_info['rating'] = clean_text(rating_tag.get_text()) if rating_tag else ""

            # 获取评论数（格式如 "85487人评价"）
            comment_tag = item.find('span', attrs={'class': 'pl'})
            if comment_tag:
                comment_match = re.search(r'(\d+)人评价', comment_tag.get_text())
                book_info['comment_count'] = comment_match.group(1) if comment_match else '0'
            else:
                book_info['comment_count'] = '0'

            # 获取推荐语
            quote_tag = item.find('span', attrs={'class': 'inq'})
            book_info['recommendation'] = clean_text(quote_tag.get_text()) if quote_tag else ""

            book_list.append(book_info)
        except Exception as e:
            print(f"  [WARN] 解析第 {idx + 1} 本书时出错: {e}")
            continue

    return book_list


# 获取多页图书信息
def crawl_all_pages(total_pages=10):
    all_books = []
    for page in range(total_pages):
        start = page * 25
        url = f"{BASE_URL}?start={start}"
        print(f"正在爬取第{page + 1}/{total_pages}页: {url}")

        html = get_html(url)
        if not html:
            print(f"  [WARN] 第 {page + 1} 页获取失败，跳过")
            continue

        page_books = extract_book_info(html)
        all_books.extend(page_books)
        print(f"  第{page + 1}页爬取完成，获取{len(page_books)}本图书")

        # 请求间隔，降低被封 IP 的风险
        if page < total_pages - 1:
            time.sleep(REQUEST_DELAY)

    print(f"\n爬取结束，共获取{len(all_books)}本图书信息")
    return all_books


# 将数据保存至excel
def save_to_excel(books):
    df = pd.DataFrame(books)
    columns_order = ['book_name', 'author', 'publisher', 'publish_date', 'price',
                     'rating', 'comment_count', 'recommendation', 'detail_url']
    df = df[columns_order]
    try:
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='豆瓣图书Top250', index=False)
        print(f"\nExcel文件已保存至: {os.path.abspath(EXCEL_PATH)}")
    except Exception as e:
        print(f"保存Excel失败: {e}")


# 创建数据库连接
def create_db_connection():
    try:
        conn = sqlite3.connect(DB_PATH)
        print(f"\nSQLite数据库连接成功（版本: {sqlite3.version}）")
        return conn
    except Error as e:
        print(f"数据库连接失败：{e}")
        return None


# 创建图书表（使用正确的数据类型）
def create_books_table(conn):
    try:
        cursor = conn.cursor()
        create_table_sql = """
        CREATE TABLE IF NOT EXISTS books (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            book_name TEXT NOT NULL,
            author TEXT,
            publisher TEXT,
            publish_date TEXT,
            price TEXT,
            rating REAL,
            comment_count INTEGER,
            recommendation TEXT,
            detail_url TEXT UNIQUE
        );
        """
        cursor.execute(create_table_sql)
        conn.commit()
        print("图书表创建成功（若不存在）")
    except Error as e:
        print(f"创建表失败: {e}")


# 插入图书数据到数据库（类型转换 + 去重）
def insert_books_to_db(conn, books):
    try:
        cursor = conn.cursor()
        insert_sql = """
        INSERT OR IGNORE INTO books (
            book_name, author, publisher, publish_date, price,
            rating, comment_count, recommendation, detail_url
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
        """
        data = []
        for book in books:
            # 类型转换：rating -> float, comment_count -> int
            try:
                rating = float(book.get('rating', 0))
            except (ValueError, TypeError):
                rating = None
            try:
                comment_count = int(book.get('comment_count', 0))
            except (ValueError, TypeError):
                comment_count = 0
            data.append((
                book.get('book_name', ''),
                book.get('author', ''),
                book.get('publisher', ''),
                book.get('publish_date', ''),
                book.get('price', ''),
                rating,
                comment_count,
                book.get('recommendation', ''),
                book.get('detail_url', ''),
            ))

        cursor.executemany(insert_sql, data)
        conn.commit()
        print(f"成功插入{cursor.rowcount}条数据（已存在的记录将跳过）")
    except Error as e:
        print(f"插入数据失败：{e}")


# 纯查询操作，不做任何修改/删除，避免污染数据
def db_query_operations(conn):
    cursor = conn.cursor()

    # 查询评分最高的前10本书
    print("\n【查询】评分最高的前10本书:")
    cursor.execute("""
        SELECT book_name, author, rating, comment_count
        FROM books WHERE rating IS NOT NULL
        ORDER BY rating DESC LIMIT 10
    """)
    top_books = cursor.fetchall()
    for i, book in enumerate(top_books, 1):
        print(f"  {i}. 《{book[0]}》 - {book[1]} (评分: {book[2]}, 评价人数: {book[3]})")

    # 查询各出版社出现次数 Top 10
    print("\n【查询】出现次数最多的前10个出版社:")
    cursor.execute("""
        SELECT publisher, COUNT(*) as cnt
        FROM books WHERE publisher != ''
        GROUP BY publisher ORDER BY cnt DESC LIMIT 10
    """)
    publishers = cursor.fetchall()
    for i, (pub, cnt) in enumerate(publishers, 1):
        print(f"  {i}. {pub} ({cnt}本)")


# 基于书名和推荐语生成词云（兼容跨平台字体）
def generate_wordcloud(books):
    """
    基于书名生成词云，按评分高低调整文字权重（评分越高，字号越大）
    """
    weighted_text = {}
    for book in books:
        book_name = clean_text(book.get('book_name', ''))
        rating_str = clean_text(book.get('rating', '0'))

        if not book_name or not rating_str:
            continue

        try:
            rating = float(rating_str)
        except ValueError:
            rating = 0.0

        # 权重公式：评分越高权重越大
        weight = max(0.1, 1.0 + (rating - 8) * 0.5)
        weighted_text[book_name] = weighted_text.get(book_name, 0) + weight

    # 跨平台字体适配
    system = platform.system()
    if system == 'Windows':
        font_path = 'C:/Windows/Fonts/simhei.ttf'
    elif system == 'Darwin':  # macOS
        font_path = '/System/Library/Fonts/STHeiti Light.ttc'
    else:  # Linux
        # 尝试常见中文字体路径
        linux_fonts = [
            '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
            '/usr/share/fonts/wqy-zenhei/wqy-zenhei.ttc',
            '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
        ]
        font_path = next((f for f in linux_fonts if os.path.exists(f)), None)

    # 如果找不到中文字体，尝试用 matplotlib 内置的
    if not font_path or not os.path.exists(font_path):
        # 设置 matplotlib 使用默认字体回退
        matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'sans-serif']
        matplotlib.rcParams['axes.unicode_minus'] = False
        font_path = None
    else:
        matplotlib.rcParams['font.sans-serif'] = []
        matplotlib.rcParams['axes.unicode_minus'] = False

    try:
        wc_kwargs = dict(
            width=1200,
            height=800,
            background_color='white',
            max_words=200,
            min_font_size=10,
            max_font_size=100,
            random_state=42,
            contour_width=2,
            contour_color='steelblue',
        )
        if font_path:
            wc_kwargs['font_path'] = font_path

        wordcloud = WordCloud(**wc_kwargs).generate_from_frequencies(weighted_text)

        plt.figure(figsize=(15, 10))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.title('豆瓣图书Top250词云', fontsize=20, pad=20, fontweight='bold')
        plt.tight_layout()

        wordcloud.to_file(WORD_CLOUD_PATH)
        print(f"\n词云图已保存至: {os.path.abspath(WORD_CLOUD_PATH)}")
    except Exception as e:
        print(f"生成词云失败: {e}")
        print("  （可能是缺少中文字体，词云图未生成）")
    finally:
        plt.close()


# 设置主函数并执行
def main():
    # 确保目录存在
    ensure_dirs()

    # 爬取所有页面数据
    all_books = crawl_all_pages()
    if not all_books:
        print("未获取到图书数据，程序终止")
        return

    # 保存到 Excel
    save_to_excel(all_books)

    # 数据库操作
    conn = create_db_connection()
    if conn:
        create_books_table(conn)
        insert_books_to_db(conn, all_books)
        db_query_operations(conn)  # 仅做查询，不修改/删除数据
        conn.close()
        print("\n数据库连接已关闭")

    # 生成词云
    generate_wordcloud(all_books)
    print("\n所有任务完成")


if __name__ == '__main__':
    main()
