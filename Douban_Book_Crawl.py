# 导包
import requests
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
import re
import pandas as pd
import os
import sqlite3
from sqlite3 import Error
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# 反爬虫，不设置反爬虫时状态码为418
ua = UserAgent()
headers = {
    "User-Agent": ua.chrome
}

# 先给出爬取网址和各存储路径
BASE_URL = 'https://book.douban.com/top250'
EXCEL_PATH = 'douban_book_top250.xlsx'
DB_PATH = 'douban_books.db'
WORD_CLOUD_PATH = 'book_wordcloud.png'

# 获取豆瓣网页源码
def get_html(url):
    try:
        response = requests.get(url, timeout=20, headers=headers)
        response.raise_for_status()
        response.encoding = 'utf-8'
        return response.text
    except:
        print("获取网页失败")
        return ""
def clean_text(text):
    if not text:
        return ""
    return re.sub(r'\s+', ' ', text.strip())

# 获取单页图书信息
def extract_book_info(html):
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('tr', attrs={'class': 'item'})
    # 全部书收集到book_list
    book_list = []
    for item in items:

        # 每本书的信息
        book_info = {}

        # 获取每本书的详细url
        detail_url = item.find('a', attrs={'class': 'nbg'})['href']
        book_info['detail_url'] = detail_url

        # 获取书名
        title_tag = item.find('div', attrs={'class': 'pl2'}).find('a')
        book_info['book_name'] = clean_text(title_tag['title'])

        # 获取作者、出版社、出版时间、定价
        pl_text = clean_text(item.find('p', attrs={'class': 'pl'}).get_text())
        pl_parts = pl_text.split('/')
        if len(pl_parts) >= 4:
            book_info['author'] = clean_text(pl_parts[0])
            book_info['publisher'] = clean_text(pl_parts[1])
            book_info['publish_date'] = clean_text(pl_parts[2])
            book_info['price'] = clean_text(pl_parts[3])
        else:
            book_info['author'] = ""
            book_info['publisher'] = ""
            book_info['publish_date'] = ""
            book_info['price'] = ""

        # 获取评分
        rating_tag = item.find('span', attrs={'class': 'rating_nums'})
        book_info['rating'] = clean_text(rating_tag.get_text()) if rating_tag else ""

        # 获取评论数
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
    return book_list

# 获取多页图书信息
def crawl_all_pages(total_pages=10):
    all_books = []
    for page in range(total_pages):
        start = page * 25
        url = f"{BASE_URL}?start={start}"
        print(f"正在爬取第{page + 1}页: {url}")

        html = get_html(url)
        if not html:
            continue

        page_books = extract_book_info(html)
        all_books.extend(page_books)
        print(f"第{page + 1}页爬取完成，获取{len(page_books)}本图书")

    print(f"\n爬取结束，共获取{len(all_books)}本图书信息")
    return all_books

# 将数据保存至excel
def save_to_excel(books):
    df = pd.DataFrame(books)

    # 调整列顺序
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
    conn = None
    try:
        conn = sqlite3.connect(DB_PATH)
        print(f"\nSQLite数据库连接成功（版本: {sqlite3.version}）")
    except Error as e:
        print(f"数据库连接失败：{e}")
    return conn

# 创建图书表
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
            rating TEXT,
            comment_count TEXT,
            recommendation TEXT,
            detail_url TEXT UNIQUE
        );
        """
        cursor.execute(create_table_sql)
        conn.commit()
        print("图书表创建成功（若不存在）")
    except Error as e:
        print(f"创建表失败: {e}")

# 插入图书数据到数据库
def insert_books_to_db(conn, books):
    try:
        cursor = conn.cursor()
        insert_sql = """
        INSERT OR IGNORE INTO books (
            book_name, author, publisher, publish_date, price,
            rating, comment_count, recommendation, detail_url
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
        """
        data = [
            (
                book['book_name'], book['author'], book['publisher'], book['publish_date'],
                book['price'], book['rating'], book['comment_count'], book['recommendation'],
                book['detail_url']
            ) for book in books
        ]
        cursor.executemany(insert_sql, data)
        conn.commit()
        print(f"成功插入{cursor.rowcount}条数据（已存在的记录将跳过）")
    except Error as e:
        print(f"插入数据失败：{e}")

# 数据库增删改查操作
"""
     为避免破坏数据，将数据修改后撤回修改，删除不存在的数据
"""
def db_crud_operations(conn):
    cursor = conn.cursor()

    # 查询评分大于9.5的图书
    print("\n【查询】评分≥9.5的图书:")
    cursor.execute("SELECT book_name, author, rating FROM books WHERE CAST(rating AS FLOAT) >= 9.5")
    high_rating_books = cursor.fetchall()
    for book in high_rating_books:
        print(f"书名: {book[0]}, 作者: {book[1]}, 评分: {book[2]}")

    # 修改《红楼梦》的推荐语
    print("\n【修改】更新《红楼梦》的推荐语")
    update_sql = """
    UPDATE books SET recommendation = '中国古典小说的巅峰之作' 
    WHERE book_name = '红楼梦';
    """
    cursor.execute(update_sql)
    conn.commit()
    print(f"修改影响行数: {cursor.rowcount}")

    # 撤回对《红楼梦》推荐语的修改
    print("\n【修改】撤回对《红楼梦》推荐语的修改")
    update_sql = """
        UPDATE books SET recommendation = '都云作者痴，谁解其中味？' 
        WHERE book_name = '红楼梦';
        """
    cursor.execute(update_sql)
    conn.commit()
    print(f"修改影响行数: {cursor.rowcount}")

    # 删除某本不存在的图书
    print("\n【删除】尝试删除不存在的图书（测试）")
    delete_sql = "DELETE FROM books WHERE book_name = '不存在的图书';"
    cursor.execute(delete_sql)
    conn.commit()
    print(f"删除影响行数: {cursor.rowcount}")

# 基于书名和推荐语生成词云
def generate_wordcloud(books):
    """
    基于书名生成词云，按评分高低调整文字权重（评分越高，字号越大）
    :param books: 图书列表（含book_name/recommendation/rating字段）
    """
    # 数据预处理：过滤无效数据，转换评分为浮点型
    weighted_text = {}
    for book in books:
        # 过滤空值
        book_name = clean_text(book.get('book_name', ''))
        rating_str = clean_text(book.get('rating', '0'))

        if not book_name or rating_str == '0':
            continue

        # 转换评分为浮点型（处理异常值）
        try:
            rating = float(rating_str)
        except ValueError:
            rating = 0.0

        # 计算权重：以8分为基准，评分越高权重越大
        # 权重公式：基础权重 = 1 + (评分-8)*0.5（评分≥8时权重递增，<8时权重降低）
        weight = 1.0 + (rating - 8) * 0.5
        weight = max(0.1, weight)  # 避免权重为负数

        # 将书名和推荐语按权重加入字典
        if book_name:
            weighted_text[book_name] = weighted_text.get(book_name, 0) + weight

    # 配置词云参数
    font_path = 'C:/Windows/Fonts/simhei.ttf'

    wordcloud = WordCloud(
        width=1200,  # 词云宽度
        height=800,  # 词云高度
        background_color='white',  # 背景色
        font_path=font_path,  # 中文字体路径
        max_words=200,  # 最大显示词数
        min_font_size=10,  # 最小字号
        max_font_size=100,  # 最大字号
        random_state=42,  # 固定随机种子（保证生成结果一致）
        contour_width=2,  # 轮廓宽度
        contour_color='steelblue'  # 轮廓颜色
    ).generate_from_frequencies(weighted_text)

    # 绘制并保存词云图
    plt.figure(figsize=(15, 10))
    plt.imshow(wordcloud, interpolation='bilinear')  # 双线性插值使图像更平滑
    plt.axis('off')  # 关闭坐标轴
    plt.title('豆瓣图书Top250词云', fontsize=20, pad=20, fontweight='bold')
    plt.tight_layout()  # 调整布局

    # 保存词云图
    try:
        wordcloud.to_file(WORD_CLOUD_PATH)
        print(f"\n词云图已保存至: {os.path.abspath(WORD_CLOUD_PATH)}")
    except Exception as e:
        print(f"保存词云图失败: {e}")
    finally:
        plt.close()
# 设置主函数并执行
def main():
    # 爬取所有页面数据
    all_books = crawl_all_pages()
    if not all_books:
        print("未获取到图书数据，程序终止")
        return

    # 保存到Excel
    save_to_excel(all_books)

    # 数据库操作
    conn = create_db_connection()
    if conn:
        create_books_table(conn)
        insert_books_to_db(conn, all_books)
        db_crud_operations(conn)
        conn.close()
        print("\n数据库连接已关闭")

    # 生成词云
    generate_wordcloud(all_books)
    print("\n所有任务完成")

if __name__ == '__main__':
    main()