import os
import sqlite3
import pandas as pd
from pathlib import Path
from tabulate import tabulate

def get_books_db_path():
    """获取 Books 应用的数据库路径"""
    home = str(Path.home())
    return os.path.join(home, 'Library/Containers/com.apple.iBooksX/Data/Documents/BKLibrary/BKLibrary-1-091020131601.sqlite')

def query_books():
    # 获取数据库路径
    db_path = get_books_db_path()
    
    if not os.path.exists(db_path):
        print(f"Error: Books 数据库文件不存在: {db_path}")
        print("请确认你的 Books 应用是否已经打开过，并且有添加过书籍。")
        return
    
    try:
        # 连接数据库
        conn = sqlite3.connect(db_path)
        
        # 查询更详细的书籍信息
        query = """
        SELECT 
            ZASSETID as ID,
            ZTITLE as 标题,
            ZAUTHOR as 作者,
            ZGENRE as 类型,
            ZYEAR as 年份,
            ZEPUBID as ISBN,
            ZLANGUAGE as 语言,
            ZPAGECOUNT as 页数,
            ZRATING as 评分,
            ZISFINISHED as 是否完成,
            ZISEXPLICIT as 是否限制级,
            ZPATH as 文件路径
        FROM ZBKLIBRARYASSET
        WHERE ZTITLE IS NOT NULL
        ORDER BY ZASSETID DESC
        """
        
        # 读取数据到 DataFrame
        df = pd.read_sql_query(query, conn)
        
        # 关闭数据库连接
        conn.close()
        
        if df.empty:
            print("没有找到任何书籍数据！")
            return
            
        # 处理布尔值
        df['是否完成'] = df['是否完成'].map({1: '是', 0: '否'})
        df['是否限制级'] = df['是否限制级'].map({1: '是', 0: '否'})
        
        # 打印基本统计信息
        print(f"\n📚 总共找到 {len(df)} 本书籍")
        print(f"📖 已完成阅读: {len(df[df['是否完成'] == '是'])} 本")
        
        # 使用 tabulate 打印美化后的表格
        print("\n📑 最近阅读的书籍:")
        recent_books = df[['标题', '作者', '类型', '页数', '评分', '是否完成']].head(10)
        print(tabulate(recent_books, headers='keys', tablefmt='pretty', showindex=False))
        
        # 按类型统计
        print("\n📊 书籍类型统计:")
        genre_stats = df['类型'].value_counts().head()
        print(tabulate(genre_stats.reset_index(), headers=['类型', '数量'], tablefmt='pretty', showindex=False))
        
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == '__main__':
    query_books()
