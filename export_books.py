import os
import sqlite3
import pandas as pd
from pathlib import Path

def get_books_db_path():
    """获取 Books 应用的数据库路径"""
    home = str(Path.home())
    return os.path.join(home, 'Library/Containers/com.apple.iBooksX/Data/Documents/BKLibrary/BKLibrary-1-091020131601.sqlite')

def export_books():
    # 获取数据库路径
    db_path = get_books_db_path()
    
    if not os.path.exists(db_path):
        print(f"Error: Books 数据库文件不存在: {db_path}")
        print("请确认你的 Books 应用是否已经打开过，并且有添加过书籍。")
        return
    
    try:
        # 连接数据库
        conn = sqlite3.connect(db_path)
        
        # 查询书籍信息
        query = """
        SELECT 
            ZASSETID as id,
            ZTITLE as title,
            ZAUTHOR as author,
            ZGENRE as genre,
            ZYEAR as year,
            ZEPUBID as isbn
        FROM ZBKLIBRARYASSET
        WHERE ZEPUBID IS NOT NULL
        """
        
        # 读取数据到 DataFrame
        df = pd.read_sql_query(query, conn)
        
        # 关闭数据库连接
        conn.close()
        
        if df.empty:
            print("没有找到任何书籍数据！")
            return
        
        # 导出到 Excel
        output_file = 'books_export.xlsx'
        df.to_excel(output_file, index=False)
        print(f"成功导出 {len(df)} 本书籍信息到 {output_file}")
        print(f"导出文件位置: {os.path.abspath(output_file)}")
        
    except Exception as e:
        print(f"导出过程中出错: {str(e)}")
        print("请确保你已经打开过 Books 应用并添加过书籍。")

if __name__ == '__main__':
    export_books()
