import os
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

def get_annotation_db_path():
    """获取 Books 应用的笔记数据库路径"""
    home = str(Path.home())
    return os.path.join(home, 'Library/Containers/com.apple.iBooksX/Data/Documents/AEAnnotation/AEAnnotation_v10312011_1727_local.sqlite')

def get_books_db_path():
    """获取 Books 应用的数据库路径"""
    home = str(Path.home())
    return os.path.join(home, 'Library/Containers/com.apple.iBooksX/Data/Documents/BKLibrary/BKLibrary-1-091020131601.sqlite')

def query_annotations():
    # 获取数据库路径
    annotation_db_path = get_annotation_db_path()
    books_db_path = get_books_db_path()
    
    if not os.path.exists(annotation_db_path):
        print(f"Error: Books 笔记数据库文件不存在: {annotation_db_path}")
        return
        
    if not os.path.exists(books_db_path):
        print(f"Error: Books 数据库文件不存在: {books_db_path}")
        return
    
    try:
        # 连接笔记数据库
        anno_conn = sqlite3.connect(annotation_db_path)
        books_conn = sqlite3.connect(books_db_path)
        
        # 查看表结构
        tables_query = """
        SELECT name FROM sqlite_master 
        WHERE type='table' AND name NOT LIKE 'sqlite_%'
        """
        tables = pd.read_sql_query(tables_query, anno_conn)
        print("\n可用的表:")
        for table in tables['name']:
            print(f"- {table}")
            
        # 查询注释数据
        query = """
        SELECT 
            ZANNOTATIONASSETID as 书籍ID,
            ZANNOTATIONSELECTEDTEXT as 划线内容,
            ZANNOTATIONNOTE as 笔记内容,
            ZANNOTATIONTYPE as 类型,
            ZFUTUREPROOFING5 as 颜色,
            datetime(ZANNOTATIONCREATIONDATE + 978307200, 'unixepoch', 'localtime') as 创建时间,
            datetime(ZANNOTATIONMODIFICATIONDATE + 978307200, 'unixepoch', 'localtime') as 修改时间,
            ZFUTUREPROOFING3 as 章节,
            ZPLLOCATIONRANGESTART as 位置开始,
            ZPLLOCATIONRANGEEND as 位置结束
        FROM ZAEANNOTATION
        WHERE ZANNOTATIONDELETED = 0
        ORDER BY ZANNOTATIONASSETID, ZPLLOCATIONRANGESTART
        """
        
        # 读取数据到 DataFrame
        df = pd.read_sql_query(query, anno_conn)
        
        # 获取书籍信息
        books_query = """
        SELECT 
            ZASSETID as 书籍ID,
            ZTITLE as 书名,
            ZAUTHOR as 作者
        FROM ZBKLIBRARYASSET
        WHERE ZTITLE IS NOT NULL
        """
        books_df = pd.read_sql_query(books_query, books_conn)
        
        # 合并书籍信息
        df = df.merge(books_df, on='书籍ID', how='left')
        
        # 关闭数据库连接
        anno_conn.close()
        books_conn.close()
        
        if df.empty:
            print("没有找到任何划线或笔记！")
            return
            
        # 统计信息
        total_annotations = len(df)
        total_books_with_annotations = df['书籍ID'].nunique()
        total_notes = len(df[df['笔记内容'].notna()])
        total_highlights = len(df[df['划线内容'].notna()])
        
        print(f"\n📚 统计信息:")
        print(f"- 总共有 {total_annotations} 条划线和笔记")
        print(f"- 涉及 {total_books_with_annotations} 本书")
        print(f"- 其中有 {total_highlights} 条划线")
        print(f"- 其中有 {total_notes} 条笔记")
        
        # 导出到Excel
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'books_annotations_{timestamp}.xlsx'
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 按书籍分组导出
            for book_id in df['书籍ID'].unique():
                book_data = df[df['书籍ID'] == book_id]
                book_name = book_data['书名'].iloc[0]
                # 确保sheet名称有效
                sheet_name = str(book_name)[:31]  # Excel限制sheet名最长31字符
                # 移除无效字符
                invalid_chars = [':', '/', '\\', '?', '*', '[', ']']
                for char in invalid_chars:
                    sheet_name = sheet_name.replace(char, '_')
                book_data.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n📝 笔记和划线已导出到: {output_file}")
        print("每本书的笔记和划线保存在单独的sheet中")
        
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == '__main__':
    query_annotations()
