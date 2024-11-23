import os
import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

def get_books_db_path():
    """获取 Books 应用的数据库路径"""
    home = str(Path.home())
    return os.path.join(home, 'Library/Containers/com.apple.iBooksX/Data/Documents/BKLibrary/BKLibrary-1-091020131601.sqlite')

def analyze_books():
    # 获取数据库路径
    db_path = get_books_db_path()
    
    if not os.path.exists(db_path):
        print(f"Error: Books 数据库文件不存在: {db_path}")
        print("请确认你的 Books 应用是否已经打开过，并且有添加过书籍。")
        return
    
    try:
        # 连接数据库
        conn = sqlite3.connect(db_path)
        
        # 查询书籍基本信息
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
            ZISEXPLICIT as 是否限制级
        FROM ZBKLIBRARYASSET
        WHERE ZTITLE IS NOT NULL
        """
        
        # 读取数据到 DataFrame
        df = pd.read_sql_query(query, conn)
        conn.close()
        
        if df.empty:
            print("没有找到任何书籍数据！")
            return
            
        # 处理布尔值
        df['是否完成'] = df['是否完成'].map({1: '是', 0: '否'})
        df['是否限制级'] = df['是否限制级'].map({1: '是', 0: '否'})
        
        # 创建统计数据
        stats = []
        
        # 1. 总体统计
        total_books = len(df)
        completed_books = len(df[df['是否完成'] == '是'])
        rated_books = len(df[df['评分'] > 0])
        
        stats.append({
            '统计类型': '总体统计',
            '指标': '总书籍数',
            '数值': total_books,
            '百分比': '100%'
        })
        stats.append({
            '统计类型': '总体统计',
            '指标': '已完成阅读',
            '数值': completed_books,
            '百分比': f'{(completed_books/total_books)*100:.1f}%'
        })
        stats.append({
            '统计类型': '总体统计',
            '指标': '已评分书籍',
            '数值': rated_books,
            '百分比': f'{(rated_books/total_books)*100:.1f}%'
        })
        
        # 2. 评分统计
        for rating in range(1, 6):
            count = len(df[df['评分'] == rating])
            stats.append({
                '统计类型': '评分统计',
                '指标': f'{rating}星书籍',
                '数值': count,
                '百分比': f'{(count/total_books)*100:.1f}%'
            })
            
        # 3. 年份统计
        df['年份'] = pd.to_numeric(df['年份'], errors='coerce')
        year_stats = df['年份'].value_counts().sort_index()
        for year, count in year_stats.items():
            if pd.notna(year) and year > 1900:  # 过滤掉无效年份
                stats.append({
                    '统计类型': '出版年份',
                    '指标': f'{int(year)}年',
                    '数值': count,
                    '百分比': f'{(count/total_books)*100:.1f}%'
                })
        
        # 4. 语言统计
        lang_stats = df['语言'].value_counts()
        for lang, count in lang_stats.items():
            if pd.notna(lang):
                stats.append({
                    '统计类型': '语言分布',
                    '指标': lang,
                    '数值': count,
                    '百分比': f'{(count/total_books)*100:.1f}%'
                })
        
        # 5. 页数统计
        df['页数'] = pd.to_numeric(df['页数'], errors='coerce')
        page_ranges = [(0, 100), (101, 300), (301, 500), (501, 1000), (1001, float('inf'))]
        range_labels = ['100页以下', '101-300页', '301-500页', '501-1000页', '1000页以上']
        
        for (start, end), label in zip(page_ranges, range_labels):
            count = len(df[(df['页数'] > start) & (df['页数'] <= end)])
            stats.append({
                '统计类型': '页数分布',
                '指标': label,
                '数值': count,
                '百分比': f'{(count/total_books)*100:.1f}%'
            })
        
        # 创建统计DataFrame
        stats_df = pd.DataFrame(stats)
        
        # 导出到Excel，创建多个sheet
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'books_analysis_{timestamp}.xlsx'
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 导出统计数据
            stats_df.to_excel(writer, sheet_name='统计分析', index=False)
            
            # 导出原始数据
            df.to_excel(writer, sheet_name='原始数据', index=False)
            
            # 按评分导出Top书籍
            top_rated = df[df['评分'] > 0].sort_values('评分', ascending=False)
            top_rated[['标题', '作者', '评分', '是否完成']].head(20).to_excel(
                writer, sheet_name='评分最高', index=False)
            
            # 导出未读书籍
            unread = df[df['是否完成'] != '是'][['标题', '作者', '类型', '页数']]
            unread.to_excel(writer, sheet_name='未读书籍', index=False)
        
        print(f"\n📊 统计分析已导出到: {output_file}")
        print("\n包含以下sheet:")
        print("1. 统计分析: 包含各类统计数据")
        print("2. 原始数据: 所有书籍的原始信息")
        print("3. 评分最高: 评分最高的20本书")
        print("4. 未读书籍: 尚未阅读完成的书籍列表")
        
    except sqlite3.Error as e:
        print(f"数据库错误: {e}")
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == '__main__':
    analyze_books()
