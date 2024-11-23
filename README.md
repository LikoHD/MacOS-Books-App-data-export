# macOS Books 导出工具

这个工具用于从 macOS 的 Books 应用程序中导出书籍的标题和 ISBN 信息到 Excel 文件。以及对 Books 应用中的书籍数据进行查询和分析。

## 使用说明

1. 确保你有 Python 3.x 环境

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 运行脚本：
```bash
python export_books.py
```

4. 查看结果：
   - 脚本会生成 `books_export.xlsx` 文件
   - Excel 文件包含两列：title（书名）和 isbn（ISBN）

## 其他脚本说明

本工具包含以下脚本：

1. `export_books.py`
   - 从 macOS Books 应用的数据库中导出书籍信息
   - 生成 `books_export.xlsx` 文件，包含书名和 ISBN 信息
   - 主要用于数据导出，格式简单，便于后续使用

2. `query_annotations.py`
   - 查询 Books 应用中的笔记和划线数据
   - 连接笔记数据库和书籍数据库
   - 可以查看所有的笔记、划线内容，以及对应的书籍信息

3. `query_books.py`
   - 查询并显示 Books 应用中的详细书籍信息
   - 包含完整信息：ID、标题、作者、类型、年份、ISBN、语言、页数、评分等
   - 以表格形式展示查询结果

4. `analyze_books.py`
   - 对 Books 应用中的书籍数据进行统计分析
   - 提供各种统计信息，如书籍总数、作者分布、类型分布等
   - 生成数据分析报告，帮助了解阅读情况

这些脚本形成了一个完整的工具集，可以帮助用户导出、查询和分析 Books 应用中的电子书数据。

## 注意事项

- 脚本会直接读取 macOS Books 应用的数据库
- 只会导出有 ISBN 的书籍信息
- 如果遇到权限问题，请确保给予终端访问权限
