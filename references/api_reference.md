# Office Pro API Reference

## WordProcessor

Word 文档处理器，提供企业级文档自动化功能。

### 初始化

```python
from office_pro import WordProcessor

# 使用默认模板目录
wp = WordProcessor()

# 指定模板目录
wp = WordProcessor(template_dir='/path/to/templates')
```

### 文档创建与加载

#### create_document()
创建新文档

```python
doc = wp.create_document()
```

#### load_document(path)
加载现有文档

```python
doc = wp.load_document('/path/to/document.docx')
```

#### load_template(template_name)
加载模板文件

```python
wp.load_template('meeting-minutes.docx')
```

### 内容添加

#### add_heading(text, level=1)
添加标题

```python
wp.add_heading('章节标题', level=1)
wp.add_heading('小节标题', level=2)
```

#### add_paragraph(text, style=None)
添加段落

```python
wp.add_paragraph('这是一段普通文本。')
wp.add_paragraph('引用文本', style='Quote')
```

#### add_table(rows, cols, style=None)
添加表格

```python
table = wp.add_table(3, 3, style='Table Grid')
table.cell(0, 0).text = '表头1'
table.cell(0, 1).text = '表头2'
```

#### add_picture(image_path, width=None, height=None)
添加图片

```python
wp.add_picture('photo.jpg', width=5.0)  # 宽度 5 英寸
```

#### add_page_break()
添加分页符

```python
wp.add_page_break()
```

### 页眉页脚

#### add_header(text, align='center')
添加页眉

```python
wp.add_header('公司机密 - 内部使用', align='center')
```

#### add_footer(text, align='center')
添加页脚

```python
wp.add_footer('© 2024 Company Name', align='center')
```

#### add_page_number(location='footer', align='center')
添加页码

```python
wp.add_page_number(location='footer', align='right')
```

### 模板渲染

#### render_template(context)
渲染模板

```python
context = {
    'company_name': 'ABC 公司',
    'date': '2024-03-15',
    'items': [
        {'name': '产品 A', 'price': 100},
        {'name': '产品 B', 'price': 200}
    ]
}
wp.render_template(context)
```

#### render_and_save(context, output_path)
渲染并保存

```python
wp.render_and_save(context, 'output/contract.docx')
```

### 文档保存

#### save(path)
保存文档

```python
wp.save('output/document.docx')
```

### 文档信息

#### get_document_info()
获取文档信息

```python
info = wp.get_document_info()
print(info['paragraph_count'])
print(info['table_count'])
```

## ExcelProcessor

Excel 表格处理器，提供企业级报表自动化功能。

### 初始化

```python
from office_pro import ExcelProcessor

# 使用默认模板目录
ep = ExcelProcessor()

# 指定模板目录
ep = ExcelProcessor(template_dir='/path/to/templates')
```

### 工作簿操作

#### create_workbook()
创建工作簿

```python
wb = ep.create_workbook()
```

#### load_workbook(path, data_only=False)
加载工作簿

```python
wb = ep.load_workbook('data.xlsx')
```

#### load_template(template_name)
加载模板

```python
ep.load_template('sales-report.xlsx')
```

#### save(path)
保存工作簿

```python
ep.save('output/report.xlsx')
```

### 工作表操作

#### get_sheet(name=None)
获取工作表

```python
# 获取活动工作表
ws = ep.get_sheet()

# 获取指定工作表
ws = ep.get_sheet('Sales')
```

#### create_sheet(title, index=None)
创建工作表

```python
ep.create_sheet('New Sheet')
ep.create_sheet('First', index=0)
```

#### remove_sheet(name)
删除工作表

```python
ep.remove_sheet('Old Sheet')
```

### 单元格操作

#### write_cell(cell, value, sheet=None)
写入单元格

```python
ep.write_cell('A1', 'Hello')
ep.write_cell('B2', 12345)
ep.write_cell('C3', datetime.now())
```

#### read_cell(cell, sheet=None)
读取单元格

```python
value = ep.read_cell('A1')
```

#### write_range(start_cell, data, sheet=None)
写入数据区域

```python
data = [
    ['Name', 'Age', 'City'],
    ['Alice', 25, 'NYC'],
    ['Bob', 30, 'LA']
]
ep.write_range('A1', data)
```

### 样式设置

#### set_cell_style(cell, font=None, fill=None, border=None, alignment=None, number_format=None, sheet=None)
设置单元格样式

```python
# 字体
ep.set_cell_style('A1', font={
    'name': 'Arial',
    'size': 14,
    'bold': True,
    'color': 'FF0000'
})

# 填充
ep.set_cell_style('B2', fill={
    'patternType': 'solid',
    'fgColor': {'rgb': 'FFFF00'}
})

# 边框
ep.set_cell_style('C3', border={
    'style': 'thin',
    'color': '000000'
})

# 对齐
ep.set_cell_style('D4', alignment={
    'horizontal': 'center',
    'vertical': 'center',
    'wrap_text': True
})

# 数字格式
ep.set_cell_style('E5', number_format='#,##0.00')
```

#### set_column_width(column, width, sheet=None)
设置列宽

```python
ep.set_column_width('A', 20)
```

#### set_row_height(row, height, sheet=None)
设置行高

```python
ep.set_row_height(1, 30)
```

#### merge_cells(range_str, sheet=None)
合并单元格

```python
ep.merge_cells('A1:C1')
```

### 图表

#### add_chart(chart_type, data_range, title=None, position=None, sheet=None)
添加图表

```python
# 柱状图
ep.add_chart('bar', 'A1:B10', title='Sales Data', position='D1')

# 折线图
ep.add_chart('line', 'A1:C20', title='Trend')

# 饼图
ep.add_chart('pie', 'A1:B5', title='Market Share')
```

### 数据导入导出

#### import_csv(csv_path, sheet=None, delimiter=',', encoding='utf-8')
从 CSV 导入

```python
ep.import_csv('data.csv', delimiter=',')
```

#### export_csv(csv_path, sheet=None, delimiter=',', encoding='utf-8')
导出到 CSV

```python
ep.export_csv('output.csv')
```

#### read_dataframe(sheet=None, header=0, range_str=None)
读取为 pandas DataFrame

```python
df = ep.read_dataframe(sheet='Sales')
```

#### write_dataframe(df, start_cell='A1', sheet=None, include_header=True, index=False)
写入 pandas DataFrame

```python
import pandas as pd

df = pd.DataFrame({
    'Name': ['Alice', 'Bob'],
    'Age': [25, 30]
})

ep.write_dataframe(df, start_cell='A1')
```

### 模板渲染

#### render_template(data)
渲染模板

```python
data = {
    'report_date': '2024-03-15',
    'company': 'ABC Corp',
    'sales': [
        {'product': 'A', 'amount': 1000},
        {'product': 'B', 'amount': 2000}
    ]
}

ep.render_template(data)
```

#### render_and_save(data, output_path)
渲染并保存

```python
ep.render_and_save(data, 'output/report.xlsx')
```

## 模板语法

### Word 模板 (docxtpl/Jinja2)

```docx
公司名称: {{ company_name }}
日期: {{ date }}

{% for item in items %}
- {{ item.name }}: {{ item.price }} 元
{% endfor %}

{% if show_total %}
总计: {{ total }} 元
{% endif %}
```

### Excel 模板 (xlsx-template 风格)

| A | B |
|---|---|
| 报告日期 | ${report_date} |
| 公司名称 | ${company_name} |
| 销售总额 | ${total_sales} |

表格数据：
| 产品 | 数量 | 金额 |
| ${table:products.name} | ${table:products.qty} | ${table:products.amount} |

## 最佳实践

### 1. 模板设计
- 在 Word/Excel 中先设计好模板
- 使用清晰的占位符命名
- 保持模板简洁，避免过多复杂格式

### 2. 数据处理
- 使用 pandas 进行复杂数据处理
- 批量操作时考虑内存使用
- 验证输入数据的有效性

### 3. 性能优化
- 大批量文档生成使用批处理
- 复用处理器实例
- 及时保存和释放资源

### 4. 错误处理
- 始终使用 try-except 包装操作
- 验证文件路径和权限
- 提供有意义的错误信息

## 示例代码

### 批量生成合同

```python
import json
from office_pro import WordProcessor

# 加载数据
with open('contracts.json') as f:
    contracts = json.load(f)

# 批量生成
wp = WordProcessor()

for contract in contracts:
    wp.load_template('contract.docx')
    wp.render_template(contract)
    
    filename = f"contracts/contract_{contract['contract_no']}.docx"
    wp.save(filename)
    print(f"生成: {filename}")
```

### 生成月度销售报表

```python
import pandas as pd
from office_pro import ExcelProcessor
from datetime import datetime

# 读取销售数据
df = pd.read_csv('sales_data.csv')

# 汇总数据
summary = {
    'report_month': datetime.now().strftime('%Y年%m月'),
    'total_sales': df['amount'].sum(),
    'total_orders': len(df),
    'avg_order_value': df['amount'].mean(),
    'top_products': df.groupby('product')['amount'].sum().nlargest(5).to_dict()
}

# 生成报表
ep = ExcelProcessor()
ep.load_template('monthly-sales.xlsx')
ep.render_template(summary)
ep.save(f"reports/sales_{datetime.now().strftime('%Y%m')}.xlsx")
```

## 相关链接

- [python-docx 文档](https://python-docx.readthedocs.io/)
- [openpyxl 文档](https://openpyxl.readthedocs.io/)
- [docxtpl 文档](https://docxtpl.readthedocs.io/)
- [Jinja2 文档](https://jinja.palletsprojects.com/)
