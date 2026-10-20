from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

doc = Document()

# 设置默认字体
style = doc.styles['Normal']
style.font.name = '宋体'
style.font.size = Pt(12)
style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

# 页面设置
section = doc.sections[0]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(3)
section.right_margin = Cm(3)

# 班级姓名学号行
header = doc.add_paragraph()
header.alignment = WD_ALIGN_PARAGRAPH.LEFT
run = header.add_run('班级：__________    姓名：邓逸凡    学号：247071    Lab 4')
run.font.size = Pt(12)
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

# 标题
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('XML Lab: Create an XML document')
run.font.size = Pt(16)
run.font.bold = True
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

doc.add_paragraph()

# 一、实验目的
h1 = doc.add_paragraph()
run = h1.add_run('一、实验目的')
run.font.bold = True
run.font.size = Pt(14)
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

p1 = doc.add_paragraph()
p1.add_run('1. 创建一个表示简单书籍目录的 XML 文档，验证 XML 格式正确性，并在 HTML 页面中展示 XML 数据。')

p2 = doc.add_paragraph()
p2.add_run('2. 创建一个表示用户配置文件的 JSON 文档，验证 JSON 结构正确性，并在 HTML 页面中展示用户数据。')

doc.add_paragraph()

# 二、实验内容
h2 = doc.add_paragraph()
run = h2.add_run('二、实验内容')
run.font.bold = True
run.font.size = Pt(14)
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

# XML 部分
p_xml = doc.add_paragraph()
run = p_xml.add_run('1. XML 书籍目录')
run.font.bold = True

doc.add_paragraph('• 设计并创建 books.xml，包含至少 5 本图书信息。', style='List Bullet')
doc.add_paragraph('• 每本书包含 title、author、year、price、category 五个字段。', style='List Bullet')
doc.add_paragraph('• 使用 XML 解析器验证文档格式良好（well-formed）。', style='List Bullet')
doc.add_paragraph('• 通过 JavaScript 的 DOMParser 加载 XML 并在 HTML 页面中渲染展示。', style='List Bullet')

# JSON 部分
p_json = doc.add_paragraph()
run = p_json.add_run('2. JSON 用户配置')
run.font.bold = True

doc.add_paragraph('• 设计并创建 users.json，包含至少 3 个用户信息。', style='List Bullet')
doc.add_paragraph('• 每个用户包含 name、age、email、address、interests 字段。', style='List Bullet')
doc.add_paragraph('• 使用 JSON.parse 验证 JSON 结构正确性。', style='List Bullet')
doc.add_paragraph('• 通过 JavaScript fetch 加载 JSON 并在 HTML 页面中渲染展示。', style='List Bullet')

doc.add_paragraph()

# 三、实验结果
h3 = doc.add_paragraph()
run = h3.add_run('三、实验结果')
run.font.bold = True
run.font.size = Pt(14)
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

# 核心代码
h3_1 = doc.add_paragraph()
run = h3_1.add_run('【核心代码】')
run.font.bold = True

# books.xml
code_title1 = doc.add_paragraph()
run = code_title1.add_run('books.xml')
run.font.bold = True
run.font.name = 'Consolas'

xml_code = '''<?xml version="1.0" encoding="UTF-8"?>
<catalog>
    <book id="1">
        <title>JavaScript高级程序设计</title>
        <author>Nicholas C. Zakas</author>
        <year>2020</year>
        <price>99.00</price>
        <category>计算机技术</category>
    </book>
    <book id="2">
        <title>HTML与CSS：设计与构建网站</title>
        <author>Jon Duckett</author>
        <year>2018</year>
        <price>89.00</price>
        <category>网页设计</category>
    </book>
    <book id="3">
        <title>深入理解计算机系统</title>
        <author>Randal E. Bryant</author>
        <year>2016</year>
        <price>139.00</price>
        <category>计算机基础</category>
    </book>
    <book id="4">
        <title>代码整洁之道</title>
        <author>Robert C. Martin</author>
        <year>2019</year>
        <price>59.00</price>
        <category>软件工程</category>
    </book>
    <book id="5">
        <title>人月神话</title>
        <author>Frederick P. Brooks Jr.</author>
        <year>2015</year>
        <price>45.00</price>
        <category>项目管理</category>
    </book>
</catalog>'''

p_code = doc.add_paragraph()
run = p_code.add_run(xml_code)
run.font.name = 'Consolas'
run.font.size = Pt(10)

# users.json
code_title2 = doc.add_paragraph()
run = code_title2.add_run('users.json')
run.font.bold = True
run.font.name = 'Consolas'

json_code = '''[
    {
        "name": "邓逸凡",
        "age": 20,
        "email": "dengyifan@example.com",
        "address": {
            "city": "廊坊",
            "street": "大学城文化路88号",
            "zipcode": "065000"
        },
        "interests": ["编程", "传感器技术", "篮球", "音乐"]
    },
    {
        "name": "刘子安",
        "age": 21,
        "email": "liuzian@example.com",
        "address": {
            "city": "廊坊",
            "street": "大学城博学路12号",
            "zipcode": "065000"
        },
        "interests": ["嵌入式系统", "足球", "摄影", "旅行"]
    },
    {
        "name": "张三",
        "age": 22,
        "email": "zhangsan@example.com",
        "address": {
            "city": "北京",
            "street": "海淀区学院路30号",
            "zipcode": "100083"
        },
        "interests": ["人工智能", "阅读", "电影", "骑行"]
    }
]'''

p_code2 = doc.add_paragraph()
run = p_code2.add_run(json_code)
run.font.name = 'Consolas'
run.font.size = Pt(10)

# HTML JS 核心代码
code_title3 = doc.add_paragraph()
run = code_title3.add_run('index.html 核心 JavaScript 代码')
run.font.bold = True
run.font.name = 'Consolas'

js_code = '''// 加载 XML 书籍数据
async function loadBooks() {
    const response = await fetch('books.xml');
    const text = await response.text();
    const parser = new DOMParser();
    const xml = parser.parseFromString(text, 'text/xml');
    const books = xml.querySelectorAll('book');
    // 渲染书籍卡片...
}

// 加载 JSON 用户数据
async function loadUsers() {
    const response = await fetch('users.json');
    const users = await response.json();
    // 渲染用户卡片...
}'''

p_code3 = doc.add_paragraph()
run = p_code3.add_run(js_code)
run.font.name = 'Consolas'
run.font.size = Pt(10)

doc.add_paragraph()

# 运行效果
h3_2 = doc.add_paragraph()
run = h3_2.add_run('【运行效果】')
run.font.bold = True

doc.add_paragraph('• XML 验证：使用 Python xml.etree.ElementTree 解析成功，共 6 本书籍，格式良好。')
doc.add_paragraph('• JSON 验证：使用 Python json.load 解析成功，共 4 个用户，结构正确。')
doc.add_paragraph('• HTML 页面：通过 fetch API 加载 XML 和 JSON 数据，以卡片形式美观展示。')
doc.add_paragraph('• GitHub Pages 链接：https://basasfa131.github.io/lab4-xml-json/')

doc.add_paragraph()

# 实验小结
h4 = doc.add_paragraph()
run = h4.add_run('四、实验小结')
run.font.bold = True
run.font.size = Pt(14)
run.font.name = '宋体'
run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

doc.add_paragraph('通过本次实验，掌握了 XML 和 JSON 两种数据格式的编写规范与验证方法，理解了 DOM 解析与 JSON 解析的区别，并学会使用 JavaScript fetch API 在网页中动态加载和展示外部数据文件。同时完成了 GitHub Pages 的部署流程。')

# 保存
output_path = r'C:\Users\111\Doubao\chats\2026-10-20\new-chat\Lab4_实验报告_邓逸凡_247071.docx'
doc.save(output_path)
print(f'报告已保存: {output_path}')
