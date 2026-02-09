"""
生成测试PDF文件用于演示
"""
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def create_test_pdf():
    """创建一个包含文本和表格的测试PDF"""
    filename = "d:/pythonProject/PDFtoWORD/test_sample.pdf"
    
    # 创建PDF文档
    doc = SimpleDocTemplate(filename, pagesize=letter)
    story = []
    
    # 获取样式
    styles = getSampleStyleSheet()
    
    # 添加标题
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#2196F3'),
        spaceAfter=30,
    )
    title = Paragraph("测试PDF文档", title_style)
    story.append(title)
    
    # 添加正文
    normal_text = Paragraph(
        "这是一个用于测试PDF转换功能的示例文档。本文档包含文本内容和表格数据。",
        styles['Normal']
    )
    story.append(normal_text)
    story.append(Spacer(1, 0.2*inch))
    
    # 添加表格
    data = [
        ['姓名', '年龄', '职位', '薪资'],
        ['张三', '28', '软件工程师', '15000'],
        ['李四', '32', '项目经理', '20000'],
        ['王五', '25', '设计师', '12000'],
        ['赵六', '30', '测试工程师', '13000'],
    ]
    
    table = Table(data, colWidths=[2*inch, 1.5*inch, 2*inch, 1.5*inch])
    
    # 设置表格样式
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2196F3')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    story.append(table)
    story.append(Spacer(1, 0.3*inch))
    
    # 添加更多文本
    more_text = Paragraph(
        "表格上方展示了一些员工信息。PDF转Word功能可以保留这些格式。",
        styles['Normal']
    )
    story.append(more_text)
    
    # 构建PDF
    doc.build(story)
    print(f"测试PDF文件已创建: {filename}")

if __name__ == "__main__":
    create_test_pdf()
