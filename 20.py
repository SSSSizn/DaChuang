import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
import os

# 读取 Excel 文件
df = pd.read_excel("data/20.xlsx", header=[0, 1])  # 假设第一行和第二行是多级表头

# 重新设置列名（根据实际Excel结构调整）
df.columns = ['行标签', '2025年_DTU', '2025年_FTU', '2025年_智能融合终端', '2025年_光缆', '2025年_ONU（套）',
              '2026年_DTU', '2026年_FTU', '2026年_智能融合终端', '2026年_光缆', '2026年_ONU（套）',
              '2027年 - 2030年_DTU', '2027年 - 2030年_FTU', '2027年 - 2030年_智能融合终端', '2027年 - 2030年_光缆',
              '2027年 - 2030年_ONU（套）',
              '求和项:_DTU汇总', '求和项:_FTU汇总', '求和项:_智能融合终端汇总', '求和项:_光缆汇总',
              '求和项:_ONU（套）汇总']

# 从第2行开始
df = df[0:]

# 重置索引
df = df.reset_index(drop=True)

# 创建输出目录
output_dir = "data_20"
os.makedirs(output_dir, exist_ok=True)

# 遍历每一行（每个网格）
for index, row in df.iterrows():
    name = row['行标签']
    if pd.isna(name):
        continue

    # 构建行数据结构
    data_rows = []
    indicators = ['DTU', 'FTU', '智能融合终端', '光缆', 'ONU（套）']
    indicator_map = {
        'DTU': 'DTU（台）',
        'FTU': 'FTU（台）',
        '智能融合终端': '智能融合终端（台）',
        '光缆': '光缆（km）',
        'ONU（套）': 'ONU（套）'
    }
    years = ['2025年', '2026年', '2027年 - 2030年']

    for ind in indicators:
        vals = []
        for y in years:
            col_name = f'{y}_{ind}'
            v = row.get(col_name, '/')
            vals.append('0' if pd.isna(v) else str(v))

        total_col_name = f'求和项:_{ind}汇总'
        total = row.get(total_col_name, '/')
        total = '0' if pd.isna(total) else str(total)

        data_rows.append([indicator_map[ind]] + vals + [total])

    # 创建 Word 文档
    doc = Document()
    title_text = f'表 20  {name}网格“十五五”配电网自动化设备建设规模'
    p = doc.add_heading(title_text, level=1)
    run = p.runs[0]
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(12)
    run.bold = True

    # 创建表格：5列
    headers = ['指标', '2025\n公用电网', '2026\n公用电网', '2027–2030\n公用电网', '十五五合计\n公用电网']
    table = doc.add_table(rows=1 + len(data_rows), cols=5)
    table.style = 'Table Grid'

    # 表头
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        run = cell.paragraphs[0].add_run(h)
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(10)

    # 数据行
    for r_idx, row_data in enumerate(data_rows, start=1):
        row_cells = table.row_cells(r_idx)
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = val
            run = row_cells[c_idx].paragraphs[0].runs[0]
            run.font.name = '宋体'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
            run.font.size = Pt(10)

    # 保存 Word 文档
    safe_name = "".join(c for c in str(name) if c.isalnum() or c in (' ', '_', '-')).rstrip()
    output_path = os.path.join(output_dir, f"{safe_name}.docx")
    doc.save(output_path)
    print(f"✅ 已保存：{output_path}")