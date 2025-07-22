import pandas as pd
import os
from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT


def process_excel_by_grid(excel_path):
    """
    按网格处理Excel文件，为每个网格创建单独的Word文档
    """
    # 确保输出目录存在
    output_dir = "data_21"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path)
        print(f"成功读取Excel文件：{excel_path}")
        print(f"数据形状：{df.shape}")
        print(f"列名：{list(df.columns)}")

        # 检查是否有网格编码和网格名称列
        if '网格编码' not in df.columns or '网格名称' not in df.columns:
            print("警告：未找到'网格编码'或'网格名称'列")
            # 如果没有这些列，假设每行代表一个网格
            df['网格编码'] = df.index
            df['网格名称'] = f"网格_{df.index}"

        # 按网格处理数据
        for index, row in df.iterrows():
            grid_code = str(row.get('网格编码', f'grid_{index}'))
            grid_name = str(row.get('网格名称', f'网格_{index}'))

            # 创建安全的文件名
            safe_name = "".join(c for c in grid_name if c.isalnum() or c in (' ', '_', '-')).rstrip()
            if not safe_name:  # 如果清理后为空，使用网格编码
                safe_name = f"grid_{grid_code}"

            output_word = f"{output_dir}/{safe_name}.docx"

            # 为当前网格创建Word文档
            doc = create_grid_word_table(row, grid_name, grid_code)
            doc.save(output_word)
            print(f"已保存：{output_word}")

    except FileNotFoundError:
        print(f"错误：找不到文件 {excel_path}")
    except Exception as e:
        print(f"处理Excel文件时出错：{str(e)}")


def create_grid_word_table(row_data, grid_name, grid_code):
    """
    为单个网格创建Word表格文档
    """
    doc = Document()
    doc.add_heading(f'电网建设项目规划表 - {grid_name}', 0)

    # 添加网格信息
    info_para = doc.add_paragraph()
    info_para.add_run('网格编码：').bold = True
    info_para.add_run(f'{grid_code}')
    info_para.add_run('\n网格名称：').bold = True
    info_para.add_run(f'{grid_name}')

    doc.add_paragraph()  # 空行

    # 定义项目结构
    project_structure = {
        "中压新建项目": [
            ("变电站配套送出", "变电站配套送出"),
            ("中压线路新建", "中压线路新建"),
            ("中压业扩配套", "中压业扩配套"),
            ("配变新增布点", "配变新增布点")
        ],
        "中压改造项目": [
            ("中压电缆线路及设备改造", "中压电缆线路及设备改造"),
            ("中压架空线路及设备改造", "中压架空线路及设备改造"),
            ("中压防灾抗灾专项改造", "中压防灾抗灾专项改造"),
            ("配变增容改造", "配变增容改造")
        ],
        "低压新建项目": [
            ("低压线路新建", "低压线路新建"),
            ("低压业扩配套", "低压业扩配套")
        ],
        "低压改造项目": [
            ("低压线路改造", "低压线路改造"),
            ("低压设备及附属设施改造", "低压设备及附属设施改造")
        ],
        "智能化项目": [
            ("站房 DTU 建设改造", "站房"),
            ("智能开关建设改造", "DTU 建设改造"),
            ("通信光缆建设改造", "智能开关建设改造"),
            ("配变融合终端改造", "通信光缆建设改造"),
            ("", "配变融合终端改造")
        ],
        "其他项目": [
            ("抢修包", "抢修包")
        ]
    }

    # 年份列
    year_columns = ["2025", "2026", "2027-2030年", "十五五合计"]

    # 计算总行数
    total_rows = 1  # 表头
    for items in project_structure.values():
        total_rows += len(items)

    # 创建表格
    table = doc.add_table(rows=total_rows, cols=5)
    table.style = 'Table Grid'

    # 设置表头
    headers = ["类型", "2025", "2026", "2027-2030", "十五五合计"]
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        cell.text = header
        # 设置表头格式
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
        if not cell.paragraphs[0].runs:
            run = cell.paragraphs[0].add_run(header)
            run.bold = True

    # 填充数据
    current_row = 1

    for category, items in project_structure.items():
        category_start_row = current_row

        for i, (item_display, item_key) in enumerate(items):
            row = table.rows[current_row]

            # 第一列：项目名称
            row.cells[0].text = item_display

            # 填充数据列
            for j, year_col in enumerate(year_columns, 1):
                cell_value = ""

                # 尝试从row_data中获取对应的值
                # 构建可能的列名
                possible_keys = [
                    f"{item_key}_{year_col}",
                    f"{category}_{item_key}_{year_col}",
                    item_key,
                    year_col
                ]

                # 查找匹配的数据
                for key in possible_keys:
                    if key in row_data and pd.notna(row_data[key]):
                        cell_value = str(row_data[key])
                        break

                row.cells[j].text = cell_value

            current_row += 1

        # 设置类别标题（合并第一列的相关行）
        if len(items) > 0:
            # 为类别的第一行添加类别标识
            first_row = table.rows[category_start_row]
            current_text = first_row.cells[0].text
            first_row.cells[0].text = f"{category}\n{current_text}" if current_text else category

            # 将类别名称设为粗体
            for paragraph in first_row.cells[0].paragraphs:
                if paragraph.runs:
                    paragraph.runs[0].bold = True
                else:
                    run = paragraph.add_run(first_row.cells[0].text)
                    run.bold = True

    return doc


def batch_process_excel_files(input_dir="data", output_dir="data_2"):
    """
    批量处理目录中的Excel文件
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    excel_files = [f for f in os.listdir(input_dir) if f.endswith(('.xlsx', '.xls'))]

    for excel_file in excel_files:
        excel_path = os.path.join(input_dir, excel_file)
        print(f"\n处理文件：{excel_path}")
        process_excel_by_grid(excel_path)


def create_single_grid_document(excel_path, grid_name_or_index=None):
    """
    为特定网格创建单个Word文档
    """
    try:
        df = pd.read_excel(excel_path)

        if grid_name_or_index is not None:
            # 根据网格名称或索引筛选数据
            if isinstance(grid_name_or_index, str):
                if '网格名称' in df.columns:
                    row_data = df[df['网格名称'] == grid_name_or_index].iloc[0]
                    grid_name = grid_name_or_index
                    grid_code = row_data.get('网格编码', 'N/A')
                else:
                    print("Excel文件中未找到'网格名称'列")
                    return None
            else:
                # 按索引选择
                row_data = df.iloc[grid_name_or_index]
                grid_name = row_data.get('网格名称', f'网格_{grid_name_or_index}')
                grid_code = row_data.get('网格编码', grid_name_or_index)
        else:
            # 使用第一行数据
            row_data = df.iloc[0]
            grid_name = row_data.get('网格名称', '网格_0')
            grid_code = row_data.get('网格编码', '0')

        # 创建文档
        doc = create_grid_word_table(row_data, grid_name, grid_code)

        # 生成安全文件名
        safe_name = "".join(c for c in grid_name if c.isalnum() or c in (' ', '_', '-')).rstrip()
        if not safe_name:
            safe_name = f"grid_{grid_code}"

        # 确保输出目录存在
        output_dir = "data_2"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        output_word = f"{output_dir}/{safe_name}.docx"
        doc.save(output_word)
        print(f"已保存：{output_word}")

        return output_word

    except Exception as e:
        print(f"创建文档时出错：{str(e)}")
        return None


# 使用示例
if __name__ == "__main__":
    # 方法1：处理单个Excel文件中的所有网格
    excel_path = "data/21.xlsx"
    process_excel_by_grid(excel_path)

    # 方法2：批量处理目录中的所有Excel文件
    # batch_process_excel_files("data", "data_2")

    # 方法3：为特定网格创建文档
    # create_single_grid_document("data/21.xlsx", "网格名称")  # 按名称
    # create_single_grid_document("data/21.xlsx", 0)  # 按索引