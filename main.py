from openpyxl import load_workbook
from openpyxl.styles import Font


def get_rows_with_red_font_color(xls_path, sheet_name):
    # 加载工作簿和指定的工作表
    wb = load_workbook(xls_path)
    ws = wb[sheet_name]

    # 存储包含红色字体的行号
    red_font_rows = []

    # 遍历工作表中的所有单元格
    for row in ws.iter_rows():
        for cell in row:
            # 检查单元格的字体颜色是否为红色
            if cell.font and cell.font.color and cell.font.color.rgb == 'FFFF0000':
                red_font_rows.append(ws.cell(row=row[0].row, column=1).row)  # 假设检查第一列的单元格颜色
                break  # 找到红色字体后，跳出当前行的循环

    return red_font_rows


if __name__ == "__main__":
    # 使用示例
    xls_path = R"D:\files\语文\六选二词对.xlsx"
    sheet_name = '潘晨光GRE终极词表'
    red_rows = get_rows_with_red_font_color(xls_path, sheet_name)
    print("Rows with red font color:", red_rows)