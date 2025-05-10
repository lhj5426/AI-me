import sys
import os
import openpyxl

def convert_excel_to_txt(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    lines = []
    for row in ws.iter_rows(values_only=True):
        # 将单元格内容转为字符串并用 TAB 拼接
        cleaned = [str(cell) if cell is not None else '' for cell in row]
        lines.append('\t'.join(cleaned))

    # 生成输出路径
    base_dir = os.path.dirname(excel_path)
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    output_path = os.path.join(base_dir, f"{base_name}_export.txt")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    print(f"已导出为: {output_path}")

def main():
    if len(sys.argv) < 2:
        print("请将Excel文件拖拽到脚本上运行。")
        return

    for excel_path in sys.argv[1:]:
        convert_excel_to_txt(excel_path)

if __name__ == "__main__":
    main()
