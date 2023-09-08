import argparse
import openpyxl
from docx import Document
import os
import shutil

def replace_text_with_style(run, new_text):
    # 复制原始样式
    font = run.font
    bold = run.bold
    italic = run.italic
    underline = run.underline
    color = run.font.color.rgb

    # 清除原始文本
    run.clear()

    # 添加新文本
    run.text = new_text

    # 应用原始样式
    run.font.size = font.size
    run.font.name = font.name
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.color.rgb = color

def replace_placeholders_in_word(word_template, data_row):
    document = Document(word_template)

    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            paragraph_text = run.text

            for key, value in data_row.items():
                placeholder = "{" + key + "}"
                if placeholder in paragraph_text:
                    replace_text_with_style(run, str(value))

    return document

def main():
    parser = argparse.ArgumentParser(description="批量替换Word文件中的占位符")
    parser.add_argument("--excel-file", help="包含数据的Excel文件")
    parser.add_argument("--word-template", help="Word模板文件")
    args = parser.parse_args()

    excel_file = args.excel_file
    word_template = args.word_template
    output_dir = "output"  # 输出目录

    # 清空输出目录
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    # 打开Excel文件
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb.active

    # 获取Excel列名
    column_names = [cell.value for cell in sheet[1]]

    # 遍历Excel数据行
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data_row = dict(zip(column_names, row))

        # 打开Word模板并替换占位符
        document = replace_placeholders_in_word(word_template, data_row)

        # 生成Word文件名，以姓名为例
        word_file_name = os.path.join(output_dir, f"output_{data_row['姓名']}.docx")
        document.save(word_file_name)

if __name__ == "__main__":
    main()