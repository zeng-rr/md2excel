import argparse

from md2excel.md_to_excel import MarkdownToExcelConverter


def main():
    parser = argparse.ArgumentParser(description='Markdown表格转Excel工具')
    parser.add_argument('input', help='输入Markdown文件路径')
    parser.add_argument('output', help='输出Excel文件路径')
    args = parser.parse_args()

    with open(args.input, 'r', encoding='utf-8') as f:
        md_content = f.read()

    converter = MarkdownToExcelConverter()
    converter.convert(md_content, args.output)

    print(f'成功生成Excel文件：{args.output}')

if __name__ == '__main__':
    main()