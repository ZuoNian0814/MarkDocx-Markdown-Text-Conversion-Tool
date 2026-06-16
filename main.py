from MarkDocx import *

if __name__ == '__main__':
    with open("测试案例.md", mode='r', encoding='utf-8') as f:
        text = f.read()

    p = PargManager()
    p.markdocx(text, "测试文档.docx", header=None, footer=None, page_number=False, page_number_f="X", page_num_alignment="center")