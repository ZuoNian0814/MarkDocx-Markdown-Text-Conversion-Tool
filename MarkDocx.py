import copy
import json
import pprint
from mistletoe import Document, ast_renderer
from OfficeEdit import DocumentDocx

class HierarchicalLogger:
    def __init__(self):
        self.Hierarchical = []

    def judge(self, name):
        return name in self.Hierarchical

    def back(self):
        self.Hierarchical = self.Hierarchical[:-1]
        return self.Hierarchical

    def into(self, name):
        self.Hierarchical.append(name)
        return self.Hierarchical

    def show(self):
        print(f'\033[96m{"/ ".join(self.Hierarchical)}\033[0m')

class PargManager:
    def __init__(self):
        self.items = {}
        self.hl = HierarchicalLogger()

        self.table_col = 0
        self.table_line = 0
        self.img_path = None
        self.img_count = 1
        self.quote_count = 1

    def init(self):
        self.items = {}
        self.hl = HierarchicalLogger()

        self.table_col = 0
        self.table_line = 0
        self.img_path = None
        self.img_count = 1
        self.quote_count = 1

    def analysis(self, item):
        line_num = self.hl.Hierarchical[0]
        if line_num not in self.items:
            self.items[line_num] = []

        self.hl.show()
        print(f"\033[90m{item}\033[92m")

        content = item["content"]
        item_type = self.hl.Hierarchical[1]
        bold = self.hl.judge("Strong")
        italic = self.hl.judge("Emphasis")
        code = self.hl.judge("InlineCode")
        img = self.hl.judge("Image")
        if img:  # 矫正主类型
            item_type = "Image"
        print(f'类型：{item_type}')
        data = {"type": item_type}
        if item_type == "Table":
            data["shape"] = (self.table_col, self.table_line)
        quote_level = self.hl.Hierarchical.count("Quote")
        if quote_level:
            if not content:
                content = "  "
            else:
                print(f"引用等级：{quote_level}")
                data["quote_level"] = quote_level
                print(f"引用计数：{self.quote_count}")
                data["quote_count"] = self.quote_count
                self.quote_count += 1
        if img:
            print(f"图像路径：{self.img_path}")
            data["img_path"] = self.img_path
            if not self.hl.judge("Table"):  # 表格中的图像不参与计数
                print(f"图像计数：{self.img_count}")
                data["img_count"] = self.img_count
                self.img_count += 1
            if self.hl.judge("Table"):
                data["type"] = "Table"
                data["shape"] = (self.table_col, self.table_line)
            else:
                content = ""    # 表格中的图像不显示图例
        print(f'内容：“{content}”')
        data["content"] = content
        if bold:
            print(f'加粗：{bold}')
            data["bold"] = bold
        if italic:
            print(f'斜体：{italic}')
            data["italic"] = italic
        if code:
            print(f'代码：{code}')
            data["code"] = code
        self.items[line_num].append(data)
        print("\033[0m", end='')

    def recursion(self, items):
        if type(items) is list:
            for i in items:
                self.recursion(i)
        elif type(items) is dict:
            item_type = items['type']
            if item_type == "ThematicBreak": # 分隔符单独处理
                self.hl.show()
                print(f"\033[90m{items}\033[92m")
                print(f'类型：{item_type}\033[m')
                self.items[items.get('line_number')] = [{"type": "ThematicBreak", "content": ""}]
            if items.get('line_number') and len(self.hl.Hierarchical) <= 1:
                self.hl.Hierarchical = [str(items['line_number'])]
            header_level = items.get('level')
            if header_level:
                item_type = f"{item_type}_{header_level}"
            self.hl.into(item_type)  # 入栈当前节点类型
            if "content" in items:
                self.analysis(items)
            if "header" in items:
                self.recursion(items['header'])
                self.table_col = len(items['header']['children'])
                self.table_line = len(items['children'])+1
            if "src" in items:
                self.img_path = items['src']
            if "children" in items:
                children = items['children']
                if children:
                    self.recursion(children)
                else:
                    self.analysis({'type': 'RawText', 'content': ''})

            self.hl.back()  # 无论是否递归子节点，都弹出当前节点类型

    def transform_config(self, items, header=None, footer=None, page_number=False, page_number_f="X", page_num_alignment="center"):
        cfg = style_configs  # 或 style_configs 直接使用
        # ---- 页眉/页脚/页码配置（优先参数内容，样式从cfg读取） ----
        header_cfg = cfg.get('header', {})
        footer_cfg = cfg.get('footer', {})
        page_num_cfg = cfg.get('page_number', {})
        configs = {
            "header": {
                "enabled": bool(header),
                "text": header or "",
                "font_name": header_cfg.get('font_name', '黑体'),
                "font_size": header_cfg.get('font_size', 11),
                "color": header_cfg.get('color', '#000000'),
                "alignment": header_cfg.get('alignment', 'right')
            },
            "footer": {
                "enabled": bool(footer),
                "text": footer or "",
                "font_name": footer_cfg.get('font_name', '黑体'),
                "font_size": footer_cfg.get('font_size', 9),
                "color": footer_cfg.get('color', '#000000'),
                "alignment": footer_cfg.get('alignment', 'right')
            },
            "page_number": {
                "enabled": page_number,
                "format": page_number_f,
                "font_name": page_num_cfg.get('font_name', 'Times New Roman'),
                "font_size": page_num_cfg.get('font_size', 9),
                "color": page_num_cfg.get('color', '#000000'),
                "alignment": page_num_alignment
            },
            "items": []
        }

        # ---- 配置读取辅助 ----
        def get_style(key, default_name='黑体', default_size=14, default_color='#000000'):
            s = cfg.get(key, {})
            return {
                'name': s.get('font_name', default_name),
                'size': s.get('font_size', default_size),
                'color': s.get('color', default_color)
            }

        for line_num, all_content in items.items():
            paragraph_config = {
                "type": "paragraph",
                "config": {
                    "line_spacing": 1.5,
                    "first_line_indent": 0,
                },
                "content": []
            }
            table_config = {
                "type": "table",
                "config": {
                    "rows": 2,
                    "cols": 3,
                    "table_align": "center",
                    "cell_vertical_align": "center",
                    "cell_contents": []
                }
            }
            table_count = 0
            list_count = 1
            item_type = None
            for item in all_content:
                i_type = item['type']
                content = item['content']
                if "Heading" in i_type:
                    item_type = "Heading"
                    level = int(i_type.split("_")[-1])
                    style = get_style('heading', '黑体', 24, '#000000')
                    paragraph_config["config"]["first_line_indent"] = 0
                    paragraph_config["content"].append({
                        "text": content,
                        "color": style['color'],
                        "name": style['name'],
                        "size": style['size'] - (level - 1) * 2,
                        "bold": "True",
                        "alignment": "center" if level <= 1 else "left"
                    })
                elif i_type == "Quote":
                    item_type = "Quote"
                    style = get_style('quote', 'Times New Roman', 14, '#000000')
                    quote_count = item.get('quote_count')
                    quote_level = item.get('quote_level')
                    prefix = ""
                    if quote_count and quote_level:
                        prefix = "\n" + "\t" * quote_level + f"[{quote_count}] "
                    paragraph_config["config"]["first_line_indent"] = 2
                    paragraph_config["content"].append({
                        "text": prefix + content,
                        "color": style['color'],
                        "name": style['name'],
                        "size": style['size'],
                        "bold": "False",
                        "alignment": "left"
                    })
                elif i_type == "Paragraph":
                    item_type = "Paragraph"
                    # 保留 code 属性对字体的特殊影响
                    is_code = item.get('code', False)
                    if is_code:
                        # 行内代码使用 code_fence 配置或硬编码默认
                        code_style = get_style('code_fence', 'Noto Sans SC', 15, '#000000')
                        name = code_style['name']
                        size = code_style['size']
                        color = code_style['color']
                    else:
                        normal_style = get_style('paragraph', '黑体', 15, '#000000')
                        name = normal_style['name']
                        size = normal_style['size']
                        color = normal_style['color']
                    paragraph_config["config"]["first_line_indent"] = 2
                    paragraph_config["content"].append({
                        "text": content,
                        "color": color,
                        "name": name,
                        "size": size,
                        "bold": "True" if item.get('bold') else "False",
                        "italic": "True" if item.get('italic') else "False",
                        "alignment": "left"
                    })
                elif i_type == "CodeFence":
                    item_type = "CodeFence"
                    style = get_style('code_fence', 'Noto Sans SC', 14, '#000000')
                    paragraph_config["config"]["first_line_indent"] = 0
                    paragraph_config["content"].append({
                        "text": f"\n{content}\n",
                        "color": style['color'],
                        "name": style['name'],
                        "size": style['size'],
                        "alignment": "left"
                    })
                elif i_type == "List":
                    item_type = "List"
                    style = get_style('list', '黑体', 14, '#000000')
                    paragraph_config["config"]["first_line_indent"] = 0
                    paragraph_config["content"].append({
                        "text": f"{list_count}. {content}\n",
                        "color": style['color'],
                        "name": style['name'],
                        "size": style['size'],
                        "alignment": "left"
                    })
                    list_count += 1
                elif i_type == "ThematicBreak":
                    configs["items"].append({
                        "type": "page_breaks",
                        "config": {}
                    })
                elif i_type == "Image":
                    img_count = item["img_count"]
                    img_path = item["img_path"]
                    configs["items"].append({
                        "type": "picture",
                        "config": {
                            "img_path": img_path,
                            "width": 10,
                            "alignment": "center",
                            "spacing_before": 0.5
                        }
                    })
                    style = get_style('image_caption', '黑体', 12, '#000000')
                    paragraph_config["content"].append({
                        "text": f"图{img_count} {content}",
                        "color": style['color'],
                        "name": style['name'],
                        "size": style['size'],
                        "alignment": "center"
                    })
                    configs["items"].append(paragraph_config)
                elif i_type == "Table":
                    item_type = "Table"
                    shape = all_content[table_count]['shape']
                    b = 0
                    while not sum(all_content[table_count + b]['shape']):
                        b += 1
                        shape = all_content[table_count + b]['shape']
                    table_config["config"]["cols"] = shape[0]
                    table_config["config"]["rows"] = shape[1]
                    cell_contents = table_config["config"]["cell_contents"]
                    if len(cell_contents) < shape[1]:
                        cell_contents.append([])
                    style = get_style('table_cell', '黑体', 14, '#000000')
                    if item.get("img_path"):
                        cell_contents[table_count // shape[1]].append({
                            "path": item.get("img_path")
                        })
                    else:
                        cell_contents[table_count // shape[1]].append({
                            "text": content,
                            "size": style['size'],
                            "color": style['color'],
                            "bold": "True" if table_count // shape[1] == 0 else "False",
                            "name": style['name']  # 补上字体名称
                        })
                    table_count += 1

            if item_type in ["Heading", "Quote", "Paragraph", "CodeFence", "List"]:
                configs["items"].append(paragraph_config)
            elif item_type == "Table":
                configs["items"].append(table_config)
                paragraph_config["content"].append({  # 表格下方添加空行
                    "text": "\n",
                })
                configs["items"].append(paragraph_config)
        return configs

    def markdocx(self, md, save_path, header=None, footer=None, page_number=False, page_number_f="X", page_num_alignment="center"):
        doc = Document(md)
        children = ast_renderer.get_ast(doc)['children']
        pprint.pprint(children)

        pm = PargManager()
        pm.recursion(children)  # 递归解析md内容
        all_items = pm.items  # 获取结果

        pprint.pprint(all_items)

        configs = pm.transform_config(all_items, header, footer, page_number, page_number_f, page_num_alignment)
        pprint.pprint(configs)

        d = DocumentDocx()
        d.write_doc(configs, save_path)

with open("style_configs.json", mode='r', encoding="utf-8") as f:
    style_configs = json.loads(f.read())

if __name__ == '__main__':
    text = """
# 标题
"""

    p = PargManager()
    p.markdocx(text, "测试文档.docx", header=None, footer=None, page_number=False, page_number_f="X", page_num_alignment="center")
