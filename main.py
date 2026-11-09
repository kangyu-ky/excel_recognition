# -*- coding: utf-8 -*-
"""
大模型识别题目 - 基础版本
用法示例: python main.py 1 2 3 5
流程: 读取 tests.xlsx 中指定 id 的题目 -> 将题目文件(PDF/图片)转为图片 ->
      视觉模型转录相关内容 -> 合计校验(Python 验算, 不符则复核) ->
      文本模型作答(thinking 题由 Python 计算) -> 将答案写入 submit.xlsx
"""
import argparse
import ast
import base64
import io
import json
import operator
import os
import re
import sys
from collections import Counter

import fitz  # PyMuPDF
import openpyxl
from openai import OpenAI
from PIL import Image

# ================== 配置常量（直接修改即可，便于更换和配置） ==================
API_KEY = os.environ.get("QWEN_API_KEY", "")            # 千问 API Key，从环境变量 QWEN_API_KEY 读取
BASE_URL = "https://maas.qianwenaiapi.com/compatible-mode/v1"  # 千问 API 网关（OpenAI 兼容）
MODEL = "qwen3.5-omni-plus"                            # 千问多模态模型（转录文件内容）
TEXT_MODEL = "qwen3.8-flash"                           # 千问文本模型（基于转录内容作答）

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TESTS_XLSX = os.path.join(BASE_DIR, "tests.xlsx")          # 题目文件
FILES_DIR = os.path.join(BASE_DIR, "files")                # 题目所需文件目录
SUBMIT_TEMPLATE = os.path.join(BASE_DIR, "submit-template.xlsx")  # 答案模板
SUBMIT_XLSX = os.path.join(BASE_DIR, "submit.xlsx")        # 答案输出文件

PDF_DPI = 300  # PDF 转图片的清晰度
MAX_IMAGE_SIDE = 2400  # 图片最长边上限, 超过时模型会返回空结果
# ============================================================================


def load_questions(ids):
    """从 tests.xlsx 读取指定 id 的题目"""
    wb = openpyxl.load_workbook(TESTS_XLSX)
    ws = wb.active
    header = [c.value for c in ws[1]]
    rows = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows[int(row[0])] = dict(zip(header, row))
    questions = []
    for i in ids:
        if i not in rows:
            raise KeyError(f"题目 id={i} 不存在于 tests.xlsx")
        questions.append(rows[i])
    return questions


def page_rotation(page):
    """根据文本行方向判断该页扫描内容是否被旋转, 返回需要顺时针旋转的角度"""
    dirs = Counter()
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            dirs[(round(line["dir"][0]), round(line["dir"][1]))] += 1
    if not dirs:
        return 0
    (dx, dy), _ = dirs.most_common(1)[0]
    if (dx, dy) == (0, -1):
        return 90
    if (dx, dy) == (0, 1):
        return 270
    if (dx, dy) == (-1, 0):
        return 180
    return 0


def file_to_images(file_name):
    """将题目文件转为图片列表，返回 [(mime, bytes), ...]"""
    path = os.path.join(FILES_DIR, file_name)
    ext = os.path.splitext(file_name)[1].lower()
    if ext == ".pdf":
        images = []
        doc = fitz.open(path)
        for page in doc:
            pix = page.get_pixmap(dpi=PDF_DPI)
            data = pix.tobytes("png")
            # 扫描件常出现整页内容被旋转 90 度, 先摆正再交给模型, 否则表头与数据行会错列
            deg = page_rotation(page)
            if deg:
                img = Image.open(io.BytesIO(data)).rotate(-deg, expand=True)
                # 旋转后长边会超过 3000 像素, 模型侧会直接返回空结果, 故压缩到 2400 以内
                if max(img.size) > MAX_IMAGE_SIDE:
                    ratio = MAX_IMAGE_SIDE / max(img.size)
                    img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                data = buf.getvalue()
            images.append(("image/png", data))
        doc.close()
        return images
    # 图片文件直接读取
    mime = "image/jpeg" if ext in (".jpg", ".jpeg") else f"image/{ext.lstrip('.')}"
    with open(path, "rb") as f:
        return [(mime, f.read())]


FORMAT_DESC = {
    "json": "以 JSON 对象格式输出",
    "json_array": "以 JSON 数组格式输出",
    "number": "只输出一个数字",
    "string": "只输出一段文本",
}

TABLE_NAME_RE = re.compile(r"表\s*\d+\s*(?:[（(]\s*续\s*[)）])?")


def norm_text(s):
    """去掉空白并统一括号写法, 便于匹配表名/表头"""
    return re.sub(r"[\s\u3000]+", "", s).replace("(", "（").replace(")", "）")


def table_row_count(transcript, question):
    """题目指定了某张表时, 统计该表的数据行数, 用于约束取列/取值的个数"""
    names = TABLE_NAME_RE.findall(question)
    if not names:
        return None
    name = norm_text(names[0])
    lines = transcript.splitlines()
    start = None
    for i, line in enumerate(lines):
        if name in norm_text(line):
            start = i
            break
    if start is None:
        return None
    rows, seen_header = 0, False
    for line in lines[start:]:
        s = line.strip()
        if not s.startswith("|"):
            if seen_header:
                break
            continue
        cells = [c.strip() for c in s.strip("|").split("|")]
        if not seen_header:
            seen_header = True
            continue
        if cells and all(set(c) <= set("-: ") for c in cells):
            continue  # 分隔行
        if cells and cells[0]:
            rows += 1
    return rows if seen_header else None


def is_separator_row(row):
    """判断 markdown 表格的分隔行"""
    return all(set(c) <= set("-: ") for c in row) if row else False


def parse_tables(transcript):
    """抽取转录中的所有 markdown 表格, 返回 [[单元格, ...], ...]"""
    tables, block = [], []
    for line in transcript.splitlines():
        s = line.strip()
        if s.startswith("|"):
            block.append([c.strip() for c in s.strip("|").split("|")])
        elif block:
            if len(block) >= 2:
                tables.append(block)
            block = []
    if len(block) >= 2:
        tables.append(block)
    return tables


def lookup_cell(transcript, question):
    """按题干中的行标签与列标签在转录表格中唯一定位单元格, 定位不到或存在多义时返回 None"""
    quoted = re.findall(r"[“\"'‘「]([^”\"'’」]{2,60})[”\"'’」]", question)
    if not quoted:
        return None
    label = quoted[0]
    m = re.search(r"\d{4}\s*[-~至]\s*\d{2,4}", label)
    period = m.group(0) if m else None
    row_key = (label[:m.start()] + label[m.end():]).strip() if m else label
    m2 = re.search(r"([^，。？、]+?)\s*列", question)
    col_key = m2.group(1).strip() if m2 else period
    if not col_key:
        return None
    row_key = norm_text(row_key)
    col_key = norm_text(col_key).replace("-", "").replace("－", "")
    for rows in parse_tables(transcript):
        header = rows[0]
        col_idx = next((i for i, h in enumerate(header)
                        if col_key in norm_text(h).replace("-", "").replace("－", "")), None)
        if col_idx is None:
            continue
        hits = [r[col_idx] for r in rows[1:] if row_key in norm_text(r[0]) and col_idx < len(r)]
        if len(hits) == 1:
            return hits[0]
    return None


def lookup_column(transcript, question):
    """题干问'某列的全部单元格内容'时, 直接从转录表格按列名取出整列, 不交给模型取列"""
    if "列" not in question or not re.search(r"内容|所有值|全部值", question):
        return None
    m = re.search(r"(\S{1,12}?)\s*列\s*(?:的)?\s*(?:单元格内容|所有值|全部值)", question)
    if not m:
        return None
    col = norm_text(m.group(1))
    table = None
    names = TABLE_NAME_RE.findall(question)
    if names:
        name = norm_text(names[0])
        lines = transcript.splitlines()
        start = next((i for i, l in enumerate(lines) if name in norm_text(l)), None)
        if start is not None:
            block = []
            for line in lines[start:]:
                if line.strip().startswith("|"):
                    block.append([c.strip() for c in line.strip().strip("|").split("|")])
                elif block:
                    break
            if len(block) >= 2:
                table = block
    for rows in ([table] if table else parse_tables(transcript)):
        idx = next((i for i, h in enumerate(rows[0]) if col in norm_text(h)), None)
        if idx is None:
            continue
        # 跳过 markdown 表格的分隔行(| --- | --- |), 否则会取到 ":---" 之类的伪单元格
        cells = [r[idx] for r in rows[1:]
                 if idx < len(r) and r[idx] and not is_separator_row(r)]
        if cells:
            return "、".join(cells)
    return None


def cell_to_number(cell):
    """单元格内容转 number 答案: 横杠表示该单元格无金额, 按 0 处理"""
    if cell is None:
        return None
    s = cell.strip()
    if s in ("-", "—", "–", "--", "——"):
        return 0
    num = to_num(s)
    return format_number(num) if num is not None else None


def build_prompt(q):
    """根据题目构造提示词"""
    parts = [
        "请根据提供的文件内容回答下面的题目。",
        f"题目类型: {q['question_type']}",
        f"题目: {q['question']}",
    ]
    if q.get("table_hint"):
        parts.append(f"表格提示: {q['table_hint']}")
    parts.append(f"答案格式要求: {FORMAT_DESC.get(q['answer_format'], q['answer_format'])}")
    parts.append(
        "作答要求:\n"
        "1. 答案必须严格来自文件原文，先定位题目中的主体，不要答成相邻行或其他条目的内容。\n"
        "2. 表格题注意行列对齐，区分每一列的表头含义（如期末余额/上期期末余额），不要串行、串列。\n"
        "3. 严格按照题目给出的项目顺序作答；表格中金额为空白(无任何字符)的项目填 null；"
        "文本、JSON 题中原样保留单元格里的横杠符号（如 —），绝对不要用其他行的数值填充。\n"
        "4. 数值答案去掉千分位逗号，只保留数字和小数点。\n"
        "5. 只输出答案本身，不要输出任何解释、思考过程或 markdown 代码块标记。\n"
        "6. 若题目问的是'是否含某类结构/元素'(如是否含合计区、是否含多级表头、是否含手写批注、"
        "是否含审核栏、是否含数量列)，必须依据图中实际结构回答'是'或'否'，不要返回 null。\n"
        "7. 只有题目所问的字段或数值在文件中确实不存在时，才返回 null 或空数组。\n"
        "8. 若题目指定了具体某一张表(如'表2（续）'、'表3')，只取该表的数据，"
        "禁止混入同一文件中其他表格的行或列，取列内容时不得多取或少取一行。\n"
        "9. 数值题中若目标单元格是横杠('-')、破折号('—')或双横杠('--')，按 0 输出。\n"
        "10. 若题目问'包含的X'(如包含的文化活动)而表中没有同名字段，"
        "则逐条列出表中相关条目的原文，不要概括改写；'由哪些项组成'类问题仍按第 7 条处理。\n"
        "11. 手写批注题须完整输出手写内容的每一段(含单位和全部数值)，不要只取其中一个数字。\n"
        "12. 题目限定范围时(如'第一行'、'前3列'、'前两行三列')，只输出该范围内的内容，"
        "不要输出题目未要求的行列(例如要求前3列时不得带出第4列)。\n"
        "13. 结构恢复题(structure)必须输出 JSON: {\"row_count\": 行数, \"col_count\": 列数, "
        "\"cells\": [每行一个数组, 按行列顺序填单元格文字]}; row_count 含表头行。"
        "题目形如'第一行和前N列'时, cells 第一行输出完整表头(全部列), 其后各行只输出前N列。"
    )
    return "\n".join(parts)


def build_transcribe_prompt():
    """第一步：完整转录文件内容"""
    return (
        "请完整转录所提供文件的全部内容。\n"
        "要求:\n"
        "1. 表格用 markdown 表格完整转录，保留每一列表头和所有行（包括内容为空白的行），数值原样保留。\n"
        "2. 文字内容按阅读顺序摘录原文，不要改写、不要遗漏。\n"
        "3. 单元格内多行内容用'、'连接，禁止使用 <br> 等 HTML 标记。\n"
        "4. 对于日程/卡片类版面，改用长表转录: 每张卡片一行，列为 | 时间段 | 天 | 标题 | 描述 | 标签(颜色) |; "
        "没有卡片的时段也要列出(如 午餐/休息)；保留所有时间标注(如 9:30、12:00、21:30)；"
        "每天的主题/副标题(通常在天的标题下方)也单独转录一行(时间段填'主题')。\n"
        "5. 只输出转录内容，不要输出任何解释。\n"
        "6. 只允许 markdown 管道表格，禁止输出 LaTeX(tabular/multicolumn 等)。\n"
        "7. 多级表头必须展开成一行: 把上层表头与下层表头用'-'连接写成同一列名(如 '主要成分-Ti'), "
        "保证每列只有一个表头、每行数据与列严格对齐。\n"
        "8. 若图片中的文字或表格是旋转的(整页横置或文字竖排), 先按正确方向摆正再转录, "
        "不要按旋转后的视觉位置错行错列。\n"
        "9. 图中若存在手写文字或手写批注, 必须单独转录一行, 格式 | 手写位置 | 手写内容 |; "
        "手写内容逐字识别并按原位置分段列出(同一单元格或多段的用'、'分隔), 保留单位和全部数值, "
        "看不清的字符用 [?] 标注, 不要只记录其中一段。"
    )


def ask_model(client, model, images, prompt):
    """调用千问模型，images 为空列表时走纯文本对话"""
    content = [{"type": "text", "text": prompt}]
    for mime, data in images:
        b64 = base64.b64encode(data).decode("ascii")
        content.append({"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}})
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": content}],
        temperature=0,
        extra_body={"enable_thinking": False},  # 关闭思考链, 避免长延时
    )
    if not resp.choices:
        raise RuntimeError(f"模型 {model} 返回无结果: {resp}")
    return resp.choices[0].message.content.strip()


def parse_json(text):
    """从模型输出中解析 JSON(可能包裹 markdown 代码块或多余文字)"""
    m = re.search(r"(\{.*\}|\[.*\])", text.strip(), re.DOTALL)
    return json.loads(m.group(1))


# ==================== 合计校验: Python 验算子项之和, 不符则复核 ====================
TOTAL_KEYWORDS = ("合计", "总计", "小计")


def build_relations_prompt(transcript):
    """从转录内容中抽取合计关系"""
    return (
        "从下面的表格转录内容中找出所有合计关系(行名包含 合计/总计/小计 的行)。\n"
        "每一组合计关系按每个数值列分别输出:\n"
        '[{"total_name": "合计行名", "column": "列名", "total_value": 数值, '
        '"items": [{"name": "子项行名", "value": 数值}]}]\n'
        "要求: 数值去掉千分位逗号; 空白单元格不要列入; 子项取该合计行之上的直接组成行; "
        "没有合计关系则输出 []。只输出 JSON。\n\n"
        f"转录内容:\n{transcript}"
    )


def to_num(value):
    """把模型返回的数值(可能带千分位逗号或为空)转成 float, 无法解析返回 None"""
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.replace(",", "").replace("，", "").strip()
        try:
            return float(text)
        except ValueError:
            return None
    return None


def check_sums(relations, tol=0.05):
    """Python 验算: 子项之和与合计值不符的返回差异列表"""
    mismatches = []
    for rel in relations:
        total = to_num(rel.get("total_value"))
        if total is None:  # 合计值无法解析, 无法验算
            continue
        s = sum(to_num(item.get("value")) or 0 for item in rel.get("items") or [])  # 空白子项按 0 计
        if abs(s - total) > tol:
            mismatches.append({**rel, "sum": s})
    return mismatches


def build_recheck_prompt(mismatches):
    """针对合计不一致的行, 让视觉模型重新辨认"""
    rows = sorted({n for m in mismatches for n in [m.get("total_name"), *(i.get("name") for i in m.get("items") or [])] if n})
    return (
        "请重新仔细辨认图片中以下表格行每一列的数值(此前识别可能有误)。\n"
        f"行名: {', '.join(rows)}\n"
        '输出 JSON: [{"name": "行名", "column": "列名", "value": 数值}], '
        "数值去掉千分位逗号, 空白填 null。只输出 JSON。"
    )


def apply_corrections(relations, corrections):
    """将复核值写回合计关系"""
    by_key = {(c.get("name"), c.get("column")): c.get("value") for c in corrections if c.get("value") is not None}
    for rel in relations:
        key = (rel.get("total_name"), rel.get("column"))
        if key in by_key:
            rel["total_value"] = by_key[key]
        for item in rel.get("items") or []:
            key = (item.get("name"), rel.get("column"))
            if key in by_key:
                item["value"] = by_key[key]
    return relations


def verify_totals(client, images, transcript):
    """合计校验主流程: 有合计关系则验算, 不符则复核并在转录后附加修正信息"""
    if not any(k in transcript for k in TOTAL_KEYWORDS):
        return transcript
    relations = parse_json(ask_model(client, TEXT_MODEL, [], build_relations_prompt(transcript)))
    if not isinstance(relations, list) or not relations:
        return transcript
    mismatches = check_sums(relations)
    if not mismatches:
        print("合计校验: 通过")
        return transcript
    print(f"合计校验: 发现 {len(mismatches)} 处不一致, 触发复核")
    corrections = parse_json(ask_model(client, MODEL, images, build_recheck_prompt(mismatches)))
    if not isinstance(corrections, list):
        corrections = []
    still = check_sums(apply_corrections(relations, corrections))
    lines = ["", "【合计校验与复核信息(以此为准)】"]
    for c in corrections:
        lines.append(f"{c.get('name')} 的 {c.get('column')} 复核值为 {c.get('value')}")
    if still:
        lines.append("注意: 复核后仍存在合计不一致, 请谨慎取值。")
    return transcript + "\n".join(lines) + "\n"


# ==================== thinking 题: 模型出计算规格, Python 真正计算 ====================
_SAFE_OPS = {ast.Add: operator.add, ast.Sub: operator.sub,
             ast.Mult: operator.mul, ast.Div: operator.truediv}


def safe_eval(expr):
    """只允许数字与 +-*/() 的安全算式求值; 表达式中的时间(如 21:30)先换算为小数小时"""
    expr = re.sub(r"\d{1,2}:\d{2}", lambda m: str(int(m.group(0).split(":")[0]) + int(m.group(0).split(":")[1]) / 60), expr)

    def _eval(node):
        if isinstance(node, ast.Expression):
            return _eval(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
            return node.value
        if isinstance(node, ast.BinOp) and type(node.op) in _SAFE_OPS:
            return _SAFE_OPS[type(node.op)](_eval(node.left), _eval(node.right))
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            return -_eval(node.operand) if isinstance(node.op, ast.USub) else _eval(node.operand)
        raise ValueError(f"非法表达式: {expr}")
    return _eval(ast.parse(expr, mode="eval"))


def build_thinking_prompt(q, transcript):
    """thinking 题: 模型只输出计算规格 JSON, 不直接给最终计算结果"""
    return (
        "请根据转录内容分析下面的题目。不要直接给出最终计算结果, 而是输出一个 JSON 对象:\n"
        '{"operation": "运算类型", "values": [...], "expression": "算式"}\n'
        "operation 取值:\n"
        '- "count": 计数, values 为符合条件的条目列表(逐条列出)\n'
        '- "sum": 求和, values 为数值列表(空白按 0 计)\n'
        '- "subtract": 减法, values=[被减数, 减数]\n'
        '- "divide": 除法, values=[被除数, 除数]\n'
        '- "average"/"max"/"min": values 为数值列表\n'
        '- "expression": 复杂四则运算, expression 为只含数字和 +-*/() 的算式, values 填 []\n'
        '- "none": 无需计算(纯提取/列举/文本), values 直接放答案本身\n'
        "operation 选择必须匹配答案格式:\n"
        '- 答案格式为 json_array 时结果必须是数组(问"哪些"用 none 并把条目数组放入 values, 禁止用 count)\n'
        '- 答案格式为 number 时结果必须是数值\n'
        '- 答案格式为 string 时 values 直接放文本\n'
        "时长/持续时间类题目必须计算: 用 subtract 或 expression 由起止时间算出差值"
        "(如 12:00 到 14:00 → expression \"(14-12)*60\" → 120 分钟), 禁止直接抄时间点。\n"
        "财务领域知识: 题目问某项目'由哪些项组成/包含哪些部分'时, 只有表中明确列出的子项(如'其中：'缩进项或合计成该行的明细行)才可作为答案; "
        "若表中没有体现组成项, 输出空数组, 不要用表外知识或相邻行推断。\n"
        "要求: 数值使用原始数字(去掉千分位逗号); 只输出 JSON。\n\n"
        f"题目: {q['question']}\n"
        f"答案格式要求: {FORMAT_DESC.get(q['answer_format'], q['answer_format'])}\n\n"
        f"转录内容:\n{transcript}"
    )


def format_number(x):
    """数值格式化: 整数去小数点, 小数保留 6 位去尾零"""
    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        return f"{x:.6f}".rstrip("0")
    return str(x)


def compute_thinking_answer(spec_raw, answer_format):
    """按计算规格用 Python 计算, 并按题目要求的格式输出"""
    spec = parse_json(spec_raw)
    if not isinstance(spec, dict):
        spec = {}
    op = spec.get("operation", "none")
    values = spec.get("values")
    if values is None:
        values = []
    numeric = [to_num(v) for v in values] if isinstance(values, list) else []

    if op == "count":
        result = len(values)
    elif op == "sum":
        result = sum(n or 0 for n in numeric)  # 空白按 0 计
    elif op == "subtract":
        result = (numeric[0] or 0) - (numeric[1] or 0)
    elif op == "divide":
        divisor = numeric[1] or 0
        result = (numeric[0] or 0) / divisor if divisor else None
    elif op == "average":
        result = sum(n or 0 for n in numeric) / len(numeric) if numeric else None
    elif op in ("max", "min"):
        valid = [n for n in numeric if n is not None]
        result = (max(valid) if op == "max" else min(valid)) if valid else None
    elif op == "expression":
        result = safe_eval(spec.get("expression") or "")
    else:  # none: 无需计算, values 即答案
        result = values
    if result is None:
        return "null"
    if isinstance(result, (int, float)):
        return format_number(result)
    # 单元素数组按 number/string 格式输出时拆包
    if isinstance(result, list) and len(result) == 1 and answer_format in ("number", "string"):
        result = result[0]
        if isinstance(result, (int, float)):
            return format_number(result)
        return str(result)
    if answer_format in ("json", "json_array"):
        return json.dumps(result, ensure_ascii=False)
    if isinstance(result, list):  # string 格式: 数组用顿号连接
        return "、".join(str(v) for v in result)
    return str(result)


# 按文件缓存转录结果: 同一文件的多道题共用一份完整转录, 保证答案一致并减少调用
TRANSCRIPT_CACHE = {}


def get_transcript(client, file_name):
    """完整转录文件(带缓存)并做合计校验"""
    if file_name not in TRANSCRIPT_CACHE:
        images = file_to_images(file_name)
        # 逐页转录: 多张大图一起发送时模型会返回空结果
        transcript = "\n\n".join(ask_model(client, MODEL, [img], build_transcribe_prompt()) for img in images)
        TRANSCRIPT_CACHE[file_name] = verify_totals(client, images, transcript)
    return TRANSCRIPT_CACHE[file_name]


def answer_question(client, q):
    """视觉模型完整转录(按文件缓存) -> 合计校验 -> 文本模型作答(thinking 题由 Python 计算)"""
    transcript = get_transcript(client, q["file_name"])
    # 整列取值的题目直接按列名从转录表格取值, 避免模型多取或少取行
    if q["answer_format"] == "string":
        column = lookup_column(transcript, q["question"])
        if column:
            return clean_answer(column, q["answer_format"])
    # 题干带引号的行标签且能唯一定位到列标签时, 直接取单元格值, 避免模型取错列
    if q["answer_format"] == "number":
        answer = cell_to_number(lookup_cell(transcript, q["question"]))
        if answer is not None:
            return answer
    if q["question_type"] == "thinking":
        spec_raw = ask_model(client, TEXT_MODEL, [], build_thinking_prompt(q, transcript))
        return compute_thinking_answer(spec_raw, q["answer_format"])
    answer_prompt = build_prompt(q) + f"\n\n以下是从文件中转录的相关内容:\n{transcript}"
    # 题目点名了某张表且要求逐行取列时, 用该表的实际行数约束答案个数, 防止混入同文件其他表格的行
    rows = table_row_count(transcript, q["question"])
    if rows and "列" in q["question"]:
        answer_prompt += f"\n\n注意: 题目所指表格的数据行数为 {rows}, 答案必须恰好包含 {rows} 个值, 不多不少。"
    answer = clean_answer(ask_model(client, TEXT_MODEL, [], answer_prompt), q["answer_format"])
    if q["question_type"] == "structure" and q["answer_format"] == "json":
        answer = trim_structure(answer, q["question"])
    return answer


CN_NUM = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
          "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}


def range_limit(question, unit):
    """从题干解析'前N列/前N行/第一列/第一行'中的数量, 解析不到返回 None"""
    pattern = rf"前([0-9]+|[一二两三四五六七八九十]+)\s*{unit}"
    m = re.search(pattern, question)
    if m:
        token = m.group(1)
        return int(token) if token.isdigit() else CN_NUM.get(token)
    m = re.search(rf"第([一二两三四五六七八九十])\s*{unit}", question)
    if m:
        return CN_NUM.get(m.group(1))
    return None


def trim_structure(text, question):
    """结构题按题干范围裁剪: 第一行保留完整表头, 数据行截到前N列、前N行"""
    obj = json.loads(text)
    if not isinstance(obj, dict) or "cells" not in obj:
        return text
    cells = obj["cells"]
    if not cells:
        return text
    col_limit = range_limit(question, "列")
    row_limit = range_limit(question, "行")
    head, body = cells[0], cells[1:]
    if col_limit:
        body = [r[:col_limit] for r in body]
    if row_limit and row_limit > 1:
        body = body[:row_limit - 1]
    cells = [head] + body
    obj["cells"] = cells
    obj["row_count"] = len(cells)
    obj["col_count"] = len(head)
    return json.dumps(obj, ensure_ascii=False)


def normalize_numbers(obj):
    """把 JSON 中的纯数字字符串转为数值, 避免答案里出现带引号的数字"""
    if isinstance(obj, list):
        return [normalize_numbers(v) for v in obj]
    if isinstance(obj, dict):
        return {k: normalize_numbers(v) for k, v in obj.items()}
    if isinstance(obj, str):
        s = obj.strip().replace(",", "")
        if re.fullmatch(r"-?(0|[1-9][0-9]*)(\.[0-9]+)?", s):
            return float(s) if "." in s else int(s)
    return obj


def clean_answer(text, answer_format):
    """清理模型输出：去除可能包裹的 markdown 代码块"""
    text = text.strip()
    m = re.match(r"^```[a-zA-Z]*\s*\n?(.*?)\n?```$", text, re.DOTALL)
    if m:
        text = m.group(1).strip()
    # json 类答案压缩为单行, 并把数字字符串转成数值
    if answer_format in ("json", "json_array"):
        text = json.dumps(normalize_numbers(json.loads(text)), ensure_ascii=False)
    return text


def write_answers(answers):
    """将答案写入 submit.xlsx: 已存在则按 id 增量更新, 否则基于模板创建"""
    src = SUBMIT_XLSX if os.path.exists(SUBMIT_XLSX) else SUBMIT_TEMPLATE
    wb = openpyxl.load_workbook(src)
    ws = wb.active
    rows = {int(r[0]): r[1] for r in ws.iter_rows(min_row=2, values_only=True) if r[0] is not None}
    rows.update(dict(answers))
    # 清空数据行后按 id 排序重写
    ws.delete_rows(2, ws.max_row)
    for qid in sorted(rows):
        ws.append([qid, rows[qid]])
    wb.save(SUBMIT_XLSX)


def main():
    parser = argparse.ArgumentParser(description="大模型识别题目 - 基础版本")
    parser.add_argument("ids", nargs="+", type=int, help="题目 id，可传多个，如: python main.py 1 2 3 5")
    args = parser.parse_args()

    if not API_KEY:
        print("错误: 环境变量 QWEN_API_KEY 未设置，请先配置 API Key", file=sys.stderr)
        sys.exit(1)

    questions = load_questions(args.ids)
    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

    done = {}
    for n, q in enumerate(questions, 1):
        qid = int(q["id"])
        print(f"--- [{n}/{len(questions)}] 题目 {qid} ({q['file_name']}) ---", flush=True)
        print(f"问题: {q['question']}", flush=True)
        answer = answer_question(client, q)
        print(f"答案: {answer}", flush=True)
        done[qid] = answer
        # 每题结束即写入; 文件被 Excel 等占用时跳过, 后续题目会再次尝试
        try:
            write_answers(done.items())
        except PermissionError:
            print("警告: submit.xlsx 被其他程序占用, 本次未写入, 稍后重试", flush=True)

    write_answers(done.items())
    print(f"\n共处理 {len(questions)} 条题目, 结果已写入 {SUBMIT_XLSX}")


if __name__ == "__main__":
    main()
