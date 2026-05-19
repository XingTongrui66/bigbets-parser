"""
Big Bets PPT 解析 API
接收 PPT 文件，提取每页的 Big Bet 信息，返回 JSON

部署到 Railway 后，Power Automate Flow 调用此 API
"""

from flask import Flask, request, jsonify
from pptx import Presentation
from io import BytesIO
import re

app = Flask(__name__)

# 需要提取的字段关键词（用于模糊匹配）
FIELD_KEYWORDS = [
    "Key issues",
    "Insights to why",
    "Objective",
    "Project Boundary",
    "Project Scope",
    "Any Hypotheses",
    "Expected output",
]

# 8 个 Big Bets 名称
BIG_BETS_NAMES = [
    "Device Management",
    "Future Restaurant",
    "Store Agent",
    "MACE",
    "McCruiser",
    "Next Gen DMB",
    "Smart Production",
    "Digital Drive Through",
]


def extract_all_text_from_slide(slide):
    """提取一页幻灯片中所有文本框的全部文本，合并为一个文本池"""
    all_texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text = paragraph.text.strip()
                if text:
                    all_texts.append(text)
        if shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    if text:
                        all_texts.append(text)
    return all_texts


def extract_title_info(all_texts):
    """从文本池中提取 Big Bets 名称和 Owner"""
    big_bet = ""
    owner = ""

    for text in all_texts:
        # 匹配标题格式：Big Bet名称 – BBO 人名 或 Big Bet名称 - BBO 人名
        for name in BIG_BETS_NAMES:
            if name.lower() in text.lower():
                big_bet = name
                # 提取 BBO 后面的人名
                bbo_match = re.search(r'BBO\s+(.+)', text, re.IGNORECASE)
                if bbo_match:
                    owner = bbo_match.group(1).strip()
                break
        if big_bet:
            break

    return big_bet, owner


def find_field_content(all_texts, field_keyword):
    """
    在文本池中找到某个字段关键词，并提取其后续内容
    逻辑：找到包含关键词的行，然后收集后续行直到遇到下一个字段关键词
    """
    # 找到关键词所在的行索引
    start_index = -1
    for i, text in enumerate(all_texts):
        # 模糊匹配：忽略大小写、忽略冒号和空格差异
        cleaned = text.lower().replace(":", "").replace("：", "").strip()
        keyword_cleaned = field_keyword.lower().replace(":", "").strip()
        if keyword_cleaned in cleaned:
            start_index = i
            break

    if start_index == -1:
        return ""

    # 检查关键词行本身是否包含内容（关键词后面跟着内容）
    keyword_line = all_texts[start_index]
    # 去掉字段标题部分，看剩下的是否有内容
    content_parts = []

    # 检查关键词行是否只是标题（如 "Key issues :" 或 "Key issues："）
    after_keyword = re.split(r'[:：]', keyword_line, maxsplit=1)
    if len(after_keyword) > 1 and after_keyword[1].strip():
        content_parts.append(after_keyword[1].strip())

    # 收集后续行，直到遇到下一个字段关键词
    for i in range(start_index + 1, len(all_texts)):
        text = all_texts[i]
        # 检查是否是下一个字段的开始
        is_next_field = False
        for kw in FIELD_KEYWORDS:
            cleaned = text.lower().replace(":", "").replace("：", "").strip()
            kw_cleaned = kw.lower().replace(":", "").strip()
            if kw_cleaned in cleaned and len(text) < len(kw) + 30:
                # 看起来像是字段标题行（不是内容中恰好包含关键词）
                is_next_field = True
                break

        if is_next_field:
            break

        content_parts.append(text)

    return "\n".join(content_parts)


def parse_ppt(file_bytes):
    """解析整个 PPT，返回所有 Big Bets 的数据"""
    prs = Presentation(BytesIO(file_bytes))
    results = []

    for slide in prs.slides:
        # 提取该页所有文本（全量聚合）
        all_texts = extract_all_text_from_slide(slide)

        if not all_texts:
            continue

        # 提取标题信息
        big_bet, owner = extract_title_info(all_texts)

        if not big_bet:
            # 如果没有匹配到 Big Bet 名称，跳过这页
            continue

        # 提取各字段内容
        record = {
            "Big Bets": big_bet,
            "Big Bets Owner": owner,
            "Key issues": find_field_content(all_texts, "Key issues"),
            "Insights to why": find_field_content(all_texts, "Insights to why"),
            "Objective": find_field_content(all_texts, "Objective"),
            "Project Boundary": find_field_content(all_texts, "Project Boundary"),
            "Project Scope": find_field_content(all_texts, "Project Scope"),
            "Any Hypotheses": find_field_content(all_texts, "Any Hypotheses"),
            "Expected output/Success": find_field_content(all_texts, "Expected output"),
        }

        results.append(record)

    return results


@app.route("/parse", methods=["POST"])
def parse_endpoint():
    """
    接收 PPT 文件，返回解析结果
    请求：POST，body 为 PPT 文件的二进制内容
    响应：JSON 数组，每个元素是一个 Big Bet 的 9 个字段
    """
    if "file" in request.files:
        # multipart/form-data 方式
        file = request.files["file"]
        file_bytes = file.read()
    else:
        # 直接发送二进制内容
        file_bytes = request.get_data()

    if not file_bytes:
        return jsonify({"error": "No file provided"}), 400

    try:
        results = parse_ppt(file_bytes)
        return jsonify({
            "success": True,
            "count": len(results),
            "data": results
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route("/health", methods=["GET"])
def health():
    """健康检查"""
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
