"""
Big Bets PPT 解析 API + MCP 协议支持
支持两种调用方式：
1. 普通 REST API: POST /parse
2. MCP 协议: /mcp (SSE 流式传输)
"""

from flask import Flask, request, jsonify, Response
from pptx import Presentation
from io import BytesIO
import re
import json
import uuid

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
        for name in BIG_BETS_NAMES:
            if name.lower() in text.lower():
                big_bet = name
                bbo_match = re.search(r'BBO\s+(.+)', text, re.IGNORECASE)
                if bbo_match:
                    owner = bbo_match.group(1).strip()
                break
        if big_bet:
            break

    return big_bet, owner


def find_field_content(all_texts, field_keyword):
    """在文本池中找到某个字段关键词，并提取其后续内容"""
    start_index = -1
    for i, text in enumerate(all_texts):
        cleaned = text.lower().replace(":", "").replace("：", "").strip()
        keyword_cleaned = field_keyword.lower().replace(":", "").strip()
        if keyword_cleaned in cleaned:
            start_index = i
            break

    if start_index == -1:
        return ""

    keyword_line = all_texts[start_index]
    content_parts = []

    after_keyword = re.split(r'[:：]', keyword_line, maxsplit=1)
    if len(after_keyword) > 1 and after_keyword[1].strip():
        content_parts.append(after_keyword[1].strip())

    for i in range(start_index + 1, len(all_texts)):
        text = all_texts[i]
        is_next_field = False
        for kw in FIELD_KEYWORDS:
            cleaned = text.lower().replace(":", "").replace("：", "").strip()
            kw_cleaned = kw.lower().replace(":", "").strip()
            if kw_cleaned in cleaned and len(text) < len(kw) + 30:
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
        all_texts = extract_all_text_from_slide(slide)

        if not all_texts:
            continue

        big_bet, owner = extract_title_info(all_texts)

        if not big_bet:
            continue

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


# ============ REST API 端点 ============

@app.route("/parse", methods=["POST"])
def parse_endpoint():
    """普通 REST API：接收 PPT 文件，返回解析结果"""
    if "file" in request.files:
        file = request.files["file"]
        file_bytes = file.read()
    else:
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


# ============ MCP 协议端点 ============

# 存储 PPT 文件（内存中，由上传工具存入）
ppt_storage = {}


def make_sse_message(event_type, data):
    """构造 SSE 消息"""
    return f"event: {event_type}\ndata: {json.dumps(data)}\n\n"


@app.route("/mcp", methods=["GET", "POST"])
def mcp_endpoint():
    """MCP SSE 端点"""
    if request.method == "GET":
        # SSE 连接建立，返回 server info
        def generate():
            # 发送 endpoint event
            yield make_sse_message("endpoint", f"/mcp/messages")
        return Response(generate(), mimetype="text/event-stream",
                       headers={"Cache-Control": "no-cache", "Connection": "keep-alive"})

    # POST - 处理 JSON-RPC 消息
    msg = request.get_json()
    method = msg.get("method", "")
    msg_id = msg.get("id")

    if method == "initialize":
        response = {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {
                    "tools": {"listChanged": False}
                },
                "serverInfo": {
                    "name": "BigBets Parser",
                    "version": "1.0.0"
                }
            }
        }
        return jsonify(response)

    elif method == "tools/list":
        response = {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "tools": [
                    {
                        "name": "parse_bigbets_ppt",
                        "description": "解析Big Bets PPT文件，提取每页的Big Bet信息（名称、Owner、Key issues、Insights to why、Objective、Project Boundary、Project Scope、Any Hypotheses、Expected output/Success）。返回所有8个Big Bets的完整数据。",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "file_url": {
                                    "type": "string",
                                    "description": "PPT文件的URL或base64编码内容"
                                },
                                "action": {
                                    "type": "string",
                                    "description": "操作类型：parse（解析PPT）",
                                    "enum": ["parse"]
                                }
                            },
                            "required": ["action"]
                        }
                    },
                    {
                        "name": "get_parsed_data",
                        "description": "获取最近一次解析的Big Bets数据。如果已经有解析结果缓存，直接返回。",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "big_bet_name": {
                                    "type": "string",
                                    "description": "可选：指定获取某个Big Bet的数据。不填则返回全部。"
                                }
                            }
                        }
                    }
                ]
            }
        }
        return jsonify(response)

    elif method == "tools/call":
        tool_name = msg.get("params", {}).get("name", "")
        arguments = msg.get("params", {}).get("arguments", {})

        if tool_name == "parse_bigbets_ppt":
            # 如果有缓存的 PPT 数据，直接解析
            if "default" in ppt_storage:
                try:
                    results = parse_ppt(ppt_storage["default"])
                    result_text = json.dumps(results, ensure_ascii=False, indent=2)
                    response = {
                        "jsonrpc": "2.0",
                        "id": msg_id,
                        "result": {
                            "content": [
                                {
                                    "type": "text",
                                    "text": f"成功解析PPT，共提取 {len(results)} 个Big Bets:\n\n{result_text}"
                                }
                            ]
                        }
                    }
                except Exception as e:
                    response = {
                        "jsonrpc": "2.0",
                        "id": msg_id,
                        "result": {
                            "content": [{"type": "text", "text": f"解析失败: {str(e)}"}],
                            "isError": True
                        }
                    }
            else:
                # 没有缓存，使用预置的解析结果
                response = {
                    "jsonrpc": "2.0",
                    "id": msg_id,
                    "result": {
                        "content": [
                            {
                                "type": "text",
                                "text": "请先通过 /upload 端点上传PPT文件，或者直接调用 get_parsed_data 获取预缓存的数据。"
                            }
                        ]
                    }
                }
            return jsonify(response)

        elif tool_name == "get_parsed_data":
            big_bet_name = arguments.get("big_bet_name", "")
            if "default" in ppt_storage:
                results = parse_ppt(ppt_storage["default"])
                if big_bet_name:
                    results = [r for r in results if r["Big Bets"].lower() == big_bet_name.lower()]
                result_text = json.dumps(results, ensure_ascii=False, indent=2)
            else:
                result_text = "暂无解析数据，请先上传PPT文件到 /upload 端点。"

            response = {
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {
                    "content": [{"type": "text", "text": result_text}]
                }
            }
            return jsonify(response)

        else:
            response = {
                "jsonrpc": "2.0",
                "id": msg_id,
                "error": {"code": -32601, "message": f"Unknown tool: {tool_name}"}
            }
            return jsonify(response)

    elif method == "notifications/initialized":
        return "", 204

    else:
        response = {
            "jsonrpc": "2.0",
            "id": msg_id,
            "error": {"code": -32601, "message": f"Method not found: {method}"}
        }
        return jsonify(response)


@app.route("/mcp/messages", methods=["POST"])
def mcp_messages():
    """MCP 消息处理端点（与 /mcp POST 相同逻辑）"""
    return mcp_endpoint()


@app.route("/upload", methods=["POST"])
def upload_ppt():
    """上传 PPT 文件到服务器缓存"""
    if "file" in request.files:
        file = request.files["file"]
        file_bytes = file.read()
    else:
        file_bytes = request.get_data()

    if not file_bytes:
        return jsonify({"error": "No file provided"}), 400

    ppt_storage["default"] = file_bytes

    # 立即解析并返回结果
    try:
        results = parse_ppt(file_bytes)
        return jsonify({
            "success": True,
            "message": f"PPT已上传并解析，共 {len(results)} 个Big Bets",
            "count": len(results),
            "data": results
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

# Force redeploy
