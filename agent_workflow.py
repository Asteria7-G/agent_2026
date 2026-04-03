import json
import os
import re
from typing import Any, Dict, List

from nebula3.Config import Config
from nebula3.gclient.net import ConnectionPool
from langchain_core.messages import HumanMessage
from langchain_core.tools import tool
from langchain_openai import AzureChatOpenAI

from webscoket_connect import websocket_client


NEBULA_HOST = os.getenv("NEBULA_HOST", "10.8.140.25")
NEBULA_PORT = int(os.getenv("NEBULA_PORT", "9669"))
NEBULA_USER = os.getenv("NEBULA_USER", "root")
NEBULA_PASS = os.getenv("NEBULA_PASS", "nebula")
NEBULA_SPACE = os.getenv("NEBULA_SPACE", "digitalSE_V5")


def get_nebula_schema() -> Dict[str, Any]:
    """获取 Nebula schema（tags/edges）。"""
    config = Config()
    config.max_connection_pool_size = 10
    pool = ConnectionPool()
    ok = pool.init([(NEBULA_HOST, NEBULA_PORT)], config)
    if not ok:
        return {"tags": [], "edges": [], "error": "Nebula 连接失败"}

    schema = {"tags": [], "edges": [], "error": ""}
    with pool.session_context(NEBULA_USER, NEBULA_PASS) as session:
        use_rs = session.execute(f"USE {NEBULA_SPACE}")
        if not use_rs.is_succeeded():
            schema["error"] = use_rs.error_msg()
            return schema

        tag_rs = session.execute("SHOW TAGS")
        if tag_rs.is_succeeded():
            for row in tag_rs:
                schema["tags"].append(str(row.values()[0])[1:-1])
        else:
            schema["error"] = tag_rs.error_msg()
            return schema

        edge_rs = session.execute("SHOW EDGES")
        if edge_rs.is_succeeded():
            for row in edge_rs:
                schema["edges"].append(str(row.values()[0])[1:-1])
        else:
            schema["error"] = edge_rs.error_msg()
            return schema

    return schema


def execute_gql(gql: str) -> Dict[str, Any]:
    """执行 GQL 并返回结构化 rows/error。"""
    config = Config()
    config.max_connection_pool_size = 10
    pool = ConnectionPool()
    ok = pool.init([(NEBULA_HOST, NEBULA_PORT)], config)
    if not ok:
        return {"rows": [], "error": "Nebula 连接失败"}

    result = {"rows": [], "error": ""}
    with pool.session_context(NEBULA_USER, NEBULA_PASS) as session:
        use_rs = session.execute(f"USE {NEBULA_SPACE}")
        if not use_rs.is_succeeded():
            result["error"] = use_rs.error_msg()
            return result

        rs = session.execute(gql)
        if not rs.is_succeeded():
            result["error"] = rs.error_msg()
            return result

        keys = rs.keys()
        for row in rs:
            row_values = row.values()
            row_dict = {}
            for k, v in zip(keys, row_values):
                row_dict[k] = str(v)[1:-1]
            result["rows"].append(row_dict)

    return result


@tool("get_nebula_schema_tool")
def get_nebula_schema_tool() -> Dict[str, Any]:
    """读取 Nebula schema（tags/edges）。"""
    return get_nebula_schema()


@tool("execute_gql_tool")
def execute_gql_tool(gql: str) -> Dict[str, Any]:
    """执行一条 Nebula GQL，返回 rows/error。"""
    return execute_gql(gql)


def _load_llm() -> AzureChatOpenAI:
    return AzureChatOpenAI(
        deployment_name=os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-5"),
        temperature=0,
        azure_endpoint=os.getenv(
            "AZURE_OPENAI_ENDPOINT", "https://openaichatgpt-me-cn.openai.azure.com/"
        ),
        openai_api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2025-01-01-preview"),
        openai_api_key=os.getenv("AZURE_OPENAI_API_KEY", ""),
    )


llm = _load_llm()


def _normalize_llm_content(content: Any) -> str:
    """兼容不同模型返回格式，统一提取为文本。"""
    if isinstance(content, str):
        return content

    if isinstance(content, list):
        parts: List[str] = []
        for item in content:
            if isinstance(item, str):
                parts.append(item)
            elif isinstance(item, dict):
                txt = item.get("text") or item.get("content") or ""
                if txt:
                    parts.append(str(txt))
            else:
                parts.append(str(item))
        return "\n".join(parts)

    return str(content)


def _strip_code_fence(text: str) -> str:
    """去掉 ```json ... ``` 或 ``` ... ``` 包裹。"""
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    return cleaned.strip()


def _safe_json_loads(raw: Any) -> Dict[str, Any]:
    """兼容 markdown fence，并尽量从文本中提取 JSON 对象。"""
    text = _strip_code_fence(_normalize_llm_content(raw))
    if not text:
        raise ValueError("模型返回为空，无法解析 JSON")

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{[\s\S]*\}", text)
        if match:
            return json.loads(match.group(0))
        raise ValueError(f"模型未返回合法 JSON：{text[:120]}")


def _strip_order_by_clause(gql: str) -> str:
    """Nebula 某些场景 ORDER BY 只能跟列名，失败时做一次降级重试。"""
    # 去掉 ORDER BY ... [LIMIT n]
    pattern = re.compile(
        r"(?is)\s+ORDER\s+BY\s+[\s\S]*?(?=(\s+LIMIT\b|\s*;?\s*$))"
    )
    rewritten = re.sub(pattern, " ", gql).strip()
    rewritten = re.sub(r"\s{2,}", " ", rewritten)
    return rewritten


def _relax_document_optional_match(gql: str) -> str:
    """当步骤/检查项是强制 MATCH 导致空结果时，降级为 OPTIONAL MATCH 重试。"""
    rewritten = gql
    rewritten = re.sub(
        r"(?i)\bMATCH\s*\(\s*d\s*\)\s*-\s*\[\s*:\s*document_record_test_operation_step\s*\]\s*->",
        "OPTIONAL MATCH (d)-[:document_record_test_operation_step]->",
        rewritten,
    )
    rewritten = re.sub(
        r"(?i)\bMATCH\s*\(\s*d\s*\)\s*-\s*\[\s*:\s*document_has_test_check_items\s*\]\s*->",
        "OPTIONAL MATCH (d)-[:document_has_test_check_items]->",
        rewritten,
    )
    return rewritten


def _extract_pn_code_from_gql(gql: str) -> str:
    match = re.search(r"""p\.code\s*={1,2}\s*['"]([^'"]+)['"]""", gql, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def _build_empty_result_hint(gql: str) -> str:
    """基于真实检查查询，避免 LLM 幻觉建议。"""
    pn_code = _extract_pn_code_from_gql(gql)
    if not pn_code:
        return "未查到匹配数据。可提供更完整的 PN/文档信息后重试。"

    pn_rs = execute_gql(
        f"""MATCH (p:PN) WHERE p.code == "{pn_code}" RETURN p.code AS pn_code LIMIT 1;"""
    )
    if pn_rs.get("error"):
        return f"未查到匹配数据，且诊断查询失败：{pn_rs['error']}"
    if not pn_rs.get("rows"):
        return f"未查到匹配数据：PN `{pn_code}` 在图中不存在或编码不一致。"

    doc_rs = execute_gql(
        f"""
MATCH (p:PN)-[:pn_reference_document]->(d:PRV_DOCUMENT)
WHERE p.code == "{pn_code}"
RETURN d.document_id AS document_id, d.document_name AS document_name, d.version AS version
LIMIT 5;
""".strip()
    )
    if doc_rs.get("error"):
        return f"PN `{pn_code}` 存在，但查询关联文档失败：{doc_rs['error']}"
    if not doc_rs.get("rows"):
        return f"PN `{pn_code}` 存在，但未关联 PRV_DOCUMENT。"

    return (
        f"PN `{pn_code}` 存在，且已关联文档（示例 {len(doc_rs['rows'])} 条），"
        "但当前组合条件未命中步骤/检查项。建议改为 OPTIONAL MATCH 或放宽筛选条件。"
    )


def _extract_graph_payload(rows: List[Dict[str, str]]) -> Dict[str, Any]:
    """根据查询结果尽量构造前端可视化子图结构。"""
    nodes = {}
    edges = []

    for row in rows:
        # 节点：先把每个单元都当作候选 node id 收集
        for v in row.values():
            if v not in nodes:
                nodes[v] = {"id": v, "label": v}

        # 边：尝试识别常见字段组合
        source = row.get("source_vid") or row.get("src") or row.get("from")
        target = row.get("destination_vid") or row.get("dst") or row.get("to")
        rel_type = row.get("relationship_type") or row.get("edge") or "related_to"
        if source and target:
            edges.append(
                {
                    "source": source,
                    "target": target,
                    "label": rel_type,
                }
            )

    return {"nodes": list(nodes.values()), "edges": edges}


@tool("build_subgraph_tool")
def build_subgraph_tool(rows: List[Dict[str, str]]) -> Dict[str, Any]:
    """将表格 rows 构造成前端图结构。"""
    return _extract_graph_payload(rows)


async def _send_ws_message(task_id: str, category: str, message: str, response_end: int):
    payload = {
        "connectionID": task_id,
        "category": category,
        "from": "",
        "to": "",
        "message": message,
        "remarks": json.dumps({"paragraph_start": 0, "response_end": response_end}),
    }
    await websocket_client(payload)


async def run_agent(task_id: str, user_question: str):
    try:
        state: Dict[str, Any] = {"user_question": user_question}
        state["schema"] = get_nebula_schema_tool.invoke({})

        if state["schema"].get("error"):
            await _send_ws_message(task_id, "text", f"数据库连接失败：{state['schema']['error']}", 1)
            return

        prompt_understand = f"""
你是图数据库智能助手。Schema: {state['schema']}
用户问题: {state['user_question']}

请分析：
1. 查询意图（intent）
2. 查询条件（field, operator, value）
3. 返回字段
4. 提示：如果问题超出知识范围或者数据不存在，请说明。

只返回 JSON 对象。
"""
        resp = llm.invoke([HumanMessage(content=prompt_understand)])
        try:
            state["structured_query"] = _safe_json_loads(resp.content)
        except ValueError:
            state["structured_query"] = {
                "intent": "graph_query",
                "question": state["user_question"],
                "note": _normalize_llm_content(resp.content)[:200],
            }

        prompt_gql = f"""
根据用户意图和 schema 生成 Nebula GQL 查询：
Schema: {state['schema']}
用户意图: {state['structured_query']}
只返回一条可执行的 GQL 语句，不要附加解释。
"""
        resp_gql = llm.invoke([HumanMessage(content=prompt_gql)])
        state["gql"] = _strip_code_fence(_normalize_llm_content(resp_gql.content))

        nebula_result = execute_gql_tool.invoke({"gql": state["gql"]})
        if (
            nebula_result.get("error")
            and "Only column name can be used as sort item" in nebula_result["error"]
        ):
            fallback_gql = _strip_order_by_clause(state["gql"])
            if fallback_gql and fallback_gql != state["gql"]:
                state["gql"] = fallback_gql
                nebula_result = execute_gql_tool.invoke({"gql": state["gql"]})

        if nebula_result.get("error"):
            await _send_ws_message(task_id, "text", f"查询失败：{nebula_result['error']}", 1)
            return

        # 结果为空时，尝试把文档下游关系改为 OPTIONAL MATCH，减少“任一分支无数据即全空”的情况
        if not nebula_result.get("rows"):
            relaxed_gql = _relax_document_optional_match(state["gql"])
            if relaxed_gql != state["gql"]:
                retry_result = execute_gql_tool.invoke({"gql": relaxed_gql})
                if not retry_result.get("error") and retry_result.get("rows"):
                    state["gql"] = relaxed_gql
                    nebula_result = retry_result

        state["table_data"] = {
            "gql": state["gql"],
            "columns": list(nebula_result["rows"][0].keys()) if nebula_result["rows"] else [],
            "rows": nebula_result["rows"],
        }
        state["subgraph"] = build_subgraph_tool.invoke({"rows": nebula_result["rows"]})

        if not nebula_result["rows"]:
            state["summary"] = _build_empty_result_hint(state["gql"])
        else:
            prompt_summary = (
                "根据以下数据生成简洁中文总结，不逐行复述：\n"
                f"{state['table_data']}"
            )
            resp_sum = llm.invoke([HumanMessage(content=prompt_summary)])
            state["summary"] = _normalize_llm_content(resp_sum.content)

        await _send_ws_message(
            task_id,
            "table",
            json.dumps(state["table_data"], ensure_ascii=False),
            0,
        )
        await _send_ws_message(
            task_id,
            "subgraph",
            json.dumps(state["subgraph"], ensure_ascii=False),
            0,
        )
        await _send_ws_message(task_id, "text", state["summary"], 1)
    except Exception as exc:
        await _send_ws_message(task_id, "text", f"Agent 执行异常：{str(exc)}", 1)
