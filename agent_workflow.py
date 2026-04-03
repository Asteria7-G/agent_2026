from nebula3.Config import Config
from nebula3.gclient.net import ConnectionPool

NEBULA_HOST = "10.8.140.25"
NEBULA_PORT = 9669
NEBULA_USER = "root"
NEBULA_PASS = "nebula"

def get_nebula_schema():
    config = Config()
    config.max_connection_pool_size = 10
    pool = ConnectionPool()
    ok = pool.init([(NEBULA_HOST, NEBULA_PORT)], config)
    if not ok:
        return {"error": "Nebula 连接失败"}

    schema = {"tags": [], "edges": []}
    with pool.session_context(NEBULA_USER, NEBULA_PASS) as session:
        session.execute("USE digitalSE_V5")
        # More robust approach
        schema["tags"] = [str(row.values()[0]) for row in session.execute("SHOW TAGS")]
        # Update your Edges line too!
        schema["edges"] = [str(row.values()[0]) for row in session.execute("SHOW EDGES")]
    return schema

def execute_gql(gql):
    config = Config()
    pool = ConnectionPool()
    pool.init([(NEBULA_HOST, NEBULA_PORT)], config)
    result = {"rows": [], "error": ""}
    with pool.session_context(NEBULA_USER, NEBULA_PASS) as session:
        session.execute("USE digitalSE_V5")
        rs = session.execute(gql)
        if rs.is_succeeded():
            keys = rs.keys()
            for row in rs.rows():
                row_dict = {k: v.as_string() for k, v in zip(keys, row.values)}
                result["rows"].append(row_dict)
        else:
            result["error"] = rs.error_msg()
    return result

# agent_workflow.py
import asyncio
import json
from langchain_openai import AzureChatOpenAI
from langchain_core.messages import HumanMessage
from webscoket_connect import websocket_client

llm = AzureChatOpenAI(
    deployment_name="gpt-5",
    temperature=1,
    azure_endpoint="https://openaichatgpt-me-cn.openai.azure.com/",
    openai_api_version="2025-01-01-preview",
    openai_api_key="a72b7770afac45d6ba000394ddde7151"
)

async def run_agent(task_id, user_question):
    state = {"user_question": user_question, "schema": get_nebula_schema()}

    # 1️⃣ 理解问题 + schema 提示
    prompt_understand = f"""
你是图数据库智能助手。Schema: {state['schema']}
用户问题: {state['user_question']}

请分析：
1. 查询意图（intent）
2. 查询条件（field, operator, value）
3. 返回字段
4. 提示：如果问题超出知识范围或者数据不存在，请说明。

用 JSON 返回。
"""
    resp = llm.invoke([HumanMessage(content=prompt_understand)])
    state["structured_query"] = json.loads(resp.content)

    # 2️⃣ 生成 GQL
    prompt_gql = f"""
根据用户意图和 schema 生成 Nebula GQL 查询：
Schema: {state['schema']}
用户意图: {state['structured_query']}
只返回 GQL 查询语句
"""
    resp_gql = llm.invoke([HumanMessage(content=prompt_gql)])
    state["gql"] = resp_gql.content.strip()

    # 3️⃣ Nebula 查询
    nebula_result = execute_gql(state["gql"])
    state["table_data"] = {"columns": list(nebula_result["rows"][0].keys()) if nebula_result["rows"] else [],
                           "rows": nebula_result["rows"]}
    # 简单提取子图 nodes/edges
    nodes, edges = set(), set()
    for r in nebula_result["rows"]:
        nodes.update(r.values())
        edges.update([k for k in r.keys() if k in state["schema"]["edges"]])
    state["nodes"] = list(nodes)
    state["edges"] = list(edges)

    # 4️⃣ summary
    prompt_summary = f"根据以下数据生成简洁中文总结，不逐行复述:\n{state['table_data']}"
    resp_sum = llm([HumanMessage(content=prompt_summary)])
    state["summary"] = resp_sum.content

    # 5️⃣ WebSocket 流式发送
    table_msg = {
        "connectionID": task_id,
        "category": "table",
        "from": "",
        "to": "",
        "message": json.dumps(state["table_data"], ensure_ascii=False),
        "remarks": json.dumps({"paragraph_start": 1, "response_end": 0})
    }
    subgraph_msg = {
        "connectionID": task_id,
        "category": "subgraph",
        "from": "",
        "to": "",
        "message": json.dumps({"nodes": state["nodes"], "edges": state["edges"]}, ensure_ascii=False),
        "remarks": json.dumps({"paragraph_start": 0, "response_end": 0})
    }
    summary_msg = {
        "connectionID": task_id,
        "category": "text",
        "from": "",
        "to": "",
        "message": state["summary"],
        "remarks": json.dumps({"paragraph_start": 0, "response_end": 1})
    }

    # 顺序发送
    await websocket_client(table_msg)
    await websocket_client(subgraph_msg)
    await websocket_client(summary_msg)