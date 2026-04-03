import os
import uvicorn
import uuid
import requests
import tempfile
import zipfile
from multiprocessing import Process
from typing import List
from fastapi import FastAPI, BackgroundTasks, APIRouter, Request,HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from data_extraction_tool import *

app = FastAPI(title='Digital_SE_Q_Q')
app.add_middleware(
            CORSMiddleware,
            allow_origins=["*"],
            allow_credentials=True,
            allow_methods=["*"],
            allow_headers=["*"],
)

# def judge_login_status(token):
#     try:
#         url = "https://szhlinvma72.apac.bosch.com:53235/api/User/GetLoginUser"
#         user_info = requests.get(url=url, headers={'Authorization': token}, verify=False)
#         status_code = user_info.status_code
#         return status_code
#     except:
#         return 401

# class GetProductInfo(BaseModel):
#     node_name: str
##############################################################
 #                 AGENT for KG 2026            #
##############################################################

agent_2026_router = APIRouter()

from agent_workflow import run_agent  # 我们的 LangGraph agent 主函数
# 请求模型
class GetAnswerFromKG2026(BaseModel):
    user_question: str

# POST 接口：返回 task_id 并启动后台 Agent
@agent_2026_router.post("/api/digital_se/KG/answer_from_agent_2026/",tags=['agent2026'])
async def post_answer_from_kg(item: GetAnswerFromKG2026, background_tasks: BackgroundTasks):
    task_id = str(uuid.uuid4())
    background_tasks.add_task(run_agent, task_id, item.user_question)
    return {"status": 1, "result": {"connectionID": task_id}}

##############################################################
 #                 DIGITAL AGE 2025               #
##############################################################
digital_age_2025_router = APIRouter()

class UserInput(BaseModel):
    five_word: str

@digital_age_2025_router.post('/api/digital_se/digital_age_2025/get_short_answer_around_words', tags=['DigitalAge2025'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_agent_reply(item: UserInput, request: Request):
    """
    demo:

    return
    """
    try:
        result = get_agent_reply(item.five_word)
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

##############################################################
 #                 DIGITAL SE NEBULA GRAPH               #
##############################################################
digital_se_nebula_graph_router = APIRouter()

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_product_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_product_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("PRODUCT")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_series_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_series_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("SERIES")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)


@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_pn_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_pn_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("PN")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)


@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_customer_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_customer_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("CUSTOMER")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_prv_document_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_prv_document_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("PRV_DOCUMENT")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)


@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_test_operation_step_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_test_operation_step_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("TEST_OPERATION_STEP")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)


@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_test_check_items_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_test_check_items_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("TEST_CHECK_ITEMS")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_module_info/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_module_info(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_node_info("MODULE")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)


@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_product_series_relationship/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_product_series_relationship(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_relationship_info('PRODUCT', 'SERIES')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)


@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_series_pn_relationship/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_series_pn_relationship(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_relationship_info('SERIES', 'PN')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_pn_module_relationship/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_pn_module_relationship(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_relationship_info('PN', 'MODULE')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_customer_pn_relationship/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_customer_pn_relationship(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_relationship_info('CUSTOMER', 'PN')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_pn_doc_relationship/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_pn_doc_relationship(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_relationship_info('PN', 'PRV_DOCUMENT')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_doc_step_relationship/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_doc_step_relationship(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_relationship_info('PRV_DOCUMENT', 'TEST_OPERATION_STEP')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_doc_check_relationship/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_doc_check_relationship(request: Request):
    """
    demo:

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = get_relationship_info('PRV_DOCUMENT', 'TEST_CHECK_ITEMS')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)


@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_customer_info_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_customer_info_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_node_info_from_customer_req("CUSTOMER")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_product_info_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_product_info_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_node_info_from_customer_req("PRODUCT")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_application_info_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_application_info_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_node_info_from_customer_req("APPLICATION")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_hw_info_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_hw_info_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_node_info_from_customer_req("HARDWARE")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_req_info_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_req_info_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_node_info_from_customer_req("REQUIREMENT")
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_customer_product_relationship_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_customer_product_relationship_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_relationship_info_from_customer_req('CUSTOMER', 'PRODUCT')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_product_application_relationship_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_product_application_relationship_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_relationship_info_from_customer_req('PRODUCT', 'APPLICATION')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_product_hw_relationship_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_product_hw_relationship_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_relationship_info_from_customer_req('PRODUCT', 'HARDWARE')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

@digital_se_nebula_graph_router.post('/api/digital_se/KG/get_hw_req_relationship_from_customer_req/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_hw_req_relationship_from_customer_req(request: Request):
    """
    demo:

    return
    """
    try:
        result = get_relationship_info_from_customer_req('HARDWARE', 'REQUIREMENT')
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

class GetTestOperationStepAndCheckFromModule(BaseModel):
     module_vid_list: List[str]
@digital_se_nebula_graph_router.post('/api/digital_se/KG/from_module_get_related_test_operation_step_and_check/', tags=['DigitalSENebulaGraph'])
# async def post_vehicle_list(item: VehicleList, request: Request, background_tasks: BackgroundTasks)
async def post_test_operation_step_and_check_from_module(item: GetTestOperationStepAndCheckFromModule, request: Request):
    """
    demo: {module_vid_list:["module1","module10"]}

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        result = reference_test_operation_step_and_check(item.module_vid_list)
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)

class GetAnswerFromKG(BaseModel):
    user_question: str
@digital_se_nebula_graph_router.post('/api/digital_se/KG/answer_from_kg/', tags=['DigitalSENebulaGraph'])
async def post_answer_from_kg(item: GetAnswerFromKG, background_tasks: BackgroundTasks, request: Request):
    """
    demo: {user_question:"0437CX001F有哪些测试步骤和参数？"}

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        # reply = extract_query_instruction(item.user_question)
        # result = llm_chat(reply)
        task_id = str(uuid.uuid4())
        background_tasks.add_task(llm_chat, task_id, item.user_question)
        result = {'connectionID': task_id}
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)

    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)



##############################################################
#   DIGITAL SE Wuj AUTO PARSING PRV AND GENERATING SPE      #
##############################################################
digital_se_auto_p_prv_and_g_spe_router = APIRouter()

# ==================== 数据模型 ====================
class AutoPRV(BaseModel):
    pdf_path: str
    excel_path: str

# ==================== 工具函数 ====================
def download_if_url(path: str, suffix: str):
    """
    如果 path 是 URL，则下载到临时文件并返回本地路径。
    如果本地文件不存在或下载失败，抛出异常。
    """
    if path.startswith("http"):
        response = requests.get(path, verify=False)
        if response.status_code != 200:
            raise FileNotFoundError(f"下载失败: {path}, 状态码: {response.status_code}")
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp_file.write(response.content)
        tmp_file.close()
        return tmp_file.name
    else:
        if not os.path.exists(path):
            raise FileNotFoundError(f"文件不存在: {path}")
        return path

# ==================== 接口 ====================
@digital_se_auto_p_prv_and_g_spe_router.post(
    '/api/digital_se/auto_prv/generate',
    tags=['DigitalSEAutoParsingPRVandGeneratingSPE']
)
async def post_auto_prv(item: AutoPRV, background_tasks: BackgroundTasks, request: Request):
    """
    demo:
    {
        "pdf_path": "http://example.com/test.pdf",
        "excel_path": "http://example.com/test.xlsx"
    }
    return: { "status": 1, "result": { "connectionID": "..." } }
    """
    try:
        print("Received data:", item.model_dump())  # 打印收到的数据

        # 下载或校验 PDF / Excel
        try:
            pdf_path = download_if_url(item.pdf_path, suffix=".pdf")
            excel_path = download_if_url(item.excel_path, suffix=".xlsx")
        except FileNotFoundError as e:
            return JSONResponse(content={'status': 2, 'result': str(e)}, status_code=404)

        # 生成任务 ID，后台执行处理
        task_id = str(uuid.uuid4())
        background_tasks.add_task(auto_prv_improve_by_action, task_id, pdf_path, excel_path)

        # 返回任务 ID
        result = {'connectionID': task_id}
        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)

    except Exception as e:
        print("接口异常:", str(e))
        return JSONResponse(content={'status': 2, 'result': str(e)}, status_code=200)

# @digital_se_auto_p_prv_and_g_spe_router.post('/api/digital_se/auto_prv/generate', tags=['DigitalSEAutoParsingPRVandGeneratingSPE'])
# async def post_auto_prv(item: AutoPRV, background_tasks: BackgroundTasks, request: Request):
#     """
#     demo: pdf_path
#           excel_path
#
#     return
#     """
#     try:
#         # token = request.headers.get("Authorization")
#         # status_code = judge_login_status(token)
#         # if status_code == 401:
#         #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)
#
#         print("Received data:", item.model_dump())  # Fixed deprecation warning
#         pdf_path = item.pdf_path
#         excel_path = item.excel_path
#
#         # If pdf URL, download the file
#         if pdf_path.startswith("http"):
#             response = requests.get(pdf_path, verify=False)
#             if response.status_code == 200:
#                 # Save to temp file
#                 with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
#                     tmp_file.write(response.content)
#                     pdf_path = tmp_file.name
#             else:
#                 return JSONResponse(content={'status': 2, 'result': 'Failed to download PDF'}, status_code=404)
#
#         # If URL, download the file
#         if not os.path.exists(pdf_path):
#             return JSONResponse(content={'status': 2, 'result': f'PDF not found: {pdf_path}'}, status_code=404)
#
#         # If excel URL, download the file
#         if excel_path.startswith("http"):
#             response = requests.get(excel_path, verify=False)
#             if response.status_code == 200:
#                 # Save to temp file
#                 with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
#                     tmp_file.write(response.content)
#                     excel_path = tmp_file.name
#             else:
#                 return JSONResponse(content={'status': 2, 'result': 'Failed to download excel'}, status_code=404)
#
#         if not os.path.exists(excel_path):
#             return JSONResponse(content={'status': 2, 'result': f'PDF not found: {excel_path}'}, status_code=404)
#
#
#         # result = auto_prv(pdf_path)  # FIXED
#         task_id = str(uuid.uuid4())
#         background_tasks.add_task(auto_prv_improve_by_action, task_id, pdf_path,excel_path)
#         result = {'connectionID': task_id}
#
#         return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
#     except Exception as e:
#         print(str(e))
#         return JSONResponse(content={'status': 2, 'result': ''}, status_code=200)



class DownloadRequest(BaseModel):
    path_list: List[str]

@digital_se_auto_p_prv_and_g_spe_router.post(
    "/api/digital_se/auto_prv/download",
    tags=['DigitalSEAutoParsingPRVandGeneratingSPE']
)
def download_file(item: DownloadRequest):
    file_paths = [os.path.normpath(p) for p in item.path_list]

    for p in file_paths:
        print("download excel path: ", p)
        if not os.path.isfile(p):
            raise HTTPException(status_code=404, detail=f"文件未找到: {p}")

    # 临时生成 zip 文件
    tmp_zip = tempfile.NamedTemporaryFile(delete=False, suffix=".zip")
    with zipfile.ZipFile(tmp_zip.name, 'w') as zf:
        for p in file_paths:
            zf.write(p, arcname=os.path.basename(p))

    return FileResponse(tmp_zip.name, filename="testing_spe.zip")

# class DownloadRequest(BaseModel):
#     path: str
#
# @digital_se_auto_p_prv_and_g_spe_router.post("/api/digital_se/auto_prv/download",  tags=['DigitalSEAutoParsingPRVandGeneratingSPE'])
# def download_file(item: DownloadRequest):
#     full_path = os.path.normpath(item.path)
#     print("download excel path: ", full_path)
#
#     if not os.path.isfile(full_path):
#         raise HTTPException(status_code=404, detail="文件未找到")
#
#     return FileResponse(full_path, filename=os.path.basename(full_path))



##############################################################
#               TCD AUTO GENERATED TEST PROGRAM              #
##############################################################
tcd_auto_generated_test_program_router = APIRouter()

class AutoTCD(BaseModel):
    pdf_path: str
@tcd_auto_generated_test_program_router.post('/api/digital_se/auto_tcd/electric_architecture', tags=['TCDAutoGeneratedTestProgram'])
async def post_auto_tcd_electric_architecture(item: AutoTCD, background_tasks: BackgroundTasks, request: Request):
    """
    demo: pdf_path

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        print("Received data:", item.model_dump())  # Fixed deprecation warning
        pdf_path = item.pdf_path

        # If URL, download the file
        if pdf_path.startswith("http"):
            response = requests.get(pdf_path, verify=False)
            if response.status_code == 200:
                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(response.content)
                    pdf_path = tmp_file.name
            else:
                return JSONResponse(content={'status': 2, 'result': 'Failed to download PDF'}, status_code=404)

        if not os.path.exists(pdf_path):
            return JSONResponse(content={'status': 2, 'result': f'PDF not found: {pdf_path}'}, status_code=404)


        # result = auto_prv(pdf_path)  # FIXED
        task_id = str(uuid.uuid4())
        start_page = find_electric_architecture_page(pdf_path)
        background_tasks.add_task(auto_tcd_electric_architecture_extract, task_id, pdf_path, start_page)
        result = {'connectionID': task_id}

        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'ressult': ''}, status_code=200)


@tcd_auto_generated_test_program_router.post('/api/digital_se/auto_tcd/interface_table', tags=['TCDAutoGeneratedTestProgram'])
async def post_auto_tcd_interface_table(item: AutoTCD, background_tasks: BackgroundTasks, request: Request):
    """
    demo: pdf_path

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        print("Received data:", item.model_dump())  # Fixed deprecation warning
        pdf_path = item.pdf_path

        # If URL, download the file
        if pdf_path.startswith("http"):
            response = requests.get(pdf_path, verify=False)
            if response.status_code == 200:
                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(response.content)
                    pdf_path = tmp_file.name
            else:
                return JSONResponse(content={'status': 2, 'result': 'Failed to download PDF'}, status_code=404)

        if not os.path.exists(pdf_path):
            return JSONResponse(content={'status': 2, 'result': f'PDF not found: {pdf_path}'}, status_code=404)


        # result = auto_prv(pdf_path)  # FIXED
        task_id = str(uuid.uuid4())
        start_page, end_page = find_interface_table_page(pdf_path)
        background_tasks.add_task(auto_tcd_interface_table_extract, task_id, pdf_path, start_page, end_page)
        result = {'connectionID': task_id}

        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'ressult': ''}, status_code=200)


@tcd_auto_generated_test_program_router.post('/api/digital_se/auto_tcd/characteristics_table', tags=['TCDAutoGeneratedTestProgram'])
async def post_auto_tcd_interface_table(item: AutoTCD, background_tasks: BackgroundTasks, request: Request):
    """
    demo: pdf_path

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        print("Received data:", item.model_dump())  # Fixed deprecation warning
        pdf_path = item.pdf_path

        # If URL, download the file
        if pdf_path.startswith("http"):
            response = requests.get(pdf_path, verify=False)
            if response.status_code == 200:
                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(response.content)
                    pdf_path = tmp_file.name
            else:
                return JSONResponse(content={'status': 2, 'result': 'Failed to download PDF'}, status_code=404)

        if not os.path.exists(pdf_path):
            return JSONResponse(content={'status': 2, 'result': f'PDF not found: {pdf_path}'}, status_code=404)


        # result = auto_prv(pdf_path)  # FIXED
        task_id = str(uuid.uuid4())
        start_page, end_page = find_char_table_page(pdf_path)
        background_tasks.add_task(auto_tcd_char_table_extract, task_id, pdf_path, start_page, end_page)
        result = {'connectionID': task_id}

        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'ressult': ''}, status_code=200)



@tcd_auto_generated_test_program_router.post('/api/digital_se/auto_tcd/pn_list', tags=['TCDAutoGeneratedTestProgram'])
async def post_auto_tcd_pn_list(item: AutoTCD, background_tasks: BackgroundTasks, request: Request):
    """
    demo: pdf_path

    return
    """
    try:
        # token = request.headers.get("Authorization")
        # status_code = judge_login_status(token)
        # if status_code == 401:
        #     return JSONResponse(content={'status': 2, 'result': ''}, status_code=401)

        print("Received data:", item.model_dump())  # Fixed deprecation warning
        pdf_path = item.pdf_path

        # If URL, download the file
        if pdf_path.startswith("http"):
            response = requests.get(pdf_path, verify=False)
            if response.status_code == 200:
                # Save to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(response.content)
                    pdf_path = tmp_file.name
            else:
                return JSONResponse(content={'status': 2, 'result': 'Failed to download PDF'}, status_code=404)

        if not os.path.exists(pdf_path):
            return JSONResponse(content={'status': 2, 'result': f'PDF not found: {pdf_path}'}, status_code=404)

        result = auto_tcd_pn_table_extract(pdf_path)

        return JSONResponse(content={'status': 1, 'result': result}, status_code=200)
    except Exception as e:
        print(str(e))
        return JSONResponse(content={'status': 2, 'ressult': ''}, status_code=200)


app.include_router(digital_se_nebula_graph_router)
app.include_router(digital_se_auto_p_prv_and_g_spe_router)
app.include_router(tcd_auto_generated_test_program_router)
app.include_router(digital_age_2025_router)
app.include_router(agent_2026_router)
#
# uvicorn.run(app=app, host='0.0.0.0', port=8000)

if __name__ == "__main__":
    uvicorn.run(app=app, host='0.0.0.0', port=8000)