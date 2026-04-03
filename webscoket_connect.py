import asyncio
import json
import websockets
import ssl


async def websocket_client(message):
    uri = "wss://szhlinvma75.apac.bosch.com:59108"  # 替换为你的 WebSocket 服务器地址
    # 创建 SSL 连接上下文，忽略证书验证
    ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    ssl_context.check_hostname = False
    ssl_context.verify_mode = ssl.CERT_NONE

    async with websockets.connect(uri, ssl=ssl_context) as websocket:
        await websocket.send(json.dumps(message, ensure_ascii=False))
        # print(f"Sent: {message}")
        #
        # response = await websocket.recv()
        # print(f"Received: {response}")




if __name__ == "__main__":
    # 运行 WebSocket 客户端
    asyncio.run(websocket_client(json.dumps({'connectionID': '111', 'category': 'bbb', 'from': 'b', 'to': 'c', 'message': 'd',
                                  'remarks': 'e'})))
