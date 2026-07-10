# -*- coding: utf-8 -*-
"""
95598会话抽检工具 - Render 部署版
提供 AI 代理 + Word 报告生成 API
"""
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.error
import urllib.parse
import json
import os
import sys
import socketserver
import re
from io import BytesIO
import requests
import time
from datetime import datetime

DEFAULT_PORT = 8899
API_URL = 'https://kb.toguide.cn:20443/v1/messages'


def resource_path(*parts):
    """Return a path that works in source and deployed environments."""
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)


WORKSPACE = resource_path()


def configured_api_key():
    """获取 API Key（Render 部署版）"""
    env_key = os.environ.get('AI_TOOL_API_KEY', '').strip()
    if env_key:
        return env_key
    key_path = os.path.join(WORKSPACE, 'api_key.txt')
    try:
        if os.path.exists(key_path):
            with open(key_path, 'r', encoding='utf-8') as f:
                value = f.read().strip()
            if value:
                return value
    except Exception as e:
        print(f'[Config] 读取API Key失败: {e}', flush=True)
    # 无输入版默认配置
    mask = [73, 22, 91, 34, 108, 7, 41, 88, 13, 77, 31, 66, 5, 99, 14, 37]
    encoded = [58, 125, 118, 75, 26, 109, 69, 15, 117, 124, 83, 55, 54, 83, 123, 105, 11, 115, 28, 116, 42, 85, 95, 59, 78, 59, 111, 33, 93, 5, 88, 20, 47, 87, 110, 102, 85, 81, 69, 51, 64, 46, 71, 117, 116, 2, 70, 119, 46, 80, 31]
    return ''.join(chr(value ^ mask[index % len(mask)]) for index, value in enumerate(encoded))


class ThreadedHTTPServer(socketserver.ThreadingMixIn, HTTPServer):
    """多线程HTTP服务器"""
    allow_reuse_address = True
    daemon_threads = True


class ProxyHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WORKSPACE, **kwargs)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path == '/api/proxy':
            self.handle_api_proxy()
        elif path == '/api/generate-report':
            self.handle_generate_report()
        else:
            self.send_error(404)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == '/' or path == '':
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'status': 'ok', 'service': '95598-backend'}).encode('utf-8'))
            return
        if path.startswith('/api/'):
            self.send_error(404)
            return
        super().do_GET()

    def log_message(self, format, *args):
        sys.stderr.write("[%s] %s - - %s\n" % (
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            self.address_string(),
            format % args
        ))

    def read_json_body(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        if not body:
            return {}
        return json.loads(body.decode('utf-8'))

    def send_json(self, status_code, payload):
        data = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_api_proxy(self):
        start_time = time.time()
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            api_key = configured_api_key() or self.headers.get('X-Api-Key', '')
            if not api_key:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': '缺少API Key'}).encode('utf-8'))
                return

            try:
                req_json = json.loads(body)
                model_name = req_json.get('model', 'unknown')
            except:
                model_name = 'unknown'

            print(f'[API Proxy] 开始调用 {model_name}，请求大小: {len(body)} bytes', flush=True)

            response = requests.post(
                API_URL,
                data=body,
                headers={
                    'Content-Type': 'application/json',
                    'x-api-key': api_key,
                    'anthropic-version': '2023-06-01'
                },
                verify=False,
                timeout=900
            )

            elapsed = time.time() - start_time
            print(f'[API Proxy] {model_name} 返回，HTTP状态码: {response.status_code}，耗时: {elapsed:.1f}秒', flush=True)

            self.send_response(response.status_code)
            self.send_header('Content-Type', response.headers.get('Content-Type', 'application/json'))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(response.content)

        except requests.exceptions.Timeout as e:
            elapsed = time.time() - start_time
            print(f'[API Proxy] 超时，耗时: {elapsed:.1f}秒，错误: {e}')
            self.send_response(504)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': f'API请求超时（{elapsed:.0f}秒）: {str(e)}'}).encode('utf-8'))
        except requests.exceptions.RequestException as e:
            elapsed = time.time() - start_time
            print(f'[API Proxy] 请求异常，耗时: {elapsed:.1f}秒，错误: {e}')
            self.send_response(502)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': f'API请求异常: {str(e)}'}).encode('utf-8'))
        except Exception as e:
            elapsed = time.time() - start_time
            print(f'[API Proxy] 未知错误，耗时: {elapsed:.1f}秒，错误: {e}')
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))

    def handle_generate_report(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            req_data = json.loads(body.decode('utf-8'))

            report_data = req_data.get('data', {})
            funnel_image = req_data.get('funnelImage', None)
            ai_suggestions = req_data.get('aiSuggestions', None)

            for eval_key in ['appealEval', 'relationEval', 'identityEval', 'resultEval']:
                if eval_key in req_data and req_data[eval_key]:
                    report_data[eval_key] = req_data[eval_key]

            if 'flowToIdentity' in req_data:
                report_data['flowToIdentity'] = req_data['flowToIdentity']
            if 'flowToBroadcast' in req_data:
                report_data['flowToBroadcast'] = req_data['flowToBroadcast']

            if funnel_image:
                print(f'[Report] 接收到漏斗图数据，长度: {len(funnel_image)} 字符', flush=True)
            else:
                print('[Report] 未接收到漏斗图数据', flush=True)

            if ai_suggestions:
                print(f'[Report] 接收到AI建议，长度: {len(ai_suggestions)} 字符', flush=True)
            else:
                print('[Report] 未接收到AI建议', flush=True)

            sys.path.insert(0, WORKSPACE)
            from word_report_generator import generate_report

            output = generate_report(
                json.dumps(report_data, ensure_ascii=False),
                funnel_image,
                ai_suggestions
            )

            file_data = output.getvalue()
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', 'attachment; filename="report.docx"')
            self.send_header('Content-Length', str(len(file_data)))
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(file_data)

        except Exception as e:
            print(f'[Report API Error] {e}')
            import traceback
            traceback.print_exc()
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()


if __name__ == '__main__':
    port = int(os.environ.get('PORT', DEFAULT_PORT))
    bind_host = '0.0.0.0'

    print('=' * 50)
    print('  95598会话抽检工具 - 后端服务')
    print('  ================================================')
    print('')
    print('  [OK] 启动成功！')
    print('  工作目录: ' + WORKSPACE)
    print('')
    print('  按 Ctrl+C 停止服务')
    print('=' * 50)

    server = ThreadedHTTPServer((bind_host, port), ProxyHandler)
    actual_port = server.server_address[1]
    print(f'  访问地址: http://0.0.0.0:{actual_port}/')
    print(json.dumps({'status': 'ready', 'port': actual_port}), flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n服务已停止')
        server.server_close()
