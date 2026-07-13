# -*- coding: utf-8 -*-
"""
95598会话抽检工具 - Flask/WSGI 版本（PythonAnywhere 部署用）
仅提供 Word 报告生成 API（AI 代理由 Cloudflare Worker 处理）
"""
from flask import Flask, request, jsonify, send_file
from io import BytesIO
import json
import os
import sys

app = Flask(__name__)

def resource_path(*parts):
    """返回资源文件的绝对路径"""
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)

WORKSPACE = resource_path()

# 确保 word_report_generator.py 在路径中
sys.path.insert(0, WORKSPACE)


@app.route('/', methods=['GET'])
def index():
    return jsonify({'status': 'ok', 'service': '95598-report-api'})


@app.route('/api/generate-report', methods=['POST', 'OPTIONS'])
def generate_report():
    """处理 Word 报告生成请求"""
    if request.method == 'OPTIONS':
        resp = app.make_default_options_response()
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp

    try:
        req_data = request.get_json(force=True)

        report_data = req_data.get('data', {})
        funnel_image = req_data.get('funnelImage', None)
        ai_suggestions = req_data.get('aiSuggestions', None)

        # 注入环节评估字段
        for eval_key in ['appealEval', 'relationEval', 'identityEval', 'resultEval']:
            if eval_key in req_data and req_data[eval_key]:
                report_data[eval_key] = req_data[eval_key]

        # 注入流转数据
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

        # 导入报告生成器
        from word_report_generator import generate_report as gen_word_report

        output = gen_word_report(
            json.dumps(report_data, ensure_ascii=False),
            funnel_image,
            ai_suggestions
        )

        file_data = output.getvalue()

        return send_file(
            BytesIO(file_data),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name='report.docx'
        )

    except Exception as e:
        print(f'[Report API Error] {e}')
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.after_request
def after_request(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    return response


if __name__ == '__main__':
    print('=' * 50)
    print('  95598会话抽检工具 - 报告生成服务')
    print('=' * 50)
    app.run(host='0.0.0.0', port=5000, debug=False)
