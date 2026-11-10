#!/usr/bin/env python3
"""
Word AI Copilot Add-in - 本地轻量静态服务器
自动生成本地自签名 SSL 证书，并以 HTTPS 协议在 https://localhost:3000 运行
满足 Office Add-in 对本地调试必须使用安全 HTTPS 协议的要求。
"""

import http.server
import ssl
import os
import sys
import subprocess
import webbrowser
import urllib.request
import urllib.error

PORT = 3000
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class CORSHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        # 允许 Office 跨域加载与通信
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'X-Requested-With, Content-Type, Authorization, X-Target-Url')
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200, "OK")
        self.end_headers()

    def do_POST(self):
        # 处理本地 API 反向代理，解决 HTTPS 页面无法直接请求 HTTP 接口的混合内容限制 (Mixed Content)
        if self.path.startswith('/api-proxy'):
            target_url = self.headers.get('X-Target-Url')
            if not target_url:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b'{"error": "Missing X-Target-Url header"}')
                return

            content_length = int(self.headers.get('Content-Length', 0))
            post_body = self.rfile.read(content_length) if content_length > 0 else None

            # 构造向上游的代理请求
            req_headers = {
                'Content-Type': self.headers.get('Content-Type', 'application/json'),
            }
            if 'Authorization' in self.headers:
                req_headers['Authorization'] = self.headers['Authorization']

            try:
                upstream_req = urllib.request.Request(
                    target_url,
                    data=post_body,
                    headers=req_headers,
                    method='POST'
                )

                with urllib.request.urlopen(upstream_req, timeout=60) as upstream_resp:
                    self.send_response(upstream_resp.status)
                    # 转发响应头
                    for k, v in upstream_resp.headers.items():
                        if k.lower() in ['content-type', 'cache-control']:
                            self.send_header(k, v)
                    self.end_headers()

                    # 流式中继传输数据
                    while True:
                        chunk = upstream_resp.read(1024)
                        if not chunk:
                            break
                        self.wfile.write(chunk)
                        self.wfile.flush()

            except urllib.error.HTTPError as e:
                self.send_response(e.code)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                err_data = e.read()
                self.wfile.write(err_data if err_data else f'{{"error": "{e.reason}"}}'.encode('utf-8'))
            except Exception as e:
                self.send_response(502)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(f'{{"error": "代理连接目标接口失败: {str(e)}"}}'.encode('utf-8'))
            return

        super().do_POST()

def generate_self_signed_cert(cert_dir):
    os.makedirs(cert_dir, exist_ok=True)
    cert_file = os.path.join(cert_dir, "localhost.crt")
    key_file = os.path.join(cert_dir, "localhost.key")

    if os.path.exists(cert_file) and os.path.exists(key_file):
        return cert_file, key_file

    print("[*] 正在生成本地测试 SSL 自签名证书...")
    try:
        cmd = [
            "openssl", "req", "-x509", "-newkey", "rsa:2048",
            "-keyout", key_file,
            "-out", cert_file,
            "-days", "365",
            "-nodes",
            "-subj", "/CN=localhost"
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        print("[✓] 自签名证书生成成功！")
        return cert_file, key_file
    except Exception as e:
        print(f"[!] 未检测到 openssl 工具或生成证书失败: {e}")
        return None, None

def main():
    os.chdir(DIRECTORY)
    cert_dir = os.path.join(DIRECTORY, "certs")
    cert_file, key_file = generate_self_signed_cert(cert_dir)

    server_address = ('localhost', PORT)
    httpd = http.server.HTTPServer(server_address, CORSHTTPRequestHandler)

    protocol = "http"
    if cert_file and key_file and os.path.exists(cert_file) and os.path.exists(key_file):
        try:
            context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
            context.load_cert_chain(certfile=cert_file, keyfile=key_file)
            httpd.socket = context.wrap_socket(httpd.socket, server_side=True)
            protocol = "https"
            print("[✓] SSL 加密已启用 (HTTPS)")
        except Exception as e:
            print(f"[!] 加载 SSL 证书失败，降级为普通 HTTP 模式: {e}")

    url = f"{protocol}://localhost:{PORT}/src/taskpane.html"
    print("=" * 60)
    print(f"🚀 Word AI Copilot Add-in 本地调试服务已就绪！")
    print(f"   访问地址: {url}")
    print(f"   清单文件: {os.path.join(DIRECTORY, 'manifest.xml')}")
    print("=" * 60)
    print("提示：首次在浏览器打开 https 地址时，若浏览器提示证书不受信任，")
    print("请点击【高级】->【继续前往 localhost】即可信任本地证书。")
    print("按 Ctrl + C 可停止服务。\n")

    try:
        # 尝试自动在浏览器中打开页面方便测试
        webbrowser.open(url)
    except Exception:
        pass

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[*] 服务已停止。")
        httpd.server_close()

if __name__ == "__main__":
    main()
