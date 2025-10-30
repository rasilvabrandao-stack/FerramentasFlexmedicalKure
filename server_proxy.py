#!/usr/bin/env python3
"""
Servidor HTTP personalizado para desenvolvimento que desabilita cache
e inclui proxy para Google Apps Script (resolvendo CORS)
"""
import http.server
import socketserver
import os
import json
import urllib.request
import urllib.parse
from urllib.parse import unquote

class ProxyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Desabilitar cache para todos os arquivos
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def do_OPTIONS(self):
        # Handle preflight requests
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With')
        self.end_headers()

    def do_POST(self):
        if self.path.startswith('/api/google-sheets'):
            self.handle_google_sheets_proxy()
        else:
            # Serve static files normally
            super().do_POST()

    def do_GET(self):
        if self.path.startswith('/api/google-sheets'):
            self.handle_google_sheets_proxy()
        else:
            # Serve static files normally
            super().do_GET()

    def handle_google_sheets_proxy(self):
        try:
            # URL do Google Apps Script
            google_apps_script_url = 'https://script.google.com/macros/s/AKfycbw7_F6_p_cnLGenGmPFbep7zHwdZ5UcAYC1OXLu8Jp7SXrdjU9Nncimkxpuvt8qRw7oBA/exec'

            # Prepare request to Google Apps Script
            if self.command == 'POST':
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length) if content_length > 0 else b''
                # Forward JSON data as-is to Google Apps Script
                req = urllib.request.Request(google_apps_script_url, data=post_data, method='POST')
                req.add_header('Content-Type', 'application/json')
            else:
                # GET request
                req = urllib.request.Request(google_apps_script_url, method='GET')

            # Add headers
            req.add_header('User-Agent', 'Python-Proxy/1.0')

            # Make request to Google Apps Script
            with urllib.request.urlopen(req) as response:
                result = response.read().decode('utf-8')

            # Send response back to client
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
            self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With')
            self.end_headers()
            self.wfile.write(result.encode('utf-8'))

        except Exception as e:
            # Handle errors
            error_response = {
                'error': str(e),
                'message': 'Erro no proxy para Google Apps Script'
            }
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(error_response).encode('utf-8'))

    def log_message(self, format, *args):
        # Log mais limpo - apenas erros
        if "error" in format.lower():
            super().log_message(format, *args)

def run_server(port=8000):
    with socketserver.TCPServer(("", port), ProxyHTTPRequestHandler) as httpd:
        print(f"Servidor rodando em http://localhost:{port}")
        print("Cache desabilitado para desenvolvimento")
        print("Proxy CORS para Google Apps Script habilitado em /api/google-sheets")
        print("Pressione Ctrl+C para parar")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor parado")

if __name__ == "__main__":
    run_server()
