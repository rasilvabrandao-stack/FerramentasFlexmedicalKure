#!/usr/bin/env python3
"""
Servidor HTTP personalizado para desenvolvimento que desabilita cache
"""
import http.server
import socketserver
import os
from urllib.parse import unquote

class NoCacheHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # Desabilitar cache para todos os arquivos
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def log_message(self, format, *args):
        # Log mais limpo
        pass

def run_server(port=8000):
    with socketserver.TCPServer(("", port), NoCacheHTTPRequestHandler) as httpd:
        print(f"Servidor rodando em http://localhost:{port}")
        print("Cache desabilitado para desenvolvimento")
        print("Pressione Ctrl+C para parar")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor parado")

if __name__ == "__main__":
    run_server()
