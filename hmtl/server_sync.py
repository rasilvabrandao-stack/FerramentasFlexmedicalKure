#!/usr/bin/env python3
"""
Servidor HTTP com API para sincronização de dados Excel
"""
import http.server
import socketserver
import json
import os
import subprocess
import sys
from urllib.parse import unquote, parse_qs
from datetime import datetime

# Adicionar o diretório atual ao path para importar módulos locais
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

class DataSyncHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path.startswith('/api/sync'):
            self.handle_sync()
        elif self.path.startswith('/api/generate_sync_excel'):
            self.handle_generate_sync_excel()
        elif self.path.startswith('/api/sync_google_sheets'):
            self.handle_sync_google_sheets()
        else:
            # Servir arquivos estáticos
            self.serve_static()

    def do_POST(self):
        if self.path.startswith('/api/sync_google_sheets'):
            self.handle_sync_google_sheets_post()
        else:
            self.send_error(404, "Endpoint não encontrado")

    def do_POST(self):
        if self.path.startswith('/api/sync_google_sheets'):
            self.handle_sync_google_sheets_post()
        else:
            self.send_error(404, "Endpoint não encontrado")

    def handle_sync_google_sheets_post(self):
        try:
            print("Sincronizando com Google Sheets via POST...")

            # Ler dados do POST
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            payload = json.loads(post_data.decode('utf-8'))

            # Importar módulo de sync
            from google_sheets_sync import sincronizar_tudo

            # Usar dados do payload ou carregar do JSON
            movimentacoes = payload.get('movimentacoes', self.load_json('movimentacoes.json'))
            ferramentas = payload.get('ferramentas', self.load_json('ferramentas.json'))
            solicitantes = payload.get('solicitantes', self.load_json('solicitantes.json'))

            # Sincronizar
            sucesso = sincronizar_tudo(movimentacoes, ferramentas, solicitantes)

            if sucesso:
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "success", "message": "Sincronização com Google Sheets realizada"}).encode())
            else:
                self.send_error(500, "Erro na sincronização com Google Sheets")

        except Exception as e:
            print(f"Erro na sincronização com Google Sheets: {e}")
            self.send_error(500, f"Erro interno: {str(e)}")

    def do_OPTIONS(self):
        # Handle CORS preflight requests
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def handle_sync(self):
        try:
            # Parse query parameters
            query = parse_qs(self.path.split('?', 1)[1] if '?' in self.path else '')
            last_sync = query.get('last_sync', [''])[0]

            # Load data from JSON files
            data = {
                'movimentacoes': self.load_json('movimentacoes.json', last_sync),
                'ferramentas': self.load_json('ferramentas.json', last_sync),
                'solicitantes': self.load_json('solicitantes.json', last_sync),
                'timestamp': datetime.now().isoformat()
            }

            # Send response
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(data, indent=2).encode())

        except Exception as e:
            self.send_error(500, f"Erro interno: {str(e)}")

    def handle_generate_sync_excel(self):
        try:
            print("Gerando Excel sincronizado...")

            # Executar o script de geração
            script_path = os.path.join(os.path.dirname(__file__), 'gerar_excel_sync_new.py')
            result = subprocess.run([sys.executable, script_path],
                                  capture_output=True, text=True, cwd=os.path.dirname(__file__))

            if result.returncode != 0:
                print(f"Erro ao executar script: {result.stderr}")
                self.send_error(500, f"Erro ao gerar Excel: {result.stderr}")
                return

            # Encontrar o arquivo gerado (mais recente)
            files = [f for f in os.listdir('.') if f.startswith('controle_ferramentas_') and f.endswith('.xlsm')]
            if not files:
                self.send_error(500, "Arquivo Excel não foi gerado")
                return

            latest_file = max(files, key=lambda f: os.path.getctime(f))
            filepath = os.path.join(os.path.dirname(__file__), latest_file)

            print(f"Servindo arquivo: {latest_file}")

            # Servir o arquivo
            self.send_response(200)
            self.send_header('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{latest_file}"')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()

            with open(filepath, 'rb') as f:
                self.wfile.write(f.read())

        except Exception as e:
            print(f"Erro ao gerar Excel sincronizado: {e}")
            self.send_error(500, f"Erro interno: {str(e)}")

    def handle_sync_google_sheets(self):
        try:
            print("Sincronizando com Google Sheets...")

            # Importar módulo de sync
            from google_sheets_sync import sincronizar_tudo

            # Carregar dados
            data = {
                'movimentacoes': self.load_json('movimentacoes.json'),
                'ferramentas': self.load_json('ferramentas.json'),
                'solicitantes': self.load_json('solicitantes.json')
            }

            # Sincronizar
            sucesso = sincronizar_tudo(
                data['movimentacoes'],
                data['ferramentas'],
                data['solicitantes']
            )

            if sucesso:
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({"status": "success", "message": "Sincronização com Google Sheets realizada"}).encode())
            else:
                self.send_error(500, "Erro na sincronização com Google Sheets")

        except Exception as e:
            print(f"Erro na sincronização com Google Sheets: {e}")
            self.send_error(500, f"Erro interno: {str(e)}")

    def load_json(self, filename, last_sync=None):
        filepath = os.path.join(os.path.dirname(__file__), filename)
        if not os.path.exists(filepath):
            return []

        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)

        if last_sync:
            try:
                last_sync_dt = datetime.fromisoformat(last_sync.replace('Z', '+00:00'))
                # Filter data modified after last_sync
                if isinstance(data, list):
                    return [item for item in data if
                           datetime.fromisoformat(item.get('dataRegistro', item.get('dataAdicao', '2000-01-01T00:00:00')).replace('Z', '+00:00')) > last_sync_dt]
                else:
                    return data
            except:
                pass

        return data

    def serve_static(self):
        try:
            # Get the file path
            path = self.path
            if path == '/':
                path = '/index.html'

            filepath = os.path.join(os.path.dirname(__file__), path[1:])

            # Security check - only serve files in current directory
            if not os.path.abspath(filepath).startswith(os.path.abspath(os.path.dirname(__file__))):
                self.send_error(403, "Acesso negado")
                return

            if os.path.exists(filepath) and os.path.isfile(filepath):
                # Determine content type
                if filepath.endswith('.html'):
                    content_type = 'text/html'
                elif filepath.endswith('.js'):
                    content_type = 'application/javascript'
                elif filepath.endswith('.css'):
                    content_type = 'text/css'
                elif filepath.endswith('.json'):
                    content_type = 'application/json'
                else:
                    content_type = 'text/plain'

                self.send_response(200)
                self.send_header('Content-type', content_type)
                # Disable cache for development
                self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                self.send_header('Pragma', 'no-cache')
                self.send_header('Expires', '0')
                self.end_headers()

                with open(filepath, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, "Arquivo não encontrado")

        except Exception as e:
            self.send_error(500, f"Erro interno: {str(e)}")

    def log_message(self, format, *args):
        # Log mais limpo - apenas requests de API
        if self.path.startswith('/api/'):
            print(f"{self.address_string()} - {self.path}")

def run_server(port=8000):
    with socketserver.TCPServer(("", port), DataSyncHandler) as httpd:
        print(f"Servidor rodando em http://localhost:{port}")
        print("API de sincronização disponível em /api/sync")
        print("Geração de Excel sincronizado em /api/generate_sync_excel")
        print("Sincronização com Google Sheets em /api/sync_google_sheets")
        print("Cache desabilitado para desenvolvimento")
        print("Pressione Ctrl+C para parar")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServidor parado")

if __name__ == "__main__":
    run_server()
