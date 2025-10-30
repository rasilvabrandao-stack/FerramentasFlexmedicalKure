#!/usr/bin/env python3
"""
Servidor Flask com integração ao banco de dados SQL
Fornece API REST para o sistema de gestão de ferramentas
"""
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import os
from database_sql import get_db_manager, init_database
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)
CORS(app)  # Habilita CORS para todas as rotas

# Inicializa o banco de dados
db_manager = None

def get_db():
    global db_manager
    if db_manager is None:
        db_manager = init_database()
    return db_manager

# Configurações de e-mail (ajuste conforme necessário)
EMAIL_CONFIG = {
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587,
    'username': 'seu_email@gmail.com',  # Substitua pelo seu e-mail
    'password': 'sua_senha_app',        # Substitua pela senha do app
    'from_email': 'seu_email@gmail.com'
}

def enviar_email_notificacao(destinatario, assunto, mensagem):
    """Envia e-mail de notificação"""
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_CONFIG['from_email']
        msg['To'] = destinatario
        msg['Subject'] = assunto

        msg.attach(MIMEText(mensagem, 'html'))

        server = smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port'])
        server.starttls()
        server.login(EMAIL_CONFIG['username'], EMAIL_CONFIG['password'])
        text = msg.as_string()
        server.sendmail(EMAIL_CONFIG['from_email'], destinatario, text)
        server.quit()

        print(f"E-mail enviado para {destinatario}")
        return True
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        return False

@app.route('/')
def index():
    """Serve o arquivo index.html"""
    return send_from_directory('.', 'index.html')

@app.route('/<path:filename>')
def serve_static(filename):
    """Serve arquivos estáticos"""
    return send_from_directory('.', filename)

# API Routes

@app.route('/api/solicitantes', methods=['GET', 'POST'])
def handle_solicitantes():
    """Gerencia solicitantes"""
    db = get_db()

    if request.method == 'GET':
        try:
            solicitantes = db.obter_solicitantes()
            return jsonify({'success': True, 'data': solicitantes})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

    elif request.method == 'POST':
        try:
            data = request.get_json()
            solicitante_id = db.adicionar_solicitante(
                nome=data['nome'],
                email=data.get('email'),
                telefone=data.get('telefone'),
                departamento=data.get('departamento')
            )
            return jsonify({'success': True, 'id': solicitante_id, 'message': 'Solicitante adicionado com sucesso'})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/solicitantes/<int:id>', methods=['PUT', 'DELETE'])
def handle_solicitante(id):
    """Gerencia solicitante específico"""
    db = get_db()

    if request.method == 'PUT':
        try:
            data = request.get_json()
            success = db.atualizar_solicitante(
                id=id,
                nome=data.get('nome'),
                email=data.get('email'),
                telefone=data.get('telefone'),
                departamento=data.get('departamento')
            )
            if success:
                return jsonify({'success': True, 'message': 'Solicitante atualizado com sucesso'})
            else:
                return jsonify({'success': False, 'error': 'Solicitante não encontrado'}), 404
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

    elif request.method == 'DELETE':
        try:
            success = db.remover_solicitante(id)
            if success:
                return jsonify({'success': True, 'message': 'Solicitante removido com sucesso'})
            else:
                return jsonify({'success': False, 'error': 'Solicitante não encontrado'}), 404
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500



@app.route('/api/ferramentas', methods=['GET', 'POST'])
def handle_ferramentas():
    """Gerencia ferramentas"""
    db = get_db()

    if request.method == 'GET':
        try:
            ferramentas = db.obter_ferramentas()
            return jsonify({'success': True, 'data': ferramentas})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

    elif request.method == 'POST':
        try:
            data = request.get_json()
            ferramenta_id = db.adicionar_ferramenta(
                nome=data['nome'],
                quantidade_total=data.get('quantidade_total', 1)
            )
            return jsonify({'success': True, 'id': ferramenta_id, 'message': 'Ferramenta adicionada com sucesso'})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/ferramentas/<int:id>', methods=['PUT', 'DELETE'])
def handle_ferramenta(id):
    """Gerencia ferramenta específica"""
    db = get_db()

    if request.method == 'PUT':
        try:
            data = request.get_json()
            success = db.atualizar_ferramenta(
                id=id,
                nome=data.get('nome'),
                quantidade_total=data.get('quantidade_total')
            )
            if success:
                return jsonify({'success': True, 'message': 'Ferramenta atualizada com sucesso'})
            else:
                return jsonify({'success': False, 'error': 'Ferramenta não encontrada'}), 404
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

    elif request.method == 'DELETE':
        try:
            success = db.remover_ferramenta(id)
            if success:
                return jsonify({'success': True, 'message': 'Ferramenta removida com sucesso'})
            else:
                return jsonify({'success': False, 'error': 'Ferramenta não encontrada'}), 404
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/movimentacoes', methods=['GET', 'POST'])
def handle_movimentacoes():
    """Gerencia movimentações"""
    db = get_db()

    if request.method == 'GET':
        try:
            status = request.args.get('status')
            movimentacoes = db.obter_movimentacoes(status=status)
            return jsonify({'success': True, 'data': movimentacoes})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

    elif request.method == 'POST':
        try:
            data = request.get_json()

            # Adiciona movimentação
            movimentacao_id = db.adicionar_movimentacao(
                tipo=data['tipo'],
                solicitante_id=data['solicitante_id'],
                ferramenta_id=data['ferramenta_id'],
                data_saida=data.get('dataSaida'),
                data_retorno=data.get('dataRetorno'),
                hora_devolucao=data.get('horaDevolucao'),
                tem_retorno=data.get('temRetorno', 'Sim'),
                observacoes=data.get('observacoes'),
                projeto_id=data.get('projeto_id'),
                email_notificacao=data.get('emailNotificacao')
            )

            # Envia e-mail de notificação se fornecido
            if data.get('emailNotificacao'):
                assunto = f"Notificação de {'Empréstimo' if data['tipo'].lower() == 'saida' else 'Devolução'} de Ferramenta"
                mensagem = f"""
                <h3>Notificação de Movimentação de Ferramenta</h3>
                <p><strong>Tipo:</strong> {data['tipo']}</p>
                <p><strong>Ferramenta:</strong> {data.get('ferramenta', 'N/A')}</p>
                <p><strong>Solicitante:</strong> {data.get('solicitante', 'N/A')}</p>
                <p><strong>Data de Saída:</strong> {data.get('dataSaida', 'N/A')}</p>
                <p><strong>Data de Retorno:</strong> {data.get('dataRetorno', 'N/A')}</p>
                <p><strong>Observações:</strong> {data.get('observacoes', 'Nenhuma')}</p>
                """
                enviar_email_notificacao(data['emailNotificacao'], assunto, mensagem)

            return jsonify({'success': True, 'id': movimentacao_id, 'message': 'Movimentação registrada e e-mail enviado'})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/movimentacoes/<int:id>', methods=['PUT'])
def handle_movimentacao(id):
    """Atualiza movimentação específica"""
    db = get_db()

    try:
        data = request.get_json()
        success = db.atualizar_movimentacao(id, **data)
        if success:
            return jsonify({'success': True, 'message': 'Movimentação atualizada com sucesso'})
        else:
            return jsonify({'success': False, 'error': 'Movimentação não encontrada'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/movimentacoes/<int:id>/concluir', methods=['POST'])
def concluir_movimentacao(id):
    """Conclui uma movimentação (retorno)"""
    db = get_db()

    try:
        success = db.concluir_movimentacao(id)
        if success:
            return jsonify({'success': True, 'message': 'Movimentação concluída com sucesso'})
        else:
            return jsonify({'success': False, 'error': 'Movimentação não encontrada'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/estatisticas', methods=['GET'])
def get_estatisticas():
    """Retorna estatísticas do sistema"""
    db = get_db()

    try:
        stats = db.obter_estatisticas()
        return jsonify({'success': True, 'data': stats})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/google-sheets', methods=['POST'])
def proxy_google_sheets():
    """Proxy para Google Sheets (mantém compatibilidade)"""
    try:
        import urllib.request
        import json

        # URL do Google Apps Script
        google_apps_url = "https://script.google.com/macros/s/AKfycbw7_F6_p_cnLGenGmPFbep7zHwdZ5UcAYC1OXLu8Jp7SXrdjU9Nncimkxpuvt8qRw7oBA/exec"

        # Obtém dados da requisição
        data = request.get_json()
        if not data:
            return jsonify({'success': False, 'error': 'Dados JSON necessários'}), 400

        # Converte para JSON string
        json_data = json.dumps(data).encode('utf-8')

        # Cria requisição para Google Apps Script
        req = urllib.request.Request(google_apps_url, data=json_data, method='POST')
        req.add_header('Content-Type', 'application/json')

        # Faz a requisição
        with urllib.request.urlopen(req) as response:
            result = response.read().decode('utf-8')

        # Retorna resposta do Google Apps Script
        return result, response.getcode(), {'Content-Type': 'application/json'}

    except Exception as e:
        print(f"Erro no proxy Google Sheets: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/backup', methods=['POST'])
def criar_backup():
    """Cria backup do banco de dados"""
    try:
        db = get_db()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = f'backup_ferramentas_{timestamp}.db'
        db.backup_database(backup_path)
        return jsonify({'success': True, 'message': f'Backup criado: {backup_path}'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/db-viewer')
def db_viewer():
    """Serve a página do visualizador de banco de dados"""
    return send_from_directory('.', 'db_viewer.html')

@app.route('/api/db/tables', methods=['GET'])
def get_db_tables():
    """Retorna lista de tabelas do banco de dados"""
    try:
        db = get_db()
        tables = db.obter_tabelas()
        return jsonify({'success': True, 'tables': tables})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/db/<table_name>', methods=['GET'])
def get_table_data(table_name):
    """Retorna dados de uma tabela específica"""
    try:
        db = get_db()
        data = db.obter_dados_tabela(table_name)
        columns = db.obter_colunas_tabela(table_name)
        total = db.contar_registros_tabela(table_name)

        return jsonify({
            'success': True,
            'data': data,
            'columns': columns,
            'total': total
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    print("Iniciando servidor Flask com banco de dados SQL...")
    print("Acesse: http://localhost:8000")
    print("Visualizador de BD: http://localhost:8000/db-viewer")
    app.run(host='0.0.0.0', port=8000, debug=True)
