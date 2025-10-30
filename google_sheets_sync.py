#!/usr/bin/env python3
"""
Módulo para sincronização com Google Sheets via Google Apps Script
"""
import json
import requests
from datetime import datetime

# URL do Google Apps Script (doGet) - Usando o ID da planilha fornecido
GOOGLE_SHEETS_URL = "https://script.google.com/macros/s/AKfycbzka9zfxb9UcVz2kVafIWmiYT12YHx0JPb3zPU8jU1PN4BuNzBXeVUe1bMxxqG21b6O0A/exec"

def sincronizar_com_google_sheets(dados, aba="Retiradas"):
    """
    Sincroniza dados com Google Sheets via HTTP POST

    Args:
        dados: Lista de dicionários com os dados a sincronizar
        aba: Nome da aba no Google Sheets ("Retiradas" ou "Estoque")
    """
    try:
        # Preparar payload
        payload = {
            "aba": aba,
            "dados": dados,
            "timestamp": datetime.now().isoformat()
        }

        # Fazer request POST
        headers = {'Content-Type': 'application/json'}
        response = requests.post(GOOGLE_SHEETS_URL, json=payload, headers=headers, timeout=30)

        if response.status_code == 200:
            print(f"Dados sincronizados com sucesso na aba '{aba}' do Google Sheets")
            return True
        else:
            print(f"Erro ao sincronizar com Google Sheets: {response.status_code} - {response.text}")
            return False

    except requests.exceptions.RequestException as e:
        print(f"Erro de conexão com Google Sheets: {e}")
        return False
    except Exception as e:
        print(f"Erro inesperado na sincronização: {e}")
        return False

def sincronizar_retiradas(movimentacoes):
    """
    Sincroniza movimentações/retiradas com Google Sheets
    """
    # Converter dados para o formato esperado
    dados_formatados = []
    for mov in movimentacoes:
        dados_formatados.append({
            "Data": mov.get('dataRetirada', ''),
            "Ferramenta": mov.get('ferramenta', ''),
            "Patrimônio": mov.get('patrimonio', ''),
            "Solicitante": mov.get('solicitante', ''),
            "Tipo": mov.get('tipo', ''),
            "Data Devolução": mov.get('dataRetorno', ''),
            "Hora Devolução": mov.get('horaRetorno', ''),
            "Tem Retorno": mov.get('temRetorno', ''),
            "Observações": mov.get('observacao', '')
        })

    return sincronizar_com_google_sheets(dados_formatados, "Retiradas")

def sincronizar_estoque(ferramentas, solicitantes):
    """
    Sincroniza estoque (ferramentas e solicitantes) com Google Sheets
    """
    dados_formatados = []

    # Adicionar ferramentas
    for ferramenta in ferramentas:
        dados_formatados.append({
            "Tipo": "Ferramenta",
            "Nome": ferramenta.get('nome', ''),
            "Patrimônios": ', '.join(ferramenta.get('patrimonios', [])),
            "Descrição": f'Quantidade: {len(ferramenta.get("patrimonios", []))}'
        })

    # Adicionar solicitantes
    for solicitante in solicitantes:
        dados_formatados.append({
            "Tipo": "Solicitante",
            "Nome": solicitante.get('nome', ''),
            "Patrimônios": '',
            "Descrição": ''
        })

    return sincronizar_com_google_sheets(dados_formatados, "Estoque")

def sincronizar_tudo(movimentacoes, ferramentas, solicitantes):
    """
    Sincroniza todos os dados com Google Sheets
    """
    print("Iniciando sincronização com Google Sheets...")

    sucesso_retiradas = sincronizar_retiradas(movimentacoes)
    sucesso_estoque = sincronizar_estoque(ferramentas, solicitantes)

    if sucesso_retiradas and sucesso_estoque:
        print("Sincronização completa com Google Sheets realizada com sucesso!")
        return True
    else:
        print("Sincronização com Google Sheets concluída com erros.")
        return False

if __name__ == "__main__":
    # Teste com dados de exemplo
    print("Testando sincronização com Google Sheets...")

    # Dados de exemplo
    movimentacoes_teste = [
        {
            "dataRetirada": "2025-10-28",
            "ferramenta": "Furadeira",
            "patrimonio": "PAT001",
            "solicitante": "TESTE USUARIO",
            "tipo": "retirada",
            "temRetorno": "sim"
        }
    ]

    ferramentas_teste = [
        {
            "nome": "Furadeira",
            "patrimonios": ["PAT001", "PAT002"]
        }
    ]

    solicitantes_teste = []

    sincronizar_tudo(movimentacoes_teste, ferramentas_teste, solicitantes_teste)
