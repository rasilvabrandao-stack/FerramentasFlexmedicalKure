#!/usr/bin/env python3
"""
Script para visualizar o conteúdo do banco de dados SQLite
"""
import sqlite3
import json

def view_database():
    conn = sqlite3.connect('ferramentas.db')
    cursor = conn.cursor()

    print('=== BANCO DE DADOS FERRAMENTAS ===')
    print()

    # Listar tabelas
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    print('Tabelas criadas:')
    for table in tables:
        print(f'  - {table[0]}')
    print()

    # Mostrar dados de cada tabela
    tables_to_show = ['solicitantes', 'projetos', 'ferramentas', 'movimentacoes']

    for table in tables_to_show:
        try:
            cursor.execute(f'SELECT COUNT(*) FROM {table}')
            count = cursor.fetchone()[0]
            print(f'{table.upper()} ({count} registros):')

            cursor.execute(f'SELECT * FROM {table} LIMIT 5')
            rows = cursor.fetchall()

            if rows:
                # Mostrar estrutura da tabela
                cursor.execute(f'PRAGMA table_info({table})')
                columns = cursor.fetchall()
                col_names = [col[1] for col in columns]
                print(f'  Colunas: {", ".join(col_names)}')

                # Mostrar primeiros registros
                for i, row in enumerate(rows):
                    print(f'  Registro {i+1}: {dict(zip(col_names, row))}')
            else:
                print('  Nenhum registro encontrado')
            print()
        except sqlite3.Error as e:
            print(f'Erro ao consultar {table}: {e}')

    conn.close()
    print('=== FIM DA VISUALIZAÇÃO ===')

if __name__ == "__main__":
    view_database()
