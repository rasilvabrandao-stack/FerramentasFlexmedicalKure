#!/usr/bin/env python3
"""
Módulo de banco de dados SQL para o sistema de gestão de ferramentas
Usa SQLite3 como backend de banco de dados
"""
import sqlite3
import json
from datetime import datetime
from typing import List, Dict, Optional, Any
import os

class DatabaseManager:
    def __init__(self, db_path: str = 'ferramentas.db'):
        self.db_path = db_path
        self.connection = None
        self.connect()

    def connect(self):
        """Conecta ao banco de dados SQLite"""
        try:
            self.connection = sqlite3.connect(self.db_path, check_same_thread=False)
            self.connection.row_factory = sqlite3.Row  # Permite acesso por nome de coluna
            self.connection.execute("PRAGMA foreign_keys = ON")  # Habilita chaves estrangeiras
            print(f"Conectado ao banco de dados: {self.db_path}")
        except sqlite3.Error as e:
            print(f"Erro ao conectar ao banco de dados: {e}")
            raise

    def initialize_database(self):
        """Inicializa o banco de dados com o schema"""
        try:
            with open('schema.sql', 'r', encoding='utf-8') as f:
                schema = f.read()

            self.connection.executescript(schema)
            self.connection.commit()
            print("Banco de dados inicializado com sucesso")
        except FileNotFoundError:
            print("Arquivo schema.sql não encontrado")
            raise
        except sqlite3.Error as e:
            print(f"Erro ao inicializar banco de dados: {e}")
            raise

    def close(self):
        """Fecha a conexão com o banco de dados"""
        if self.connection:
            self.connection.close()
            print("Conexão com banco de dados fechada")

    # Métodos para Solicitantes
    def adicionar_solicitante(self, nome: str, email: str = None, telefone: str = None, departamento: str = None) -> int:
        """Adiciona um novo solicitante"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                INSERT INTO solicitantes (nome, email, telefone, departamento)
                VALUES (?, ?, ?, ?)
            """, (nome, email, telefone, departamento))
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Erro ao adicionar solicitante: {e}")
            raise

    def obter_solicitantes(self) -> List[Dict]:
        """Retorna todos os solicitantes"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT * FROM solicitantes ORDER BY nome")
            rows = cursor.fetchall()
            return [dict(row) for row in rows]
        except sqlite3.Error as e:
            print(f"Erro ao obter solicitantes: {e}")
            raise

    def atualizar_solicitante(self, id: int, nome: str = None, email: str = None, telefone: str = None, departamento: str = None) -> bool:
        """Atualiza um solicitante"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                UPDATE solicitantes
                SET nome = COALESCE(?, nome),
                    email = COALESCE(?, email),
                    telefone = COALESCE(?, telefone),
                    departamento = COALESCE(?, departamento),
                    atualizado_em = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (nome, email, telefone, departamento, id))
            self.connection.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Erro ao atualizar solicitante: {e}")
            raise

    def remover_solicitante(self, id: int) -> bool:
        """Remove um solicitante"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM solicitantes WHERE id = ?", (id,))
            self.connection.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Erro ao remover solicitante: {e}")
            raise



    # Métodos para Ferramentas
    def adicionar_ferramenta(self, nome: str, quantidade_total: int = 1) -> int:
        """Adiciona uma nova ferramenta"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                INSERT INTO ferramentas (nome, quantidade_total, quantidade_disponivel)
                VALUES (?, ?, ?)
            """, (nome, quantidade_total, quantidade_total))
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Erro ao adicionar ferramenta: {e}")
            raise

    def obter_ferramentas(self) -> List[Dict]:
        """Retorna todas as ferramentas"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT * FROM ferramentas ORDER BY nome")
            rows = cursor.fetchall()
            return [dict(row) for row in rows]
        except sqlite3.Error as e:
            print(f"Erro ao obter ferramentas: {e}")
            raise

    def atualizar_ferramenta(self, id: int, nome: str = None, quantidade_total: int = None) -> bool:
        """Atualiza uma ferramenta"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                UPDATE ferramentas
                SET nome = COALESCE(?, nome),
                    quantidade_total = COALESCE(?, quantidade_total),
                    atualizado_em = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (nome, quantidade_total, id))
            self.connection.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Erro ao atualizar ferramenta: {e}")
            raise

    def remover_ferramenta(self, id: int) -> bool:
        """Remove uma ferramenta"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM ferramentas WHERE id = ?", (id,))
            self.connection.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Erro ao remover ferramenta: {e}")
            raise

    # Métodos para Movimentações
    def adicionar_movimentacao(self, tipo: str, solicitante_id: int, ferramenta_id: int,
                             data_saida: str = None, data_retorno: str = None, hora_devolucao: str = None,
                             tem_retorno: str = 'Sim', observacoes: str = None,
                             email_notificacao: str = None) -> int:
        """Adiciona uma nova movimentação"""
        try:
            cursor = self.connection.cursor()

            # Se for saída, decrementa quantidade disponível
            if tipo.lower() == 'saida':
                cursor.execute("""
                    UPDATE ferramentas
                    SET quantidade_disponivel = quantidade_disponivel - 1
                    WHERE id = ? AND quantidade_disponivel > 0
                """, (ferramenta_id,))

                if cursor.rowcount == 0:
                    raise ValueError("Ferramenta não disponível para empréstimo")

            cursor.execute("""
                INSERT INTO movimentacoes (tipo, solicitante_id, ferramenta_id,
                                        data_saida, data_retorno, hora_devolucao, tem_retorno,
                                        observacoes, email_notificacao)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (tipo, solicitante_id, ferramenta_id, data_saida, data_retorno,
                  hora_devolucao, tem_retorno, observacoes, email_notificacao))

            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Erro ao adicionar movimentação: {e}")
            raise

    def obter_movimentacoes(self, status: str = None) -> List[Dict]:
        """Retorna todas as movimentações"""
        try:
            cursor = self.connection.cursor()
            query = """
                SELECT m.*, s.nome as solicitante_nome, f.nome as ferramenta_nome
                FROM movimentacoes m
                JOIN solicitantes s ON m.solicitante_id = s.id
                JOIN ferramentas f ON m.ferramenta_id = f.id
            """
            params = []

            if status:
                query += " WHERE m.status = ?"
                params.append(status)

            query += " ORDER BY m.criado_em DESC"

            cursor.execute(query, params)
            rows = cursor.fetchall()
            return [dict(row) for row in rows]
        except sqlite3.Error as e:
            print(f"Erro ao obter movimentações: {e}")
            raise

    def atualizar_movimentacao(self, id: int, **kwargs) -> bool:
        """Atualiza uma movimentação"""
        try:
            cursor = self.connection.cursor()
            fields = []
            values = []

            allowed_fields = ['tipo', 'solicitante_id', 'ferramenta_id', 'projeto_id',
                            'data_saida', 'data_retorno', 'hora_devolucao', 'tem_retorno',
                            'observacoes', 'status', 'email_notificacao']

            for field in allowed_fields:
                if field in kwargs:
                    fields.append(f"{field} = ?")
                    values.append(kwargs[field])

            if not fields:
                return False

            fields.append("atualizado_em = CURRENT_TIMESTAMP")
            query = f"UPDATE movimentacoes SET {', '.join(fields)} WHERE id = ?"
            values.append(id)

            cursor.execute(query, values)
            self.connection.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Erro ao atualizar movimentação: {e}")
            raise

    def concluir_movimentacao(self, id: int) -> bool:
        """Conclui uma movimentação (retorno de ferramenta)"""
        try:
            cursor = self.connection.cursor()

            # Busca a movimentação
            cursor.execute("SELECT * FROM movimentacoes WHERE id = ?", (id,))
            mov = cursor.fetchone()

            if not mov:
                return False

            # Se for retorno, incrementa quantidade disponível
            if mov['tipo'].lower() == 'saida':
                cursor.execute("""
                    UPDATE ferramentas
                    SET quantidade_disponivel = quantidade_disponivel + 1
                    WHERE id = ?
                """, (mov['ferramenta_id'],))

            # Atualiza status da movimentação
            cursor.execute("""
                UPDATE movimentacoes
                SET status = 'concluido', atualizado_em = CURRENT_TIMESTAMP
                WHERE id = ?
            """, (id,))

            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Erro ao concluir movimentação: {e}")
            raise

    # Métodos utilitários
    def obter_estatisticas(self) -> Dict:
        """Retorna estatísticas do sistema"""
        try:
            cursor = self.connection.cursor()

            stats = {}

            # Contagem de ferramentas
            cursor.execute("SELECT COUNT(*) as total FROM ferramentas")
            stats['total_ferramentas'] = cursor.fetchone()['total']

            # Ferramentas disponíveis
            cursor.execute("SELECT COUNT(*) as total FROM ferramentas WHERE quantidade_disponivel > 0")
            stats['ferramentas_disponiveis'] = cursor.fetchone()['total']

            # Movimentações ativas
            cursor.execute("SELECT COUNT(*) as total FROM movimentacoes WHERE status = 'ativo'")
            stats['movimentacoes_ativas'] = cursor.fetchone()['total']

            # Total de solicitantes
            cursor.execute("SELECT COUNT(*) as total FROM solicitantes")
            stats['total_solicitantes'] = cursor.fetchone()['total']

            return stats
        except sqlite3.Error as e:
            print(f"Erro ao obter estatísticas: {e}")
            raise

    def backup_database(self, backup_path: str):
        """Cria backup do banco de dados"""
        try:
            with sqlite3.connect(backup_path) as backup_conn:
                self.connection.backup(backup_conn)
            print(f"Backup criado em: {backup_path}")
        except sqlite3.Error as e:
            print(f"Erro ao criar backup: {e}")
            raise

    # Métodos para visualização do banco de dados
    def obter_tabelas(self) -> List[Dict]:
        """Retorna lista de tabelas do banco de dados com contagem de registros"""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
            tables = cursor.fetchall()

            result = []
            for table in tables:
                table_name = table['name']
                try:
                    cursor.execute(f"SELECT COUNT(*) as count FROM {table_name}")
                    count = cursor.fetchone()['count']
                    result.append({'name': table_name, 'records': count})
                except sqlite3.Error:
                    result.append({'name': table_name, 'records': 0})

            return result
        except sqlite3.Error as e:
            print(f"Erro ao obter tabelas: {e}")
            raise

    def obter_colunas_tabela(self, table_name: str) -> List[str]:
        """Retorna lista de colunas de uma tabela"""
        try:
            cursor = self.connection.cursor()
            cursor.execute(f"PRAGMA table_info({table_name})")
            columns = cursor.fetchall()
            return [col['name'] for col in columns]
        except sqlite3.Error as e:
            print(f"Erro ao obter colunas da tabela {table_name}: {e}")
            raise

    def obter_dados_tabela(self, table_name: str, limit: int = None) -> List[Dict]:
        """Retorna dados de uma tabela"""
        try:
            cursor = self.connection.cursor()
            query = f"SELECT * FROM {table_name}"
            if limit:
                query += f" LIMIT {limit}"
            cursor.execute(query)
            rows = cursor.fetchall()
            return [dict(row) for row in rows]
        except sqlite3.Error as e:
            print(f"Erro ao obter dados da tabela {table_name}: {e}")
            raise

    def contar_registros_tabela(self, table_name: str) -> int:
        """Conta registros de uma tabela"""
        try:
            cursor = self.connection.cursor()
            cursor.execute(f"SELECT COUNT(*) as count FROM {table_name}")
            return cursor.fetchone()['count']
        except sqlite3.Error as e:
            print(f"Erro ao contar registros da tabela {table_name}: {e}")
            raise

# Instância global do gerenciador de banco de dados
db_manager = None

def get_db_manager() -> DatabaseManager:
    """Retorna a instância global do gerenciador de banco de dados"""
    global db_manager
    if db_manager is None:
        db_manager = DatabaseManager()
    return db_manager

def init_database():
    """Inicializa o banco de dados"""
    manager = get_db_manager()
    manager.initialize_database()
    return manager

if __name__ == "__main__":
    # Inicialização do banco de dados quando executado diretamente
    init_database()
    print("Banco de dados SQL inicializado com sucesso!")
