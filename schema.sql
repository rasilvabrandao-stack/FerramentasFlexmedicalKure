-- Schema SQL para o sistema de gestão de ferramentas
-- Tabelas principais: ferramentas, movimentacoes, solicitantes

-- Tabela de solicitantes
CREATE TABLE IF NOT EXISTS solicitantes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE,
    email TEXT,
    telefone TEXT,
    departamento TEXT,
    criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
    atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Tabela de ferramentas
CREATE TABLE IF NOT EXISTS ferramentas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL,
    quantidade_total INTEGER DEFAULT 1,
    quantidade_disponivel INTEGER DEFAULT 1,
    status TEXT DEFAULT 'disponivel',
    criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
    atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- Tabela de movimentações
CREATE TABLE IF NOT EXISTS movimentacoes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo TEXT NOT NULL, -- 'saida' ou 'retorno'
    solicitante_id INTEGER NOT NULL,
    ferramenta_id INTEGER NOT NULL,
    quantidade INTEGER DEFAULT 1,
    data_saida DATE,
    data_retorno DATE,
    hora_devolucao TIME,
    tem_retorno TEXT DEFAULT 'Sim',
    observacoes TEXT,
    status TEXT DEFAULT 'ativo', -- 'ativo', 'concluido', 'cancelado'
    email_notificacao TEXT,
    criado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
    atualizado_em DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (solicitante_id) REFERENCES solicitantes (id),
    FOREIGN KEY (ferramenta_id) REFERENCES ferramentas (id)
);

-- Índices para melhor performance
CREATE INDEX IF NOT EXISTS idx_movimentacoes_solicitante ON movimentacoes(solicitante_id);
CREATE INDEX IF NOT EXISTS idx_movimentacoes_ferramenta ON movimentacoes(ferramenta_id);
CREATE INDEX IF NOT EXISTS idx_ferramentas_status ON ferramentas(status);

-- Inserir dados iniciais de solicitantes
INSERT OR IGNORE INTO solicitantes (nome) VALUES
('BRUNO GOMES DA SILVA'),
('CARLOS EDUARDO'),
('DANIEL SILVA'),
('EDUARDO SANTOS'),
('FERNANDO OLIVEIRA'),
('GABRIEL COSTA'),
('HENRIQUE ALVES'),
('IGOR PEREIRA'),
('JOÃO PEDRO'),
('LUCAS MARTINS');
