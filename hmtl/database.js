// Inicializar o banco de dados (simples, sempre resolve)
async function initDB() {
    try {
        console.log('Iniciando banco de dados localStorage...');

        // Verifica se localStorage está disponível
        if (typeof Storage === "undefined") {
            throw new Error('localStorage não está disponível neste navegador');
        }

        // Garante chaves básicas no localStorage
        const keys = ['ferramentas', 'movimentacoes', 'solicitantes', 'projetos'];
        keys.forEach(k => {
            try {
                if (localStorage.getItem(k) === null) {
                    localStorage.setItem(k, '[]');
                    console.log(`Chave ${k} inicializada`);
                }
            } catch (e) {
                console.error(`Erro ao inicializar chave ${k}:`, e);
            }
        });

        // Semeia solicitantes iniciais se disponível via CONFIG
        try {
            if (typeof window.APP_CONFIG !== 'undefined' && Array.isArray(window.APP_CONFIG.SOLICITANTES_INICIAIS)) {
                const existentes = getStore('solicitantes');
                if (existentes.length === 0 && window.APP_CONFIG.SOLICITANTES_INICIAIS.length > 0) {
                    const semeados = window.APP_CONFIG.SOLICITANTES_INICIAIS.map(nome => ({ nome, id: generateId() }));
                    setStore('solicitantes', semeados);
                    console.log('Solicitantes iniciais semeados');
                }
            }
        } catch (e) {
            console.warn('Falha ao semear solicitantes iniciais:', e);
        }

        // Marca DB como pronto para evitar erros de inicialização
        window.db = { ready: true };
        console.log('Banco de dados localStorage inicializado com sucesso');
        return true;
    } catch (error) {
        console.error('Erro fatal ao inicializar banco de dados:', error);
        // Mesmo com erro, marca como pronto para não travar a aplicação
        window.db = { ready: true, error: error.message };
        return false;
    }
}

// Funções auxiliares para localStorage
function getStore(key) {
    return JSON.parse(localStorage.getItem(key) || '[]');
}

function setStore(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
}

function generateId() {
    return Date.now() + Math.random();
}

// Funções para gerenciar ferramentas
async function adicionarFerramenta(ferramenta) {
    const ferramentas = getStore('ferramentas');
    const newFerramenta = { ...ferramenta, id: generateId() };
    ferramentas.push(newFerramenta);
    setStore('ferramentas', ferramentas);
    console.log('Ferramenta adicionada com sucesso');
    return newFerramenta.id;
}

async function obterFerramentas() {
    return getStore('ferramentas');
}

async function atualizarFerramenta(id, ferramenta) {
    const ferramentas = getStore('ferramentas');
    const index = ferramentas.findIndex(f => f.id === id);
    if (index !== -1) {
        ferramentas[index] = { ...ferramenta, id };
        setStore('ferramentas', ferramentas);
        console.log('Ferramenta atualizada com sucesso');
        return id;
    } else {
        throw new Error('Ferramenta não encontrada');
    }
}

async function removerFerramenta(id) {
    const ferramentas = getStore('ferramentas');
    const filtered = ferramentas.filter(f => f.id !== Number(id));
    setStore('ferramentas', filtered);
    console.log('Ferramenta removida com sucesso');
}

async function removerFerramentas(ids) {
    if (typeof ids === 'string') {
        ids = ids.split(',').map(id => Number(id.trim()));
    }
    const ferramentas = getStore('ferramentas');
    const filtered = ferramentas.filter(f => !ids.includes(f.id));
    setStore('ferramentas', filtered);
    console.log('Ferramentas removidas com sucesso');
}

// Funções para gerenciar movimentações
async function adicionarMovimentacao(movimentacao) {
    const movimentacoes = getStore('movimentacoes');
    const newMovimentacao = { ...movimentacao, id: generateId() };
    movimentacoes.push(newMovimentacao);
    setStore('movimentacoes', movimentacoes);
    console.log('Movimentação adicionada com sucesso');
    return newMovimentacao.id;
}

async function obterMovimentacoes() {
    return getStore('movimentacoes');
}

async function atualizarMovimentacao(id, movimentacao) {
    const movimentacoes = getStore('movimentacoes');
    const index = movimentacoes.findIndex(m => m.id === id);
    if (index !== -1) {
        movimentacoes[index] = { ...movimentacao, id };
        setStore('movimentacoes', movimentacoes);
        console.log('Movimentação atualizada com sucesso');
        return id;
    } else {
        throw new Error('Movimentação não encontrada');
    }
}

// Funções para gerenciar solicitantes
async function adicionarSolicitante(solicitante) {
    const solicitantes = getStore('solicitantes');
    const newSolicitante = { ...solicitante, id: generateId() };
    solicitantes.push(newSolicitante);
    setStore('solicitantes', solicitantes);
    console.log('Solicitante adicionado com sucesso');
    return newSolicitante.id;
}

async function obterSolicitantes() {
    return getStore('solicitantes');
}

async function atualizarSolicitante(id, solicitante) {
    const solicitantes = getStore('solicitantes');
    const index = solicitantes.findIndex(s => s.id === id);
    if (index !== -1) {
        solicitantes[index] = { ...solicitante, id };
        setStore('solicitantes', solicitantes);
        console.log('Solicitante atualizado com sucesso');
        return id;
    } else {
        throw new Error('Solicitante não encontrado');
    }
}

async function removerSolicitante(id) {
    const solicitantes = getStore('solicitantes');
    const filtered = solicitantes.filter(s => s.id !== Number(id));
    setStore('solicitantes', filtered);
    console.log('Solicitante removido com sucesso');
}

// Funções para gerenciar projetos (usando localStorage por enquanto)
async function adicionarProjeto(projeto) {
    const projetos = JSON.parse(localStorage.getItem('projetos') || '[]');
    const newProjeto = { ...projeto, id: Date.now() + Math.random() };
    projetos.push(newProjeto);
    localStorage.setItem('projetos', JSON.stringify(projetos));
    console.log('Projeto adicionado com sucesso');
    return newProjeto.id;
}

async function obterProjetos() {
    return JSON.parse(localStorage.getItem('projetos') || '[]');
}

async function atualizarProjeto(id, projeto) {
    const projetos = JSON.parse(localStorage.getItem('projetos') || '[]');
    const index = projetos.findIndex(p => p.id === id);
    if (index !== -1) {
        projetos[index] = { ...projeto, id };
        localStorage.setItem('projetos', JSON.stringify(projetos));
        console.log('Projeto atualizado com sucesso');
        return id;
    } else {
        throw new Error('Projeto não encontrado');
    }
}

async function removerProjeto(id) {
    const projetos = JSON.parse(localStorage.getItem('projetos') || '[]');
    const filtered = projetos.filter(p => p.id !== Number(id));
    localStorage.setItem('projetos', JSON.stringify(filtered));
    console.log('Projeto removido com sucesso');
}

// Criar o objeto dbManager global para acesso unificado
window.dbManager = {
    init: initDB,
    ferramentas: {
        adicionar: adicionarFerramenta,
        obterTodas: obterFerramentas,
        atualizar: atualizarFerramenta,
        remover: removerFerramenta,
        removerVarios: removerFerramentas
    },
    movimentacoes: {
        adicionar: adicionarMovimentacao,
        obterTodas: obterMovimentacoes,
        atualizar: atualizarMovimentacao
    },
    solicitantes: {
        adicionar: adicionarSolicitante,
        obterTodos: obterSolicitantes,
        atualizar: atualizarSolicitante,
        remover: removerSolicitante
    },
    projetos: {
        adicionar: adicionarProjeto,
        obterTodos: obterProjetos,
        atualizar: atualizarProjeto,
        remover: removerProjeto
    }
};


