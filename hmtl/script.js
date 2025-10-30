// === script.js (Unificado para Painel e Formulário) ===

// O evento DOMContentLoaded garante que o HTML foi completamente carregado antes de o script rodar.
document.addEventListener('DOMContentLoaded', () => {

    // --- DETECÇÃO DE PÁGINA ---
    // Verificamos a existência de um elemento único de cada página para saber onde estamos.
    const isPaginaFormulario = document.getElementById('ferramentaForm') !== null;
    const isPaginaPainel = document.getElementById('painel') !== null;

    if (isPaginaFormulario) {
        // Se encontrarmos o formulário, inicializamos a lógica da página de retirada.
        inicializarPaginaFormulario();
    }

    if (isPaginaPainel) {
        // Se encontrarmos o painel, inicializamos a lógica da página de administração.
        inicializarPaginaPainel();
    }
});


// ===================================================================================
// === FUNÇÃO PARA ENVIAR DADOS PARA GOOGLE SHEETS ==================================
// ===================================================================================

// Função para enviar dados para Google Sheets via Apps Script com retry
async function doPost(payload, maxRetries = 3, baseDelay = 1000) {
    let lastError;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            // Converter payload JSON para form-encoded
            console.log(`Tentativa ${attempt}/${maxRetries} - Enviando dados para Google Sheets:`, payload);

            const response = await fetch(window.APP_CONFIG.API_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(payload)
            });

            console.log(`Tentativa ${attempt} - Resposta do Google Sheets:`, response.status, response.statusText);

            if (!response.ok) {
                const errorText = await response.text();
                console.error(`Tentativa ${attempt} - Erro na resposta:`, errorText);

                // Se for erro 429 (Too Many Requests) ou 503 (Service Unavailable), tentar novamente
                if (response.status === 429 || response.status === 503 || response.status === 502) {
                    lastError = new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
                    if (attempt < maxRetries) {
                        const delay = baseDelay * Math.pow(2, attempt - 1); // Exponential backoff
                        console.log(`Tentativa ${attempt} falhou. Aguardando ${delay}ms antes de tentar novamente...`);
                        await new Promise(resolve => setTimeout(resolve, delay));
                        continue;
                    }
                } else {
                    throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
                }
            }

            const result = await response.text(); // Google Apps Script retorna texto
            console.log(`Tentativa ${attempt} - Resultado do Google Sheets:`, result);
            return { status: 'success', message: result };
        } catch (error) {
            console.error(`Tentativa ${attempt} - Erro na requisição doPost:`, error);
            lastError = error;

            if (attempt < maxRetries) {
                const delay = baseDelay * Math.pow(2, attempt - 1);
                console.log(`Tentativa ${attempt} falhou. Aguardando ${delay}ms antes de tentar novamente...`);
                await new Promise(resolve => setTimeout(resolve, delay));
            }
        }
    }

    // Se todas as tentativas falharam
    console.error('Todas as tentativas falharam. Último erro:', lastError);
    throw lastError;
}

// ===================================================================================
// === FUNÇÕES GLOBAIS PARA EXCEL ====================================================
// ===================================================================================

// Função auxiliar para obter chave mensal do Excel
function getMesAtual() {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

// Função para salvar workbook no localStorage
function salvarWorkbookMensal(wb, mes) {
    try {
        localStorage.setItem('excel_mes_' + mes, JSON.stringify(wb));
        console.log('Workbook salvo para mês:', mes);
    } catch (error) {
        console.error('Erro ao salvar workbook:', error);
    }
}

// Função para carregar workbook do localStorage
function carregarWorkbookMensal(mes) {
    try {
        const data = localStorage.getItem('excel_mes_' + mes);
        if (data) {
            const wb = XLSX.utils.book_new();
            Object.assign(wb, JSON.parse(data));
            return wb;
        }
    } catch (error) {
        console.error('Erro ao carregar workbook:', error);
    }
    return null;
}

// Função para atualizar workbook mensal com novos dados
async function atualizarWorkbookMensal() {
    const mes = getMesAtual();
    let wb = carregarWorkbookMensal(mes);
    if (!wb) {
        // Se não existe, gerar novo
        wb = await gerarWorkbookCompleto();
        salvarWorkbookMensal(wb, mes);
    } else {
        // Atualizar com dados atuais
        const newWb = await gerarWorkbookCompleto();
        // Substituir abas
        wb.Sheets = newWb.Sheets;
        wb.SheetNames = newWb.SheetNames;
        salvarWorkbookMensal(wb, mes);
    }
}

// Função para gerar workbook completo (extraída)
async function gerarWorkbookCompleto() {
    // Obter dados do sistema
    const movimentacoes = await window.dbManager.movimentacoes.obterTodas();
    const ferramentas = await window.dbManager.ferramentas.obterTodas();
    const solicitantes = await window.dbManager.solicitantes.obterTodos();


    // Criar workbook
    const wb = XLSX.utils.book_new();

    // Função auxiliar para aplicar estilos básicos
    function applyBasicStyling(ws, headers, title) {
        // Adicionar título na primeira linha
        XLSX.utils.sheet_add_aoa(ws, [[title]], { origin: 'A1' });
        // Adicionar headers na segunda linha
        XLSX.utils.sheet_add_aoa(ws, [headers], { origin: 'A2' });

        // Calcular o range atual da planilha
        const range = XLSX.utils.decode_range(ws['!ref']);

        // Mover dados para baixo (para dar espaço ao título e headers)
        for (let R = range.e.r; R >= 0; --R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const addr = XLSX.utils.encode_cell({ r: R + 2, c: C });
                const addr_old = XLSX.utils.encode_cell({ r: R, c: C });
                if (ws[addr_old]) ws[addr] = ws[addr_old];
                delete ws[addr_old];
            }
        }

        // Atualizar range
        range.e.r += 2;
        ws['!ref'] = XLSX.utils.encode_range(range);

        // Aplicar estilos
        if (!ws['!cols']) ws['!cols'] = [];
        headers.forEach((_, i) => {
            if (!ws['!cols'][i]) ws['!cols'][i] = { wch: 15 };
        });

        // Estilo para título (linha 1)
        const titleCell = ws['A1'];
        if (titleCell) {
            titleCell.s = {
                font: { bold: true, sz: 16, color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "4F81BD" } },
                alignment: { horizontal: "center" }
            };
        }

        // Estilo para headers (linha 2)
        headers.forEach((_, i) => {
            const cellAddr = XLSX.utils.encode_cell({ r: 1, c: i });
            const cell = ws[cellAddr];
            if (cell) {
                cell.s = {
                    font: { bold: true, color: { rgb: "FFFFFF" } },
                    fill: { fgColor: { rgb: "9BC2E6" } },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } }
                    }
                };
            }
        });

        // Bordas para todas as células de dados
        const dataRange = XLSX.utils.decode_range(ws['!ref']);
        for (let R = 2; R <= dataRange.e.r; ++R) {
            for (let C = dataRange.s.c; C <= dataRange.e.c; ++C) {
                const cellAddr = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[cellAddr];
                if (cell) {
                    if (!cell.s) cell.s = {};
                    cell.s.border = {
                        top: { style: "thin", color: { rgb: "D9D9D9" } },
                        bottom: { style: "thin", color: { rgb: "D9D9D9" } },
                        left: { style: "thin", color: { rgb: "D9D9D9" } },
                        right: { style: "thin", color: { rgb: "D9D9D9" } }
                    };
                }
            }
        }

        // Mesclar células para o título
        ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];
    }

    // Aba 1: Retiradas (dados das solicitações do index.html)
    const retiradasHeaders = ['Data', 'Ferramenta', 'Patrimônio', 'Solicitante', 'Tipo', 'Data Devolução', 'Hora Devolução', 'Tem Retorno', 'Observações'];
    const retiradasData = movimentacoes.map(mov => [
        mov.dataRetirada || '',
        mov.ferramenta || '',
        mov.patrimonio || '',
        mov.solicitante || '',
        mov.tipo || '',
        mov.dataRetorno || '',
        mov.horaRetorno || '',
        mov.temRetorno || '',
        mov.observacoes || ''
    ]);
    const wsRetiradas = XLSX.utils.aoa_to_sheet([retiradasHeaders, ...retiradasData]);
    applyBasicStyling(wsRetiradas, retiradasHeaders, 'Relatório de Retiradas de Ferramentas');
    XLSX.utils.book_append_sheet(wb, wsRetiradas, 'Retiradas');

    // Aba 2: Painel (dados administrativos)
    const painelHeaders = ['Tipo', 'Nome', 'Patrimônios', 'Descrição'];
    const painelData = [];

    // Adicionar ferramentas
    ferramentas.forEach(f => {
        painelData.push(['Ferramenta', f.nome || '', f.patrimonios ? f.patrimonios.join(', ') : '', f.descricao || '']);
    });



    // Adicionar solicitantes
    solicitantes.forEach(s => {
        painelData.push(['Solicitante', s.nome || '', '', '']);
    });

    const wsPainel = XLSX.utils.aoa_to_sheet([painelHeaders, ...painelData]);
    applyBasicStyling(wsPainel, painelHeaders, 'Dados Administrativos do Sistema');
    XLSX.utils.book_append_sheet(wb, wsPainel, 'Painel');

    // Aba 3: Análise Gráfica (dados agregados para gráficos)
    const analiseHeaders = ['Tipo', 'Categoria', 'Valor', 'Porcentagem', 'Descrição'];
    const analiseData = [];

    // KPIs principais
    const totalRetiradas = movimentacoes.length;
    const totalFerramentas = ferramentas.reduce((acc, f) => acc + (f.patrimonios ? f.patrimonios.length : 0), 0);
    const totalEmUso = movimentacoes.filter(mov => {
        if (mov.tipo !== 'quebrada') {
            if (mov.temRetorno === 'sim') {
                return !mov.dataRetorno || mov.dataRetorno === '';
            }
            return true;
        }
        return false;
    }).length;
    const totalQuebradas = movimentacoes.filter(mov => mov.tipo === 'quebrada').length;

    analiseData.push(['KPI', 'Total de Retiradas', totalRetiradas, '', 'Número total de solicitações de retirada']);
    analiseData.push(['KPI', 'Ferramentas no Estoque', totalFerramentas, '', 'Total de ferramentas disponíveis']);
    analiseData.push(['KPI', 'Ferramentas em Uso', totalEmUso, totalFerramentas > 0 ? ((totalEmUso / totalFerramentas) * 100).toFixed(1) + '%' : '0%', 'Ferramentas atualmente emprestadas']);
    analiseData.push(['KPI', 'Ferramentas Quebradas', totalQuebradas, totalFerramentas > 0 ? ((totalQuebradas / totalFerramentas) * 100).toFixed(1) + '%' : '0%', 'Ferramentas danificadas']);

    // Dados para gráfico de barras: Retiradas por ferramenta
    const retiradasPorFerramenta = {};
    movimentacoes.forEach(mov => {
        if (mov.ferramenta) {
            retiradasPorFerramenta[mov.ferramenta] = (retiradasPorFerramenta[mov.ferramenta] || 0) + 1;
        }
    });

    Object.entries(retiradasPorFerramenta).forEach(([ferramenta, quantidade]) => {
        analiseData.push(['Gráfico de Barras', 'Retiradas por Ferramenta', quantidade, totalRetiradas > 0 ? ((quantidade / totalRetiradas) * 100).toFixed(1) + '%' : '0%', ferramenta]);
    });

    // Dados para gráfico de pizza: Retiradas por solicitante
    const retiradasPorSolicitante = {};
    movimentacoes.forEach(mov => {
        if (mov.solicitante) {
            retiradasPorSolicitante[mov.solicitante] = (retiradasPorSolicitante[mov.solicitante] || 0) + 1;
        }
    });

    Object.entries(retiradasPorSolicitante).forEach(([solicitante, quantidade]) => {
        analiseData.push(['Gráfico de Pizza', 'Retiradas por Solicitante', quantidade, totalRetiradas > 0 ? ((quantidade / totalRetiradas) * 100).toFixed(1) + '%' : '0%', solicitante]);
    });

    // Dados para gráfico de pizza: Status das ferramentas
    analiseData.push(['Gráfico de Pizza', 'Status das Ferramentas', totalFerramentas - totalEmUso - totalQuebradas, totalFerramentas > 0 ? (((totalFerramentas - totalEmUso - totalQuebradas) / totalFerramentas) * 100).toFixed(1) + '%' : '0%', 'Disponíveis']);
    analiseData.push(['Gráfico de Pizza', 'Status das Ferramentas', totalEmUso, totalFerramentas > 0 ? ((totalEmUso / totalFerramentas) * 100).toFixed(1) + '%' : '0%', 'Em Uso']);
    analiseData.push(['Gráfico de Pizza', 'Status das Ferramentas', totalQuebradas, totalFerramentas > 0 ? ((totalQuebradas / totalFerramentas) * 100).toFixed(1) + '%' : '0%', 'Quebradas']);

    // Dados para gráfico de barras: Retiradas por mês
    const retiradasPorMes = {};
    movimentacoes.forEach(mov => {
        if (mov.dataRetirada) {
            const data = new Date(mov.dataRetirada);
            const mesAno = `${data.getMonth() + 1}/${data.getFullYear()}`;
            retiradasPorMes[mesAno] = (retiradasPorMes[mesAno] || 0) + 1;
        }
    });

    Object.entries(retiradasPorMes).forEach(([mesAno, quantidade]) => {
        analiseData.push(['Gráfico de Barras', 'Retiradas por Mês', quantidade, '', mesAno]);
    });

    const wsAnalise = XLSX.utils.aoa_to_sheet([analiseHeaders, ...analiseData]);
    applyBasicStyling(wsAnalise, analiseHeaders, 'Análise Gráfica e KPIs');
    XLSX.utils.book_append_sheet(wb, wsAnalise, 'Análise Gráfica');

    return wb;
}

// Função global para baixar Excel completo
window.baixarExcelCompleto = async function() {
    try {
        // Gerar workbook sempre com dados atuais (sem cache)
        const wb = await gerarWorkbookCompleto();

        // Baixar o arquivo com timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
        XLSX.writeFile(wb, `controle_ferramentas_${timestamp}.xlsx`, { cellStyles: true });

    } catch (error) {
        console.error('Erro ao gerar Excel:', error);
        alert('Erro ao gerar o arquivo Excel. Verifique o console para mais detalhes.');
    }
};

// Função global para atualizar dashboard
async function atualizarDashboard() {
    try {
        const ferramentas = await window.dbManager.ferramentas.obterTodas();
        const movimentacoes = await window.dbManager.movimentacoes.obterTodas();

        // Ferramentas no estoque: soma dos patrimônios disponíveis
        const estoqueCount = ferramentas.reduce((acc, f) => acc + (f.patrimonios ? f.patrimonios.length : 0), 0);

        // Ferramentas em uso: movimentações ativas com retorno esperado e sem retorno registrado
        const usoCount = movimentacoes.filter(mov => mov.tipo !== 'quebrada' && mov.temRetorno === 'sim' && (!mov.dataRetorno || mov.dataRetorno === '')).length;

        // Ferramentas quebradas: movimentações do tipo 'quebrada'
        const quebradasCount = movimentacoes.filter(mov => mov.tipo === 'quebrada').length;

        document.getElementById('estoqueCount').textContent = estoqueCount;
        document.getElementById('usoCount').textContent = usoCount;
        document.getElementById('quebradasCount').textContent = quebradasCount;
    } catch (error) {
        console.error('Erro ao atualizar dashboard:', error);
    }
}

// ===================================================================================
// === LÓGICA PARA A PÁGINA DO FORMULÁRIO (index.html) ===============================
// ===================================================================================
async function inicializarPaginaFormulario() {
    console.log("Modo Formulário Ativado.");

    // --- ELEMENTOS DO FORMULÁRIO ---
    const form = document.getElementById('ferramentaForm');
    const mensagemEl = document.getElementById('mensagem');
    const ferramentaSelect = document.getElementById('ferramenta');
    const patrimonioSelect = document.getElementById('patrimonio');
    const solicitanteSelect = document.getElementById('solicitante');
    const projetoSelect = document.getElementById('projeto');

    // Controles de devolução
    const btnDevolucaoHoras = document.getElementById('btnDevolucaoHoras');
    const btnDevolucaoDias = document.getElementById('btnDevolucaoDias');
    const devolucaoHorasContainer = document.getElementById('devolucaoHorasContainer');
    const devolucaoDiasContainer = document.getElementById('devolucaoDiasContainer');
    const dataDevolucaoInput = document.getElementById('dataDevolucao');
    const horaDevolucaoInput = document.getElementById('horaDevolucao');



    // --- ESTADO DA APLICAÇÃO ---
    const state = {
        ferramentas: [],
        solicitantes: []
    };

    // --- FUNÇÕES DE DADOS E UI ---
    async function carregarDados() {
        state.ferramentas = await window.dbManager.ferramentas.obterTodas();
        state.solicitantes = await window.dbManager.solicitantes.obterTodos();
        // Add default solicitantes if they don't already exist
        const defaultSolicitantes = [
            { nome: "BRUNO GOMES DA SILVA" },
            { nome: "JOSÉ ADRIANO DE SIQUEIRA ARAÚJO" },
            { nome: "MANUEL PEREIRA ALENCAR JUNIOR" },
            { nome: "NEUSVALDO NOVAIS RODRIGUES" },
            { nome: "RONALDO GONÇALVES DA SILVA" },
            { nome: "TIAGO FELIPE DOS SANTOS COELHO" },
            { nome: "DENILSON DE SOUZA SANTOS" },
            { nome: "FELIPE DE LIMA PEREIRA" },
            { nome: "JOSÉ CARLOS FIGUEIRA DA SILVA" },
            { nome: "JOSÉ GENILSON MARTINS SOARES" },
            { nome: "NETANIS DOS SANTOS" },
            { nome: "THIAGO PACHECO ALMEIDA" },
            { nome: "ROBERTO CARLOS DA SILVA" },
            { nome: "WESLEY ALEKSANDER ALCANTI DA SILVA" },
            { nome: "VALDEMIRO GOMES JUNIOR" }
        ];

        const existingNames = state.solicitantes.map(s => s.nome.toUpperCase());
        for (const solicitante of defaultSolicitantes) {
            if (!existingNames.includes(solicitante.nome.toUpperCase())) {
                await window.dbManager.solicitantes.adicionar(solicitante);
            }
        }
        state.solicitantes = await window.dbManager.solicitantes.obterTodos();
        popularSelects();
    }

    async function popularSelects() {
        if (ferramentaSelect) {
            ferramentaSelect.innerHTML = '<option value="" disabled selected>Selecione uma ferramenta</option>';
            const ferramentasAgrupadas = {};
            state.ferramentas.forEach(f => {
                if (f.patrimonios && f.patrimonios.length > 0) {
                    if (!ferramentasAgrupadas[f.nome]) {
                        ferramentasAgrupadas[f.nome] = { total: 0, ids: [] };
                    }
                    ferramentasAgrupadas[f.nome].total += f.patrimonios.length;
                    ferramentasAgrupadas[f.nome].ids.push(f.id);
                }
            });
            Object.keys(ferramentasAgrupadas).forEach(nome => {
                const total = ferramentasAgrupadas[nome].total;
                ferramentaSelect.add(new Option(`${nome} (${total} disponíveis)`, nome));
            });

            atualizarPatrimonios();
        }

        if (solicitanteSelect) {
            solicitanteSelect.innerHTML = '<option value="" disabled selected>Selecione um solicitante</option>';
            state.solicitantes.forEach(s => {
                solicitanteSelect.add(new Option(s.nome, s.nome));
            });
        }

        if (projetoSelect) {
            projetoSelect.innerHTML = '<option value="" selected>Nenhum projeto específico</option>';
        }
    }

    function atualizarPatrimonios() {
        const nomeFerramenta = ferramentaSelect.value;
        patrimonioSelect.innerHTML = '<option value="">Qualquer unidade disponível</option>';
        const ferramentasComNome = state.ferramentas.filter(f => f.nome === nomeFerramenta);
        ferramentasComNome.forEach(ferramenta => {
            if (ferramenta.patrimonios) {
                ferramenta.patrimonios.forEach(p => {
                    patrimonioSelect.add(new Option(p, p));
                });
            }
        });
    }

    // --- EVENT LISTENERS ---
    ferramentaSelect.addEventListener('change', () => {
        atualizarPatrimonios();
    });

    btnDevolucaoHoras.addEventListener('click', () => {
        devolucaoHorasContainer.style.display = 'block';
        devolucaoDiasContainer.style.display = 'none';
        dataDevolucaoInput.value = '';
        btnDevolucaoHoras.classList.add('active');
        btnDevolucaoDias.classList.remove('active');
    });

    btnDevolucaoDias.addEventListener('click', () => {
        devolucaoHorasContainer.style.display = 'none';
        devolucaoDiasContainer.style.display = 'block';
        horaDevolucaoInput.value = '';
        btnDevolucaoDias.classList.add('active');
        btnDevolucaoHoras.classList.remove('active');
    });

    document.getElementById('submitBtn').addEventListener('click', async () => {
        const submitBtn = document.getElementById('submitBtn');
        const originalText = submitBtn.textContent;

        const formData = new FormData(form);
        const data = Object.fromEntries(formData.entries());

        if (!data.solicitante || !data.ferramenta || !data.dataRetirada || !data.horaRetirada) {
            mensagemEl.textContent = 'Preencha todos os campos obrigatórios.';
            mensagemEl.className = 'mt-4 p-4 rounded-lg text-center font-semibold error-message';
            return;
        }

        // Validate return fields based on toggle
        if (btnDevolucaoHoras.classList.contains('active') && !data.horaDevolucao) {
            mensagemEl.textContent = 'Preencha a hora de devolução.';
            mensagemEl.className = 'mt-4 p-4 rounded-lg text-center font-semibold error-message';
            return;
        }
        if (btnDevolucaoDias.classList.contains('active') && !data.dataDevolucao) {
            mensagemEl.textContent = 'Preencha a data de devolução.';
            mensagemEl.className = 'mt-4 p-4 rounded-lg text-center font-semibold error-message';
            return;
        }

        // Prepare dataRetirada and dataDevolucao
        data.dataRetirada = data.dataRetirada + ' ' + data.horaRetirada;
        if (btnDevolucaoHoras.classList.contains('active')) {
            data.dataDevolucao = data.dataRetirada.split(' ')[0] + ' ' + data.horaDevolucao;
        }
        // For days, data.dataDevolucao is already the date

        // Disable button and set to "Enviando..."
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Enviando...';

        // Save locally first
        data.temRetorno = 'sim'; // Assume que retiradas têm retorno esperado
        const movimentacao = { ...data, tipo: 'retirada', dataRegistro: new Date().toISOString() };
        await window.dbManager.movimentacoes.adicionar(movimentacao);

        if(data.patrimonio) {
            const ferramentasComNome = state.ferramentas.filter(f => f.nome === data.ferramenta);
            for (const ferramenta of ferramentasComNome) {
                const index = ferramenta.patrimonios.indexOf(data.patrimonio);
                if (index > -1) {
                    ferramenta.patrimonios.splice(index, 1);
                    await window.dbManager.ferramentas.atualizar(ferramenta.id, ferramenta);
                    break; // Remove from the first one found
                }
            }
        }

        // Try to send to Google Sheets using doPost
        let enviadoParaPlanilha = false;
        try {
            // Format data according to Google Apps Script expectations
            const payload = {
                tipo: "movimentacao",
                solicitante: data.solicitante,
                ferramenta: data.ferramenta,
                dataSaida: data.dataRetirada.split(' ')[0], // date part only
                dataRetorno: data.dataDevolucao ? data.dataDevolucao.split(' ')[0] : "",
                horaDevolucao: btnDevolucaoHoras.classList.contains('active') ? data.horaDevolucao : "",
                temRetorno: data.temRetorno,
                observacoes: data.observacoes || ""
            };

            console.log("Sending payload to Google Sheets:", payload);

            const result = await doPost(payload);
            console.log('Dados enviados para Google Sheets:', result);
            enviadoParaPlanilha = true;
        } catch (error) {
            console.error("Erro ao enviar para Google Sheets (continuando localmente):", error);
            // Continue, since local save worked
        }

        // Update dashboard with new data
        await atualizarDashboard();

        // Change to "Enviado"
        submitBtn.innerHTML = '<i class="fas fa-check mr-2"></i>Enviado';

        if (enviadoParaPlanilha) {
            mensagemEl.textContent = 'Solicitação enviada com sucesso para a planilha!';
            mensagemEl.className = 'mt-4 p-4 rounded-lg text-center font-semibold success-message';
        } else {
            mensagemEl.textContent = 'Solicitação salva localmente. Verifique a conexão com a planilha.';
            mensagemEl.className = 'mt-4 p-4 rounded-lg text-center font-semibold warning-message';
        }

        // After 2 seconds, reset form and button
        setTimeout(() => {
            form.reset();
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="fas fa-paper-plane mr-2"></i>Enviar Solicitação';
            mensagemEl.textContent = '';
        }, 2000);

        await carregarDados();
    });

    // --- INICIALIZAÇÃO ---
    try {
        console.log('Iniciando aplicação...');
        const dbInitResult = await window.dbManager.init();
        if (!dbInitResult) {
            console.warn('Banco de dados inicializado com avisos, mas continuando...');
        }
        await carregarDados();
        await atualizarDashboard();
        console.log('Aplicação inicializada com sucesso');
    } catch (error) {
        console.error('Falha ao iniciar ou carregar dados:', error);
        mensagemEl.textContent = 'Erro ao carregar dados. Verifique o console e tente recarregar a página.';
        mensagemEl.className = 'mt-4 p-4 rounded-lg text-center font-semibold error-message';
        // Não travar a aplicação, permitir uso limitado
    }
}


// ===================================================================================
// === LÓGICA PARA A PÁGINA DO PAINEL (painel.html) ==================================
// ===================================================================================
async function inicializarPaginaPainel() {
    console.log("Modo Painel Ativado.");

    // --- ELEMENTOS DO PAINEL ---
    const estoqueDiv = document.getElementById('estoqueDiv');
    const listaSolicitantes = document.getElementById('listaSolicitantes');
    const mensagemLogin = document.getElementById('mensagemLogin');
    const loginContainer = document.getElementById('login');
    const painelContainer = document.getElementById('painel');
    const formAddFerramenta = document.getElementById('formAddFerramenta');
    const devolucaoForm = document.getElementById('devolucaoForm');
    const ferramentaDevolucaoSelect = document.getElementById('ferramentaDevolucaoSelect');
    const patrimonioDevolucaoSelect = document.getElementById('patrimonioDevolucaoSelect');
    const feedbackDevolucao = document.getElementById('feedbackDevolucao');

    // --- ESTADO DA APLICAÇÃO ---
    const state = {
        ferramentas: [],
        solicitantes: []
    };

    // --- FUNÇÕES DE RENDERIZAÇÃO ---
    function renderFerramentas() {
        estoqueDiv.innerHTML = '';
        if (state.ferramentas.length === 0) {
            estoqueDiv.innerHTML = '<p class="text-white/70 text-center col-span-full">Nenhuma ferramenta cadastrada.</p>';
            return;
        }
        state.ferramentas.forEach(item => {
            const card = document.createElement('div');
            card.className = 'tool-card rounded-xl p-4 card-hover';
            card.innerHTML = `
                <div class="flex justify-between items-start">
                    <h3 class="text-lg font-semibold gradient-text">${item.nome}</h3>
                    <span class="bg-purple-100 text-purple-800 text-xs font-medium px-2.5 py-0.5 rounded-full">Qtd: ${item.patrimonios.length}</span>
                </div>
                <div class="mt-3 text-sm text-gray-600">
                    <p class="mb-2"><i class="fas fa-hashtag mr-1"></i>Patrimônios:</p>
                    <div class="flex flex-wrap gap-1">${item.patrimonios.map(p => `<span class="bg-gray-100 text-gray-800 text-xs px-2 py-1 rounded">${p}</span>`).join('')}</div>
                </div>
                <button class="mt-4 btn-danger text-white text-sm px-3 py-1 rounded-lg delete-tool-btn" data-id-ferramenta="${item.id}" data-nome-ferramenta="${item.nome}">
                    <i class="fas fa-trash-alt mr-1"></i>Remover
                </button>
            `;
            estoqueDiv.appendChild(card);
        });
    }

    // --- FUNÇÕES DE DADOS ---
    async function carregarDados() {
        state.ferramentas = await window.dbManager.ferramentas.obterTodas();
        renderFerramentas();
    }

    // --- MODAIS (precisam ser globais por causa do onclick no HTML) ---
    window.abrirModal = (id) => document.getElementById(id).classList.remove('hidden');
    window.fecharModais = () => {
        document.querySelectorAll('.modal-overlay').forEach(modal => modal.classList.add('hidden'));
        formAddFerramenta.reset();
        formAddProjeto.reset();
    };

    // Função para abrir modal de exclusão de ferramenta
    window.abrirModalExclusao = (id, nome) => {
        window.ferramentaParaExcluir = { id, nome };
        document.getElementById('modalExcluirFerramenta').classList.remove('hidden');
    };

    // Função para abrir modal de seleção de patrimônio
    window.abrirModalSelecionarPatrimonio = async () => {
        const ferramenta = window.ferramentaParaExcluir;
        const ferramentas = await window.dbManager.ferramentas.obterTodas();
        const ferramentaObj = ferramentas.find(f => f.id === ferramenta.id);
        const select = document.getElementById('selectPatrimonioExcluir');
        select.innerHTML = '<option value="">Selecione um patrimônio</option>';
        if (ferramentaObj && ferramentaObj.patrimonios) {
            ferramentaObj.patrimonios.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p;
                opt.textContent = p;
                select.appendChild(opt);
            });
        }
        document.getElementById('modalExcluirFerramenta').classList.add('hidden');
        document.getElementById('modalSelecionarPatrimonio').classList.remove('hidden');
    };

    // Event listeners para os botões de exclusão
    document.getElementById('btnExcluirTodaFerramenta').addEventListener('click', async () => {
        const ferramenta = window.ferramentaParaExcluir;
        if (confirm(`Tem certeza que deseja excluir toda a ferramenta "${ferramenta.nome}"?`)) {
            await window.dbManager.ferramentas.remover(ferramenta.id);
            await carregarDados();
            window.fecharModais();
        }
    });

    document.getElementById('btnExcluirPatrimonio').addEventListener('click', () => {
        window.abrirModalSelecionarPatrimonio();
    });

    document.getElementById('btnConfirmarExcluirPatrimonio').addEventListener('click', async () => {
        const ferramenta = window.ferramentaParaExcluir;
        const patrimonio = document.getElementById('selectPatrimonioExcluir').value;
        if (!patrimonio) {
            alert('Selecione um patrimônio para excluir.');
            return;
        }
        if (confirm(`Tem certeza que deseja excluir o patrimônio "${patrimonio}" da ferramenta "${ferramenta.nome}"?`)) {
            const ferramentas = await window.dbManager.ferramentas.obterTodas();
            const ferramentaObj = ferramentas.find(f => f.id === ferramenta.id);
            if (ferramentaObj) {
                ferramentaObj.patrimonios = ferramentaObj.patrimonios.filter(p => p !== patrimonio);
                await window.dbManager.ferramentas.atualizar(ferramenta.id, ferramentaObj);
                await carregarDados();
                window.fecharModais();
            }
        }
    });

    // --- EVENT LISTENERS ---
    document.getElementById('btnLogin').addEventListener('click', async () => {
        const user = document.getElementById('usuario').value.trim();
        const pass = document.getElementById('senha').value.trim();
        if (user === "admin" && pass === "1234") {
            loginContainer.classList.add('hidden');
            painelContainer.classList.remove('hidden');
            await carregarDados();
        } else {
            mensagemLogin.textContent = "Usuário ou senha inválidos!";
        }
    });

    formAddFerramenta.addEventListener('submit', async (e) => {
        e.preventDefault();
        const bulkText = document.getElementById('ferramentasBulk').value.trim();
        if (!bulkText) {
            alert("Preencha as ferramentas no formato correto.");
            return;
        }
        const lines = bulkText.split('\n').map(line => line.trim()).filter(line => line);
        for (const line of lines) {
            const parts = line.split(':');
            if (parts.length !== 2) {
                alert(`Formato inválido na linha: ${line}. Use: NomeFerramenta:patrimonio1,patrimonio2`);
                return;
            }
            const nome = parts[0].trim();
            const patrimonios = parts[1].split(',').map(p => p.trim()).filter(p => p);
            if (!nome || patrimonios.length === 0) {
                alert(`Dados inválidos na linha: ${line}`);
                return;
            }
            // Check if tool already exists
            const ferramentasExistentes = await window.dbManager.ferramentas.obterTodas();
            if (ferramentasExistentes.some(f => f.nome.toLowerCase() === nome.toLowerCase())) {
                alert(`Ferramenta "${nome}" já existe.`);
                return;
            }
            await window.dbManager.ferramentas.adicionar({ nome, patrimonios });
        }
        await carregarDados();
        window.fecharModais();
    });

    formAddProjeto.addEventListener('submit', async (e) => {
        e.preventDefault();
        const bulkText = document.getElementById('nomesProjetos').value.trim();
        if (!bulkText) {
            alert("Preencha os nomes dos projetos.");
            return;
        }
        const nomes = bulkText.split('\n').map(line => line.trim()).filter(line => line);
        for (const nome of nomes) {
            if (!nome) continue;
            // Check if project already exists
            const projetosExistentes = await window.dbManager.projetos.obterTodos();
            if (projetosExistentes.some(p => p.nome.toLowerCase() === nome.toLowerCase())) {
                alert(`Projeto "${nome}" já existe.`);
                return;
            }
            await window.dbManager.projetos.adicionar({ nome });
        }
        await carregarDados();
        window.fecharModais();
    });

    document.body.addEventListener('click', async (e) => {
        const button = e.target.closest('button');
        if (!button) return;

        const idProjeto = button.dataset.idProjeto;
        if (idProjeto) {
             if (confirm(`Tem certeza que deseja excluir o projeto "${button.dataset.nomeProjeto}"?`)) {
                await window.dbManager.projetos.remover(idProjeto);
                await carregarDados();
            }
        }

        // Handle delete tool button
        if (button.classList.contains('delete-tool-btn')) {
            const idFerramenta = button.dataset.idFerramenta;
            const nomeFerramenta = button.dataset.nomeFerramenta;
            if (idFerramenta && nomeFerramenta) {
                window.abrirModalExclusao(idFerramenta, nomeFerramenta);
            }
        }
    });

    // --- INICIALIZAÇÃO DO PAINEL ---
    try {
        await window.dbManager.init();
        if (document.getElementById('estoqueCount')) {
            await atualizarDashboard();
        }
    } catch (error) {
        console.error('Falha fatal ao iniciar o DB:', error);
        alert('Não foi possível carregar o banco de dados. A aplicação não pode continuar.');
    }

    // Funções de download
    function baixarRetiradas() {
      const table = document.querySelector('#tabelaRetiradas table');
      if (!table) return;
      const wb = XLSX.utils.table_to_book(table);
      XLSX.writeFile(wb, 'retiradas.xlsx');
    }

    function baixarDevolucoes() {
      const table = document.querySelector('#tabelaDevolucoes table');
      if (!table) return;
      const wb = XLSX.utils.table_to_book(table);
      XLSX.writeFile(wb, 'devolucoes.xlsx');
    }

    function baixarQuebradas() {
      const table = document.querySelector('#tabelaQuebradas table');
      if (!table) return;
      const wb = XLSX.utils.table_to_book(table);
      XLSX.writeFile(wb, 'ferramentas_quebradas.xlsx');
    }
}
