# Sistema de Controle de Ferramentas - Kure Fleximedical

Sistema web completo para controle de retirada, devolução e gerenciamento de ferramentas, desenvolvido para a empresa Kure Fleximedical.

## 🚀 Funcionalidades

### Página Principal (index.html)
- ✅ Formulário de retirada de ferramentas
- ✅ Seleção de solicitante e ferramenta
- ✅ Controle de datas e horários de retirada/devolução
- ✅ Dashboard com estatísticas em tempo real
- ✅ Salvamento automático no banco local
- ✅ Integração com Google Sheets (opcional)

### Painel Administrativo (painel.html)
- ✅ Login seguro (admin/1234)
- ✅ Cadastro de ferramentas com patrimônios

- ✅ Sistema de devoluções
- ✅ Visualização de planilhas Excel
- ✅ Relatórios com filtros e KPIs

### Relatórios Excel
- ✅ Geração automática de Excel completo
- ✅ Múltiplas abas: Retiradas, Devoluções, Quebradas, Dashboard
- ✅ KPIs calculados automaticamente
- ✅ Formatação profissional com cores e estilos

## 📋 Pré-requisitos

- Navegador web moderno (Chrome, Firefox, Edge)
- Python 3.7+ (para executar o servidor local)
- Conexão com internet (opcional, para Google Sheets)

## 🛠️ Instalação e Execução

### Método 1: Arquivo Batch (Recomendado)
1. Clique duas vezes no arquivo `start_app.bat`
2. O sistema instalará automaticamente as dependências
3. O navegador abrirá automaticamente

### Método 2: Manual
1. Instale o Python 3.7+ se não tiver
2. Abra o terminal/cmd na pasta raiz do projeto
3. Execute: `pip install pandas openpyxl`
4. Entre na pasta hmtl: `cd hmtl`
5. Execute: `python -m http.server 8000`
6. Abra o navegador em: `http://localhost:8000/index.html`

## 📁 Estrutura do Projeto

```
/
├── start_app.bat          # Inicializador automático
├── hmtl/
│   ├── index.html         # Página principal (formulário)
│   ├── painel.html        # Painel administrativo
│   ├── script.js          # Lógica JavaScript unificada
│   ├── database.js        # Gerenciamento do banco localStorage
│   ├── config.js          # Configurações da aplicação
│   ├── style.css          # Estilos CSS
│   ├── gerar_excel.py     # Script de geração de relatórios Excel
│   ├── requirements.txt   # Dependências Python
│   └── README.md          # Este arquivo
```

## 🔧 Configuração

### Google Sheets (Opcional)
Para integrar com Google Sheets, configure a URL da API no arquivo `config.js`:

```javascript
const CONFIG = {
    API_URL: "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec",
    // ... outras configurações
};
```

### Credenciais de Acesso
- **Painel Admin**: usuário `admin`, senha `1234`

## 📊 Funcionalidades dos Relatórios

### Excel Completo (`relatorio_completo_kure_*.xlsx`)
- **Aba Retiradas**: Todas as solicitações registradas
- **Aba Devoluções**: Histórico de devoluções
- **Aba Ferramentas Quebradas**: Equipamentos danificados
- **Aba Dashboard**: KPIs e estatísticas

### KPIs Calculados
- Total de retiradas
- Total de devoluções
- Ferramentas quebradas
- Taxa de devolução
- Taxa de quebra
- Ferramentas mais utilizadas

- Solicitantes mais ativos

## 🔒 Segurança

- Dados salvos localmente no navegador (localStorage)
- Login seguro para painel administrativo
- Timeout automático de sessão
- Validação de formulários

## 🐛 Solução de Problemas

### Botão "Enviar" não funciona
1. Verifique se todos os campos obrigatórios estão preenchidos
2. Recarregue a página (F5)
3. Verifique o console do navegador (F12)

### Erro ao gerar Excel
1. Execute: `pip install pandas openpyxl`
2. Certifique-se de estar na pasta `hmtl`
3. Execute: `python gerar_excel.py`

### Servidor não inicia
1. Verifique se a porta 8000 está livre
2. Execute como administrador se necessário
3. Use `start_app.bat` para instalação automática

## 📞 Suporte

Para suporte técnico ou dúvidas:
- Verifique os logs no console do navegador (F12)
- Execute o script de diagnóstico: `python teste_webapp.py`

## 📝 Licença

Este projeto é propriedade da Kure Fleximedical.

---

**Desenvolvido para otimizar o controle de ferramentas e aumentar a produtividade da equipe.**
