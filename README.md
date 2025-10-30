a Sistema de Controle de Ferramentas - Kure Fleximedical

Este é um sistema web para gerenciamento de ferramentas, desenvolvido para a empresa Kure Fleximedical. O sistema permite registrar retiradas de ferramentas, gerenciar estoque e movimentações.

## Funcionalidades

### Página de Formulário (index.html)
- Formulário para registro de retirada de ferramentas
- Seleção de solicitante, ferramenta e patrimônio
- Campos para datas e horários de retirada e devolução
- Toggle entre modo horas e dias
- Integração com IndexedDB para dados locais

### Painel Administrativo (painel.html)
- Login para acesso (usuário: admin, senha: 1234)
- Gerenciamento de ferramentas (adicionar, visualizar, excluir)

- Registro de movimentações (ferramentas quebradas)
- Notificações de movimentações
- Integração com Google Apps Script para sincronização de dados

## Arquivos do Sistema

- `index.html` - Página principal do formulário
- `painel.html` - Painel administrativo
- `database.js` - Gerenciamento do IndexedDB local
- `config.js` - Configurações do sistema (URLs, usuários)
- `script.js` - Lógica unificada para ambas as páginas
- `style.css` - Estilos CSS adicionais
- `teste_webapp.py` - Script de teste para API do Google Apps Script

## Como Executar

### Opção 1: Servidor Local Simples
1. Abra um terminal na pasta `hmtl`
2. Execute um servidor HTTP simples:
   - Python: `python -m http.server 8000`
   - Node.js: `npx http-server -p 8000`
   - PHP: `php -S localhost:8000`
3. Abra o navegador em `http://localhost:8000/index.html`

### Opção 2: Abrir Diretamente no Navegador
- Clique duplo nos arquivos HTML para abrir no navegador
- Nota: Alguns recursos podem não funcionar corretamente devido a restrições de CORS

## Estrutura de Dados

### IndexedDB
- **ferramentas**: Armazena informações das ferramentas (nome, patrimônios)

- **movimentacoes**: Registros de movimentações (quebradas, empréstimos)

### Google Apps Script
- Sincronização de dados com planilha Google Sheets
- Endpoint: `https://script.google.com/macros/s/AKfycbzka9zfxb9UcVz2kVafIWmiYT12YHx0JPb3zPU8jU1PN4BuNzBXeVUe1bMxxqG21b6O0A/exec`

## Desenvolvimento

### Tecnologias Utilizadas
- HTML5, CSS3, JavaScript ES6+
- IndexedDB para armazenamento local
- Tailwind CSS para estilização
- Font Awesome para ícones
- Google Apps Script para backend

### Estrutura do Código
- `database.js`: Funções para manipulação do IndexedDB
- `script.js`: Lógica de interface e integração
- `config.js`: Configurações globais

## Teste da API

Execute o script Python para testar a integração com o Google Apps Script:

```bash
python teste_webapp.py
```

## Notas de Segurança

- Credenciais de login hardcoded (apenas para demonstração)
- Dados sensíveis devem ser protegidos em produção
- Implementar autenticação adequada para ambiente real

## Suporte

Para dúvidas ou problemas, entre em contato com a equipe de desenvolvimento.
