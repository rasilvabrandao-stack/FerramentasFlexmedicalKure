@echo off
echo ============================================
echo    SISTEMA DE CONTROLE DE FERRAMENTAS
echo           KURE FLEXIMEDICAL
echo ============================================
echo.

REM Verificar se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python não está instalado ou não está no PATH.
    echo Por favor, instale o Python e tente novamente.
    pause
    exit /b 1
)

REM Verificar se estamos no diretório correto
if not exist "hmtl" (
    echo ERRO: Pasta 'hmtl' não encontrada.
    echo Execute este arquivo no diretório raiz do projeto.
    pause
    exit /b 1
)

REM Entrar na pasta hmtl
cd hmtl

REM Instalar dependências se necessário
if not exist "requirements.txt" (
    echo Criando requirements.txt...
    echo pandas>=1.3.0 > requirements.txt
    echo openpyxl>=3.0.7 >> requirements.txt
)

echo Instalando dependências Python...
pip install -r requirements.txt

REM Verificar se as dependências foram instaladas
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo AVISO: Algumas dependências Python não puderam ser instaladas.
    echo O sistema funcionará, mas alguns recursos podem não estar disponíveis.
)

echo.
echo ============================================
echo    Iniciando servidor web...
echo ============================================
echo.
echo O sistema estará disponível em:
echo http://localhost:8000/index.html
echo.
echo Para acessar o painel administrativo:
echo http://localhost:8000/painel.html
echo.
echo Pressione Ctrl+C para parar o servidor.
echo.

REM Iniciar servidor personalizado sem cache
python server.py

pause
