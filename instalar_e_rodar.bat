@echo off
chcp 65001 >nul
echo ============================================
echo  Ferramenta de Implantação Contábil - BHub
echo ============================================
echo.

REM Caminho do Python instalado via winget
set PYTHON_PATH=%LOCALAPPDATA%\Programs\Python\Python313\python.exe

REM Verificar se Python existe nesse caminho
if not exist "%PYTHON_PATH%" (
    REM Tentar Python do PATH
    where python >nul 2>nul
    if %ERRORLEVEL% NEQ 0 (
        echo [ERRO] Python nao encontrado.
        echo Instale o Python em python.org ou via Microsoft Store.
        pause
        exit /b 1
    )
    set PYTHON_PATH=python
)

echo Python encontrado: %PYTHON_PATH%
echo.

echo [1/3] Criando ambiente virtual...
if not exist ".venv" (
    "%PYTHON_PATH%" -m venv .venv
)

echo [2/3] Instalando dependencias...
call .venv\Scripts\activate.bat
pip install -r requirements.txt --quiet --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org

echo [3/3] Iniciando a ferramenta...
echo.
echo ============================================
echo  Acesse no navegador: http://localhost:8501
echo  Para encerrar: pressione CTRL+C aqui
echo ============================================
echo.
streamlit run app.py --server.headless false --server.port 8501 --browser.gatherUsageStats false

pause
