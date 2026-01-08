@echo off
echo --- Configurando SARA PowerBI Server ---

if not exist "venv" (
    echo Criando ambiente virtual (venv)...
    python -m venv venv
)

echo Ativando venv e instalando dependencias...
call venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo --- INSTALACAO CONCLUIDA ---
echo Para rodar, use o arquivo 'run.bat' ou configure sua IDE apontando para:
echo %CD%\venv\Scripts\python.exe
echo.
pause
