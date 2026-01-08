@echo off
set "TARGET_DIR=%USERPROFILE%\mcp-servers\sara-powerbi"

echo ===================================================
echo   INSTALADOR SARA POWER BI SERVER
echo ===================================================
echo.
echo Este script vai instalar o servidor em:
echo %TARGET_DIR%
echo.
pause

echo.
echo [1/4] Criando pasta de destino...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

echo.
echo [2/4] Copiando arquivos...
copy /Y "sara_mcp.py" "%TARGET_DIR%\" >nul
copy /Y "requirements.txt" "%TARGET_DIR%\" >nul
copy /Y "README.md" "%TARGET_DIR%\" >nul

echo.
echo [3/4] Criando Ambiente Virtual (isso pode levar 1 min)...
cd /d "%TARGET_DIR%"
if not exist "venv" python -m venv venv

echo.
echo [4/4] Instalando Dependencias...
call venv\Scripts\activate
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt >nul

echo.
echo ===================================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo ===================================================
echo.
echo O servidor foi instalado em: %TARGET_DIR%
echo.
echo --- CONFIGURACAO PARA SEU CLAUDE DESKTOP / IDE ---
echo Copie e cole o seguinte JSON no seu arquivo de configuracao:
echo.
echo {
echo   "mcpServers": {
echo     "sara-powerbi": {
echo       "command": "%TARGET_DIR:\=\\%\\venv\\Scripts\\python.exe",
echo       "args": ["-u", "%TARGET_DIR:\=\\%\\sara_mcp.py"]
echo     }
echo   }
echo }
echo.
echo ===================================================
echo Pressione qualquer tecla para abrir a pasta instalada...
pause >nul
start .
