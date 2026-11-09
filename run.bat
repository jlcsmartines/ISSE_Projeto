@echo off
echo Iniciando servidor ISSE...

call venv\Scripts\activate

REM Abre o frontend no navegador apos 2 segundos (tempo pro servidor subir)
start "" cmd /c "timeout /t 2 /nobreak >nul && start frontend\index.html"

REM Sobe o backend (fica rodando ate fechar a janela)
python backend\main.py

pause
