@echo off
cd /d "%~dp0"
python interface.py
if errorlevel 1 (
    echo Nao foi possivel iniciar o bot. Confira a mensagem acima.
    pause
)
