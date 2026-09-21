@echo off
cd /d %~dp0
python -c "import docx" >nul 2>&1
if errorlevel 1 (
  echo Instalando dependencia python-docx...
  python -m pip install python-docx
)
python tools\atualizar_conteudo.py
if errorlevel 1 pause
