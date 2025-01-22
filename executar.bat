@echo off
Title: INSTALACAO DE DEPENDENCIAS...
python --version 2>nul | findstr /i "3.13" >nul
if errorlevel 1 (
    echo Python 3.13 não encontrado. Baixando Python 3.13...
    curl -o python-installer.exe https://www.python.org/ftp/python/3.13.1/python-3.13.1-amd64.exe
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
)
python --version 2>nul | findstr /i "3.13" >nul
if errorlevel 1 (
    echo Falha na instalação do Python 3.13. Por favor, instale manualmente.
    pause
    exit /b
)
echo Instalando bibliotecas para automacao...
python -m pip install -r requirements.txt
echo Instalacao concluida!
pause

