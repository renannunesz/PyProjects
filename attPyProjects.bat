@echo off
setlocal

set REPO_DIR=C:\PyProjects
set REPO_URL=https://github.com/renannunesz/PyProjects.git

:: Se a pasta existir, remove tudo
if exist "%REPO_DIR%" (
    echo Apagando conteudo existente em %REPO_DIR%...
    rmdir /s /q "%REPO_DIR%"
)

:: Clona o repositório remoto
echo Clonando repositório...
git clone %REPO_URL% "%REPO_DIR%"

if %ERRORLEVEL%==0 (
    echo Clonagem concluida com sucesso!
) else (
    echo Erro ao clonar o repositório!
)

pause
