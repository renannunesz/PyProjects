@echo off
cd C:\PyProjects

:: Descarta qualquer modificação local
git reset --hard

:: Garante que você tem os últimos dados do repositório remoto
git fetch --all

:: Força a sincronização com o repositório remoto
git reset --hard origin/main

if %ERRORLEVEL%==0 (
    echo Atualizacao concluida com sucesso!
) else (
    echo Houve um erro durante a atualizacao!
)
pause