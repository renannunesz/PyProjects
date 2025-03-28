@echo off
cd C:\PyProjects
git pull origin master
if %ERRORLEVEL%==0 (
    echo Atualização concluída com sucesso!
) else (
    echo Houve um erro durante a atualização.
)
pause