@echo off
chcp 65001 >nul
setlocal EnableExtensions

title Relatório SKUPDV em Python
cd /d "%~dp0"

set "PYTHON_CODEX=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%PYTHON_CODEX%" (
    set "PYTHON_EXE=%PYTHON_CODEX%"
) else (
    set "PYTHON_EXE=python"
)

"%PYTHON_EXE%" -X utf8 ".\gerar_relatorio_skupdv.py"
set "CODIGO_SAIDA=%ERRORLEVEL%"

echo.
if "%CODIGO_SAIDA%"=="0" (
    echo Relatório concluído com sucesso.
) else (
    echo O relatório não foi concluído. Leia a mensagem exibida acima.
)
echo.
pause
exit /b %CODIGO_SAIDA%
