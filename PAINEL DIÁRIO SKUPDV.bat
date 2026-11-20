@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion

rem ============================================================================
rem PAINEL DIÁRIO SKUPDV
rem Centraliza os três processos oficiais sem duplicar regras ou fontes.
rem Cada processo continua sendo responsável pelas próprias validações.
rem ============================================================================

title Painel Diário SKUPDV
set "RAIZ_SKUPDV=%~dp0"
set "GERADOR_RELATORIO=%RAIZ_SKUPDV%SKUPDV python\Executar Relatorio SKUPDV Python.cmd"
set "GERADOR_COMUNICADOS=%RAIZ_SKUPDV%Disparo SKUPDV aos RNs\scripts\Gerar_Comunicados_SKUPDV.bat"
set "PAINEL_DISPAROS=%RAIZ_SKUPDV%Disparo SKUPDV aos RNs\scripts\Abrir Painel de Disparos SKUPDV.cmd"

:MENU
cls
echo.
echo ============================================================================
echo                           PAINEL DIÁRIO SKUPDV
echo ============================================================================
echo.
echo  1 - Gerar o relatório diário de SKUPDV
echo  2 - Gerar comunicados diários de SKUPDV
echo  3 - Disparar comunicados diários de SKUPDV
echo.
echo  0 - Sair
echo.
set "OPCAO="
rem A variável abaixo existe apenas para teste automatizado do menu. No uso
rem normal ela não existe e a escolha é solicitada diretamente ao usuário.
if defined SKUPDV_OPCAO_AUTOMATICA (
    set "OPCAO=%SKUPDV_OPCAO_AUTOMATICA%"
) else (
    set /p "OPCAO=Escolha uma opção: "
)

if "%OPCAO%"=="1" (
    call :EXECUTAR "%GERADOR_RELATORIO%" "Gerador do relatório diário"
    goto MENU
)
if "%OPCAO%"=="2" (
    call :EXECUTAR "%GERADOR_COMUNICADOS%" "Gerador de comunicados"
    goto MENU
)
if "%OPCAO%"=="3" (
    call :EXECUTAR "%PAINEL_DISPAROS%" "Painel supervisionado de disparos"
    goto MENU
)
if "%OPCAO%"=="0" goto FIM

echo.
echo Opção inválida. Informe 1, 2, 3 ou 0.
pause
goto MENU

:EXECUTAR
rem Confere o ponto de entrada antes de iniciar qualquer uma das etapas.
if not exist "%~1" (
    echo.
    echo ERRO: %~2 não foi localizado.
    echo Caminho esperado: %~1
    echo.
    pause
    goto :eof
)

echo.
echo Iniciando: %~2
echo.
call "%~1"
set "CODIGO_SAIDA=%ERRORLEVEL%"

echo.
echo ============================================================================
if "%CODIGO_SAIDA%"=="0" (
    echo Processo finalizado. Consulte as mensagens exibidas pelo processo acima.
) else (
    echo Processo encerrado com código %CODIGO_SAIDA%. Leia a mensagem exibida acima.
)
echo ============================================================================
echo.
pause
goto :eof

:FIM
endlocal
exit /b 0
