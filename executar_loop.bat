@echo off
:: ============================================================
:: Loop de execução do monitoramento de relatórios
:: Reinicia automaticamente se o Java fechar ou der erro
:: ============================================================

:: Muda para a pasta onde o script está localizado
cd /d "%~dp0"

:: --- Arquivo de log (salvo na mesma pasta do script) ---
set LOG=%~dp0relatorio_log.txt

:: --- Localiza o Java automaticamente ---
:: Tenta o Java do PATH primeiro; se não achar, usa caminhos comuns do servidor
where java >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set JAVA_EXE=java
) else if exist "C:\Program Files\Java\jdk-17\bin\java.exe" (
    set JAVA_EXE="C:\Program Files\Java\jdk-17\bin\java.exe"
) else if exist "C:\Program Files\Java\jdk-11\bin\java.exe" (
    set JAVA_EXE="C:\Program Files\Java\jdk-11\bin\java.exe"
) else if exist "C:\Program Files\Eclipse Adoptium\jdk-17\bin\java.exe" (
    set JAVA_EXE="C:\Program Files\Eclipse Adoptium\jdk-17\bin\java.exe"
) else (
    echo [%date% %time%] ERRO: Java nao encontrado! Verifique a instalacao. >> "%LOG%"
    echo ERRO: Java nao encontrado! Verifique se o Java esta instalado.
    timeout /t 30 /nobreak >nul
    goto loop
)

:loop
echo [%date% %time%] Iniciando monitoramento... >> "%LOG%"
echo [%date% %time%] Iniciando monitoramento...

:: Executa o programa Java
%JAVA_EXE% -cp "libs/*;." Main

:: Captura o código de saída
set EXIT_CODE=%ERRORLEVEL%
echo [%date% %time%] Programa encerrou (codigo: %EXIT_CODE%). Reiniciando em 5s... >> "%LOG%"
echo [%date% %time%] Programa encerrou (codigo: %EXIT_CODE%). Reiniciando em 5s...

:: Aguarda 5 segundos e reinicia automaticamente
timeout /t 5 /nobreak >nul
goto loop
