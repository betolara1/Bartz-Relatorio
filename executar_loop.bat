@echo off
:: Muda para a pasta onde o script está localizado
cd /d "%~dp0"

:loop
echo [%date% %time%] Iniciando aplicacao...

:: Substitua pelo comando que roda o seu codigo:
:: Exemplo sem Java JAR:
java -cp "libs/*;." Main

:: Exemplo com Java JAR:
:: java -cp "libs/*;." Main


:: Ou exemplo com Maven / Java direto:
:: java -cp ... com.bartz.ui.TelaPrincipal

:: Se o programa fechar ou cair por erro, espera 5 segundos e reinicia:
timeout /t 5 /nobreak >nul
goto loop
