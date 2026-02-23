@echo off
:: Garante que o terminal entenda acentos (UTF-8)
chcp 65001 > nul

set NODE_TLS_REJECT_UNAUTHORIZED=0

:: Navega até a pasta do projeto
cd /d "C:\Caminho\Para\Seu\Diretorio\Sincronizado"

echo 🤖 Iniciando o Robô Auvo...
echo ------------------------------------------

:: Executa o python sem herdar privilégios elevados (se possível) 
:: ou simplesmente executa o comando padrão se já estiver em modo normal
python automacao_auvo.py

echo ------------------------------------------
echo ⚠️ O processo terminou.

:: Aguarda 5 segundos e fecha automaticamente
echo Fechando em 5 segundos...
timeout /t 5 /nobreak > nul
exit