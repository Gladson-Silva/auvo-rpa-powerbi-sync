# 🤖 RPA Data Pipeline: Auvo Desk to Power BI Cloud Sync

Este repositório apresenta uma solução de **RPA (Robotic Process Automation)** desenvolvida em Python para automatizar o ciclo completo de dados: extração de relatórios de um ERP de gestão, processamento local e sincronização nativa com dashboards no Power BI Service.

O projeto resolve o desafio de manter dashboards em nuvem atualizados a partir de fontes de dados que exigem navegação web complexa e recálculo de fórmulas em motores de planilha (Excel/WPS).

## 🌟 Funcionalidades e Diferenciais

* **Extração Automatizada (Playwright):** Navegação *headless* para autenticação, filtragem dinâmica de períodos e download de múltiplos relatórios trimestrais.
* **Merge Cirúrgico de Dados (Pandas):** Algoritmo que realiza o *upsert* de dados brutos exclusivamente no intervalo de colunas `A:AH`, garantindo a integridade de fórmulas complexas e KPIs personalizados localizados a partir da coluna `AI`.
* **Engine-Agnostic Recalculation (Pywin32):** Integração com a API COM do Windows para forçar o recálculo de fórmulas em segundo plano (suporta Microsoft Excel e WPS Office).
* **Sincronização Cloud-Native:** Arquitetura desenhada para operar via diretórios sincronizados (SharePoint/OneDrive), eliminando a necessidade de Gateways locais.
* **Resiliência Operacional:** Sistema de logs detalhados e travas de segurança que impedem a corrupção da base de dados histórica.

## ⚖️ Conformidade e LGPD

> **Aviso de Privacidade:** Em conformidade com a **LGPD (Lei Geral de Proteção de Dados)**, todas as credenciais de acesso, links de diretórios corporativos e nomes de empresas foram removidos ou substituídos por variáveis genéricas neste repositório para garantir a privacidade das informações.

## 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* **Playwright** (Navegação Web)
* **Pandas** (Tratamento de Dados)
* **Openpyxl** (Edição de .xlsx)
* **Pywin32** (Integração Windows COM)
* **Power BI Service** (Cloud Analytics)

## 📋 Pré-requisitos

Instale as dependências e os binários do navegador antes de executar o script:

```bash
pip install pandas playwright openpyxl inputimeout pywin32
playwright install chromium
```
## ⚙️ Como Configurar
Para adaptar este script ao seu ambiente, edite as variáveis no bloco de configuração do script Python:
```bash
# CONFIGURAÇÕES DE AMBIENTE (Substitua pelos seus dados)
USER = "seu_usuario@dominio.com"
PASSWORD = "sua_password_segura"
BASE_DIR = r"C:\Caminho\Para\Seu\Diretorio\Sincronizado"
NOME_ARQUIVO_MESTRE = "Seu_Relatorio_Geral.xlsx"
```
## 🚀 Arquitetura do Fluxo
1. **Extraction:** O bot realiza login e extrai dados de forma assíncrona.
2. **Transformation:** O script limpa o range de dados antigos e injeta os novos registros, preservando a estrutura de colunas calculadas.
3. **Validation:** O motor de planilha é acionado de forma invisível via pywin32 para validar fórmulas e garantir que o Power BI receba dados calculados.
4. **Loading:** O arquivo é salvo no diretório sincronizado e o Power BI Service atualiza os visuais automaticamente via conexão Web.

## 🕒 Automação e Execução

1. Execução via Arquivo de Lote (.bat)
Para que usuários sem conhecimento técnico em Python possam operar o script, utilize um arquivo .bat:

```bash
@echo off
:: Garante que o terminal entenda acentos (UTF-8)
chcp 65001 > nul

set NODE_TLS_REJECT_UNAUTHORIZED=0

:: Navega até a pasta do projeto
cd /d "C:\Users\Dell\Desktop\Automação"

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
```
## 2. Agendador de Tarefas do Windows
Para automação total (sem cliques):

1. No Agendador de Tarefas, crie uma Tarefa Básica.
2. Defina o disparador como Diário e escolha o horário (ex: 07:00).
3. Na ação Iniciar um programa, selecione o seu arquivo .bat.
4. Garanta que o PC esteja ligado ou em modo de espera no horário definido.

3. Script de Preparação de Pastas
Execute o código abaixo em um arquivo .bat para criar automaticamente a estrutura de diretórios necessária:

```bash
@echo off
setlocal
title Configurador de Estrutura - Automacao RPA

:: Localiza automaticamente a pasta Documentos do usuário atual
set "ROOT_DIR=%USERPROFILE%\Documents\AUTOMACAO"

echo ======================================================
echo    PREPARANDO AMBIENTE PARA O ROBÔ DE DADOS
echo ======================================================
echo.

:: Cria a pasta principal
if not exist "%ROOT_DIR%" (
    mkdir "%ROOT_DIR%"
    echo [+] Pasta PRINCIPAL criada em: %ROOT_DIR%
) else (
    echo [!] A pasta PRINCIPAL ja existe.
)

:: Cria a subpasta para os arquivos do Auvo
if not exist "%ROOT_DIR%\downloads_temporarios" (
    mkdir "%ROOT_DIR%\downloads_temporarios"
    echo [+] Subpasta DOWNLOADS_TEMPORARIOS criada.
) else (
    echo [!] A subpasta DOWNLOADS_TEMPORARIOS ja existe.
)

echo.
echo ======================================================
echo    ESTRUTURA PRONTA! COLOQUE O SCRIPT PYTHON NA PASTA:
echo    %ROOT_DIR%
echo ======================================================
echo.
pause
```

## 💡 Dicas de Manutenção e Solução de Problemas
* Power BI Desktop: Mantenha o arquivo .pbix fechado durante a execução do script para evitar erros de permissão de escrita no Excel.
* Credenciais na Nuvem: Ao publicar o relatório, configure as credenciais no Power BI Service usando o método OAuth2 e nível de privacidade Organizacional para fontes Web/SharePoint.
* Cache Local: Caso precise baixar dados novos após já ter rodado o script, basta não responder ao prompt de 10 segundos ou selecionar a opção de limpeza no terminal.
* Timeouts: Se o site de origem estiver lento, ajuste o tempo de espera nas funções do Playwright dentro do script.



⭐ Desenvolvido para automação de processos e eficiência operacional.
