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

## 💡 Dicas de Manutenção e Solução de Problemas
* Power BI Desktop: Mantenha o arquivo .pbix fechado durante a execução do script para evitar erros de permissão de escrita no Excel.
* Credenciais na Nuvem: Ao publicar o relatório, configure as credenciais no Power BI Service usando o método OAuth2 e nível de privacidade Organizacional para fontes Web/SharePoint.
* Timeouts: Se o site de origem estiver lento, ajuste o tempo de espera nas funções do Playwright dentro do script.



⭐ Desenvolvido para automação de processos e eficiência operacional.
