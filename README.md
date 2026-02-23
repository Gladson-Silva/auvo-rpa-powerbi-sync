# 🤖 RPA Data Pipeline: Auvo Desk to Power BI Cloud Sync

Este repositório apresenta uma solução de **RPA (Robotic Process Automation)** desenvolvida em Python para automatizar o ciclo completo de dados: extração de relatórios de um ERP de gestão (Auvo Desk), processamento local e sincronização nativa com dashboards no Power BI Service.

O projeto resolve o desafio de manter dashboards em nuvem atualizados a partir de fontes de dados que exigem navegação web complexa e recálculo de fórmulas em motores de planilha (Excel/WPS).



## 🌟 Funcionalidades e Diferenciais

- **Extração Automatizada (Playwright):** Navegação *headless* para autenticação, filtragem dinâmica de períodos e download de múltiplos relatórios trimestrais.
- **Merge Cirúrgico de Dados (Pandas):** Algoritmo que realiza o *upsert* de dados brutos exclusivamente no intervalo de colunas `A:AH`, garantindo a integridade de fórmulas complexas e KPIs personalizados localizados a partir da coluna `AI`.
- **Engine-Agnostic Recalculation (Pywin32):** Integração com a API COM do Windows para forçar o recálculo de fórmulas em segundo plano (suporta Microsoft Excel e WPS Office), essencial para que o Power BI Service leia metadados já processados.
- **Sincronização Cloud-Native:** Arquitetura desenhada para operar via diretórios sincronizados (SharePoint/OneDrive), permitindo a atualização automática do Power BI Web via conexão Web segura, eliminando a necessidade de Gateways locais.
- **Resiliência Operacional:** Sistema de logs detalhados e travas de segurança que impedem a corrupção da base de dados histórica em caso de instabilidade no serviço de origem.

## ⚖️ Conformidade e LGPD

> **Aviso de Privacidade:** Em conformidade com a **LGPD (Lei Geral de Proteção de Dados)**, todas as credenciais de acesso, links de diretórios corporativos, nomes de empresas e dados de clientes foram removidos ou substituídos por variáveis genéricas e *placeholders* neste repositório. O código fornecido é para fins de demonstração técnica de arquitetura de automação.

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**
- **Playwright** (Navegação e Automação Web)
- **Pandas** (Data Wrangling e Manipulação de DataFrames)
- **Openpyxl** (Edição de arquivos .xlsx)
- **Pywin32** (Interoperabilidade com Windows COM)
- **Power BI Service** (Cloud Analytics & Visualization)

## 📋 Pré-requisitos

1. **Instalar as dependências do projeto:**
   ```bash
   pip install pandas playwright openpyxl inputimeout pywin32
2. **Instalar os binários do navegador para o RPA:**
   ```bash
   playwright install chromium

## 📋 Como Configurar
