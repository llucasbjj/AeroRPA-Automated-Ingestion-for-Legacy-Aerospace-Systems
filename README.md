# 🚀 AeroRPA: Automated Ingestion for Legacy Aerospace Systems

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge)
![OpenCV](https://img.shields.io/badge/OpenCV-Computer_Vision-5C3EE8?style=for-the-badge)
![Pandas](https://img.shields.io/badge/Pandas-Data_Engineering-150458?style=for-the-badge)

Um sistema de Automação Robótica de Processos (RPA) híbrido, focado na ingestão de dados de suprimentos aeronáuticos (abastecimento de QAV/AVGAS) e logs de voo para dentro de sistemas ERP/Legacy isolados em Máquinas Virtuais (VDI).

## 🛰️ O Desafio do Setor
Na engenharia aeroespacial e de defesa, muitos sistemas de controle de frota e manutenção operam em ambientes *air-gapped* ou através de Terminal Services (Remote Desktop) por questões de segurança da informação (InfoSec). 

Isso impossibilita o uso de integrações modernas (APIs REST) ou automação web (Selenium/Playwright). Inserir dados de abastecimento de aeronaves e horas de voo nesses sistemas exigia trabalho manual exaustivo e sujeito a falhas humanas — o que é inaceitável para o controle de manutenções preventivas baseadas em ciclos de motor (TBO).

## 🏗️ Arquitetura da Solução
Desenvolvi uma solução "agnóstica" que atua em duas camadas:

1. **Camada Analítica (ETL & Inteligência):** - Utiliza `Pandas` e Expressões Regulares (Regex) para extrair dados de manifestos de voo desestruturados (PDFs, planilhas sujas).
   - Implementa **Lógica Fuzzy (RapidFuzz)** para sanitizar e auto-corrigir prefixos e matrículas de aeronaves digitadas incorretamente, validando contra o banco de dados da frota.
2. **Camada de Atuação Físico-Virtual:** - Através de Visão Computacional (`OpenCV`) e simulação de hardware (`PyAutoGUI`), o robô "lê" os pixels da tela da máquina virtual e assume o teclado para injetar os dados no sistema blindado, reduzindo o tempo de processamento em 83%.

## 💡 Habilidades Aplicadas
- **Programação Defensiva:** Tratamento de dados numéricos corrompidos em notação científica (E+).
- **Tolerância a Falhas:** Motor de regras de negócio para recalcular médias de consumo parciais automaticamente.
- **Desenvolvimento de UI:** Interface reativa em Streamlit para supervisão humana (Cockpit de controle).

Este projeto demonstra a capacidade de unir Engenharia de Software com as necessidades rigorosas de auditoria de dados no setor logístico e aeroespacial.
