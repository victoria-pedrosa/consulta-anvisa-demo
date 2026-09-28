# Demonstração — Consulta de produtos ANVISA para a LC 214/2025

> Projeto de portfólio de **Victória Pedrosa**. **Demonstração** de consulta de produtos ANVISA para a LC 214/2025 — versão com dados fictícios (nomes, CNPJs, e-mails e IDs internos substituídos).

## Problema de negócio
Para aplicar o benefício da LC 214/2025, é preciso provar que o produto é regularizado na ANVISA e enquadrado no anexo da lei — duas checagens trabalhosas.

## Antes x depois
| | Antes | Depois |
|---|---|---|
| Como é feito | Consulta manual no site da ANVISA e leitura dos anexos da lei, produto a produto. | Motor valida as duas etapas (ANVISA + enquadramento na LC 214) com base oficial embutida e gera evidência. |

## Ganho
- Enquadramento comprovado e documentado, com menos risco de autuação.

## Tecnologias
APIs REST, Gatilhos agendados, Gemini API, Google Apps Script, Google Drive, Google Sheets, HTML/JavaScript, Web App (HtmlService)

## Arquivos
- `Codigo.gs`
- `Index.html`
- `baselc214.gs`

## Como usar
Crie um projeto no Google Apps Script, copie os arquivos `.gs`/`.html` e configure as Propriedades do script indicadas no código.

## Autora
Victória Pedrosa — Product Owner do Time de IA, automação de processos contábeis e fiscais.
