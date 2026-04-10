# Painel de Contratos de Iguape

Painel estatico para consulta e acompanhamento de contratos administrativos da Prefeitura Municipal de Iguape/SP.

Todos os dados publicados no site sao gerados a partir de uma planilha Excel mantida fora do repositorio.

## Estrutura

- `index.html`: entrada usada pelo GitHub Pages.
- `painel_contratos.html`: painel principal.
- `contratos-data.js`: base de dados gerada a partir da planilha Excel.
- `scripts/gerar-dados-contratos.ps1`: gerador oficial dos dados.
- `.gitignore`: impede o versionamento da planilha e de arquivos temporarios.

## Atualizar os dados

1. Atualize a planilha Excel de controle de prazos.
2. Gere o arquivo de dados com um destes formatos:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\gerar-dados-contratos.ps1 -WorkbookPath "CAMINHO\PARA\CONTROLE DE PRAZOS 2026.xlsx"
```

```powershell
# Se a planilha estiver na raiz do projeto, o script localiza automaticamente o primeiro .xlsx valido.
powershell -ExecutionPolicy Bypass -File .\scripts\gerar-dados-contratos.ps1
```

3. Revise o arquivo `contratos-data.js`.
4. Publique as alteracoes no GitHub.

## O que o gerador faz

- Le as 8 abas da planilha.
- Normaliza datas para `YYYY-MM-DD`.
- Corrige a alternancia entre fornecedor e valor na aba `PREGAO ELETRONICO`.
- Preserva `status_excel` e `valor_texto` quando o valor original nao e numerico.
- Deriva `tipo` automaticamente como `Ata` ou `Contrato`.
- Escreve `window.PAINEL_CONTRATOS_DATA = { ultimaAtualizacao, origemArquivo, contratos }`, expondo apenas o nome do arquivo de origem.

## Publicacao

O projeto foi preparado para publicar no GitHub Pages pela branch `main`, usando a raiz do repositorio.

## Validacao recomendada

- Conferir a contagem total de registros gerados.
- Validar contratos criticos como `005/2025`, `ATA 20/2025`, `037/2025`, `005/2022`, `CE 003/2026`, `CE 009/2025` e `CE 008/2025`.
- Abrir `index.html` ou `painel_contratos.html` localmente para revisar filtros, cards e exportacao CSV.
