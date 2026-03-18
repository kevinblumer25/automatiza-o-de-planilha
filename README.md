# Projeto de Geração e Tratamento de Pedidos (CSV/Excel)

Este repositório contém scripts para gerar base fictícia de pedidos e formatar/extrair planilhas com base em regras definidas.

- `gerandoarquivo.py`: gera dados sintéticos de pedidos e exporta para Excel.
- `main.py`: lá o Excel gerado, cria abas específicas (PANELA e CASTRO), aplica filtros e formata visualmente (bordas, cabeçalho colorido, etc), salva resultado final.

---

##  Scripts

### `gerandoarquivo.py`

- Gera 1000 pedidos sintéticos.
- Colunas criadas:
  - `Pedido ID`, `Griffe`, `Gênero` (imputado por sufixo em `Griffe`), `Linha`, `ReferÊncia`, `Descrição`
  - datas: `Data Emissão`, `Data Confirmação`, `Data Original Entrega`, `Data Entrega Prevista`, `Data Entrega Real`
  - valores: `Quantidade`, `Valor Unitário`, `Desconto`, `Valor Subtotal`, `Valor Total`
  - metadados: `Status Pedido`, `Canal de Venda`, `Condição de Pagamento`, `Prioridade`, `Transportadora`, `Cliente`, `Documento Cliente`, `Observações`
  - `Dias Atraso` calculado
- Cria somente a aba:
  - `Pedidos` (com todos os dados)
- Usa `pandas` para gerar e gravar `.xlsx`.
- Usa `openpyxl` ou `xlsxwriter` conforme disponibilidade.

### `main.py`

- Lá `pedidos_griffes_ficticias.xlsx`.
- Renomeia a aba principal de `Pedidos` para `PANELA`.
- Remove colunas desnecessárias.
- Cria (ou substitui) abas:
  - `PANELA`: filtro específico `Tricot Fem` + `Underwear Masc` (a partir de `Griffe`)
  - `CASTRO`: filtro por griffes/marca e linha
- Aplica formatação:
  - tamanho de colunas ajustado
  - bordas em todas as células (por aba)
  - cabeçalho com preenchimento colorido
  - filtros automáticos via `ws.auto_filter.ref`
- Salva em `Carteira_Fictícia.xlsx`.

---

##  Como executar

1. Instale dependências:
   ```bash
   pip install pandas openpyxl xlsxwriter
   ```
2. Gerar dados:
   ```bash
   python gerandoarquivo.py
   ```
3. Processar e formatar:
   ```bash
   python main.py
   ```

---

##  Observações e peculiaridades

- `xlsxwriter` não suporta modo append (`mode='a'`); no projeto usar `engine='openpyxl'` para esse caso.
- Se a aba já existe (nome igual), o script pode criar `PANELA1` quando há conflito de nomes.
- O filtro de borda/cores no cabeçalho deve usar `ws.max_row`/`ws.max_column` por planilha.
- Caso precise `zero à esquerda`, use:
  - `df["Pedido ID"] = df["Pedido ID"].astype(str).str.zfill(4)` (no pandas)
  - ou `cell.number_format = "0000"` (openpyxl quando grava final).

---

##  Estrutura de saída final

- `pedidos_griffes_ficticias.xlsx` (base gerada)
- `Carteira_Fictícia.xlsx` (arquivo final com aba PANELA e CASTRO + formatação)

---

##  Extensões fáceis

- criar aba `Resumo Financeiro` (agrupamento por `Griffe`, `Status Pedido`, `Valor Total`)
- exportar arquivos CSV separados por linha/griffe
- adicionar validação de dados (PQ/SDC) antes de salvar
---

##  Próximos passos

- Trabalhando em uma interface gráfica (GUI) para tornar o uso mais atrativo e acessível para o usuário final. 
- Principal foco: facilitar escolha de filtros, execução (botões) e visualização de resultados sem abrir código diretamente.
