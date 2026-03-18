# Projeto de Gera��o e Tratamento de Pedidos (CSV/Excel)

Este reposit�rio cont�m scripts para gerar base fict�cia de pedidos e formatar/extrair planilhas com base em regras definidas.

- `gerandoarquivo.py`: gera dados sint�ticos de pedidos e exporta para Excel.
- `main.py`: l� o Excel gerado, cria abas espec�ficas (PANELA e CASTRO), aplica filtros e formata visualmente (bordas, cabe�alho colorido, etc), salva resultado final.

---

##  Scripts

### `gerandoarquivo.py`

- Gera 1000 pedidos sint�ticos.
- Colunas criadas:
  - `Pedido ID`, `Griffe`, `G�nero` (imputado por sufixo em `Griffe`), `Linha`, `Refer�ncia`, `Descri��o`
  - datas: `Data Emiss�o`, `Data Confirma��o`, `Data Original Entrega`, `Data Entrega Prevista`, `Data Entrega Real`
  - valores: `Quantidade`, `Valor Unit�rio`, `Desconto`, `Valor Subtotal`, `Valor Total`
  - metadados: `Status Pedido`, `Canal de Venda`, `Condi��o de Pagamento`, `Prioridade`, `Transportadora`, `Cliente`, `Documento Cliente`, `Observa��es`
  - `Dias Atraso` calculado
- Cria somente a aba:
  - `Pedidos` (com todos os dados)
- Usa `pandas` para gerar e gravar `.xlsx`.
- Usa `openpyxl` ou `xlsxwriter` conforme disponibilidade.

### `main.py`

- L� `pedidos_griffes_ficticias.xlsx`.
- Renomeia a aba principal de `Pedidos` para `PANELA`.
- Remove colunas desnecess�rias.
- Cria (ou substitui) abas:
  - `PANELA`: filtro espec�fico `Tricot Fem` + `Underwear Masc` (a partir de `Griffe`)
  - `CASTRO`: filtro por griffes/marca e linha
- Aplica formata��o:
  - tamanho de colunas ajustado
  - bordas em todas as c�lulas (por aba)
  - cabe�alho com preenchimento colorido
  - filtros autom�ticos via `ws.auto_filter.ref`
- Salva em `Carteira_Fict�cia.xlsx`.

---

##  Como executar

1. Instale depend�ncias:
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

##  Observa��es e peculiaridades

- `xlsxwriter` n�o suporta modo append (`mode='a'`); no projeto usar `engine='openpyxl'` para esse caso.
- Se a aba j� existe (nome igual), o script pode criar `PANELA1` quando h� conflito de nomes.
- O filtro de borda/cores no cabe�alho deve usar `ws.max_row`/`ws.max_column` por planilha.
- Caso precise `zero � esquerda`, use:
  - `df["Pedido ID"] = df["Pedido ID"].astype(str).str.zfill(4)` (no pandas)
  - ou `cell.number_format = "0000"` (openpyxl quando grava final).

---

##  Estrutura de sa�da final

- `pedidos_griffes_ficticias.xlsx` (base gerada)
- `Carteira_Fict�cia.xlsx` (arquivo final com aba PANELA e CASTRO + formata��o)

---

##  Extens�es f�ceis

- criar aba `Resumo Financeiro` (agrupamento por `Griffe`, `Status Pedido`, `Valor Total`)
- exportar arquivos CSV separados por linha/griffe
- adicionar valida��o de dados (PQ/SDC) antes de salvar
---

##  Próximos passos

- Trabalhando em uma interface gráfica (GUI) para tornar o uso mais atrativo e acessível para o usuário final. 
- Principal foco: facilitar escolha de filtros, execução (botões) e visualização de resultados sem abrir código diretamente.