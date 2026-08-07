# Projeto de Geração e Tratamento de Pedidos (Excel)

Automação para gerar uma base fictícia de pedidos, processar a planilha em abas específicas (PANELA e CASTRO) e aplicar formatação visual. O fluxo também conta com uma interface gráfica para execução mais simples.

---

## Arquivos do Projeto

### `gerandoarquivo.py`
Gera uma base fictícia com 1000 pedidos e exporta para `pedidos_griffes_ficticias.xlsx`.

**Saída:**
- Aba `Pedidos` contendo:
  - `Pedido ID`, `Griffe`, `Linha`, `Referência`, `Descrição`
  - Datas: `Data Emissão`, `Data Confirmação`, `Data Original Entrega`, `Data Entrega Prevista`, `Data Entrega Real`
  - Valores: `Quantidade`, `Valor Unitário`, `Desconto`, `Valor Subtotal`, `Valor Total`
  - Metadados: `Status Pedido`, `Canal de Venda`, `Condição de Pagamento`, `Prioridade`, `Transportadora`, `Cliente`, `Documento Cliente`, `Observações`
  - `Dias Atraso` (calculado automaticamente)

### `main.py`
Processa a planilha de entrada e gera um arquivo final em Excel.

**Comportamento atual:**
- Procura automaticamente por arquivos como `Carteira_Fictícia.xlsx`, `Carteira_Fictícia_<data>.xlsx` ou `pedidos_griffes_ficticias.xlsx`
- Usa a mesma pasta do executável/script para arquivos de entrada e saída
- Aceita tanto a aba `Pedidos` quanto a aba `PANELA` já existente
- Cria/atualiza as abas `PANELA` e `CASTRO`
- Aplica formatação visual e filtros automáticos

**Saída:** `Carteira_Fictícia_{dd.mm}.xlsx`

### `Interface.py`
Interface gráfica em `customtkinter` para executar o processamento com um clique. Ela chama `main.py`, cria o arquivo `dados.txt` na primeira execução e exibe a mensagem de conclusão no próprio painel.

### `interface.exe`
Executável Windows gerado a partir de `Interface.py` para uso direto, sem precisar abrir o terminal ou o Python.

### `rodar_planilha.bat`
Script auxiliar para rodar o fluxo com um clique em ambiente Windows.

---

## Como Executar

### Opção 1: Executável Windows (mais simples)
```text
1. Coloque o arquivo interface.exe e a planilha de entrada na pasta de trabalho do programa.
2. O programa utilizará a mesma pasta do executável para arquivos de saída e confirmação.
3. Clique duas vezes em interface.exe.
4. O programa irá processar a planilha e gerar a saída em Excel.
```

### Opção 2: Interface em Python
```bash
python Interface.py
```

### Opção 3: Via terminal
```bash
# Instalar dependências (primeira vez)
pip install pandas openpyxl xlsxwriter customtkinter

# Gerar dados
python gerandoarquivo.py

# Processar e formatar
python main.py
```

---

## Fluxo de Execução

```text
gerandoarquivo.py
    ↓
pedidos_griffes_ficticias.xlsx (base com 1000 pedidos)
    ↓
Interface.py ou interface.exe
    ↓
main.py (lê, processa, formata)
    ↓
Carteira_Fictícia_{data}.xlsx (resultado final)
    ↓
dados.txt / app_data (arquivos de confirmação e saída)
```

---

## Notas Técnicas

- **Engine:** usa `openpyxl` para leitura/escrita com suporte a append
- **Formato de data:** `dd.mm` no nome do arquivo (ex.: `Carteira_Fictícia_24.03.xlsx`)
- **Dependências:** `pandas`, `openpyxl`, `xlsxwriter`, `customtkinter`
- **Python:** 3.7+
- **Execução empacotada:** o executável usa uma pasta `app_data` para manter os arquivos de saída em um local previsível

---

## Estrutura de Saída

| Arquivo | Conteúdo | Quando |
|---------|----------|--------|
| `pedidos_griffes_ficticias.xlsx` | Base bruta com 1000 pedidos, aba única `Pedidos` | Após `gerandoarquivo.py` |
| `Carteira_Fictícia_{data}.xlsx` | Abas `PANELA` e `CASTRO` formatadas | Após `main.py` |
| `dados.txt` / `app_data` | Arquivos de confirmação e saída usados pela interface | Após a execução pela interface |

---

## Próximas Melhorias

- Exportação em múltiplos formatos (CSV, PDF)
- Dashboard interativo com resumos financeiros
- Agrupamento por Griffe e Status Pedido
- Melhorias visuais na interface gráfica
- Estilização mais avançada do aplicativo
- Opção para uso individual por pessoa com login
