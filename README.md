# Projeto de Geração e Tratamento de Pedidos (Excel)

Automação para gerar dados fictícios de pedidos, processar a base em abas específicas (PANELA e CASTRO) e aplicar formatação visual. O fluxo agora também conta com uma interface gráfica para execução mais simples.

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
Processa `pedidos_griffes_ficticias.xlsx` e gera `Carteira_Fictícia_{data}.xlsx`.

**Operações:**
1. Renomeia aba `Pedidos` para `PANELA`
2. Remove colunas: `Data Entrega Prevista`, `Data Entrega Real`, `Documento Cliente`, `Transportadora`, `Condição de Pagamento`
3. Cria duas abas com filtros:
   - **PANELA**: Griffes `Aura & Co`, `L'Éclat`, `Vanguardia` com linhas `Tricot Fem` + `Underwear Masc`
   - **CASTRO**: Griffes `Aura & Co Fem`, `L'Éclat Fem` com linhas `Malha`, `Malha Black`, `Moletom`, `Underwear`
4. Formatações aplicadas:
   - Ajuste automático de largura de colunas
   - Bordas em todas as células (estilo thin)
   - Cabeçalho com preenchimento laranja (#FABF8F)
   - Filtros automáticos ativados
   - `Pedido ID` com formato de 6 dígitos com zeros à esquerda (`000000`)

**Saída:** `Carteira_Fictícia_{dd.mm}.xlsx`

### `Interface.py`
Nova interface gráfica em `customtkinter` que executa o processamento diretamente ao clicar em um botão. Ela chama `main.py`, cria o arquivo `dados.txt` na primeira execução e exibe a mensagem de conclusão no próprio painel.

### `interface.exe`
Executável Windows gerado a partir de `Interface.py` para uso direto, sem precisar abrir o terminal ou o Python.

### `rodar_planilha.bat`
Script auxiliar para rodar o fluxo com um clique em ambiente Windows.

---

## Como Executar

### Opção 1: Executável Windows (mais simples)
```text
1. Coloque o arquivo interface.exe e a planilha pedidos_griffes_ficticias.xlsx na mesma pasta.
2. Clique duas vezes em interface.exe.
3. O programa irá processar a planilha e gerar a saída em Excel.
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
dados.txt (arquivo de confirmação criado pela interface)
```

---

## Notas Técnicas

- **Engine:** Por padrão usa `openpyxl` (suporta modo append)
- **Formato de data:** `dd.mm` no nome do arquivo (ex: `Carteira_Fictícia_24.03.xlsx`)
- **Dependências:** `pandas`, `openpyxl`, `xlsxwriter`, `customtkinter`
- **Python:** 3.7+

---

## Estrutura de Saída

| Arquivo | Conteúdo | Quando |
|---------|----------|--------|
| `pedidos_griffes_ficticias.xlsx` | Base bruta com 1000 pedidos, aba única `Pedidos` | Após `gerandoarquivo.py` |
| `Carteira_Fictícia_{data}.xlsx` | Abas `PANELA` e `CASTRO` formatadas | Após `main.py` |
| `dados.txt` | Arquivo de confirmação criado pela interface após o processamento | Após a primeira execução pela interface |

---

## Próximas Melhorias

- Exportação em múltiplos formatos (CSV, PDF)
- Dashboard interativo com resumos financeiros
- Agrupamento por Griffe e Status Pedido
- Melhorias visuais na interface gráfica
