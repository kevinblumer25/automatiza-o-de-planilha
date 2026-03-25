# Projeto de Geração e Tratamento de Pedidos (Excel)

Automação de geração de dados fictícios de pedidos e processamento em abas específicas (PANELA e CASTRO) com formatação visual.

---

## 📋 Arquivos do Projeto

### `gerandoarquivo.py`
Gera base fictícia com 1000 pedidos e exporta para `pedidos_griffes_ficticias.xlsx`.

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

### `rodar_planilha.bat` ⭐ (NOVO)
Executável Windows para rodar `main.py` com um clique.

**Funcionalidade:**
- Duplo-clique para iniciar
- Mensagem "Executando o script, aguarde..."
- Exibe "Concluído com sucesso" ou "Erro na execução" ao final
- Aguarda finalização completa (sem timeout)
- Pressione qualquer tecla para fechar

---

## ⚙️ Como Executar

### Opção 1: Via `.bat` (Recomendado - sem terminal)
```
Duplo-clique em: rodar_planilha.bat
```

### Opção 2: Via Python (terminal)
```bash
# Instalar dependências (primeira vez)
pip install pandas openpyxl xlsxwriter

# Gerar dados
python gerandoarquivo.py

# Processar e formatar
python main.py
```

---

## 📁 Fluxo de Execução

```
gerandoarquivo.py
    ↓
pedidos_griffes_ficticias.xlsx (base com 1000 pedidos)
    ↓
main.py (lê, processa, formata)
    ↓
Carteira_Fictícia_{data}.xlsx (resultado final)
```

---

## 📝 Notas Técnicas

- **Engine:** Por padrão usa `openpyxl` (suporta modo append)
- **Formato de data:** `dd.mm` no nome do arquivo (ex: `Carteira_Fictícia_24.03.xlsx`)
- **Dependências:** `pandas`, `openpyxl`, `xlsxwriter` (opcional)
- **Python:** 3.7+

---

## 📊 Estrutura de Saída

| Arquivo | Conteúdo | Quando |
|---------|----------|--------|
| `pedidos_griffes_ficticias.xlsx` | Base bruta com 1000 pedidos, aba única `Pedidos` | Após `gerandoarquivo.py` |
| `Carteira_Fictícia_{data}.xlsx` | Abas `PANELA` e `CASTRO` formatadas | Após `main.py` |

---

## 🔄 Próximas Melhorias

- Interface gráfica (GUI) para facilitar uso final
- Exportação em múltiplos formatos (CSV, PDF)
- Dashboard interativo com resumos financeiros
- Agrupamento por Griffe e Status Pedido
