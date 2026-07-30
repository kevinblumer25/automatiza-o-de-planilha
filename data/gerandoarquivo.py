import pandas as pd
import random
from datetime import datetime, timedelta

# Configurações iniciais
griffes = ['Aura & Co Masc', "Aura & Co Fem", "L'Éclat Masc", "L'Éclat Fem", 'Vanguardia', 'Selva Urbana Masc', 'Selva Urbana Fem', 'Minimalist Lab Masc', 'Minimalist Lab Fem', 'Essence Wear Masc', 'Essence Wear Fem']
linhas = ['Acessório', 'Tricot', 'Underwear', 'Calçado', 'Sarja', 'Malha', 'Malha Black', 'Moletom', 'Praia', 'Alfaiataria']
produtos_base = {
    'Acessório': ['Cinto Couro', 'Bolsa Tote', 'Boné Canvas'],
    'Tricot': ['Suéter Soft', 'Cardigan Slim', 'Colete Textura'],
    'Underwear': ['Cueca Boxer', 'Top Comfort', 'Kit Meias'],
    'Calçado': ['Tênis Urbano', 'Bota Chelsea', 'Sandália Clean'],
    'Sarja': ['Calça Chino', 'Bermuda Cargo', 'Jaqueta Trucker'],
    'Malha': ['Camiseta Basic', 'Polo Piquet', 'Regata Rib'],
    'Malha Black': ['T-shirt Premium', 'Henley Dark', 'Longline Black'],
    'Moletom': ['Hoodie Oversized', 'Calça Jogger', 'Blusão Fleece'],
    'Praia': ['Sunga Estampada', 'Biquíni Aro', 'Short De Água'],
    'Alfaiataria': ['Blazer Modern', 'Calça Reta', 'Colete Social']
}

status_pedido = ['Em Aberto', 'Confirmado', 'Em Produção', 'Enviado', 'Concluído', 'Cancelado']
canal_venda = ['E-commerce', 'Varejo', 'Atacado', 'Marketplace']
pagamento = ['À Vista', '30 dias', '60 dias', 'Cartão', 'Boleto']
prioridade = ['Baixa', 'Normal', 'Alta', 'Urgente']
transportadoras = ['Correios', 'Total Express', 'Jadlog', 'Loggi', 'DHL']
no_clientes = ['Ana Silva', 'Bruno Costa', 'Carla Lima', 'Diego Souza', 'Eduarda Neto']

data_lista = []
referencias_cache = {}  # Para garantir que a mesma ref tenha a mesma descrição

for pedido_id in range(1, 1001):  # 1000 pedidos
    griffe = random.choice(griffes)
    
    linha = random.choice(linhas)

    # Lógica de Referência
    ref_key = f"{griffe}_{linha}"
    if ref_key not in referencias_cache:
        ref_num = random.randint(1000, 9999)
        desc_base = random.choice(produtos_base[linha])
        referencias_cache[ref_key] = (ref_num, f"{desc_base}")

    ref_id, descricao = referencias_cache[ref_key]

    # Lógica de Datas
    emissao = datetime(2025, random.randint(1, 6), random.randint(1, 28))
    confirmacao = emissao + timedelta(days=random.randint(1, 5))
    original = confirmacao + timedelta(days=random.randint(15, 45))
    entrega_prevista = original + timedelta(days=random.randint(2, 8))
    entrega_real = entrega_prevista + timedelta(days=random.randint(-2, 3))
    if entrega_real < confirmacao:
        entrega_real = confirmacao

    quantidade = random.randint(10, 100)
    valor_unitario = round(random.uniform(49.9, 899.0), 2)
    desconto = round(random.choice([0.0, 0.05, 0.1, 0.15, 0.2]), 2)
    subtotal = round(quantidade * valor_unitario, 2)
    valor_total = round(subtotal * (1 - desconto), 2)

    cliente = random.choice(no_clientes)
    documento = f"{random.randint(100,999)}.{random.randint(100,999)}.{random.randint(100,999)}-{random.randint(10,99)}"

    data_lista.append({
        'Pedido ID': pedido_id,
        'Griffe': griffe,
        'Linha': linha,
        'Referência': ref_id,
        'Descrição': descricao,
        'Data Emissão': emissao.strftime('%d/%m/%Y'),
        'Data Confirmação': confirmacao.strftime('%d/%m/%Y'),
        'Data Original Entrega': original.strftime('%d/%m/%Y'),
        'Data Entrega Prevista': entrega_prevista.strftime('%d/%m/%Y'),
        'Data Entrega Real': entrega_real.strftime('%d/%m/%Y'),
        'Quantidade': quantidade,
        'Valor Unitário': valor_unitario,
        'Desconto': desconto,
        'Valor Subtotal': subtotal,
        'Valor Total': valor_total,
        'Status Pedido': random.choice(status_pedido),
        'Canal de Venda': random.choice(canal_venda),
        'Condição de Pagamento': random.choice(pagamento),
        'Prioridade': random.choice(prioridade),
        'Transportadora': random.choice(transportadoras),
        'Cliente': cliente,
        'Documento Cliente': documento,
        'Observações': random.choice(['', 'Pedido urgente', 'Revisar tabela de preços', 'Atenção no acabamento'])
    })

# Planilha principal
pedidos_df = pd.DataFrame(data_lista)

# Calcular dados extras possíveis no mesmo DataFrame
pedidos_df['Data Entrega Prevista'] = pd.to_datetime(pedidos_df['Data Entrega Prevista'], dayfirst=True).dt.strftime('%d/%m/%Y')
pedidos_df['Data Entrega Real'] = pd.to_datetime(pedidos_df['Data Entrega Real'], dayfirst=True).dt.strftime('%d/%m/%Y')

# Adicionar coluna de dias de atraso diretamente
pedidos_df['Dias Atraso'] = (pd.to_datetime(pedidos_df['Data Entrega Real'], dayfirst=True) - pd.to_datetime(pedidos_df['Data Entrega Prevista'], dayfirst=True)).dt.days

# Gravar em uma única aba com todos os dados
try:
    import xlsxwriter
    writer_engine = 'xlsxwriter'
except ImportError:
    writer_engine = 'openpyxl'

with pd.ExcelWriter('pedidos_griffes_ficticias.xlsx', engine=writer_engine) as writer:
    pedidos_df.to_excel(writer, sheet_name='Pedidos', index=False)

print(f"Arquivo gerado com sucesso: pedidos_griffes_ficticias.xlsx (engine={writer_engine})")