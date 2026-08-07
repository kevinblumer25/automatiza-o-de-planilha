from datetime import datetime
import os
import sys
from pathlib import Path


def get_app_dir():
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        if exe_dir and os.path.isdir(exe_dir):
            return exe_dir

        if getattr(sys, "_MEIPASS", None):
            meipass_dir = os.path.abspath(sys._MEIPASS)
            if meipass_dir and os.path.isdir(meipass_dir):
                return meipass_dir

        return os.getcwd()

    return os.path.dirname(os.path.abspath(__file__))


def get_work_dir():
    return Path(get_app_dir())


app_dir = get_work_dir()
app_dir.mkdir(parents=True, exist_ok=True)


def find_input_workbook():
    today = datetime.now().strftime('%d-%m-%y')
    candidates = []

    for base in [Path.cwd(), Path(app_dir), Path(__file__).resolve().parent, Path(get_app_dir())]:
        candidates.append(base / f"Carteira_Fictícia_{today}.xlsx")
        candidates.append(base / "Carteira_Fictícia.xlsx")
        candidates.append(base / "pedidos_griffes_ficticias.xlsx")

    for pattern in ["Carteira_Fictícia*.xlsx", "Carteira*.xlsx", "pedidos_griffes_ficticias.xlsx"]:
        for base in [Path.cwd(), Path(app_dir), Path(__file__).resolve().parent, Path(get_app_dir())]:
            candidates.extend(base.glob(pattern))

    seen = set()
    results = []
    for path in candidates:
        if not path.exists():
            continue
        if path in seen:
            continue
        seen.add(path)
        results.append(path)

    if not results:
        return None

    return max(results, key=lambda p: p.stat().st_mtime)


def get_file_path(filename):
    workbook = find_input_workbook()
    if workbook is not None:
        return str(workbook)

    return os.path.join(app_dir, filename)


def main():
    import pandas as pd
    from openpyxl import load_workbook

    workbook_path = get_file_path("pedidos_griffes_ficticias.xlsx")

    # Carregar o workbook
    wb = load_workbook(workbook_path)

    # Aceitar tanto a aba original "Pedidos" quanto a aba já existente "PANELA"
    if 'Pedidos' in wb.sheetnames:
        ws = wb['Pedidos']
    elif 'PANELA' in wb.sheetnames:
        ws = wb['PANELA']
    else:
        ws = wb.active

    if ws.title != 'PANELA':
        ws.title = 'PANELA'

    wb.save(workbook_path)

    # importar o arquivo Excel
    df = pd.read_excel(workbook_path)

    # Excluindo colunas
    colunas = ['Data Entrega Prevista', 'Data Entrega Real', 'Documento Cliente', 'Transportadora', 'Condição de Pagamento']
    df = df.drop(columns=colunas)

    # Alterando o nome das planilhas (usa openpyxl, pois xlsxwriter não aceita modo append)
    with pd.ExcelWriter(workbook_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='PANELA', index=False)
        df.to_excel(writer, sheet_name='CASTRO', index=False)
    
    df_novo = pd.read_excel(workbook_path, sheet_name=['PANELA', 'CASTRO'])

    df_panela = df_novo['PANELA']
    df_castro = df_novo['CASTRO']

    # Filtrando Griffes Panela
    griffes_panela = ['Aura & Co Fem', 'Aura & Co Masc', 'Vanguardia', 'L\'Éclat Fem', 'L\'Éclat Masc']
    linhas_excluir_panela = ['Malha', 'Malha Black', 'Moletom']
    df_panela = df_panela.drop(df_panela[df_panela['Linha'].isin(linhas_excluir_panela)].index)
    # Mantém apenas Tricot Feminino e Underwear Masculino na aba PANELA
    filtro_panela = (
        (df_panela['Linha'] != 'Tricot') & (df_panela['Griffe'].isin(['Aura & Co Masc', 'L\'Éclat Masc', 'Vanguardia']))
    ) | (
       (df_panela['Linha'] != 'Underwear') & (df_panela['Griffe'].isin(['Aura & Co Fem', 'L\'Éclat Fem']))
    )

    df_panela = df_panela[df_panela['Griffe'].isin(griffes_panela) & filtro_panela]

    # Filtrando Griffes Castro
    griffes_castro = ['Aura & Co Fem', 'L\'Éclat Fem']
    linhas_castro = ['Malha', 'Malha Black', 'Moletom', 'Underwear']
    df_castro = df_castro[df_castro['Griffe'].isin(griffes_castro) & df_castro['Linha'].isin(linhas_castro)]

    with pd.ExcelWriter(workbook_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_panela.to_excel(writer, sheet_name='PANELA', index=False)
        df_castro.to_excel(writer, sheet_name='CASTRO', index=False)


    from openpyxl.utils import get_column_letter


    wb = load_workbook(workbook_path)
    ws = wb.active

    # Filtros automáticos
    for ws in wb.worksheets:
        if ws.max_row > 1 and ws.max_column > 0:
            first_col = 1
            last_col = ws.max_column
            ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}1"



    # Ajustar a largura das colunas para todas as sheets
    for ws in wb.worksheets:
        dims = {}
        for row in ws.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column] = max((dims.get(cell.column, 0), len(str(cell.value))))
        for col, value in dims.items():
            ws.column_dimensions[get_column_letter(col)].width = value + 2



    # Adicionando bordas
    from openpyxl.styles import Border, Side

    borda = Border(
        left=Side(border_style='thin'),
        right=Side(border_style='thin'),
        top=Side(border_style='thin'),
        bottom=Side(border_style='thin')
    )


    for ws in wb.worksheets:
        for linha in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for celula in linha:
                celula.border = borda



    # Adicionando cor de fundo no cabeçalho somente
    from openpyxl.styles import PatternFill

    fill = PatternFill(start_color='FABF8F', end_color='FABF8F', fill_type='solid')

    for ws in wb.worksheets:
        # Preenche apenas a primeira linha (cabeçalho), até a última coluna com dados
        header_row = 1
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=header_row, column=col)
            if cell.value is not None:
                cell.fill = fill

    # Adicionando um (ou mais) 0 à esquerda
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = '000000'

    output_path = os.path.join(app_dir, f'Carteira_Fictícia_{datetime.now().strftime("%d.%m")}.xlsx')
    wb.save(output_path)

if __name__ == "__main__":
    main()