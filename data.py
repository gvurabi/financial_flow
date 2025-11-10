# export_categorias.py
import os
import json
import requests
import pandas as pd  # (você importa, mas não usa; pode remover se quiser)
from datetime import datetime, date
from decimal import Decimal
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

API_BASE = "https://app.base44.com/api"
API_KEY = os.getenv("BASE44_API_KEY")  # Lê do ambiente

if not API_KEY:
    raise RuntimeError("BASE44_API_KEY não está definida. Configure a variável de ambiente.")

def criar_planilha(dados, nome_arquivo_saida):
    caminho = str(nome_arquivo_saida)
    if not caminho.lower().endswith(".xlsx"):
        caminho += ".xlsx"
    dirpath = os.path.dirname(os.path.abspath(caminho))
    if dirpath and not os.path.exists(dirpath):
        os.makedirs(dirpath, exist_ok=True)

    headers = None
    rows_matrix = None
    dataset = None

    if isinstance(dados, dict) and "headers" in dados and "rows" in dados and isinstance(dados["rows"], list):
        headers = list(dados["headers"])
        rows_matrix = list(dados["rows"])
    else:
        if isinstance(dados, dict):
            for key in ("data", "results", "items", "value"):
                if isinstance(dados.get(key), list):
                    dataset = dados[key]
                    break
            if dataset is None:
                dataset = [dados]
        elif isinstance(dados, list):
            dataset = dados
        else:
            dataset = [dados]

        if all(isinstance(x, dict) for x in dataset):
            columns = []
            flattened_rows = []
            for item in dataset:
                flat = {}
                stack = [(None, item)]
                while stack:
                    prefix, obj = stack.pop()
                    if isinstance(obj, dict):
                        for k, v in obj.items():
                            key = f"{prefix}.{k}" if prefix else str(k)
                            stack.append((key, v))
                    elif isinstance(obj, (list, tuple)):
                        keyname = prefix if prefix is not None else "lista"
                        flat[keyname] = json.dumps(obj, ensure_ascii=False)
                    else:
                        val = obj
                        if isinstance(obj, (datetime, date)):
                            val = obj.isoformat()
                        elif isinstance(obj, Decimal):
                            try:
                                val = float(obj)
                            except Exception:
                                val = str(obj)
                        keyname = prefix if prefix is not None else "valor"
                        flat[keyname] = val

                for k in flat.keys():
                    if k not in columns:
                        columns.append(k)
                flattened_rows.append(flat)

            headers = columns
            rows_matrix = [[row.get(col) for col in columns] for row in flattened_rows]
        elif all(isinstance(x, (list, tuple)) for x in dataset):
            max_len = max((len(x) for x in dataset), default=0)
            headers = [f"col_{i+1}" for i in range(max_len)]
            rows_matrix = [list(x) + [None] * (max_len - len(x)) for x in dataset]
        else:
            headers = ["valor"]
            rows_matrix = [[x] for x in dataset]

    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"

    # Cabeçalho
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)

    ws.freeze_panes = "A2"

    col_widths = [len(str(h)) if h is not None else 0 for h in headers]

    for r_idx, row in enumerate(rows_matrix, start=2):
        if len(row) < len(headers):
            row = list(row) + [None] * (len(headers) - len(row))
        for c_idx, val in enumerate(row, start=1):
            cell_val = val
            if isinstance(cell_val, (datetime, date)):
                cell_val = cell_val.isoformat()
            elif isinstance(cell_val, Decimal):
                try:
                    cell_val = float(cell_val)
                except Exception:
                    cell_val = str(cell_val)

            ws.cell(row=r_idx, column=c_idx, value=cell_val)

            tam = len(str(cell_val)) if cell_val is not None else 0
            if tam > col_widths[c_idx - 1]:
                col_widths[c_idx - 1] = tam

    # Autofiltro (ajuste do range)
    from_col = "A"
    to_col = get_column_letter(len(headers))  # corrigido: sem +1
    last_row = len(rows_matrix) + 1
    ws.auto_filter.ref = f"{from_col}1:{to_col}{last_row}"

    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(10, min(w + 2, 60))

    wb.save(caminho)
    return caminho

def make_api_request(api_path, method='GET', data=None, timeout=60):
    url = f'{API_BASE}/{api_path}'
    headers = {
        'api_key': API_KEY,
        'Content-Type': 'application/json'
    }
    if method.upper() == 'GET':
        response = requests.request(method, url, headers=headers, params=data, timeout=timeout)
    else:
        response = requests.request(method, url, headers=headers, json=data, timeout=timeout)
    response.raise_for_status()
    return response.json()

# Execução: salvar em 'outputs/'
BASE_DIR = Path(__file__).parent
OUT_DIR = BASE_DIR / "outputs"
OUT_DIR.mkdir(exist_ok=True)

entities = make_api_request('apps/68f5182879c5fe5a86e409ee/entities/Category')
print(criar_planilha(entities, OUT_DIR / "Category.xlsx"))

entities = make_api_request('apps/68f5182879c5fe5a86e409ee/entities/Transaction')
print(criar_planilha(entities, OUT_DIR / "Transaction.xlsx"))

entities = make_api_request('apps/68f5182879c5fe5a86e409ee/entities/BankAccount')
print(criar_planilha(entities, OUT_DIR / "BankAccount.xlsx"))

entities = make_api_request('apps/68f5182879c5fe5a86e409ee/entities/StatementTransaction')
print(criar_planilha(entities, OUT_DIR / "StatementTransaction.xlsx"))
