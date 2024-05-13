import openpyxl as op
from tabulate import tabulate
from typing import List, Tuple

doct = op.load_workbook('Detalle_hotmart_Enero24.xlsx')
sheet = doct.active


def valid_name_cell():
    if sheet['J1'].value == 'Currency':
        return True
    else:
        print('No es posible continuar la celda J no es "Currency"')
        return False


def sum_values(tupla) -> float:
    price: float = 0.0
    for money, amount in tupla:
        price += amount
    return round(price, 2)


def validate_results(total, euro, dolar):
    if total == (euro + dolar):
        print(f"la suma de : {euro}(euros) +  {dolar}(dolares) es igual a {total}")
        print("Operación realizada con éxito")
    else:
        print("Error en el cálculo")


def cal_total_values():

    total: float = 0.0
    total_eur: float = 0.0
    total_usd: float = 0.0
    data: List[Tuple[str, float]] = []

    if valid_name_cell():
        for row in sheet.iter_rows(min_col=10, max_col=11, values_only=True):
            data.append(row)

    # eliminr el header de la columja J para poder hacer el calculo
    data.pop(0)

    # Lista de EUR
    tuplas_eur: List[Tuple[str, float]] = [tupla for tupla in data if tupla[0] == "EUR"]
    # Lista de USD
    tuplas_usd: List[Tuple[str, float]] = [tupla for tupla in data if tupla[0] == "USD"]

    # Total
    total = sum_values(data)
    print(f"total: {total}")
    print("============")

    total_eur = sum_values(tuplas_eur)
    print(f"total_eur: {total_eur}")

    total_usd = sum_values(tuplas_usd)
    print(f"total_usd: {total_usd}")

    validate_results(total, total_eur, total_usd)


cal_total_values()