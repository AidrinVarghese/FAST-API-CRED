from fastapi import FastAPI
from typing import Dict
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

app = FastAPI()
file_path = "D:\ILP\Python Training\FAST API CRED OPERATIONS\FASTAPI.xlsx"


def read_excel_file():
    return pd.read_excel(file_path)


def write_to_excel(df):
    wb = load_workbook(file_path)
    ws = wb.active
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)
    wb.save(file_path)


def delete_from_excel(row_to_delete):
    df = read_excel_file()
    if row_to_delete in df.index:
        df = df[df.index != row_to_delete]
        df.to_excel(file_path, index=False)
    else:
        raise ValueError(f"Row {row_to_delete} does not exist in the Excel file")


def update_excel(row_to_update, updated_data):
    df = read_excel_file()
    if row_to_update in df.index:
        df.loc[row_to_update] = updated_data
        df.to_excel(file_path, index=False)
    else:
        raise ValueError(f"Row {row_to_update} does not exist in the Excel file")


@app.post("/write")
async def write_data_to_excel(data: Dict[str, str]):
    row_data = {key: [value] for key, value in data.items()}
    df = pd.DataFrame(row_data)
    write_to_excel(df)
    return {"message": "Data written to Excel file"}


@app.delete("/delete/{row}")
async def delete_data_from_excel(row: int):
    delete_from_excel(row)
    return {"message": f"Row {row} deleted from Excel file"}


@app.put("/update/{row}")
async def update_data_in_excel(row: int, updated_data: dict):
    update_excel(row, updated_data)
    return {"message": f"Row {row} updated in Excel file"}
