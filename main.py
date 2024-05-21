from fastapi import FastAPI
from typing import Dict
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import os

app = FastAPI()

base_dir = os.getenv('BASE_DIR', '/app')

file_path = os.path.join(base_dir, 'FASTAPI.xlsx')

def read_excel_file():
    """
    Read data from the Excel file and return as a DataFrame.
    """
    return pd.read_excel(file_path)

def write_to_excel(df):
    """
    Write DataFrame to the Excel file.
    """
    wb = load_workbook(file_path)
    ws = wb.active
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)
    wb.save(file_path)

def delete_from_excel(row_to_delete):
    """
    Delete a row from the Excel file based on the given index.
    """
    df = read_excel_file()
    if row_to_delete in df.index:
        df = df[df.index != row_to_delete]
        df.to_excel(file_path, index=False)
    else:
        raise ValueError(f"Row {row_to_delete} does not exist in the Excel file")

def update_excel(row_to_update, updated_data):
    """
    Update a row in the Excel file based on the given index with the provided data.
    """
    df = read_excel_file()
    if row_to_update in df.index:
        df.loc[row_to_update] = updated_data
        df.to_excel(file_path, index=False)
    else:
        raise ValueError(f"Row {row_to_update} does not exist in the Excel file")

@app.post("/write")
async def write_data_to_excel(data: Dict[str, str]):
    """
    Endpoint to write data to the Excel file.
    """
    row_data = {key: [value] for key, value in data.items()}
    df = pd.DataFrame(row_data)
    write_to_excel(df)
    return {"message": "Data written to Excel file"}

@app.delete("/delete/{row}")
async def delete_data_from_excel(row: int):
    """
    Endpoint to delete a row from the Excel file.
    """
    delete_from_excel(row)
    return {"message": f"Row {row} deleted from Excel file"}

@app.put("/update/{row}")
async def update_data_in_excel(row: int, updated_data: dict):
    """
    Endpoint to update a row in the Excel file.
    """
    update_excel(row, updated_data)
    return {"message": f"Row {row} updated in Excel file"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
