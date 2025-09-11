import os
import json 
import openpyxl

def excel_to_json(path):
    wb = openpyxl.load_workbook(path, data_only=True)

    for sheet in wb.sheetnames:
        print(wb[sheet]['BY278'].value)

excel_to_json("madeira.xlsx")
