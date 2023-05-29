from connector.mysql import mysql_query
import pandas as pd
from utils.workbook import createWorkbook, createTitle, setCellText, setHeader, setFormula, setData, setSheetColumnsWidth

cities = mysql_query("SELECT * FROM world.city;")
df = pd.DataFrame(cities)
wb  = createWorkbook(sheet_titles=["Cities"])
ws = wb["Cities"]
row = 1
createTitle(ws, row, row, 5, "Cities in the world data")
row += 2
setHeader(ws, start_row = row, end_row= row, start_col = 1, end_col= 1, text = "No.")
setHeader(ws, start_row = row, end_row= row, start_col = 2, end_col= 2, text = "City name")
setHeader(ws, start_row = row, end_row= row, start_col = 3, end_col= 3, text = "Country Code")
setHeader(ws, start_row = row, end_row= row, start_col = 4, end_col= 4, text = "District")
setHeader(ws, start_row = row, end_row= row, start_col = 5, end_col= 5, text = "Population")
setHeader(ws, start_row = row, end_row= row, start_col = 6, end_col= 6, text = "Created At")
row += 1
setData(df,ws, start_row=4, start_col=0)
row += len(df)
setHeader(ws, start_row= row, end_row= row, start_col= 1, end_col= 4, text="Total", cellColor="FFF000")
setFormula(ws, start_row= row, end_row= row, start_col= 5, end_col= 5, formula= f"SUM(E4:E{row - 1})", cellColor="FFF000")
setSheetColumnsWidth(ws, [10, 20, 20, 20, 20, 20])
wb.save("test1.xlsx")