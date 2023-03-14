# dataAnalisisExcel.py

#!pip install openpyxl
#!pip install xlsxwriter
#Analizando datos de un libro de Excel con Python
# Lectura del libro de Excel y almacenamiento en dataframe
import pandas as pd

# Calcular el tital de ventas y equipos vendidos
excelPath = r"dataset.xlsx"
dataframe = pd.read_excel(excelPath, "Ventas")
#print(dataframe)
#dataframe.style

print(f"Unidades vendidas: {totalUnits}")
print("Unidades totales: " + "${:0,.2f}".format(totalSales))


startMessage = "Ventas totales ${sales:0,.2f} y Unidades vendidas {units}".format(sales=totalSales, units=totalUnits)
print(startMessage)


#Convertir a DataFrame
valuesDict = {"Ventas totales:" :[totalSales], "Unidades vendidas: ":[totalUnits] }
#valuesDict
resultFrame =pd.DataFrame.from_dict(valuesDict)
resultFrame


#Ventas por marca
salesByBranchFrame =  dataframe.groupby("Marca")[["Cantidad","Total"]].sum()
salesByBranchFrame.style

#Marca con mayor ventas
opOneBranchFrame = salesByBranchFrame["Total"].sort_values(ascending=False).head(1)
topOneBranchFrame = topOneBranchFrame.to_frame()
topOneBranchFrame.style

#Guardar dataframes en un nuevo libro de Excel
writer = pd.ExcelWriter("NuevoReporte.xlsx", engine="xlsxwriter")

resultFrame.to_excel(writer,sheet_name="Resultados", index=False)
salesByBranchFrame.to_excel(writer,sheet_name="Resumen por marca", startcol=1, startrow=1)
topOneBranchFrame.to_excel(writer, sheet_name="Marca más vendida")

writer.save()