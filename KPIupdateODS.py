import pandas as pd # Importa la biblioteca pandas, que se usa para manipulacion y anlisis de datos en Python.
import os # crear carpetas, manejar archivos, obtener rutas, etc


# Ruta donde estan todos los archivos de Excel
folder_path_ods = "C:\\Users\\146767\\OneDrive - Arrow Electronics, Inc\\Desktop\\All python files\\ODS"


# Lista para almacenar los DataFrames
dataFrames = []

# Recorrer todos los archivos en la carpeta
for file in os.listdir(folder_path_ods):
    if file.endswith(".xlsx") or file.endswith(".xls"): # Solo filtre archivos de Excel
        file_path = os.path.join(folder_path_ods, file)
        print(f"Abriendo el archivo: {file_path}") # Para verificar que archivos se estan abriendo

        # Que me lea todas las hojas del archivo Excel
        sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

        # Que recorra todas las hojas y las agregaue a la lista
        for sheet_name, df in sheets.items():
            df["Archivo"] = file # Agregue una columna con el numbre del archivo
            df["Hoja"] = sheet_name # Agregue una columna con el nombre de la hoja
            dataFrames.append(df)


# Combinar todos los archivos en un solo DataFrame
if dataFrames:
    df_final = pd.concat(dataFrames, ignore_index=True)
    print(df_final.head()) # Mostrar las primeras filas


# Guardar el DataFrame final en un archivo CSV para poder pasarlo a Power BI
    df_final.to_csv("datos_combinados.csv", index=False) # Aqui se guarda el archivo CSV
    print("Datos exportados a 'datos_combinados.csv'.")    
else:
    print("No se encontraron archivos Excel en la carpeta.")


