import glob
import pandas as pd

# 1. Obtiene la lista de todos los archivos Excel en la carpeta actual
archivos_excel = glob.glob("*.xlsx")

lista_dataframes = []

# 2. Lee cada archivo y su hoja principal ('2012-2022')
for archivo in archivos_excel:
  try:
    df = pd.read_excel(archivo)
    lista_dataframes.append(df)
  except Exception as e:
    print(f"Error al leer {archivo}: {e}")

# 3. Concatena todos los DataFrames verticalmente en uno solo
df_maestro = pd.concat(lista_dataframes, ignore_index=True)

# 4. (Opcional) Filtra exclusivamente los registros del estado de Yucatán
df_yucatan = df_maestro[df_maestro["ESTADO"].str.upper() == "YUCATÁN"]

# 5. Guarda el resultado en un nuevo archivo CSV o Excel maestro
df_yucatan.to_csv("master_calidad_agua_yucatan.csv", index=False)
df_yucatan.to_excel("master_calidad_agua_yucatan.xlsx", index=False)

print(
    f"¡Listo! Se unificaron {len(archivos_excel)} archivos. Total de registros"
    f" para Yucatán: {len(df_yucatan)}"
)
