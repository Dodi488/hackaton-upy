import glob
import pandas as pd

# 1. Lee y concatena todos los archivos Excel disponibles en la carpeta
archivos_excel = glob.glob("*.xlsx")
lista_dataframes = []

for archivo in archivos_excel:
  try:
    df = pd.read_excel(archivo, sheet_name="2012-2022")
    lista_dataframes.append(df)
  except Exception as e:
    print(f"Error al leer {archivo}: {e}")

df_maestro = pd.concat(lista_dataframes, ignore_index=True)

# 2. Limpieza previa: eliminamos espacios en blanco en los nombres de los parámetros y claves
df_maestro["Parámetro"] = df_maestro["Parámetro"].astype(str).str.strip()
df_maestro["CLAVE"] = df_maestro["CLAVE"].astype(str).str.strip()

# 3. Columnas que identifican de forma única a cada sitio de monitoreo
columnas_id = ["N.P.", "CLAVE", "SITIO", "ESTADO", "MUNICIPIO", "RHA"]

# Limpiamos duplicados exactos si los hubiera por combinación de sitio y parámetro
df_limpio = df_maestro.drop_duplicates(subset=columnas_id + ["Parámetro"])

# 4. Pivoteamos la tabla para que cada parámetro sea una columna independiente
# Los valores que no existan para un sitio se llenarán automáticamente con NaN
df_pivot = df_limpio.pivot_table(
    index=columnas_id, columns="Parámetro", values="Valor del Parámetro", aggfunc="first"
).reset_index()

# Opcional: Filtramos estrictamente para que solo queden registros del estado de Yucatán
if "ESTADO" in df_pivot.columns:
  df_pivot = df_pivot[df_pivot["ESTADO"].str.upper() == "YUCATÁN"]

# 5. Exportamos el resultado final a CSV y Excel listos para Tableau
df_pivot.to_csv("matriz_calidad_agua_yucatan.csv", index=False)
df_pivot.to_excel("matriz_calidad_agua_yucatan.xlsx", index=False)

print(
    "¡Transformación exitosa! Archivo guardado como 'matriz_calidad_agua_yucatan.xlsx'."
)
print(f"Total de sitios únicos en Yucatán: {len(df_pivot)}")
print(f"Parámetros convertidos en columnas: {list(df_pivot.columns[len(columnas_id):])}")
