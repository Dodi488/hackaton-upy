import pandas as pd

path = "TODOSLOSMONITOREOS.xlsb"
df = pd.read_excel(path)
df = df[df["ESTADO"] == "YUCATÁN"]
# df = df[(df["SUBTIPO CUERPO AGUA"] == "POZO") |(df["SUBTIPO CUERPO AGUA"] == "CENOTE") |(df["SUBTIPO CUERPO AGUA"] == "BOMBEO CENOTE")]
#print(df.shape)
#print(df.head())
#print(df.info())

#print()
#print(df[df["CUENCA"] == 'YUCATÁN (RESTO DE RH 32)'])
df.to_excel("DATOS_YUCATAN.xlsx", sheet_name="Table 1", index=False)
