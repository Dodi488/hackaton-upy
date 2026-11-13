import pandas as pd

path = "datos_tidy.csv"
df = pd.read_csv(path)
#df = df[df["ESTADO"] == "YUCATÁN"]
# df = df[(df["SUBTIPO CUERPO AGUA"] == "POZO") |(df["SUBTIPO CUERPO AGUA"] == "CENOTE") |(df["SUBTIPO CUERPO AGUA"] == "BOMBEO CENOTE")]
print(df.shape)
print(df.head())
print(df.info())

print()
#print(len(df["CLAVE"].unique()))
final = df[df["CLAVE"] == "OCPYU4909"]
final.to_csv("test.csv")
