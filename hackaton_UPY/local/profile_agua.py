#!/usr/bin/env python3
"""
profile_agua.py - First look at the merged RENAMECA water-quality file.

Input : the merged .xlsx made by merge_xlsx.py (v2, with the Hoja column)
Output: a folder with
          datos_tidy.csv          one row per site x parameter, numeric value
                                  parsed, censored values ("<0.01") flagged,
                                  period (Periodo_hoja) and water type
                                  (Tipo_agua) added. Connect this in Tableau.
          perfil_calidad_datos.xlsx
                                  coverage and data-quality tables (sheets
                                  listed in the console output). Start with
                                  the 'chequeos' and 'archivos_y_codigos'
                                  sheets.

Usage:

    python profile_agua.py merged_calidad_agua.xlsx perfil_out

Nothing is imputed or dropped silently. Rows from files that are not the
per-parameter exports (e.g. master_*, matriz_*) are ignored and counted.

Requires: pandas, openpyxl
"""

import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

PREFIX = "Calidad_de_agua_Estado_"
REQUIRED = ["Source_File", "CLAVE", "SITIO", "ESTADO", "MUNICIPIO", "RHA",
            "Parámetro", "Valor del Parámetro", "Calidad"]

# Sites are split into groundwater / surface water by which of these two
# files they appear in (the file names say so explicitly).
FILE_SUBTERRANEA = "Coliformes Fecales (Subterraneos)"
FILE_SUPERFICIAL = "Coliformes Fecales (Superficiales)"

# Sanity check: text that should appear (accent-free, lower case) in the file
# name for each parameter code. A mismatch is flagged for human review only.
CODE_KEYWORDS = {
    "alc": ["alcalin"], "as_": ["arsen"], "cd": ["cadmio"],
    "coli_fec": ["coliform"], "conduct": ["conductiv"], "cr": ["cromo"],
    "dbo": ["biolog", "dbo"], "dqo": ["quimica", "dqo"], "dur": ["dureza"],
    "e_coli": ["escherichia"], "enteroc": ["enteroc"], "fe": ["hierro"],
    "fluoruros": ["fluor", "flour"], "hg": ["mercurio"], "mn": ["mangan"],
    "n_no3": ["nitrat"], "od": ["oxigeno"], "od_fon": ["oxigeno"],
    "od_med": ["oxigeno"], "od_sup": ["oxigeno"], "pb": ["plomo"],
    "sdt": ["solidos disueltos"], "sst": ["solidos suspendidos"],
    "tox_d_48": ["daphnia"], "tox_d_48_fon": ["daphnia"],
    "tox_d_48_sup": ["daphnia"], "tox_fis_fon_15": ["vibrio"],
    "tox_fis_sup_15": ["vibrio"], "tox_v_15": ["vibrio"],
}


def strip_accents(s: str) -> str:
    import unicodedata
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower()


def param_name_from_file(fname: str) -> str:
    """'Calidad_de_agua_Estado_Oxígeno Disuelto _.xlsx' -> 'Oxígeno Disuelto'.

    The file name is the reliable parameter key: the short code in the
    'Parámetro' column is NOT unique (both coliform files use 'coli_fec').
    """
    s = fname[len(PREFIX):] if fname.startswith(PREFIX) else fname
    s = re.sub(r"\.xlsx$", "", s, flags=re.I)
    return re.sub(r"[\s_]+$", "", s).strip()


def parse_value(v):
    """Return (censura, limite_deteccion, valor_num).

    censura: '='     -> normal number
             '<'     -> below detection limit (value reported as '<0.01')
             '>'     -> above a limit (value reported as '>2400')
             'vacio' -> empty cell (site listed but no result)
             'error' -> text that could not be read as a number
    """
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "vacio", np.nan, np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return "=", np.nan, float(v)
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "nd", "n/d", "s/d"):
        return "vacio", np.nan, np.nan
    mark = "="
    if s[0] in "<>":
        mark, s = s[0], s[1:].strip()
    try:
        x = float(s.replace(",", ""))
    except ValueError:
        return "error", np.nan, np.nan
    if mark != "=":
        return mark, x, np.nan
    return "=", np.nan, x


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__)
        return 1
    src, outdir = Path(sys.argv[1]), Path(sys.argv[2])
    outdir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(src, dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        print(f"ERROR: missing columns {missing}\nFound: {list(df.columns)}")
        return 1

    n_all = len(df)
    df = df[df["Source_File"].astype(str).str.startswith(PREFIX)].copy()
    print(f"Rows in file: {n_all} | kept (per-parameter exports): {len(df)} "
          f"| ignored (master/matriz/other): {n_all - len(df)}")

    # --- tidy table -------------------------------------------------------
    parsed = df["Valor del Parámetro"].map(parse_value)
    has_hoja = "Hoja" in df.columns
    t = pd.DataFrame({
        "Source_File": df["Source_File"],
        "Periodo_hoja": df["Hoja"] if has_hoja else "sin dato",
        "Parametro_archivo": df["Source_File"].map(param_name_from_file),
        "Parametro_codigo": df["Parámetro"],
        "CLAVE": df["CLAVE"].astype(str).str.strip(),
        "SITIO": df["SITIO"],
        "MUNICIPIO": df["MUNICIPIO"],
        "ESTADO": df["ESTADO"],
        "RHA": df["RHA"],
        "Valor_original": df["Valor del Parámetro"],
        "Censura": [p[0] for p in parsed],
        "Limite_deteccion": [p[1] for p in parsed],
        "Valor_num": [p[2] for p in parsed],
        "Calidad_CONAGUA": df["Calidad"],
    })
    # Common convention for non-detects: half the detection limit. It is a
    # modelling choice - state it in the Evidence Board if you use it.
    t["Valor_LD_medio"] = t["Valor_num"].where(t["Censura"] == "=",
                                               t["Limite_deteccion"] / 2)
    t.loc[t["Censura"].isin(["vacio", "error", ">"]), "Valor_LD_medio"] = np.nan
    t["tiene_valor"] = t["Censura"] != "vacio"

    # Groundwater vs surface water, from file membership (see constants).
    in_sub = set(t.loc[t.Parametro_archivo == FILE_SUBTERRANEA, "CLAVE"])
    in_sup = set(t.loc[t.Parametro_archivo == FILE_SUPERFICIAL, "CLAVE"])
    def tipo(c):
        if c in in_sub and c not in in_sup:
            return "subterranea"
        if c in in_sup and c not in in_sub:
            return "superficial"
        return "sin_clasificar"
    t["Tipo_agua"] = t["CLAVE"].map(tipo)
    t.to_csv(outdir / "datos_tidy.csv", index=False, encoding="utf-8-sig")

    w = t[t["tiene_valor"]]

    # --- tables -----------------------------------------------------------
    def cats(s):
        return " | ".join(sorted(set(s.dropna().astype(str))))

    g = t.groupby("Parametro_archivo", sort=True)
    resumen_param = pd.DataFrame({
        "filas": g.size(),
        "sitios": g["CLAVE"].nunique(),
        "filas_con_valor": g["tiene_valor"].sum(),
        "sitios_con_valor": w.groupby("Parametro_archivo")["CLAVE"].nunique(),
        "n_censurados_<": w[w.Censura == "<"].groupby("Parametro_archivo").size(),
        "n_censurados_>": w[w.Censura == ">"].groupby("Parametro_archivo").size(),
        "min": w.groupby("Parametro_archivo")["Valor_num"].min(),
        "mediana": w.groupby("Parametro_archivo")["Valor_num"].median(),
        "max": w.groupby("Parametro_archivo")["Valor_num"].max(),
        "categorias_Calidad": w.groupby("Parametro_archivo")["Calidad_CONAGUA"].agg(cats),
    }).fillna({"sitios_con_valor": 0, "n_censurados_<": 0, "n_censurados_>": 0,
               "categorias_Calidad": ""})
    for c in ["filas", "sitios", "filas_con_valor", "sitios_con_valor",
              "n_censurados_<", "n_censurados_>"]:
        resumen_param[c] = resumen_param[c].astype(int)
    resumen_param["pct_sin_valor"] = (
        100 * (1 - resumen_param["filas_con_valor"] / resumen_param["filas"])).round(1)
    resumen_param["pct_censurado_<"] = (
        100 * resumen_param["n_censurados_<"]
        / resumen_param["filas_con_valor"].replace(0, np.nan)).round(1)
    resumen_param = resumen_param.sort_values("sitios_con_valor", ascending=False)

    site_names = t.groupby("CLAVE").agg(
        SITIO=("SITIO", "first"), MUNICIPIO=("MUNICIPIO", "first"),
        n_nombres_sitio=("SITIO", "nunique"),
        n_municipios=("MUNICIPIO", "nunique"))
    resumen_sitio = site_names.join(
        w.groupby("CLAVE")["Parametro_archivo"].nunique().rename("n_parametros_con_valor")
    ).fillna({"n_parametros_con_valor": 0}).sort_values(
        "n_parametros_con_valor", ascending=False)

    resumen_mun = t.groupby("MUNICIPIO").agg(sitios=("CLAVE", "nunique")).join(
        w.groupby("MUNICIPIO").agg(sitios_con_algun_valor=("CLAVE", "nunique"),
                                   mediciones_con_valor=("CLAVE", "size"))
    ).fillna(0).sort_values("sitios", ascending=False)

    calidad_x_param = pd.crosstab(w["Parametro_archivo"], w["Calidad_CONAGUA"])

    cob = (w.assign(v=1).pivot_table(index="CLAVE", columns="Parametro_archivo",
                                     values="v", aggfunc="max", fill_value=0))
    cob = site_names[["SITIO", "MUNICIPIO"]].join(cob, how="right")

    # --- checks -----------------------------------------------------------
    dup = t.duplicated(["CLAVE", "Parametro_archivo"], keep=False).sum()
    shared_codes = (t.groupby("Parametro_codigo")["Parametro_archivo"].nunique()
                    .loc[lambda s: s > 1])
    date_cols = [c for c in df.columns
                 if re.search(r"fecha|date|a[nñ]o|periodo|\bmes\b", c, re.I)]

    # file name vs parameter code
    n_files_per_code = t.groupby("Parametro_codigo")["Parametro_archivo"].nunique()
    arch = t.groupby("Parametro_archivo").agg(
        codigo=("Parametro_codigo", lambda s: ",".join(sorted(set(map(str, s))))),
        periodo_hoja=("Periodo_hoja", lambda s: ",".join(sorted(set(map(str, s))))),
        filas=("CLAVE", "size"), filas_con_valor=("tiene_valor", "sum"))
    def flag(name, codes):
        notes = []
        for c in codes.split(","):
            kws = CODE_KEYWORDS.get(c)
            if kws is None:
                notes.append(f"codigo '{c}' no reconocido")
            elif not any(k in strip_accents(name) for k in kws):
                notes.append(f"REVISAR: el archivo se llama '{name}' pero el "
                             f"codigo es '{c}'")
            if n_files_per_code.get(c, 1) > 1:
                notes.append(f"codigo '{c}' compartido por varios archivos: "
                             f"usar el nombre del archivo")
        return " | ".join(notes)
    arch["revisar"] = [flag(n, r.codigo) for n, r in arch.iterrows()]
    arch["filas_con_valor"] = arch["filas_con_valor"].astype(int)
    flagged = arch[arch["revisar"].str.contains("REVISAR")]

    tipo_counts = t.drop_duplicates("CLAVE")["Tipo_agua"].value_counts()
    checks = pd.DataFrame([
        ("Period (sheet names = 'Hoja')",
         ", ".join(sorted(set(map(str, t["Periodo_hoja"])))) if has_hoja
         else "no 'Hoja' column - merge with merge_xlsx.py v2"),
        ("Columns that look like a date/period",
         ", ".join(date_cols) or "none (no sampling dates in the data)"),
        ("Sites by water type (file membership)",
         ", ".join(f"{k}: {v}" for k, v in tipo_counts.items())),
        ("File name vs code mismatches",
         "; ".join(f"{n} (code {r.codigo})" for n, r in flagged.iterrows()) or "none"),
        ("Distinct ESTADO values", ", ".join(map(str, t["ESTADO"].dropna().unique()))),
        ("Distinct RHA values", ", ".join(map(str, t["RHA"].dropna().unique()))),
        ("Rows duplicated on (CLAVE, parameter file)", int(dup)),
        ("CLAVE with more than one SITIO name", int((site_names.n_nombres_sitio > 1).sum())),
        ("CLAVE with more than one MUNICIPIO", int((site_names.n_municipios > 1).sum())),
        ("Values that could not be parsed (error)", int((t.Censura == "error").sum())),
        ("Parameter files with zero values", ", ".join(
            resumen_param.index[resumen_param.filas_con_valor == 0]) or "none"),
        ("Parameter codes shared by 2+ files",
         ", ".join(f"{k} ({v} files)" for k, v in shared_codes.items()) or "none"),
        ("Distinct sites (CLAVE) overall", int(t["CLAVE"].nunique())),
        ("Distinct municipalities", int(t["MUNICIPIO"].nunique())),
        ("Parameter files", int(t["Parametro_archivo"].nunique())),
    ], columns=["check", "result"])

    sheets = {
        "chequeos": checks,
        "archivos_y_codigos": arch.reset_index(),
        "resumen_parametros": resumen_param.reset_index(),
        "resumen_sitios": resumen_sitio.reset_index(),
        "resumen_municipios": resumen_mun.reset_index(),
        "calidad_x_parametro": calidad_x_param.reset_index(),
        "cobertura_sitio_x_param": cob.reset_index(),
    }
    with pd.ExcelWriter(outdir / "perfil_calidad_datos.xlsx", engine="openpyxl") as xw:
        for name, d in sheets.items():
            d.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0
                            for c in col[:60])
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 8), 60)

    # --- console summary --------------------------------------------------
    pd.set_option("display.width", 200, "display.max_columns", 20,
                  "display.max_colwidth", 38)
    print("\n== CHECKS ==")
    print(checks.to_string(index=False))
    print("\n== PARAMETERS (sorted by sites with a value) ==")
    print(resumen_param[["sitios", "sitios_con_valor", "pct_sin_valor",
                         "pct_censurado_<"]].to_string())
    print(f"\nSaved: {outdir / 'datos_tidy.csv'}")
    print(f"Saved: {outdir / 'perfil_calidad_datos.xlsx'}  (sheets: {', '.join(sheets)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
