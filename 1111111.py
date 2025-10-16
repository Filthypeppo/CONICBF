# -*- coding: utf-8 -*-
"""
Created on Tue Oct 14 17:00:25 2025
Refactor y correcciones por ChatGPT
"""
import os
from pathlib import Path
import requests
import certifi
from dateutil.parser import parse
import datetime
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter.messagebox import askyesno, showinfo

# --------- Utilidades ----------
def fetch_holidays_chile(years=None):
    """
    Intenta obtener feriados desde apis.digital.gob.cl; si falla por SSL usa Nager.Date.
    Devuelve DataFrame con columnas ['fecha'] (datetime)
    """
    if years is None:
        years = [datetime.date.today().year]
    # Try primary API
    url_primary = "https://apis.digital.gob.cl/fl/feriados"
    try:
        resp = requests.get(url_primary, headers={"User-Agent":"My User Agent 1.0"}, verify=certifi.where(), timeout=10)
        resp.raise_for_status()
        data = resp.json()
        fechas = [parse(d["fecha"]).date() for d in data]
        return pd.DataFrame({"fecha": fechas})
    except Exception:
        # Fallback: Nager.Date
        all_dates = []
        for y in years:
            try:
                url = f"https://date.nager.at/api/v3/PublicHolidays/{y}/CL"
                r = requests.get(url, timeout=10)
                r.raise_for_status()
                js = r.json()
                all_dates.extend([parse(i["date"]).date() for i in js])
            except Exception:
                continue
        return pd.DataFrame({"fecha": all_dates})

def working_holidays_frequency(holidays_df):
    """
    Devuelve DataFrame con Year, Month, n_feriados_en_mes (solo dias entre lunes-viernes)
    """
    if holidays_df.empty:
        return pd.DataFrame(columns=["Year","Month","Holidays"])
    holidays_df = holidays_df.copy()
    holidays_df["weekday"] = holidays_df["fecha"].apply(lambda d: d.weekday())
    # keep only Mon-Fri
    holidays_df = holidays_df[holidays_df["weekday"] <= 4]
    holidays_df["Year"] = holidays_df["fecha"].apply(lambda d: d.year)
    holidays_df["Month"] = holidays_df["fecha"].apply(lambda d: d.month)
    freq = holidays_df.groupby(["Year","Month"]).size().reset_index(name="Holidays")
    return freq

def month_number(mes_str):
    meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"]
    try:
        return meses.index(str(mes_str).lower()) + 1
    except ValueError:
        raise ValueError(f"Mes no reconocido: {mes_str}")

# --------- Recolección de archivos ----------
def find_xlsx_files(root_dir):
    links = []
    for root, dirs, files in os.walk(root_dir):
        for f in files:
            if f.lower().endswith(".xlsx"):
                links.append(Path(root) / f)
    print(f"se encontraron {len(links)} posibles registros HH\n")
    print("------------------------------------------------------------\n")
    return links

# --------- Verificación de formato ----------
def verify_format(links):
    """
    Verifica que la celda [1,15] (fila 2, columna 16 en base 0) tenga el texto esperado.
    Devuelve lista de archivos con formato incorrecto.
    """
    bad = []
    for p in links:
        try:
            df = pd.read_excel(p, engine="openpyxl")
            # iloc[row_index, col_index]
            if df.iloc[1, 15] != "D E S G L O S E    P O R    P R O Y E C T O":
                bad.append(p)
        except Exception as e:
            bad.append(p)
    return bad

# --------- Eliminación de duplicados por usuario+mes+año ----------
def resolve_duplicates(links, links_interferencias_df):
    """
    links_interferencias_df: DataFrame con columnas ['Links','Identificador']
    Pide confirmación al usuario y elimina archivos más antiguos dejando el más reciente.
    Devuelve lista de links resultantes.
    """
    root = tk.Tk()
    root.withdraw()
    # Construir mensaje
    files = links_interferencias_df['Links'].apply(lambda p: Path(p).name).tolist()
    message = "Los siguientes archivos corresponden a los mismos usuarios y fechas:\n" + "\n".join(f"* {n}" for n in files)
    res = askyesno(title="Error!", message=message + "\n\nDesea dejar solo los registros mas recientes?")
    root.destroy()
    if not res:
        showinfo(title="INFO", message="No se han eliminado registros duplicados, por favor verificar y volver a ejecutar")
        return links  # sin cambios

    # Para cada identificador, conservar el más reciente (por ctime)
    to_delete = set()
    for ident in links_interferencias_df['Identificador'].unique():
        sub = links_interferencias_df[links_interferencias_df['Identificador'] == ident]
        ctimes = [(str(p), os.path.getctime(str(p))) for p in sub['Links']]
        # ordenar por ctime desc -> primer elemento es el mas reciente
        ctimes.sort(key=lambda x: x[1], reverse=True)
        # conservar el primero, eliminar el resto
        for path_str, _ in ctimes[1:]:
            to_delete.add(Path(path_str))
    # eliminar archivos
    for p in to_delete:
        try:
            p.unlink()
        except Exception as e:
            print(f"Error eliminando {p}: {e}")
    # return remaining links
    remaining = [p for p in links if Path(p) not in to_delete]
    return remaining

# --------- Análisis principal ----------
def analyze(links, holidays_freq):
    """
    Procesa los archivos y retorna (Alfa, Total_HH, Suma)
    """
    # Verificación inicial
    bad = verify_format(links)
    if bad:
        print(">>> Las siguientes planillas no tienen el formato correcto, verificar. <<<\n")
        for idx, b in enumerate(bad):
            print(f"N {idx} {b}")
        # removemos los archivos malos de la lista
        links = [l for l in links if l not in bad]
        if not links:
            raise SystemExit("No hay archivos válidos para procesar.")

    Rg = pd.DataFrame()
    # Primera pasada: construir Rg con columnas por proyecto y enlazar Links
    for p in links:
        df = pd.read_excel(p, engine="openpyxl")
        Name = df.columns[4]
        Month_raw = df.iloc[1,3]
        Month = month_number(Month_raw)
        Year = int(df.iloc[1,11])
        Proyectos = df.iloc[3, 15:34].reset_index(drop=True).tolist()
        Total_Hours = df.iloc[37, 15:34].reset_index(drop=True).tolist()
        # armar columnas y fila
        cols = ["Name","Month","Year"] + Proyectos
        row = [Name, Month, Year] + Total_Hours
        # insertar link como columna 'Links' en la posición relativa (separé al final)
        row_dict = dict(zip(cols, row))
        row_dict["Links"] = str(p)
        betaux = pd.DataFrame([row_dict])
        # agrupar columnas repetidas (si las hay)
        betaux = betaux.loc[:, betaux.columns.notna()]
        betaux = betaux.groupby(betaux.columns, axis=1).sum()
        # quitar columna "TOTAL" si existe
        if "TOTAL" in betaux.columns:
            betaux = betaux.loc[:, betaux.columns != "TOTAL"]
        Rg = pd.concat([Rg, betaux], ignore_index=True, sort=False)
        Rg.fillna(0, inplace=True)

    # Detectar duplicados por Aux = Name+Month+Year
    Rg["Aux"] = Rg["Name"].astype(str) + Rg["Month"].astype(str) + Rg["Year"].astype(str)
    aux_count = Rg["Aux"].value_counts().to_frame(name="count").reset_index().rename(columns={"index":"Aux"})
    dup = aux_count[aux_count["count"] > 1]
    if not dup.empty:
        # construir links_interferencias
        links_inter = pd.DataFrame(columns=["Links","Identificador"])
        for idx, row in dup.reset_index(drop=True).iterrows():
            aux_val = row["Aux"]
            subset = Rg[Rg["Aux"] == aux_val][["Links"]].copy()
            subset["Identificador"] = idx
            links_inter = pd.concat([links_inter, subset], ignore_index=True)
        # Resolver con ventana
        links = resolve_duplicates(links, links_inter)

    # Segunda pasada: construir Alfa, Total_HH, Suma
    Alfa = pd.DataFrame()
    for p in links:
        df = pd.read_excel(p, engine="openpyxl")
        Name = df.columns[4]
        Month = month_number(df.iloc[1,3])
        Year = int(df.iloc[1,11])
        Proyectos = df.iloc[3, 15:34].reset_index(drop=True).tolist()
        Total_Hours = df.iloc[37, 15:34].reset_index(drop=True).tolist()
        cols = ["Name","Month","Year"] + Proyectos
        row = [Name, Month, Year] + Total_Hours
        beta = pd.DataFrame([row], columns=cols)
        # eliminar columna TOTAL si existe
        if "TOTAL" in beta.columns:
            beta = beta.loc[:, beta.columns != "TOTAL"]
        beta = beta.groupby(beta.columns, axis=1).sum()
        Alfa = pd.concat([Alfa, beta], ignore_index=True, sort=False)
        Alfa.fillna(0, inplace=True)

    # Reorganizar columnas: Name,Month,Year,...
    if Alfa.empty:
        raise SystemExit("No hay datos en Alfa.")
    columnas = [c for c in Alfa.columns if c not in ["Name","Month","Year"]]
    columnas = ["Name","Month","Year"] + columnas
    Alfa = Alfa.loc[:, columnas]

    Total_HH = Alfa.loc[:, ["Name","Month","Year"]].copy()
    Total_HH["Horas Realizadas"] = Alfa.loc[:, Alfa.columns.difference(["Name","Month","Year"])].sum(axis=1).round(2)

    # Calcular Workdays y Horas objetivo
    Whours = []
    Suma = Alfa.sum(axis=0).to_frame(name="Total")
    if set(["Name","Month","Year"]).issubset(Suma.index):
        Suma.drop(labels=["Name","Month","Year"], inplace=True, errors='ignore')

    for _, row in Total_HH.iterrows():
        Year = int(row["Year"])
        Month = int(row["Month"])
        start = datetime.date(year=Year, month=Month, day=1)
        if Month < 12:
            end = datetime.date(year=Year, month=(Month + 1), day=1)
        else:
            end = datetime.date(year=(Year + 1), month=1, day=1)
        Workdays = np.busday_count(start, end)
        # obtener feriados en ese mes (si existen)
        hol_row = holidays_freq[(holidays_freq["Year"] == Year) & (holidays_freq["Month"] == Month)]
        Holydays = int(hol_row["Holidays"].iloc[0]) if not hol_row.empty else 0
        Whours.append(8 * (Workdays - Holydays))

    Total_HH["Horas objetivo*"] = np.round(Whours, 2)
    Total_HH["Horas Adicionales"] = ""
    Total_HH["Horas Autorizaadas"] = ""
    Total_HH["Horas Pagadas"] = ""
    Total_HH["Diferencia"] = ""

    Alfa.iloc[:, 2:] = Alfa.iloc[:, 2:].round(2)
    Total_HH.iloc[:, 2:] = Total_HH.iloc[:, 2:].round(2)

    return Alfa.reset_index(drop=True), Total_HH.reset_index(drop=True), Suma

# --------- Main ----------
def main():
    try:
        # Obtener feriados (consultar año actual y próximos por si hay archivos de distintos años)
        current_year = datetime.date.today().year
        holidays = fetch_holidays_chile(years=[current_year, current_year+1])
        holidays_freq = working_holidays_frequency(holidays)

        # Buscar archivos
        links = find_xlsx_files(os.getcwd())
        if not links:
            print("No se encontraron archivos .xlsx en el directorio actual.")
            return

        # Analizar
        Alfa, Resumen1, Suma = analyze(links, holidays_freq)

        # Consolidar Omega
        Omega = Alfa.copy(deep=True)
        Omega["Aux"] = Omega["Name"].astype(str) + Omega["Year"].astype(str)
        # eliminar columnas originales
        Omega = Omega.drop(columns=["Name","Year","Month"], errors='ignore')
        Omega = Omega.groupby("Aux", as_index=False).sum()
        Omega.insert(0, "Name", Omega["Aux"].str[:-4])
        Omega.insert(1, "Year", Omega["Aux"].str[-4:])
        Omega["Year"] = pd.to_numeric(Omega["Year"], errors='coerce').fillna(0).astype(int)
        Omega = Omega.drop(columns=["Aux"], errors='ignore')
        Omega.reset_index(drop=True, inplace=True)

        # Escribir Excel
        with pd.ExcelWriter("Resumen.xlsx", engine="openpyxl") as writer:
            Omega.to_excel(writer, sheet_name="Registro", index=False)
            for name in Resumen1["Name"].unique():
                aux = Resumen1[Resumen1["Name"] == name].dropna(how='all', axis=0)
                aux.to_excel(writer, sheet_name=str(name)[:31], index=False)  # sheet name max 31 chars

        print("Resumen.xlsx creado correctamente.")
    except Exception as e:
        print("Error durante la ejecución:", e)

if __name__ == "__main__":
    main()
