# -*- coding: utf-8 -*-
"""
Resumen completo con Tkinter y validación de formatos de datos
Solo verifica mes y año, muestra errores en pantalla y guarda log.txt
"""
import os
from pathlib import Path
import datetime
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilenames
from tkinter.messagebox import showinfo
from dateutil.parser import parse
import requests, certifi

# ---------------- utilidades de feriados ----------------
def fetch_holidays_chile(years=None):
    if years is None:
        years = [datetime.date.today().year]
    url_primary = "https://apis.digital.gob.cl/fl/feriados"
    try:
        resp = requests.get(url_primary, headers={"User-Agent":"My User Agent 1.0"}, verify=certifi.where(), timeout=10)
        resp.raise_for_status()
        data = resp.json()
        fechas = [parse(d["fecha"]).date() for d in data]
        return pd.DataFrame({"fecha": fechas})
    except Exception:
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
    if holidays_df.empty:
        return pd.DataFrame(columns=["Year","Month","Holidays"])
    holidays_df = holidays_df.copy()
    holidays_df["weekday"] = holidays_df["fecha"].apply(lambda d: d.weekday())
    holidays_df = holidays_df[holidays_df["weekday"] <= 4]
    holidays_df["Year"] = holidays_df["fecha"].apply(lambda d: d.year)
    holidays_df["Month"] = holidays_df["fecha"].apply(lambda d: d.month)
    return holidays_df.groupby(["Year","Month"]).size().reset_index(name="Holidays")

# ---------------- conversión de mes ----------------
def month_number(mes_str):
    meses = {
        "enero":1,"ene":1,
        "febrero":2,"feb":2,
        "marzo":3,"mar":3,
        "abril":4,"abr":4,
        "mayo":5,"may":5,
        "junio":6,"jun":6,
        "julio":7,"jul":7,
        "agosto":8,"ago":8,
        "septiembre":9,"setiembre":9,"sep":9,"set":9,
        "octubre":10,"oct":10,
        "noviembre":11,"nov":11,
        "diciembre":12,"dic":12
    }
    mes = str(mes_str).strip().lower()
    if mes not in meses:
        raise ValueError(f"Mes no reconocido: {mes_str}")
    return meses[mes]

# ---------------- GUI ----------------
class LogWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Análisis planillas HH")
        self.geometry("900x650")
        self.minsize(700,500)
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", padx=8, pady=6)
        ttk.Label(frm_top, text="Progreso", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.btn_quit = ttk.Button(frm_top, text="Salir", command=self.on_quit)
        self.btn_quit.pack(side="right")
        self.text = tk.Text(self, wrap="word")
        self.text.pack(fill="both", expand=True, padx=8, pady=(0,8))
        self.scroll = ttk.Scrollbar(self.text, command=self.text.yview)
        self.text.configure(yscrollcommand=self.scroll.set)
        self.scroll.pack(side="right", fill="y")
        self.status = ttk.Label(self, text="Esperando inicio...", anchor="w")
        self.status.pack(fill="x", padx=8, pady=(0,8))
        self.protocol("WM_DELETE_WINDOW", self.on_quit)
        self._closing = False
        self.log_file = open("log.txt","w",encoding="utf-8")

    def log(self, msg):
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.log_file.write(msg + "\n")
        self.log_file.flush()
        self.update_idletasks()

    def set_status(self, msg):
        self.status.config(text=msg)
        self.update_idletasks()

    def on_quit(self):
        if tk.messagebox.askyesno("Salir", "¿Deseas cerrar la aplicación?"):
            self._closing = True
            self.log_file.close()
            self.destroy()

# ---------------- inspección de archivos ----------------
def a1_notation(row, col):
    """Convertir fila y columna base 0 a notación Excel"""
    letters = ''
    while col >= 0:
        letters = chr(col % 26 + ord('A')) + letters
        col = col // 26 - 1
    return f"{letters}{row+1}"

def inspect_sheet_for_errors(path: Path):
    errors = []
    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as e:
        errors.append(f"{path.name}: Error al abrir archivo: {e}")
        return None, errors

    # Mes (D2)
    try:
        month_raw = df.iloc[1,3]
        month_number(month_raw)
    except Exception as e:
        errors.append(f"{path.name}: error en celda {a1_notation(1,3)} -> {e}")

    # Año (L2)
    try:
        year_raw = df.iloc[1,11]
        int(year_raw)
    except Exception as e:
        errors.append(f"{path.name}: error en celda {a1_notation(1,11)} -> {e}")

    if errors:
        return None, errors
    return df, []

# ---------------- análisis principal ----------------
def analyze(selected_files, holidays_freq, log_ui):
    valid_links = []
    dfs = []
    all_errors = []

    for p in selected_files:
        log_ui.log(f"Inspeccionando: {p.name}")
        df, errors = inspect_sheet_for_errors(p)
        if errors:
            all_errors.extend(errors)
            for e in errors:
                log_ui.log(f"  - {e}")
        if df is not None:
            valid_links.append(p)
            dfs.append(df)
            log_ui.log(f"  → Archivo válido: {p.name}")
        else:
            log_ui.log(f"  → Archivo omitido: {p.name}")

    if not valid_links:
        log_ui.log("No hay archivos válidos para procesar. Fin del análisis.")
        return None, None

    # Construir Omega y Alfa
    Rg = pd.DataFrame()
    for idx, p in enumerate(valid_links):
        df = dfs[idx]
        try:
            Name = df.columns[4]
            Month = month_number(df.iloc[1,3])
            Year = int(df.iloc[1,11])
            Proyectos = df.iloc[3, 15:34].reset_index(drop=True).tolist()
            Total_Hours = df.iloc[37, 15:34].reset_index(drop=True).tolist()
            cols = ["Name","Month","Year"] + Proyectos
            row = [Name, Month, Year] + Total_Hours
            row_dict = dict(zip(cols,row))
            row_dict["Links"] = str(p)
            betaux = pd.DataFrame([row_dict])
            if "TOTAL" in betaux.columns:
                betaux = betaux.drop(columns=["TOTAL"])
            Rg = pd.concat([Rg, betaux], ignore_index=True, sort=False)
            Rg.fillna(0, inplace=True)
        except Exception as e:
            log_ui.log(f"Error procesando {p.name}: {e}")

    # Segunda pasada: Alfa
    Alfa = pd.DataFrame()
    for idx, p in enumerate(valid_links):
        df = dfs[idx]
        Name = df.columns[4]
        Month = month_number(df.iloc[1,3])
        Year = int(df.iloc[1,11])
        Proyectos = df.iloc[3, 15:34].reset_index(drop=True).tolist()
        Total_Hours = df.iloc[37, 15:34].reset_index(drop=True).tolist()
        cols = ["Name","Month","Year"] + Proyectos
        row = [Name, Month, Year] + Total_Hours
        beta = pd.DataFrame([row], columns=cols)
        if "TOTAL" in beta.columns:
            beta = beta.drop(columns=["TOTAL"])
        beta = beta.groupby(beta.columns, axis=1).sum()
        Alfa = pd.concat([Alfa, beta], ignore_index=True, sort=False)
        Alfa.fillna(0,inplace=True)

    # Total_HH
    Total_HH = Alfa.loc[:, ["Name","Month","Year"]].copy()
    Total_HH["Horas Realizadas"] = Alfa.loc[:, Alfa.columns.difference(["Name","Month","Year"])].sum(axis=1).round(2)

    # Horas objetivo
    Whours = []
    for _, row in Total_HH.iterrows():
        Year = int(row["Year"])
        Month = int(row["Month"])
        start = datetime.date(year=Year, month=Month, day=1)
        if Month < 12:
            end = datetime.date(year=Year, month=(Month+1), day=1)
        else:
            end = datetime.date(year=Year+1, month=1, day=1)
        Workdays = np.busday_count(start,end)
        hol_row = holidays_freq[(holidays_freq["Year"]==Year)&(holidays_freq["Month"]==Month)]
        Holydays = int(hol_row["Holidays"].iloc[0]) if not hol_row.empty else 0
        Whours.append(8*(Workdays-Holydays))
    Total_HH["Horas objetivo*"] = np.round(Whours,2)

    # Omega
    Omega = Alfa.copy()
    Omega["Aux"] = Omega["Name"].astype(str) + Omega["Year"].astype(str)
    Omega = Omega.drop(columns=["Name","Year","Month"], errors='ignore')
    Omega = Omega.groupby("Aux", as_index=False).sum()
    Omega.insert(0,"Name",Omega["Aux"].str[:-4])
    Omega.insert(1,"Year",Omega["Aux"].str[-4:])
    Omega["Year"] = pd.to_numeric(Omega["Year"], errors='coerce').fillna(0).astype(int)
    Omega = Omega.drop(columns=["Aux"], errors='ignore')
    Omega.reset_index(drop=True,inplace=True)

    return Alfa, Total_HH, Omega, all_errors

# ---------------- proceso principal con UI ----------------
def main_process(log_ui):
    try:
        log_ui.set_status("Seleccionando archivos .xlsx...")
        files = askopenfilenames(title="Seleccionar planillas a analizar", filetypes=[("Excel files","*.xlsx")])
        if not files:
            log_ui.log("No se seleccionaron archivos. Abortando.")
            return
        selected_files = [Path(f) for f in files]

        log_ui.set_status("Cargando feriados...")
        current_year = datetime.date.today().year
        holidays = fetch_holidays_chile(years=[current_year,current_year+1])
        holidays_freq = working_holidays_frequency(holidays)
        log_ui.log("Feriados cargados.")

        log_ui.set_status("Analizando archivos...")
        Alfa, Total_HH, Omega, errors = analyze(selected_files, holidays_freq, log_ui)

        if errors:
            log_ui.log(f"\nSe detectaron {len(errors)} errores de formato:")
            for e in errors:
                log_ui.log(f"  {e}")

        if Omega is not None:
            log_ui.set_status("Guardando Resumen.xlsx...")
            try:
                with pd.ExcelWriter("Resumen.xlsx", engine="openpyxl") as writer:
                    Omega.to_excel(writer, sheet_name="Registro", index=False)
                    for name in Total_HH["Name"].unique():
                        aux = Total_HH[Total_HH["Name"]==name].dropna(how='all')
                        aux.to_excel(writer, sheet_name=str(name)[:31], index=False)
                log_ui.log("Resumen.xlsx creado correctamente.")
            except Exception as e:
                log_ui.log(f"Error guardando Resumen.xlsx: {e}")

        log_ui.set_status("Proceso finalizado.")
        showinfo("Finalizado", "Análisis completado. Revisa la ventana y el log.txt para más detalles.")

    except Exception as e:
        log_ui.log(f"Error durante la ejecución: {e}")

# ---------------- arranque ----------------
def main():
    ui = LogWindow()
    ui.after(200, lambda: main_process(ui))
    ui.mainloop()

if __name__ == "__main__":
    main()
