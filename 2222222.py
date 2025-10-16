# -*- coding: utf-8 -*-
"""
Refactor con Tkinter UI y detección de celdas erradas
"""

import os
from pathlib import Path
import requests, certifi
from dateutil.parser import parse
import datetime
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showinfo, askyesno

# =============================================================
# --- Utilidades ---
# =============================================================

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
    holidays_df["weekday"] = holidays_df["fecha"].apply(lambda d: d.weekday())
    holidays_df = holidays_df[holidays_df["weekday"] <= 4]
    holidays_df["Year"] = holidays_df["fecha"].apply(lambda d: d.year)
    holidays_df["Month"] = holidays_df["fecha"].apply(lambda d: d.month)
    return holidays_df.groupby(["Year","Month"]).size().reset_index(name="Holidays")

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

# =============================================================
# --- Tkinter GUI Helper ---
# =============================================================

class LogWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Análisis de planillas HH")
        self.geometry("800x600")
        self.resizable(True, True)

        ttk.Label(self, text="Progreso del análisis", font=("Segoe UI", 12, "bold")).pack(pady=5)
        self.text = tk.Text(self, wrap="word", height=30)
        self.text.pack(fill="both", expand=True, padx=10, pady=10)
        self.scroll = ttk.Scrollbar(self.text, command=self.text.yview)
        self.text.configure(yscrollcommand=self.scroll.set)
        self.scroll.pack(side="right", fill="y")
        self.update()

    def log(self, message):
        self.text.insert("end", f"{message}\n")
        self.text.see("end")
        self.update()

# =============================================================
# --- Lógica de análisis ---
# =============================================================

def find_xlsx_files(root_dir):
    return [Path(root)/f for root,_,files in os.walk(root_dir) for f in files if f.lower().endswith(".xlsx")]

def verify_and_load_excel(path, log_ui):
    """Carga un Excel y detecta celdas erradas"""
    try:
        df = pd.read_excel(path, engine="openpyxl")
        # validar la celda clave
        if df.iloc[1,15] != "D E S G L O S E    P O R    P R O Y E C T O":
            log_ui.log(f"⚠️ {path.name}: formato incorrecto (celda B2 esperada con título de desglose)")
            return None
        return df
    except Exception as e:
        log_ui.log(f"❌ Error leyendo '{path.name}': {e}")
        # intentar identificar celda conflictiva
        try:
            wb = pd.read_excel(path, header=None, engine="openpyxl")
            # buscar valores no válidos en las filas clave
            for row in range(0, 5):
                for col in range(0, 20):
                    val = str(wb.iloc[row, col])
                    if "error" in val.lower() or val.strip() == "nan":
                        log_ui.log(f"   → Posible problema en celda ({row+1},{col+1}) → '{val}'")
        except Exception:
            log_ui.log("   (no fue posible analizar el contenido por error grave)")
        return None

def analyze_with_ui(log_ui):
    current_year = datetime.date.today().year
    holidays = fetch_holidays_chile(years=[current_year, current_year+1])
    holidays_freq = working_holidays_frequency(holidays)

    log_ui.log("🔍 Buscando archivos Excel...")
    links = find_xlsx_files(os.getcwd())
    log_ui.log(f"→ {len(links)} archivos encontrados.\n")

    if not links:
        showinfo("Sin archivos", "No se encontraron planillas .xlsx en el directorio actual.")
        return

    valid_links, dataframes = [], []
    for p in links:
        df = verify_and_load_excel(p, log_ui)
        if df is not None:
            valid_links.append(p)
            dataframes.append(df)
        else:
            log_ui.log(f"⚠️ Archivo omitido: {p.name}\n")

    if not valid_links:
        showinfo("Error", "No se encontraron planillas válidas para analizar.")
        return

    log_ui.log(f"✅ {len(valid_links)} archivos válidos procesando...\n")

    # Ejemplo simplificado del cálculo principal
    summary_rows = []
    for i, df in enumerate(dataframes):
        try:
            name = df.columns[4]
            mes = month_number(df.iloc[1,3])
            año = int(df.iloc[1,11])
            log_ui.log(f"Procesando: {valid_links[i].name} ({mes}/{año})")
            summary_rows.append([name, mes, año])
        except Exception as e:
            log_ui.log(f"❌ Error extrayendo datos en '{valid_links[i].name}': {e}")

    resumen = pd.DataFrame(summary_rows, columns=["Name","Month","Year"])
    resumen.to_excel("Resumen_simple.xlsx", index=False)
    log_ui.log("\n💾 Archivo 'Resumen_simple.xlsx' creado con información resumida.")
    showinfo("Proceso completado", "El análisis ha finalizado correctamente.")

# =============================================================
# --- Main ---
# =============================================================

def main():
    ui = LogWindow()
    ui.after(100, lambda: analyze_with_ui(ui))
    ui.mainloop()

if __name__ == "__main__":
    main()
