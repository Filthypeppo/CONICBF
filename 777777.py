# -*- coding: utf-8 -*-
"""
Versión extendida: permite seleccionar carpetas completas con planillas .xlsx
"""

import os
from pathlib import Path
import datetime
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilenames, asksaveasfilename, askdirectory
from tkinter.messagebox import showinfo, askyesno
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
MESES_NOMBRES = {
    1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril", 5:"Mayo", 6:"Junio",
    7:"Julio", 8:"Agosto", 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"
}

def month_number(mes_str):
    meses = {
        "enero":1,"ene":1,"febrero":2,"feb":2,"marzo":3,"mar":3,"abril":4,"abr":4,
        "mayo":5,"may":5,"junio":6,"jun":6,"julio":7,"jul":7,"agosto":8,"ago":8,
        "septiembre":9,"setiembre":9,"sep":9,"set":9,"octubre":10,"oct":10,
        "noviembre":11,"nov":11,"diciembre":12,"dic":12
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

        self.btn_add = ttk.Button(frm_top, text="➕ Agregar archivos", command=self.add_files)
        self.btn_add.pack(side="right", padx=5)

        self.btn_run = ttk.Button(frm_top, text="▶️ Procesar archivos", command=self.run_analysis)
        self.btn_run.pack(side="right", padx=5)

        self.btn_export = ttk.Button(frm_top, text="💾 Exportar resultados", command=self.export_results)
        self.btn_export.pack(side="right", padx=5)

        self.btn_open = ttk.Button(frm_top, text="📂 Abrir archivos", command=self.load_new_files)
        self.btn_open.pack(side="right", padx=5)

        self.text = tk.Text(self, wrap="word")
        self.text.pack(fill="both", expand=True, padx=8, pady=(0,8))
        self.scroll = ttk.Scrollbar(self.text, command=self.text.yview)
        self.text.configure(yscrollcommand=self.scroll.set)
        self.scroll.pack(side="right", fill="y")

        self.status = ttk.Label(self, text="Esperando inicio...", anchor="w")
        self.status.pack(fill="x", padx=8, pady=(0,8))

        self.protocol("WM_DELETE_WINDOW", self.on_quit)
        self._closing = False

        # --- log en memoria ---
        self.log_lines = []
        self.Alfa = None
        self.Total_HH = None
        self.Omega = None
        self.selected_files = []

    def log(self, msg):
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.log_lines.append(msg)
        self.update_idletasks()

    def set_status(self, msg):
        self.status.config(text=msg)
        self.update_idletasks()

    def on_quit(self):
        if askyesno("Salir", "¿Deseas cerrar la aplicación?"):
            self._closing = True
            self.destroy()

    def choose_files_or_folder(self):
        win = tk.Toplevel(self)
        win.title("Seleccionar tipo de entrada")
        ttk.Label(win, text="¿Qué deseas seleccionar?", font=("Segoe UI", 10)).pack(padx=20, pady=10)

        result = tk.StringVar(value="")

        def choose_files():
            result.set("files")
            win.destroy()

        def choose_folder():
            result.set("folder")
            win.destroy()

        ttk.Button(win, text="Archivos individuales", command=choose_files).pack(fill="x", padx=20, pady=5)
        ttk.Button(win, text="Carpeta completa", command=choose_folder).pack(fill="x", padx=20, pady=5)

        win.wait_window()
        return result.get()

    def find_excel_in_folder(self, folder):
        excel_files = []
        for root, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".xlsx"):
                    excel_files.append(Path(root)/f)
        return excel_files

    def load_new_files(self):
        self.selected_files.clear()
        self.add_files()

    def add_files(self):
        choice = self.choose_files_or_folder()
        if not choice:
            self.log("Selección cancelada.")
            return

        new_paths = []
        if choice == "files":
            files = askopenfilenames(title="Seleccionar planillas", filetypes=[("Excel files","*.xlsx")])
            new_paths = [Path(f) for f in files]
        elif choice == "folder":
            folder = askdirectory(title="Seleccionar carpeta con planillas")
            if not folder:
                self.log("Selección de carpeta cancelada.")
                return
            new_paths = self.find_excel_in_folder(folder)
            self.log(f"📁 Carpeta seleccionada: {folder}")
            self.log(f"  → {len(new_paths)} archivos .xlsx encontrados.")

        if not new_paths:
            self.log("No se encontraron archivos válidos.")
            return

        added = 0
        for p in new_paths:
            if p not in self.selected_files:
                self.selected_files.append(p)
                added += 1

        self.log(f"🧾 Total de archivos seleccionados: {len(self.selected_files)} (añadidos: {added})")

    def run_analysis(self):
        if not self.selected_files:
            showinfo("Sin archivos", "No hay archivos seleccionados.")
            return
        self.after(100, lambda: main_process(self, self.selected_files))

    def export_results(self):
        if self.Omega is None:
            showinfo("Sin datos", "No hay resultados disponibles para exportar.")
            return
        dest = asksaveasfilename(
            title="Guardar archivo de resumen (.xlsx)",
            defaultextension=".xlsx",
            filetypes=[("Excel files","*.xlsx")],
            initialfile="Resumen.xlsx"
        )
        if not dest:
            self.log("Exportación cancelada.")
            return
        try:
            Omega_clean = self.Omega.loc[:, self.Omega.columns.notna()]
            Omega_clean = Omega_clean.loc[:, Omega_clean.columns != ""]

            with pd.ExcelWriter(dest, engine="openpyxl") as writer:
                Omega_clean.to_excel(writer, sheet_name="Registro", index=False)
                for name in self.Total_HH["Name"].unique():
                    aux = self.Total_HH[self.Total_HH["Name"]==name].dropna(how='all')
                    aux.to_excel(writer, sheet_name=str(name)[:31], index=False)

            dest_path = Path(dest).resolve().parent
            log_path = dest_path / "log.txt"
            with open(log_path, "w", encoding="utf-8") as f:
                for line in self.log_lines:
                    f.write(line + "\n")

            self.log(f"✅ Exportado: {Path(dest).name} y log.txt (en {dest_path})")
            showinfo("Exportación completada", f"Archivos exportados:\n{dest}\n{log_path}")
        except Exception as e:
            self.log(f"❌ Error al exportar: {e}")
            showinfo("Error", f"No fue posible exportar: {e}")

# ---------------- funciones auxiliares ----------------
def a1_notation(row, col):
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
    try:
        month_number(df.iloc[1,3])
    except Exception as e:
        errors.append(f"{path.name}: error en celda {a1_notation(1,3)} -> {e}")
    try:
        int(df.iloc[1,11])
    except Exception as e:
        errors.append(f"{path.name}: error en celda {a1_notation(1,11)} -> {e}")
    return (df if not errors else None), errors

# ---------------- análisis ----------------
def analyze(selected_files, holidays_freq, log_ui):
    valid_links, dfs, all_errors = [], [], []

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
        return None, None, None, all_errors

    Rg = pd.DataFrame()
    for idx, p in enumerate(valid_links):
        df = dfs[idx]
        try:
            Name = df.columns[4]
            Month = month_number(df.iloc[1,3])
            Year = int(df.iloc[1,11])
            Proyectos = df.iloc[3,15:34].reset_index(drop=True).tolist()
            Total_Hours = df.iloc[37,15:34].reset_index(drop=True).tolist()
            cols = ["Name","Month","Year"] + Proyectos
            row = [Name, Month, Year] + Total_Hours
            Rg = pd.concat([Rg, pd.DataFrame([dict(zip(cols,row))])], ignore_index=True)
        except Exception as e:
            log_ui.log(f"Error procesando {p.name}: {e}")

    if Rg.empty:
        return None, None, None, all_errors

    Rg = Rg.loc[:, Rg.columns.notna()]
    Rg = Rg.loc[:, Rg.columns != ""]

    Alfa = Rg.groupby(["Name","Month","Year"], as_index=False).sum(numeric_only=True)
    Total_HH = Alfa.loc[:, ["Name","Month","Year"]].copy()
    Total_HH["Horas Realizadas"] = Alfa.loc[:, Alfa.columns.difference(["Name","Month","Year"])].sum(axis=1).round(2)

    Whours = []
    for _, row in Total_HH.iterrows():
        Year, Month = int(row["Year"]), int(row["Month"])
        start = datetime.date(Year, Month, 1)
        end = datetime.date(Year + (Month==12), Month % 12 + 1, 1)
        Workdays = np.busday_count(start,end)
        hol_row = holidays_freq[(holidays_freq["Year"]==Year)&(holidays_freq["Month"]==Month)]
        Holydays = int(hol_row["Holidays"].iloc[0]) if not hol_row.empty else 0
        Whours.append(8*(Workdays-Holydays))
    Total_HH["Horas objetivo*"] = np.round(Whours,2)

    Omega = Alfa.copy()
    Omega["Aux"] = Omega["Name"].astype(str) + Omega["Year"].astype(str)
    Omega = Omega.drop(columns=["Name","Year","Month"], errors='ignore')
    Omega = Omega.groupby("Aux", as_index=False).sum()
    Omega.insert(0,"Name",Omega["Aux"].str[:-4])
    Omega.insert(1,"Year",Omega["Aux"].str[-4:])
    Omega["Year"] = pd.to_numeric(Omega["Year"], errors='coerce').fillna(0).astype(int)
    Omega = Omega.drop(columns=["Aux"], errors='ignore')
    Omega = Omega.loc[:, Omega.columns.notna()]
    Omega = Omega.loc[:, Omega.columns != ""]

    return Alfa, Total_HH, Omega, all_errors

# ---------------- proceso principal ----------------
def main_process(log_ui, selected_files):
    try:
        log_ui.set_status("Cargando feriados...")
        years = set()
        for f in selected_files:
            try:
                yy = int(pd.read_excel(f, engine="openpyxl").iloc[1,11])
                years.add(yy)
            except Exception:
                continue
        if not years:
            years = {datetime.date.today().year}
        years = sorted(years)

        holidays = fetch_holidays_chile(years)
        holidays_freq = working_holidays_frequency(holidays)
        log_ui.log("Feriados cargados correctamente.\n")

        for y in years:
            log_ui.log(f"📅 Feriados año {y}:")
            for m in range(1,13):
                row = holidays_freq[(holidays_freq["Year"]==y)&(holidays_freq["Month"]==m)]
                count = int(row["Holidays"].iloc[0]) if not row.empty else 0
                if count>0:
                    log_ui.log(f"  {MESES_NOMBRES[m]}: {count} día(s) hábil(es) feriado(s)")
            log_ui.log("")

        log_ui.set_status("Analizando archivos...")
        Alfa, Total_HH, Omega, errors = analyze(selected_files, holidays_freq, log_ui)

        if errors:
            log_ui.log(f"\n⚠️ {len(errors)} error(es) detectado(s):")
            for e in errors:
                log_ui.log(f"  {e}")

        if Omega is not None:
            log_ui.Alfa, log_ui.Total_HH, log_ui.Omega = Alfa, Total_HH, Omega
            log_ui.log("\n✅ Análisis completado. Usa '💾 Exportar resultados' para guardar.")
        else:
            log_ui.log("\n❌ No se pudo generar el resumen. Revisa el log.")
        log_ui.set_status("Proceso finalizado.")
        showinfo("Finalizado", "Análisis completado. Puedes exportar los resultados.")
    except Exception as e:
        log_ui.log(f"Error durante la ejecución: {e}")

# ---------------- arranque ----------------
def main():
    ui = LogWindow()
    ui.mainloop()

if __name__ == "__main__":
    main()
