# -*- coding: utf-8 -*-
"""
Resumen completo con Tkinter y detección de errores por celda
Guarda:
 - Resumen.xlsx  (Omega + hojas por usuario)
 - Errores_planillas.xlsx (detalle de errores detectados)
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

# ---------------- GUI de log ----------------
class LogWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Análisis de planillas HH - Resumen completo")
        self.geometry("900x650")
        self.minsize(700,500)
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", padx=8, pady=6)
        ttk.Label(frm_top, text="Progreso", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.btn_quit = ttk.Button(frm_top, text="Salir", command=self.on_quit)
        self.btn_quit.pack(side="right")
        # Text con scroll
        self.text = tk.Text(self, wrap="word")
        self.text.pack(fill="both", expand=True, padx=8, pady=(0,8))
        self.scroll = ttk.Scrollbar(self.text, command=self.text.yview)
        self.text.configure(yscrollcommand=self.scroll.set)
        self.scroll.pack(side="right", fill="y")
        # status bar
        self.status = ttk.Label(self, text="Esperando inicio...", anchor="w")
        self.status.pack(fill="x", padx=8, pady=(0,8))
        self.protocol("WM_DELETE_WINDOW", self.on_quit)
        self._closing = False

    def log(self, msg):
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.update_idletasks()

    def set_status(self, msg):
        self.status.config(text=msg)
        self.update_idletasks()

    def on_quit(self):
        if askyesno("Salir", "¿Deseas cerrar la aplicación?"):
            self._closing = True
            self.destroy()

# ---------------- búsqueda de archivos ----------------
def find_xlsx_files(root_dir):
    return [Path(root)/f for root,_,files in os.walk(root_dir) for f in files if f.lower().endswith(".xlsx")]

# ---------------- inspección detallada / validación ----------------
def inspect_sheet_for_errors(path: Path):
    """
    Intenta leer el archivo y devuelve:
     - df (DataFrame) si pudo cargar y pasar validaciones básicas (celda guía),
     - errors: lista de dicts { 'file', 'row', 'col', 'excel_cell', 'issue', 'value' }
    Si df es None, se omitirá el archivo en el análisis.
    """
    errors = []
    try:
        # Leer sin encabezado rígido para inspección si falla
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as e:
        # No pudo leerse - intentamos abrir sin interpretar headers para dar mayor info
        try:
            raw = pd.read_excel(path, engine="openpyxl", header=None)
            # marcar primera region 5x20 para inspección
            for r in range(min(8, raw.shape[0])):
                for c in range(min(20, raw.shape[1])):
                    val = raw.iloc[r, c]
                    if pd.isna(val) or str(val).strip().lower() in ("nan","none",""):
                        # no lo marcamos como error automáticamente, solo apuntamos posibles nulos
                        continue
            errors.append({'file': str(path), 'row': None, 'col': None, 'excel_cell': None,
                           'issue': f'Error al leer archivo: {e}', 'value': ''})
        except Exception:
            errors.append({'file': str(path), 'row': None, 'col': None, 'excel_cell': None,
                           'issue': f'Error al leer archivo y no se pudo inspeccionar (archivo corrupto)', 'value': ''})
        return None, errors

    # Validaciones específicas (las mismas que tu lógica original):
    # 1) Celda [1,15] (fila 2, col 16 en Excel 1-based) debe contener string exacto
    try:
        cell_val = df.iloc[1, 15]
        expected = "D E S G L O S E    P O R    P R O Y E C T O"
        if str(cell_val) != expected:
            errors.append({'file': str(path), 'row': 2, 'col': 16, 'excel_cell': 'R2C16 (fila2,col16)',
                           'issue': 'Texto esperado no coincide', 'value': str(cell_val)})
    except Exception as e:
        errors.append({'file': str(path), 'row': 2, 'col': 16, 'excel_cell': 'R2C16',
                       'issue': f'Error accediendo a celda esperada: {e}', 'value': ''})

    # 2) Celda Month: df.iloc[1,3] -> debe ser mes reconocible
    try:
        raw_month = df.iloc[1, 3]
        try:
            month_number(raw_month)
        except Exception as mne:
            errors.append({'file': str(path), 'row': 2, 'col': 4, 'excel_cell': 'D2',
                           'issue': 'Mes no reconocido', 'value': str(raw_month)})
    except Exception as e:
        errors.append({'file': str(path), 'row': 2, 'col': 4, 'excel_cell': 'D2',
                       'issue': f'Error accediendo D2: {e}', 'value': ''})

    # 3) Celda Year: df.iloc[1,11] -> debe ser entero convertible
    try:
        raw_year = df.iloc[1, 11]
        try:
            int(raw_year)
        except Exception:
            errors.append({'file': str(path), 'row': 2, 'col': 12, 'excel_cell': 'L2',
                           'issue': 'Año no convertible a entero', 'value': str(raw_year)})
    except Exception as e:
        errors.append({'file': str(path), 'row': 2, 'col': 12, 'excel_cell': 'L2',
                       'issue': f'Error accediendo L2: {e}', 'value': ''})

    # 4) Proyectos row: df.iloc[3,15:34] (fila 4, columnas 16..34) -> comprobar longitud y no-nulos en los nombres
    try:
        proyectos = df.iloc[3, 15:34].reset_index(drop=True).tolist()
        if len(proyectos) == 0:
            errors.append({'file': str(path), 'row': 4, 'col': 'P:AL', 'excel_cell': 'P4:AL4',
                           'issue': 'No se detectaron nombres de proyectos en el rango', 'value': ''})
        else:
            # detectar nombres vacíos
            for idx, val in enumerate(proyectos):
                if pd.isna(val) or str(val).strip() == "":
                    errors.append({'file': str(path), 'row': 4, 'col': 16 + idx, 'excel_cell': f'{16+idx} (fila4)',
                                   'issue': 'Nombre de proyecto vacío', 'value': str(val)})
    except Exception as e:
        errors.append({'file': str(path), 'row': 4, 'col': 'P:AL', 'excel_cell': 'P4:AL4',
                       'issue': f'Error leyendo fila proyectos: {e}', 'value': ''})

    # 5) Total_Hours row: df.iloc[37,15:34] -> longitud coincide y valores numéricos (o NaN -> 0)
    try:
        totals = df.iloc[37, 15:34].reset_index(drop=True).tolist()
        if len(totals) == 0:
            errors.append({'file': str(path), 'row': 38, 'col': 'P:AL', 'excel_cell': 'P38:AL38',
                           'issue': 'No se detectaron totales de horas en el rango', 'value': ''})
        else:
            # comprobar tipos
            for idx, val in enumerate(totals):
                if pd.isna(val):
                    # treat as 0 but register a warning
                    errors.append({'file': str(path), 'row': 38, 'col': 16 + idx, 'excel_cell': f'{16+idx} (fila38)',
                                   'issue': 'Total horas es NaN (será tratado como 0)', 'value': str(val)})
                else:
                    try:
                        float(val)
                    except Exception:
                        errors.append({'file': str(path), 'row': 38, 'col': 16 + idx, 'excel_cell': f'{16+idx} (fila38)',
                                       'issue': 'Total horas no es numérico', 'value': str(val)})
    except Exception as e:
        errors.append({'file': str(path), 'row': 38, 'col': 'P:AL', 'excel_cell': 'P38:AL38',
                       'issue': f'Error leyendo fila totales: {e}', 'value': ''})

    # Si hay errores, devolvemos None, errores. Si solo hay "warnings" leves como NaN en totales, el archivo puede procesarse
    # Distinción: si existen errores graves en Month/Year/Celda guía/Proyectos vacíos => consideramos archivo inválido
    severe_keys = ('Mes no reconocido','Año no convertible a entero','Error al leer archivo','Error accediendo','No se detectaron nombres de proyectos','No se detectaron totales de horas')
    severe = any([any(k in e['issue'] for k in severe_keys) for e in errors])
    if severe:
        return None, errors
    else:
        return df, errors

# ---------------- análisis principal (igual que antes pero con manejo de archivos válidos) ----------------
def analyze(links, holidays_freq, log_ui):
    valid_links = []
    dfs = []
    all_errors = []

    # inspeccionar cada archivo
    for p in links:
        log_ui.log(f"Inspeccionando: {p.name}")
        df, errors = inspect_sheet_for_errors(p)
        if errors:
            # anotar errores y mostrarlos
            for er in errors:
                all_errors.append(er)
                log_ui.log(f"  - ERROR: {Path(er['file']).name} | {er['excel_cell'] or ''} -> {er['issue']} -> [{er['value']}]")
        if df is not None:
            valid_links.append(p)
            dfs.append(df)
            log_ui.log(f"  → Archivo válido: {p.name}")
        else:
            log_ui.log(f"  → Archivo omitido: {p.name}")

    if not valid_links:
        log_ui.log("No hay archivos válidos para procesar. Se generará reporte de errores si corresponde.")
        # guardar errores si existen y salir
        if all_errors:
            df_err = pd.DataFrame(all_errors)
            df_err.to_excel("Errores_planillas.xlsx", index=False)
            log_ui.log("Se generó Errores_planillas.xlsx")
        return None, None, None, all_errors

    log_ui.log(f"\nProcesando {len(valid_links)} archivos válidos...")

    # reconstruir Rg como antes
    Rg = pd.DataFrame()
    for idx, p in enumerate(valid_links):
        df = dfs[idx]
        try:
            Name = df.columns[4]
            Month_raw = df.iloc[1,3]
            Month = month_number(Month_raw)
            Year = int(df.iloc[1,11])
            Proyectos = df.iloc[3, 15:34].reset_index(drop=True).tolist()
            Total_Hours = df.iloc[37, 15:34].reset_index(drop=True).tolist()
            cols = ["Name","Month","Year"] + Proyectos
            row = [Name, Month, Year] + Total_Hours
            row_dict = dict(zip(cols, row))
            row_dict["Links"] = str(valid_links[idx])
            betaux = pd.DataFrame([row_dict])
            if "TOTAL" in betaux.columns:
                betaux = betaux.loc[:, betaux.columns != "TOTAL"]
            Rg = pd.concat([Rg, betaux], ignore_index=True, sort=False)
            Rg.fillna(0, inplace=True)
            log_ui.log(f"  - Agregado: {valid_links[idx].name}")
        except Exception as e:
            all_errors.append({'file': str(valid_links[idx]), 'row': None, 'col': None, 'excel_cell': None,
                               'issue': f'Error procesando fila principal: {e}', 'value': ''})
            log_ui.log(f"  - ERROR procesando {valid_links[idx].name}: {e}")

    # detectar duplicados y resolver con ventana (como antes)
    if not Rg.empty:
        Rg["Aux"] = Rg["Name"].astype(str) + Rg["Month"].astype(str) + Rg["Year"].astype(str)
        aux_count = Rg["Aux"].value_counts().to_frame(name="count").reset_index().rename(columns={"index":"Aux"})
        dup = aux_count[aux_count["count"] > 1]
        links_current = valid_links.copy()
        if not dup.empty:
            # construir links_interferencias
            links_inter = pd.DataFrame(columns=["Links","Identificador"])
            for idx_dup, row in dup.reset_index(drop=True).iterrows():
                aux_val = row["Aux"]
                subset = Rg[Rg["Aux"] == aux_val][["Links"]].copy()
                subset["Identificador"] = idx_dup
                links_inter = pd.concat([links_inter, subset], ignore_index=True)
            # preguntamos al usuario (ventana simples)
            keep = askyesno("Duplicados", "Se detectaron registros duplicados (mismo Name+Month+Year). ¿Desea conservar solo los registros más recientes (por fecha de creación) y eliminar el resto?")
            if keep:
                # eliminar archivos más antiguos
                to_delete = set()
                for ident in links_inter['Identificador'].unique():
                    sub = links_inter[links_inter['Identificador']==ident]
                    ctimes = [(str(p), os.path.getctime(str(p))) for p in sub['Links']]
                    ctimes.sort(key=lambda x: x[1], reverse=True)
                    for path_str, _ in ctimes[1:]:
                        to_delete.add(Path(path_str))
                for p in to_delete:
                    try:
                        os.remove(p)
                        log_ui.log(f"  - Eliminado (duplicado): {p.name}")
                        # quitar de valid_links
                        if p in links_current:
                            links_current.remove(p)
                    except Exception as e:
                        log_ui.log(f"  - Error eliminando duplicado {p.name}: {e}")
                # actualizar Rg para conservar solo los links_current
                Rg = Rg[Rg["Links"].isin([str(x) for x in links_current])]
            else:
                log_ui.log("Se optó por no eliminar duplicados. Continuando con todos los archivos detectados.")

        # segunda pasada: construir Alfa sumando por columna
        Alfa = pd.DataFrame()
        for p in links_current:
            try:
                df = pd.read_excel(p, engine="openpyxl")
                Name = df.columns[4]
                Month = month_number(df.iloc[1,3])
                Year = int(df.iloc[1,11])
                Proyectos = df.iloc[3, 15:34].reset_index(drop=True).tolist()
                Total_Hours = df.iloc[37, 15:34].reset_index(drop=True).tolist()
                cols = ["Name","Month","Year"] + Proyectos
                row = [Name, Month, Year] + Total_Hours
                beta = pd.DataFrame([row], columns=cols)
                if "TOTAL" in beta.columns:
                    beta = beta.loc[:, beta.columns != "TOTAL"]
                beta = beta.groupby(beta.columns, axis=1).sum()
                Alfa = pd.concat([Alfa, beta], ignore_index=True, sort=False)
                Alfa.fillna(0, inplace=True)
            except Exception as e:
                all_errors.append({'file': str(p), 'row': None, 'col': None, 'excel_cell': None,
                                   'issue': f'Error en segunda pasada leyendo {p.name}: {e}', 'value': ''})
                log_ui.log(f"  - ERROR en segunda pasada {p.name}: {e}")

        if Alfa.empty:
            log_ui.log("No se pudieron construir datos Alfa. Abortando.")
            return None, None, None, all_errors

        columnas = [c for c in Alfa.columns if c not in ["Name","Month","Year"]]
        columnas = ["Name","Month","Year"] + columnas
        Alfa = Alfa.loc[:, columnas]

        Total_HH = Alfa.loc[:, ["Name","Month","Year"]].copy()
        Total_HH["Horas Realizadas"] = Alfa.loc[:, Alfa.columns.difference(["Name","Month","Year"])].sum(axis=1).round(2)

        # calcular horas objetivo por mes usando holidays_freq
        Whours = []
        for _, row in Total_HH.iterrows():
            Year = int(row["Year"])
            Month = int(row["Month"])
            start = datetime.date(year=Year, month=Month, day=1)
            if Month < 12:
                end = datetime.date(year=Year, month=(Month + 1), day=1)
            else:
                end = datetime.date(year=(Year + 1), month=1, day=1)
            Workdays = np.busday_count(start, end)
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

        # Suma
        Suma = Alfa.sum(axis=0).to_frame(name="Total")
        if set(["Name","Month","Year"]).issubset(Suma.index):
            Suma.drop(labels=["Name","Month","Year"], inplace=True, errors='ignore')

        # Consolidar Omega
        Omega = Alfa.copy(deep=True)
        Omega["Aux"] = Omega["Name"].astype(str) + Omega["Year"].astype(str)
        Omega = Omega.drop(columns=["Name","Year","Month"], errors='ignore')
        Omega = Omega.groupby("Aux", as_index=False).sum()
        Omega.insert(0, "Name", Omega["Aux"].str[:-4])
        Omega.insert(1, "Year", Omega["Aux"].str[-4:])
        Omega["Year"] = pd.to_numeric(Omega["Year"], errors='coerce').fillna(0).astype(int)
        Omega = Omega.drop(columns=["Aux"], errors='ignore')
        Omega.reset_index(drop=True, inplace=True)

        return Alfa.reset_index(drop=True), Total_HH.reset_index(drop=True), Suma, all_errors, Omega, links_current
    else:
        return None, None, None, all_errors

# ---------------- función principal con UI ----------------
def main_process(log_ui):
    try:
        log_ui.set_status("Obteniendo feriados...")
        current_year = datetime.date.today().year
        holidays = fetch_holidays_chile(years=[current_year, current_year+1])
        holidays_freq = working_holidays_frequency(holidays)
        log_ui.log("Feriados cargados.")

        log_ui.set_status("Buscando archivos .xlsx...")
        links = find_xlsx_files(os.getcwd())
        log_ui.log(f"Se encontraron {len(links)} archivos .xlsx")

        if not links:
            showinfo("Sin archivos", "No se encontraron archivos .xlsx en el directorio actual.")
            return

        log_ui.set_status("Inspeccionando y procesando archivos...")
        resultado = analyze(links, holidays_freq, log_ui)
        if resultado is None:
            log_ui.log("No se devolvió resultado del análisis.")
            return

        # Resultado puede venir con distinto formato (ver retorno)
        if len(resultado) >= 6:
            Alfa, Resumen1, Suma, all_errors, Omega, processed_links = resultado
        else:
            Alfa, Resumen1, Suma, all_errors = resultado
            Omega = None
            processed_links = []

        # Guardar Resumen.xlsx completo (si se generó Omega)
        if Omega is not None:
            log_ui.set_status("Guardando Resumen.xlsx...")
            try:
                with pd.ExcelWriter("Resumen.xlsx", engine="openpyxl") as writer:
                    Omega.to_excel(writer, sheet_name="Registro", index=False)
                    # Resumen1 puede no estar en el formato exacto si no fue construido; ajustamos existencia
                    try:
                        for name in Resumen1["Name"].unique():
                            aux = Resumen1[Resumen1["Name"] == name].dropna(how='all', axis=0)
                            aux.to_excel(writer, sheet_name=str(name)[:31], index=False)
                    except Exception:
                        # si Resumen1 no tiene "Name" o estructura, escribimos Alfa en otra hoja
                        Alfa.to_excel(writer, sheet_name="Alfa", index=False)
                log_ui.log("Resumen.xlsx creado correctamente.")
            except Exception as e:
                log_ui.log(f"Error guardando Resumen.xlsx: {e}")
        else:
            log_ui.log("No se generó Omega; no se crea Resumen.xlsx")

        # Guardar errores si existen
        if all_errors:
            df_err = pd.DataFrame(all_errors)
            df_err.to_excel("Errores_planillas.xlsx", index=False)
            log_ui.log(f"Errores detectados: {len(all_errors)}. Se guardó Errores_planillas.xlsx")
        else:
            log_ui.log("No se detectaron errores en las planillas procesadas.")

        log_ui.set_status("Proceso finalizado.")
        showinfo("Finalizado", "Análisis completado. Revisa la ventana para detalles y los archivos generados.")
    except Exception as e:
        log_ui.log(f"Error durante la ejecución principal: {e}")
        showinfo("Error", f"Error durante la ejecución: {e}")

# ---------------- arranque ----------------
def main():
    ui = LogWindow()
    # ejecutar el proceso poco después de iniciar la UI para que la ventana aparezca
    ui.after(200, lambda: main_process(ui))
    ui.mainloop()

if __name__ == "__main__":
    main()
