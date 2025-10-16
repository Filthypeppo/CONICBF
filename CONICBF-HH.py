# -*- coding: utf-8 -*-
"""
Created on Tue Oct 14 17:00:25 2025

@author: filth
"""
import glob, requests, sys, openpyxl, openpyxl.cell._writer
from dateutil.parser import parse
import datetime, numpy as np, os, pandas as pd, tkinter as tk
from tkinter import messagebox
from tkinter.messagebox import askyesno
try:
    url = "https://apis.digital.gob.cl/fl/feriados"
    headers = requests.utils.default_headers()
    headers.update({"User-Agent": "My User Agent 1.0"})
    response = requests.get(url, headers=headers)
    data = response.json()
    fechas = [d["fecha"] for d in data]
    fechas = [parse(fecha) for fecha in fechas]
    dia = [x.weekday() for x in fechas]
    Data1 = pd.DataFrame((list(zip(fechas, dia))), columns=["fechas", "dia"])
    Data1.drop((Data1[Data1["dia"] > 4].index), inplace=True)
    Data1.reset_index(inplace=True, drop=True)
    Data1["Month"] = pd.DatetimeIndex(Data1["fechas"]).month
    Data1["Year"] = pd.DatetimeIndex(Data1["fechas"]).year
    frecuency = Data1.groupby(["Year", "Month"]).size()
    frecuency = frecuency.to_frame()
    frecuency.reset_index(inplace=True)
    links = []
    for root, dirs, files in os.walk(os.getcwd()):
        for file in files:
            if file.endswith(".xlsx"):
                links.append(os.path.join(root, file))

        print("se encontraron " + str(len(links)) + " posibles registros HH \n")
        print("------------------------------------------------------------\n")

        def verificacion(links):
            errorlinks = []
            for i in links:
                VerData = pd.read_excel((i.encode("unicode_escape").decode()), engine="openpyxl")
                if VerData.iloc[(1, 15)] != "D E S G L O S E    P O R    P R O Y E C T O":
                    errorlinks.append(i)
                    continue
                return errorlinks


        def call(message, links, links_interferencias):
            global array_aux
            global fechamax
            global list_del
            global zzzlinks
            root = tk.Tk()
            root.withdraw()
            res = askyesno(title="Error!", message=("Los siguientes archivos corresponden a los mismos usuarios y fechas: \n " + message + "\n Desea dejar solo los registros mas recientes ? "))
            if res == True:
                print("OK!")
                for i in links_interferencias.Identificador.unique():
                    print("i es igual a : " + str(i))
                    A = links_interferencias[links_interferencias.Identificador == i]["Links"]
                    array_aux = np.empty([0, 2])
                    for j in A:
                        ti_m = os.path.getctime(j)
                        array_aux = np.append(array_aux, [[j, ti_m]], axis=0)

                    fechamax = max(array_aux[:, 1])
                    array_aux = array_aux[array_aux[:, 1] != fechamax]
                    list_del = array_aux[:, 0].tolist().copy()
                    for k in list_del:
                        os.remove(k)

                    links = [e for e in links if e not in list_del]

                zzzlinks = links
                root.destroy()
            else:
                zzzlinks = []
                messagebox.showinfo(message="No se han eliminado registros duplicados, porfavor verificar los elementos y volver a ejecutar", title="INFO")
                root.destroy()
            return zzzlinks


        def Month_number(mes):
            Meses = [
             [
              "enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"],
             [
              1,2,3,4,5,6,7,8,9,10,11,12]]
            index = Meses[0].index(mes.lower())
            return Meses[1][index]


        def Analisis(links):
            global Alfa
            global Aux_count
            global Beta
            global Betaux
            global Data
            global Month
            global Name
            global Proyectos
            global Rg
            global Total_HH
            global arr
            global links_interferencias
            global linksaux
            global lista
            global message
            global zzz
            try:
                errores = verificacion(links)
                print(" >>> Las siguientes planillas no tienen el formato correcto, verificar. <<<\n")
                for auxiliare in range(len(errores)):
                    print("N  " + str(auxiliare) + " " + errores[auxiliare] + "\n")

                root.destroy()
            except:
                pass
            else:
                Rg = pd.DataFrame()
                for i in range(len(links)):
                    Data = pd.read_excel((links[i].encode("unicode_escape").decode()), engine="openpyxl")
                    Name = Data.columns[4]
                    Month = Data.iloc[1][3]
                    print(Month)
                    Month = Month_number(Month)
                    Year = Data.iloc[1][11]
                    link = links[i]
                    print(i)
                    Proyectos = Data.iloc[3, 15:34].reset_index(drop=True).tolist()
                    Total_Hours = Data.iloc[37, 15:34].reset_index(drop=True).tolist()
                    Proyectos.insert(0, "Name")
                    Proyectos.insert(1, "Month")
                    Proyectos.insert(2, "Year")
                    Proyectos.insert(4, "Links")
                    Total_Hours.insert(0, Name)
                    Total_Hours.insert(1, Month)
                    Total_Hours.insert(2, Year)
                    Total_Hours.insert(4, link)
                    Betaux = pd.DataFrame([Total_Hours], columns=Proyectos)
                    Betaux = Betaux.loc[:, Betaux.columns.notna()]
                    Betaux.iloc[:, 3:] = Betaux.iloc[:, 3:].round(2)
                    Betaux.reset_index(drop=True)
                    Betaux = Betaux.groupby((Betaux.columns), axis=1).sum()
                    Betaux = Betaux.loc[:, Betaux.columns.notna()]
                    Betaux = Betaux.loc[:, Betaux.columns != "TOTAL"]
                    print("OK")
                    Rg = pd.concat((Rg, Betaux), axis=0)
                    Rg = Rg.loc[:, Rg.columns.notna()]
                    Rg.fillna(0, inplace=True)

                Rg["Aux"] = Rg.Name + Rg.Month.astype(str) + Rg.Year.astype(str)
                Rg.reset_index(drop=True, inplace=True)
                Aux_count = Rg["Aux"].value_counts().to_frame()
                try:
                    Aux_count = Aux_count[Aux_count.Aux > 1]
                    Aux_count.reset_index(inplace=True)
                    Aux_count.rename({"Aux": "count"}, axis="columns", inplace=True)
                    Aux_count.rename({"index": "Aux"}, axis="columns", inplace=True)
                    if len(Aux_count) == 0:
                        a = 1 / 0
                    else:
                        pass
                    links_interferencias = pd.DataFrame()
                    arr = np.empty([0, 2])
                    for i in range(len(Aux_count)):
                        linksaux = Rg[Rg["Aux"] == Aux_count.iloc[(i, 0)]]["Links"].to_frame()
                        linksaux["Identificador"] = i
                        links_interferencias = pd.concat((links_interferencias, linksaux), axis=0)

                    links_interferencias.reset_index(drop=True, inplace=True)
                    for i in range(len(links_interferencias)):
                        Filename = os.path.split(links_interferencias.Links[i])[1]
                        lista = [[Filename, links_interferencias.Identificador[i]]]
                        arr = np.append(arr, lista, axis=0)

                    message = ""
                    for i in arr[:, 0]:
                        aux_mess = "* " + str(i) + " \n "
                        message = message + aux_mess

                    try:
                        zzz = call(message, links, links_interferencias)
                        print(zzz)
                        links = zzz
                    except:
                        1 / 0

                except:
                    pass
                else:
                    if len(links) == 0:
                        sys.exit()
                    else:
                        pass
                    Alfa = pd.DataFrame()
                    Whours = []
                    for i in range(len(links)):
                        Data = pd.read_excel((links[i].encode("unicode_escape").decode()), engine="openpyxl")
                        Name = Data.columns[4]
                        Month = Data.iloc[1][3]
                        Month = Month_number(Month)
                        Year = Data.iloc[1][11]
                        Proyectos = Data.iloc[3, 15:34].reset_index(drop=True).tolist()
                        Total_Hours = Data.iloc[37, 15:34].reset_index(drop=True).tolist()
                        Proyectos.insert(0, "Name")
                        Proyectos.insert(1, "Month")
                        Proyectos.insert(2, "Year")
                        Total_Hours.insert(0, Name)
                        Total_Hours.insert(1, Month)
                        Total_Hours.insert(2, Year)
                        Beta = pd.DataFrame([Total_Hours], columns=Proyectos)
                        Beta = Beta.loc[:, Beta.columns.notna()]
                        Beta = Beta.loc[:, Beta.columns != "TOTAL"]
                        Beta = Beta.groupby((Beta.columns), axis=1).sum()
                        Alfa = pd.concat((Alfa, Beta), axis=0)
                        Alfa = Alfa.loc[:, Alfa.columns.notna()]
                        Alfa.fillna(0, inplace=True)
                        columnas = Alfa.columns.difference(["Name", "Month", "Year"]).to_list()
                        columnas.insert(0, "Name")
                        columnas.insert(1, "Month")
                        columnas.insert(2, "Year")
                        Alfa = Alfa[columnas]
                        Total_HH = Alfa.iloc[:, :3].copy()
                        Total_HH["Horas Realizadas"] = Alfa.iloc[:, 3:].sum(axis=1)
                        print(Total_HH)
                        try:
                            start = datetime.date(year=Year, month=Month, day=1)
                            if Month < 12:
                                end = datetime.date(year=Year, month=(Month + 1), day=1)
                            else:
                                end = datetime.date(year=(Year + 1), month=1, day=1)
                                navidad = datetime.date(year=Year, month=12, day=24)
                                anuevo = datetime.date(year=Year, month=12, day=31)
                            Workdays = np.busday_count(start, end)
                            print("Workdays =  " + str(Workdays))
                            Holydays = frecuency[(frecuency["Year"] == Year) & (frecuency["Month"] == Month)].iloc[(0,
                                                                                                                    2)]
                            print("Holydays = " + str(Holydays))
                        except:
                            start = datetime.date(year=Year, month=Month, day=1)
                            end = datetime.date(year=Year, month=(Month + 1), day=1)
                            Workdays = np.busday_count(start, end)
                            print("Workdays =  " + str(Workdays))
                            Holydays = 0
                            print("Holydays = " + str(Holydays))
                        else:
                            Whours.append(8 * (Workdays - Holydays))
                            Total_HH["Horas objetivo*"] = Whours
                            Total_HH["Horas Adicionales"] = " "
                            Total_HH["Horas Autorizaadas"] = " "
                            Total_HH["Horas Pagadas"] = " "
                            Total_HH["Diferencia"] = " "
                            Alfa.iloc[:, 2:] = Alfa.iloc[:, 2:].round(2)
                            Total_HH.iloc[:, 2:] = Total_HH.iloc[:, 2:].round(2)
                            Suma = Alfa.sum(axis=0).to_frame()
                            Suma.drop(labels=["Name", "Month", "Year"], inplace=True, axis=0)
                            Suma.rename({0: "Total"}, axis="columns", inplace=True)
                            print(Alfa)

                    return (Alfa, Total_HH, Suma)


        (Datos, Resumen1, Suma) = Analisis(links)
        Omega = Datos.copy(deep=True)
        Omega["Aux"] = Omega["Name"] + Omega["Year"].astype(str)
        Omega.drop(["Name", "Year", "Month"], axis=1, inplace=True)
        Omega = Omega.groupby("Aux", as_index=False).sum()
        Omega.insert(0, "Name", Omega["Aux"].str[:-4])
        Omega.insert(1, "Year", Omega["Aux"].str[-4:])
        Omega["Year"] = pd.to_numeric(Omega["Year"])
        Omega.drop(["Aux"], axis=1, inplace=True)
        Omega.reset_index(inplace=True, drop=True)

    with pd.ExcelWriter("Resumen.xlsx") as writer:
        Omega.to_excel(writer, sheet_name="Registro", index=False)
        for i in Resumen1["Name"].unique():
            print(Resumen1.where(Resumen1["Name"] == i).dropna(axis=0))
            Auxiliarwriter = Resumen1.where(Resumen1["Name"] == i).dropna(axis=0)
            Auxiliarwriter.to_excel(writer, sheet_name=i, index=False)

except Exception as e:
    try:
        print(f"Error: {e}")
    finally:
        e = None
        del e

else:
    input("presionar enter para salir ... ")

