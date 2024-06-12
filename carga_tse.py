from tkinter import ttk, filedialog
import tkinter as tk
import pandas as pd
import numpy as np
import datetime
import pyodbc
import os
import re
import time


class Selector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.mes = ""
        self.directorio = ""
        self.mes_int = None
        self.year = None
        self.__meses = [
            "Enero",
            "Febrero",
            "Marzo",
            "Abril",
            "Mayo",
            "Junio",
            "Julio",
            "Agosto",
            "Septiembre",
            "Octubre",
            "Noviembre",
            "Diciembre",
        ]
        self.title("Seleccione el directorio y el mes de análisis")
        self.geometry(self.__centrar(570, 210))
        self.protocol("WM_DELETE_WINDOW", self.__al_cerrar)

        self.__frame1 = tk.Frame(master=self, bg="#dff9fb")
        self.__frame1.pack(fill=tk.BOTH, side=tk.TOP, expand=True)

        self.__frame2 = tk.Frame(master=self, bg="#c7ecee")
        self.__frame2.pack(fill=tk.BOTH, side=tk.TOP)

        self.__frame3 = tk.Frame(master=self.__frame1, bg="#dff9fb")
        self.__frame3.pack(fill=tk.BOTH, side=tk.LEFT, padx=30, pady=30)

        self.__frame4 = tk.Frame(master=self.__frame1, bg="#dff9fb")
        self.__frame4.pack(fill=tk.BOTH, side=tk.LEFT, padx=30, pady=30, expand=True)

        self.__label1 = tk.Label(
            master=self.__frame3, bg="#dff9fb", text="Mes:", anchor=tk.W
        )
        self.__label1.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__combobox1 = ttk.Combobox(
            master=self.__frame4, state="readonly", values=self.__meses
        )
        self.__combobox1.pack(fill=tk.X, side=tk.TOP, expand=True)
        self.__combobox1.bind("<<ComboboxSelected>>", self.__actualizar_mes)

        self.__label4 = tk.Label(
            master=self.__frame3,
            bg="#dff9fb",
            text="Año:",
            pady=3.8,
            anchor=tk.W,
        )
        self.__label4.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__combobox2 = ttk.Combobox(
            master=self.__frame4, state="readonly", values=list(range(2014, 2035))
        )
        self.__combobox2.pack(fill=tk.X, side=tk.TOP, expand=True)
        self.__combobox2.bind("<<ComboboxSelected>>", self.__actualizar_year)

        self.__label2 = tk.Label(
            master=self.__frame3,
            bg="#dff9fb",
            text="Directorio:",
            pady=3.8,
            anchor=tk.W,
        )
        self.__label2.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__frame5 = tk.Frame(master=self.__frame4, bg="#dff9fb")
        self.__frame5.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__button1 = tk.Button(
            master=self.__frame5,
            text="Seleccionar...",
            padx=10,
            command=self.__actualizar_directorio,
        )
        self.__button1.pack(fill=tk.BOTH, side=tk.LEFT)

        self.__label3 = tk.Label(
            master=self.__frame5,
            bg="white",
            relief="flat",
            highlightthickness=1,
            highlightbackground="gray",
            anchor=tk.W,
        )
        self.__label3.pack(fill=tk.BOTH, side=tk.LEFT, expand=True)

        self.__button2 = tk.Button(
            master=self.__frame2, text="Aceptar", padx=10, command=self.destroy
        )
        self.__button2.pack(side=tk.RIGHT, padx=10, pady=10)

        self.mainloop()

    def __centrar(self, ancho, largo):
        ancho_s = self.winfo_screenwidth()
        largo_s = self.winfo_screenheight()
        x = int((ancho_s / 2) - (ancho / 2))
        y = int((largo_s / 2) - (largo / 2))
        return f"{ancho}x{largo}+{x}+{y}"

    def __actualizar_mes(self, _):
        self.mes = self.__combobox1.get()[0:3]
        self.mes_int = self.__combobox1.current() + 1

    def __actualizar_directorio(self):
        self.directorio = filedialog.askdirectory(title="Seleccione el directorio")
        self.__label3.configure(text=self.directorio)

    def __actualizar_year(self, _):
        self.year = int(self.__combobox2.get())

    def __al_cerrar(self):
        self.mes = ""
        self.directorio = ""
        self.destroy()


class CargaDia:
    def __init__(
        self,
        archivo: str,  # Ruta del archivo Excel
        base_datos: ManejoBaseDatos,  # Objeto para manipular la base de datos
        hoja_tension=18,  # Posición de la hoja de tension en el archivo
        hoja_potencia=19,  # Posición de la hoja de potencia en el archivo
        fila_subestaciones=0,  # Fila de las cabeceras de las subestaciones en la hoja de potencia
        col_subestaciones=1,  # Columna en la que empiezan las cabeceras de las subestaciones en la hoja de potencia
        fila_tension_nominal=1,  # Fila de las cabeceras de las subestaciones en la hoja tension
        col_tension_nominal=1,  # Columna en la que empiezan las cabeceras de las subestaciones en la hoja de tension
        fila_potencia=3,  # Fila en la que empiezan las lecturas de la potencia
        fila_tension_servicio=3,  # Fila en la que empiezan las lecturas de la tensión
        lecturas=96,  # Número de filas de potencia o de tensión a leer
    ):
        self.archivo = archivo
        self.base_datos = base_datos
        self.hoja_tension = hoja_tension
        self.hoja_potencia = hoja_potencia
        self.horas = None
        self.subestaciones_raw_potencia = None
        self.subestaciones_raw_tension = None
        self.alimentadores_raw = None
        self.tensiones_nominal_raw = None
        self.fecha = None
        self.cols_subestacion = []
        self.subestaciones_potencia = []
        self.cols_tension_nominal = []
        self.subestaciones_tension = []
        self.subestacion_tension_nominal = {}
        self.leer_encabezado(
            fila_subestaciones,
            col_subestaciones,
            fila_tension_nominal,
            col_tension_nominal,
        )
        self.generar_horas()
        self.leer_tensiones_nominal()
        self.extraer_subestaciones(col_subestaciones)
        self.extraer_cols_tension(col_tension_nominal)
        self.cargar_potencias(fila_potencia, lecturas)
        self.cargar_tension(fila_tension_servicio, lecturas)

    def leer_encabezado(
        self, fila_potencia, columna_potencia, fila_tension, columna_tension
    ):
        encabezados_potencia = pd.read_excel(
            self.archivo,
            self.hoja_potencia,
            header=None,
            skiprows=fila_potencia,
            nrows=2,
        )
        self.subestaciones_raw_potencia = (
            encabezados_potencia.iloc[0][columna_potencia:].fillna("").tolist()
        )
        self.alimentadores_raw = encabezados_potencia.iloc[1].fillna("").tolist()
        self.fecha = encabezados_potencia.iat[0, 0].to_pydatetime()

        encabezados_tension = pd.read_excel(
            self.archivo,
            self.hoja_tension,
            header=None,
            skiprows=fila_tension,
            nrows=2,
        )
        self.subestaciones_raw_tension = (
            encabezados_tension.iloc[0][columna_tension:].fillna("").tolist()
        )
        self.tensiones_nominal_raw = encabezados_tension.iloc[1].fillna("").tolist()

    def generar_horas(self):
        self.horas = Util.generar_horas()

    def leer_tensiones_nominal(self):
        self.subestacion_tension_nominal = self.base_datos.obtener_tensiones_nominal()

    def extraer_subestaciones(self, columna):
        for i, celda in enumerate(self.subestaciones_raw_potencia):
            if not celda:
                continue

            self.cols_subestacion.append(i + columna)

            if not celda.startswith("SUBESTACION"):
                break

            self.subestaciones_potencia.append(Util.parse_subestacion(celda))

    def extraer_cols_tension(self, columna):
        cols_subestacion = []

        for i, celda in enumerate(self.subestaciones_raw_tension):
            if not celda:
                continue

            self.subestaciones_tension.append(Util.parse_subestacion(celda))
            cols_subestacion.append(i + columna)

        cols_subestacion.append(i + 2)

        for i in range(len(self.subestaciones_tension)):
            col_ini = cols_subestacion[i]
            col_fin = cols_subestacion[i + 1] - 1

            tension_nom = {
                Util.parse_tension(self.tensiones_nominal_raw[j]): j
                for j in range(col_ini, col_fin)
                if self.tensiones_nominal_raw[j]
            }

            self.cols_tension_nominal.append(
                tension_nom.get(
                    self.subestacion_tension_nominal.get(self.subestaciones_tension[i])
                )
            )

    def cargar_potencias(self, fila, lecturas):
        print("    Cargando potencias...")

        for i, subestacion in enumerate(self.subestaciones_potencia):
            col_ini = self.cols_subestacion[i]
            col_fin = self.cols_subestacion[i + 1] - 3

            alimentadores = [
                Util.parse_alimentador(self.alimentadores_raw[j])
                for j in range(col_ini, col_fin)
            ]

            for j in range(0, col_fin - col_ini, 2):
                alimentador = alimentadores[j]

                if not alimentador or alimentador == "TOTAL":
                    continue

                self.base_datos.insertar_alimentador(subestacion, alimentador)

                df = pd.read_excel(
                    self.archivo,
                    self.hoja_potencia,
                    header=None,
                    names=["potencia_activa", "potencia_reactiva"],
                    usecols=[col_ini + j, col_ini + j + 1],
                    dtype=float,
                    skiprows=fila,
                    nrows=lecturas,
                )

                for k, x in df.iterrows():
                    self.base_datos.insertar_potencias(
                        self.fecha,
                        self.horas[k],
                        alimentador,
                        x["potencia_activa"],
                        x["potencia_reactiva"],
                    )

    def cargar_tension(self, fila, lecturas):
        print("    Cargando tensión...")

        for i, subestacion in enumerate(self.subestaciones_tension):
            columna = self.cols_tension_nominal[i]
            tension_nominal = self.subestacion_tension_nominal.get(subestacion)

            if not columna:
                df = pd.DataFrame({"tension_servicio": [tension_nominal] * lecturas})
            else:
                df = pd.read_excel(
                    self.archivo,
                    self.hoja_tension,
                    header=None,
                    names=["tension_servicio"],
                    usecols=[columna],
                    dtype=float,
                    skiprows=fila,
                    nrows=lecturas,
                ).fillna(tension_nominal)

            for j, x in df.iterrows():
                self.base_datos.insertar_tension(
                    self.fecha,
                    self.horas[j],
                    subestacion,
                    x["tension_servicio"],
                )


class CargaMes:
    def __init__(self, year, mes, mes_int, directorio):
        if not mes:
            exit(0)

        if not directorio:
            exit(0)

        if not year:
            exit(0)

        self.year = year
        self.mes = mes_int
        self.base_datos = ManejoBaseDatos()
        self.cargar(mes, directorio)
        self.horas = Util.generar_horas()
        self.depurar_potencias()
        self.depurar_tensiones()
        self.base_datos.desconectar()

    def cargar(self, mes, directorio):
        sub_dirs = [
            os.path.join(directorio, d)
            for d in os.listdir(directorio)
            if os.path.isdir(os.path.join(directorio, d)) and d.startswith(mes)
        ]

        for sub_dir in sub_dirs:
            archivo = [
                d
                for d in os.listdir(sub_dir)
                if os.path.isfile(os.path.join(sub_dir, d))
                and d.endswith(".xlsx")
                and not d.startswith("~")
            ][0]
            print(f"Cargando {archivo}...")

            CargaDia(os.path.join(sub_dir, archivo), self.base_datos)

        self.base_datos.commit()

    def depurar_potencias(self):
        print("Depurando potencias...")

        alimentadores = self.base_datos.obtener_alimentadores()

        for dia in range(7):
            print("    Día", dia)
            for alimentador in alimentadores:
                for hora in self.horas:
                    fechas, activas, reactivas = self.base_datos.obtener_potencias(
                        self.year, self.mes, hora, alimentador, dia
                    )

                    inferior, q2, superior = Util.calcular_quartil(activas)
                    activas = Util.reemplazar(activas, inferior, q2, superior)

                    inferior, q2, superior = Util.calcular_quartil(reactivas)
                    reactivas = Util.reemplazar(reactivas, inferior, q2, superior)

                    for i in range(len(activas)):
                        self.base_datos.insertar_potencias_dep(
                            fechas[i],
                            hora,
                            alimentador,
                            dia,
                            activas[i],
                            reactivas[i],
                        )

        self.base_datos.commit()

    def depurar_tensiones(self):
        print("Depurando tensiones...")

        subestaciones = self.base_datos.obtener_subestaciones()

        for dia in range(7):
            print("    Día", dia)
            for subestacion in subestaciones:
                for hora in self.horas:
                    fechas, tensiones = self.base_datos.obtener_tension(
                        self.year, self.mes, hora, subestacion, dia
                    )
                    inferior, q2, superior = Util.calcular_quartil(tensiones)
                    tensiones = Util.reemplazar(tensiones, inferior, q2, superior)

                    for i in range(len(tensiones)):
                        self.base_datos.insertar_tension_dep(
                            fechas[i],
                            hora,
                            subestacion,
                            dia,
                            tensiones[i],
                        )

        self.base_datos.commit()


class Util:
    @staticmethod
    def parse_subestacion(subestacion):
        encontrado = re.findall(r"\d+", str(subestacion))

        if not encontrado:
            return subestacion

        return int(encontrado[0])

    @staticmethod
    def parse_alimentador(alimentador):
        if isinstance(alimentador, float):
            return f"{alimentador:04.0f}"

        if isinstance(alimentador, int):
            return f"{alimentador:04}"

        return alimentador

    @staticmethod
    def parse_tension(tension):
        tension = tension.lower()

        encontrado = re.findall(r"\d+\.\d+|\d+", tension)

        if not encontrado:
            return tension

        valor = float(encontrado[0])

        if "kv" in tension:
            return valor

        return valor / 1000

    @staticmethod
    def calcular_quartil(datos):
        q1 = np.quantile(datos, 0.25)
        q2 = np.quantile(datos, 0.5)
        q3 = np.quantile(datos, 0.75)

        iqr = q3 - q1

        superior = q3 + 1.5 * iqr
        inferior = q1 - 1.5 * iqr

        return inferior, q2, superior

    @staticmethod
    def generar_horas(minutos=15):
        horas = []
        hora_actual = datetime.datetime(1, 1, 1)

        for _ in range(96):
            horas.append(hora_actual.time())
            hora_actual += datetime.timedelta(minutes=minutos)

        return horas

    @staticmethod
    def reemplazar(datos, inferior, media, superior):
        for i in range(len(datos)):
            if datos[i] < inferior or datos[i] > superior:
                datos[i] = media

        return datos


selector = Selector()
t1 = time.time()
CargaMes(selector.year, selector.mes, selector.mes_int, selector.directorio)
t2 = time.time()

print(f"{round((t2-t1)/60,1)} m")
