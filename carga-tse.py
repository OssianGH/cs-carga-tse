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
    """Clase/Módulo, que hereda de una interfaz, el cual se encarga de recibir y
    almacenar, en atributos de clase, los parámetros del script: el mes, el año, el
    directorio con los archivos de análisis y los números de hoja en las que se
    encuentran las lecturas de potencia y tensión en dichos archivos.
    """

    def __init__(self):
        super().__init__()
        self.mes = None
        self.directorio = None
        self.mes_int = None
        self.year = None
        self.hoja_potencia = None
        self.hoja_tension = None
        self.__placeholder = "Ingresar número de hoja a procesar..."
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
        self.title("Carga de archivos | Ingrese la información solicitada")
        self.geometry(self.__centrar(700, 300))
        self.protocol("WM_DELETE_WINDOW", self.__al_cerrar)
        self.bind("<Return>", self.__presionado_enter)
        self.bind("<Escape>", self.__presionado_escape)

        self.__frm_superior = tk.Frame(master=self, bg="#dff9fb")
        self.__frm_superior.pack(fill=tk.BOTH, side=tk.TOP, expand=True)

        self.__frm_inferior = tk.Frame(master=self, bg="#c7ecee")
        self.__frm_inferior.pack(fill=tk.BOTH, side=tk.TOP)

        self.__frm_labels = tk.Frame(master=self.__frm_superior, bg="#dff9fb")
        self.__frm_labels.pack(fill=tk.BOTH, side=tk.LEFT, padx=30, pady=30)

        self.__frm_inputs = tk.Frame(master=self.__frm_superior, bg="#dff9fb")
        self.__frm_inputs.pack(
            fill=tk.BOTH, side=tk.LEFT, padx=30, pady=30, expand=True
        )

        self.__lbl_mes = tk.Label(
            master=self.__frm_labels, bg="#dff9fb", text="Mes:", anchor=tk.W
        )
        self.__lbl_mes.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__cbx_mes = ttk.Combobox(
            master=self.__frm_inputs, state="readonly", values=self.__meses
        )
        self.__cbx_mes.pack(fill=tk.X, side=tk.TOP, expand=True)
        self.__cbx_mes.bind("<<ComboboxSelected>>", self.__actualizar_mes)

        self.__lbl_year = tk.Label(
            master=self.__frm_labels,
            bg="#dff9fb",
            text="Año:",
            pady=3.8,
            anchor=tk.W,
        )
        self.__lbl_year.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__cbx_year = ttk.Combobox(
            master=self.__frm_inputs, state="readonly", values=list(range(2014, 2035))
        )
        self.__cbx_year.pack(fill=tk.X, side=tk.TOP, expand=True)
        self.__cbx_year.bind("<<ComboboxSelected>>", self.__actualizar_year)

        self.__lbl_hoja_potencia = tk.Label(
            master=self.__frm_labels,
            bg="#dff9fb",
            text="Número de hoja de potencia:",
            pady=3.8,
            anchor=tk.W,
        )
        self.__lbl_hoja_potencia.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__txt_hoja_potencia = tk.Entry(
            master=self.__frm_inputs,
            relief="flat",
            highlightthickness=1,
            highlightbackground="gray",
            fg="gray",
        )
        self.__txt_hoja_potencia.insert(0, self.__placeholder)
        self.__txt_hoja_potencia.pack(fill=tk.X, side=tk.TOP, expand=True)
        self.__txt_hoja_potencia.bind("<KeyRelease>", self.__actualizar_hoja_potencia)
        self.__txt_hoja_potencia.bind("<FocusIn>", self.__al_poner_foco)
        self.__txt_hoja_potencia.bind("<FocusOut>", self.__al_quitar_foco)

        self.__lbl_hoja_tension = tk.Label(
            master=self.__frm_labels,
            bg="#dff9fb",
            text="Número de hoja de tension:",
            pady=3.8,
            anchor=tk.W,
        )
        self.__lbl_hoja_tension.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__txt_hoja_tension = tk.Entry(
            master=self.__frm_inputs,
            relief="flat",
            highlightthickness=1,
            highlightbackground="gray",
            fg="gray",
        )
        self.__txt_hoja_tension.insert(0, self.__placeholder)
        self.__txt_hoja_tension.pack(fill=tk.X, side=tk.TOP, expand=True)
        self.__txt_hoja_tension.bind("<KeyRelease>", self.__actualizar_hoja_tension)
        self.__txt_hoja_tension.bind("<FocusIn>", self.__al_poner_foco)
        self.__txt_hoja_tension.bind("<FocusOut>", self.__al_quitar_foco)

        self.__lbl_directorio = tk.Label(
            master=self.__frm_labels,
            bg="#dff9fb",
            text="Directorio:",
            pady=3.8,
            anchor=tk.W,
        )
        self.__lbl_directorio.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__frm_directorio = tk.Frame(master=self.__frm_inputs, bg="#dff9fb")
        self.__frm_directorio.pack(fill=tk.X, side=tk.TOP, expand=True)

        self.__btn_directorio = tk.Button(
            master=self.__frm_directorio,
            text="Seleccionar...",
            padx=10,
            command=self.__actualizar_directorio,
        )
        self.__btn_directorio.pack(fill=tk.BOTH, side=tk.LEFT)

        self.__lbl_input_directorio = tk.Label(
            master=self.__frm_directorio,
            bg="white",
            relief="flat",
            highlightthickness=1,
            highlightbackground="gray",
            anchor=tk.W,
        )
        self.__lbl_input_directorio.pack(fill=tk.BOTH, side=tk.LEFT, expand=True)

        self.__btn_aceptar = tk.Button(
            master=self.__frm_inferior, text="Aceptar", padx=10, command=self.destroy
        )
        self.__btn_aceptar.pack(side=tk.RIGHT, padx=10, pady=10)

        self.mainloop()

    def __centrar(self, ancho, largo):
        """Centra la ventana en la pantalla del usuario."""

        ancho_s = self.winfo_screenwidth()
        largo_s = self.winfo_screenheight()

        x = int((ancho_s / 2) - (ancho / 2))
        y = int((largo_s / 2) - (largo / 2))

        return f"{ancho}x{largo}+{x}+{y}"

    def __actualizar_mes(self, _):
        """Se ejecuta al cambiar de selección en `__cbx_mes`."""

        # Almacena las tres primeras letras del nombre del mes seleccionado.
        self.mes = self.__cbx_mes.get()[0:3]

        # Almacena el número de mes seleccionado ([1,12]).
        self.mes_int = self.__cbx_mes.current() + 1

    def __actualizar_directorio(self):
        """Se ejecuta al dar click en `__btn_directorio`."""

        # Solicita el directorio mediante un dialogo de selección.
        self.directorio = filedialog.askdirectory(title="Seleccione el directorio")

        # Escribe el directorio escogido en el label correspondiente.
        self.__lbl_input_directorio.configure(text=self.directorio)

    def __actualizar_year(self, _):
        """Se ejecuta al cambiar de selección en `__cbx_year` y dicha selección es
        almacenada.
        """
        self.year = self.__cbx_year.get()

    def __al_cerrar(self):
        """Se ejecuta al cerrar la ventana o presionar `Escape`."""

        self.mes = ""
        self.destroy()

    def __presionado_enter(self, _):
        self.destroy()

    def __presionado_escape(self, _):
        self.__al_cerrar()

    def __actualizar_hoja_potencia(self, _):
        """Se ejecuta al modificar el contenido de `__txt_hoja_potencia`. El valor
        ingresado es almacenado, eliminando espacios al principio y al final.
        """
        self.hoja_potencia = self.__txt_hoja_potencia.get().strip()

    def __actualizar_hoja_tension(self, _):
        """Se ejecuta al modificar el contenido de `__txt_hoja_tension`. El valor
        ingresado es almacenado, eliminando espacios al principio y al final.
        """
        self.hoja_tension = self.__txt_hoja_tension.get().strip()

    def __al_poner_foco(self, event):
        """Se ejecuta al enfocar el cursor de texto en `__txt_hoja_potencia` o
        `__txt_hoja_tension`. Si el texto en ese momento es el placeholder, borra el
        mismo.
        """
        entry = event.widget

        if entry.get() == self.__placeholder:
            entry.config(fg="black")
            entry.delete(0, "end")

    def __al_quitar_foco(self, event):
        """Se ejecuta al quitar el cursor de texto de `__txt_hoja_tension` o
        `__txt_hoja_potencia`. Si está vacío en ese momento, escribe el placeholder.
        """
        entry = event.widget

        if entry.get() == "":
            entry.config(fg="gray")
            entry.insert(0, self.__placeholder)


class CargaDia:
    """Clase/Módulo el cual se encarga de la lectura de archivos con lecturas de
    potencia y tensión, interpretando el número y posiciones de tablas de forma
    dinámica, manteniendo las relaciones (alimentador/subestación); posteriormente se
    cargan los datos leidos en la base de datos usando el módulo correspondiente.

    Parámetros del constructor:
    * `archivo`: Ruta del archivo Excel.
    * `base_datos`: Objeto para manipular la base de datos.
    * `hoja_potencia`: Posición [0,1,...] de la hoja de potencia en el archivo.
    * `hoja_tension`: Posición [0,1,...] de la hoja de tension en el archivo.
    * `fila_subestaciones`: Fila de las cabeceras de las subestaciones en la hoja de
    potencia.
    * `col_subestaciones`: Columna en la que empiezan las cabeceras de las
    subestaciones en la hoja de potencia.
    * `fila_tension_nominal`: Fila de las cabeceras de las subestaciones en la hoja
    tension.
    * `col_tension_nominal`: Columna en la que empiezan las cabeceras de las
    subestaciones en la hoja de tension.
    * `fila_potencia`: Fila en la que empiezan las lecturas de la potencia.
    * `fila_tension_servicio`: Fila en la que empiezan las lecturas de la tensión.
    * `lecturas`: Número de filas de potencia o de tensión a leer.
    """

    def __init__(
        self,
        archivo: str,
        base_datos: ManejoBaseDatos,
        hoja_potencia: int,
        hoja_tension: int,
        fila_subestaciones=0,
        col_subestaciones=1,
        fila_tension_nominal=1,
        col_tension_nominal=1,
        fila_potencia=3,
        fila_tension_servicio=3,
        lecturas=96,
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
        """Lee los encabezados en las hojas de potencia y tensión."""

        # Lee la fila con los encabezados, de las subestaciones y de los alimentadores,
        # en la hoja de potencia.
        encabezados_potencia = pd.read_excel(
            self.archivo,
            self.hoja_potencia,
            header=None,
            skiprows=fila_potencia,
            nrows=2,
        )

        # Almacena como lista la primera fila leída (encabezados de las subestaciones),
        # omitiendo las celdas antes de `columna_potencia`. Las celdas vacías se
        # reemplazan con una cadena vacía.
        self.subestaciones_raw_potencia = (
            encabezados_potencia.iloc[0, columna_potencia:].fillna("").tolist()
        )

        # Almacena como lista la segunda fila leída (encabezados de los alimentadores).
        self.alimentadores_raw = encabezados_potencia.iloc[1].fillna("").tolist()

        # Almacena la fecha, ubicada en la primera fila y la primera columna.
        self.fecha = encabezados_potencia.iat[0, 0].to_pydatetime()

        # Lee la fila con los encabezados, de las subestaciones y de las tensiones
        # nominales, en la hoja de tensión.
        encabezados_tension = pd.read_excel(
            self.archivo,
            self.hoja_tension,
            header=None,
            skiprows=fila_tension,
            nrows=2,
        )

        # Almacena como lista la primera fila leída (encabezados de las subestaciones),
        # omitiendo las celdas antes de `columna_tension`.
        self.subestaciones_raw_tension = (
            encabezados_tension.iloc[0, columna_tension:].fillna("").tolist()
        )

        # Almacena como lista la segunda fila leída (encabezados de las tensiones
        # nominales).
        self.tensiones_nominal_raw = encabezados_tension.iloc[1].fillna("").tolist()

    def generar_horas(self):
        """Genera y almacena la lista de horas."""

        self.horas = Util.generar_horas()

    def leer_tensiones_nominal(self):
        """Almacena el diccionario con las tensiones nominales de cada subestación."""

        self.subestacion_tension_nominal = self.base_datos.obtener_tensiones_nominal()

    def extraer_subestaciones(self, columna):
        """Encuentra los números de la primera columna de cada tabla de subestación en
        la hoja de potencia, así como los números de identificación de cada una de
        ellas. El argumento `columna` representa el número de celdas omitidas en
        `subestaciones_raw_potencia`.
        """
        # Itera por cada celda en la fila de encabezados de las subestaciones.
        for i, celda in enumerate(self.subestaciones_raw_potencia):
            if not celda:
                # Si la celda actual está vacía salta a la siguente iteración.
                continue

            # Si la celda no está vacía quiere decir que en la columna actual empieza
            # una tabla de subestación. Agrega a la lista de columnas, la columna
            # absoluta del archivo excel sumando la columna relativa `i` más las
            # celdas omitidas en la lista.
            self.cols_subestacion.append(i + columna)

            if not celda.startswith("SUBESTACION"):
                # Termina el bucle si el texto de la celda actual no empieza con
                # "SUBESTACION" (esto sucede cuando se han alcanzado las tablas extras
                # con registros que están fuera del análisis buscado).
                break

            # Agrega a la lista de números de subestaciones (en la hoja de potencia) el
            # número de subestacion transformado apartir del texto en la celda actual.
            self.subestaciones_potencia.append(Util.parse_subestacion_potencia(celda))

    def extraer_cols_tension(self, columna):
        """Encuentra los números de las columnas cuyo encabezado corresponde a la
        tensión nominal de cada tabla de subestación en la hoja de tensión, así como
        los números de identificación de cada una de ellas. El argumento `columna`
        representa el número de celdas omitidas en `subestaciones_raw_tension`.
        """
        cols_subestacion = []

        # Itera por cada celda en la fila de encabezados de las subestaciones.
        for i, celda in enumerate(self.subestaciones_raw_tension):
            if not celda:
                # Si la celda actual está vacía salta a la siguente iteración.
                continue

            # Si la celda no está vacía. Agrega a la lista de columnas, la columna
            # absoluta del archivo excel sumando la columna relativa `i` más las celdas
            # omitidas en la lista.
            cols_subestacion.append(i + columna)

            try:
                # Intenta transformar el texto en la celda actual en un número de
                # subestación
                subestacion = Util.parse_subestacion_tension(celda)
            except ValueError:
                # Termina el bucle si no se puede transformar (esto sucede cuando se
                # han alzanzado las tablas extras con registros que están fuera del
                # análisis buscado).
                break

            # Agrega a la lista de nombres de subestaciones (en la hoja de tensión) el
            # número de subestación.
            self.subestaciones_tension.append(subestacion)
        else:
            # Si se realizaron todas las iteraciones, es decir, si se finalizó sin un
            # `break`. Agrega a la lista de columnas la última columna de la última
            # tabla.
            cols_subestacion.append(i + columna + 2)

        # Itera por cada subestación en la hoja de tensión.
        for i in range(len(self.subestaciones_tension)):
            # Almacena la primera columna de la tabla de la subestación actual.
            col_ini = cols_subestacion[i]

            # Almacena la última columna de la tabla de la subestación actual
            # (debido a que por formato las tablas están separadas por una columna,
            # la última columna de una tabla es la primera columna de la tabla
            # siguiente menos 1).
            col_fin = cols_subestacion[i + 1] - 1

            # Define un diccionario (tension_nominal: columna) para cada encabezado de
            # tension nominal en la tabla de la subestación actual.
            try:
                tension_nom = {
                    Util.parse_tension(self.tensiones_nominal_raw[j]): j
                    for j in range(col_ini, col_fin)
                    if self.tensiones_nominal_raw[j]
                }
            except Exception as e:
                print(e, "\narchvivo: ", self.archivo)

            # Agrega a la lista de columnas de tensiones nominales, la columna que
            # contenga la tensión nominal de la subestación actual. Es decir,
            # subestacion_actual -> tension_nominal -> columa_a_leer. Si no se
            # encuentra se agrega `None`.
            self.cols_tension_nominal.append(
                tension_nom.get(
                    self.subestacion_tension_nominal.get(self.subestaciones_tension[i])
                )
            )

    def cargar_potencias(self, fila, lecturas):
        """Carga en la base de datos cada lectura de potencia activa y reactiva para
        cada alimentador de cada subestación.
        """
        # Itera por cada subestación en la hoja de potencia.
        for i, subestacion in enumerate(self.subestaciones_potencia):
            # Almacena la primera columna de la tabla de la subestación actual.
            col_ini = self.cols_subestacion[i]

            # Almacena la última columna de la tabla de la subestación actual
            # (debido a que por formato las tablas están separadas por una columna,
            # la última columna de una tabla es la primera columna de la tabla
            # siguiente menos 1 y menos las dos columnas de totales).
            col_fin = self.cols_subestacion[i + 1] - 3

            # Define la lista de alimentadores de la subestación actual, transformando
            # el texto en cada celda del encabezado en un código de alimentador válido.
            alimentadores = [
                Util.parse_alimentador(self.alimentadores_raw[j])
                for j in range(col_ini, col_fin)
            ]

            # Itera por cada tabla de alimentador a leer.
            for j in range(0, col_fin - col_ini, 2):
                alimentador = alimentadores[j]

                if not alimentador or alimentador == "TOTAL":
                    # Si el alimentador está vacío o es una tabla de total, salta a la
                    # siguiente iteración.
                    continue

                # Agrega, el alimentador actual y la subestación a la que pertenece, a
                # la base de datos, si no existe aún.
                self.base_datos.insertar_alimentador(subestacion, alimentador)

                # Lee y almacena las lecturas de potencia activa y reactiva de la tabla
                # del alimentador actual. Cualquier celda vacía se reemplaza con 0.
                df = pd.read_excel(
                    self.archivo,
                    self.hoja_potencia,
                    header=None,
                    names=["potencia_activa", "potencia_reactiva"],
                    usecols=[col_ini + j, col_ini + j + 1],
                    dtype=float,
                    skiprows=fila,
                    nrows=lecturas,
                ).fillna(0)

                # Agrega cada lectura a la base de datos.
                for k, x in df.iterrows():
                    try:
                        self.base_datos.insertar_potencias(
                            self.fecha,
                            self.horas[k],
                            alimentador,
                            x["potencia_activa"],
                            x["potencia_reactiva"],
                        )
                    except:
                        print(
                            "Error al insertar en la tabla de potencias:",
                            f"  Fecha: {self.fecha.strftime('%d/%m/%Y')}",
                            f"  Hora: {self.horas[k]}",
                            f"  Alimentador: {alimentador}",
                            f"  Subestacion: {subestacion}",
                            f"  Potencia activa: {x['potencia_activa']}",
                            f"  Potencia reactiva: {x['potencia_reactiva']}",
                            f"Revisar hoja {self.hoja_potencia + 1} del archivo ",
                            f"{self.archivo}.",
                            sep="\n",
                        )
                        exit(1)

    def cargar_tension(self, fila, lecturas):
        """Carga en la base de datos cada lectura de tensión de servicio para cada
        subestación.
        """
        # Itera por cada subestación en la hoja de tensión.
        for i, subestacion in enumerate(self.subestaciones_tension):
            # Asigna la columna a leer.
            columna = self.cols_tension_nominal[i]

            # Asigna la tensión nominal de la subestación actual.
            tension_nominal = self.subestacion_tension_nominal.get(subestacion)

            if not columna:
                # Si la columna a leer no existe, considera todas las lecturas de
                # tensión de servicio como las de la tensión nominal.
                df = pd.DataFrame({"tension_servicio": [tension_nominal] * lecturas})
            else:
                # Si la columna a leer existe, lee y almacena las lecturas de tensión
                # de servicio de la subestación actual. Cualquier celda vacía se
                # reemplaza con 0.
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

            # Agrega cada lectura a la base de datos.
            for j, x in df.iterrows():
                try:
                    self.base_datos.insertar_tension(
                        self.fecha,
                        self.horas[j],
                        subestacion,
                        x["tension_servicio"],
                    )
                except:
                    print(
                        "Error al insertar en la tabla de tension:",
                        f"  Fecha: {self.fecha.strftime('%d/%m/%Y')}",
                        f"  Hora: {self.horas[j]}",
                        f"  Subestacion: {subestacion}",
                        f"  Tensión de servicio: {x['tension_servicio']}",
                        f"Revisar hoja {self.hoja_tension + 1} del archivo ",
                        f"{self.archivo}.",
                        sep="\n",
                    )
                    exit(1)


class CargaMes:
    """Clase/Módulo el cual se encarga de la carga de cada archivo, correspondiente a
    un mes, en la base de datos, por medio del análisis y carga que realiza el módulo
    `CargaDia` en cada archivo.

    Parámetros del constructor:
    * `year`: Año al cual corresponden las lecturas en todos los archivos.
    * `mes`: Primeras tres letras del nombre del mes al cual corresponden las lecturas
    en todos los archivos.
    * `mes_int`: Número del mes al cual corresponden las lecturas en todos los
    archivos.
    * `directorio`: Ruta en la que se encuentran las carpetas con los archivos a leer.
    * `hoja_potencia`: Número [1,2,...] de la hoja de potencia en todos los archivos.
    * `hoja_tension`: Número [1,2,...] de la hoja de tension en todos los archivos.
    """

    def __init__(self, year, mes, mes_int, directorio, hoja_potencia, hoja_tension):
        if not mes:
            return

        if not year:
            return

        if not directorio:
            return

        if not hoja_potencia:
            return

        if not hoja_tension:
            return

        try:
            int_hoja_potencia = int(hoja_potencia)
        except:
            return

        try:
            int_hoja_tension = int(hoja_tension)
        except:
            return

        self.year = int(year)
        self.mes = mes_int
        self.base_datos = ManejoBaseDatos()
        self.cargar(mes, directorio, int_hoja_potencia, int_hoja_tension)
        self.horas = Util.generar_horas()
        self.depurar_potencias()
        self.depurar_tensiones()
        self.base_datos.desconectar()

    def cargar(self, mes, directorio, hoja_potencia, hoja_tension):
        """Carga en la base de datos la información de análisis a partir de cada
        archivo en `directorio` de un `mes` específico.
        """
        print("Cargando datos...")

        # Establece la lista de carpetas dentro de `directorio` que contienen los
        # archivos correspondientes a cada día del mes.
        sub_dirs = [
            os.path.join(directorio, d)
            for d in os.listdir(directorio)
            if os.path.isdir(os.path.join(directorio, d))
            and d.lower().startswith(mes.lower())
        ]

        # Itera por cada elemento dentro de cada carpeta de cada mes.
        for sub_dir in sub_dirs:
            # Selecciona el primer elemento que sea un archivo, que sea xlsx, que
            # empiece con 'tse' y que no contenga espacios en la carpeta actual.
            try:
                archivo = [
                    d
                    for d in os.listdir(sub_dir)
                    if os.path.isfile(os.path.join(sub_dir, d))
                    and d.lower().endswith(".xlsx")
                    and d.lower().startswith("tse")
                    and not " " in d
                ][0]
            except:
                print(f"Error: No se encontró el archivo esperado en {sub_dir}.")
                exit(1)

            print(archivo)

            # Carga a la base de datos el archivo encontrado correspondiente a un día.
            CargaDia(
                os.path.join(sub_dir, archivo),
                self.base_datos,
                hoja_potencia - 1,
                hoja_tension - 1,
            )

        # Confirma los cambios en la base de datos.
        self.base_datos.commit()

    def depurar_potencias(self):
        """Realiza un análisis intercuartil sobre los registros de potencias un mes, ya
        cargados en la base de datos. Se agrupan los registros de un día, alimentador y
        hora específica, para cada grupo, los registros que se encuentren fuera del
        rango intercuartil son reemplazados por la media.
        """
        print("Depurando datos...")

        # Asigna la lista de todos los alimentadores.
        alimentadores = self.base_datos.obtener_alimentadores()

        # Itera por cada día de la semana, cada alimentador y cada hora.
        for dia in range(7):
            for alimentador in alimentadores:
                for hora in self.horas:
                    # Se recogen los registros del rango actual desde la base de datos.
                    fechas, activas, reactivas = self.base_datos.obtener_potencias(
                        self.year, self.mes, hora, alimentador, dia
                    )

                    # Si no se encontraron registros para el rango actual, salta a la
                    # siguiente iteración.
                    if len(fechas) == 0:
                        continue

                    # Calcula el rango intercuartil para las potencias activas y
                    # reemplaza las lecturas fuera del rango.
                    inferior, q2, superior = Util.calcular_quartil(activas)
                    activas = Util.reemplazar(activas, inferior, q2, superior)

                    # Calcula el rango intercuartil para las potencias reactivas y
                    # reemplaza las lecturas fuera del rango.
                    inferior, q2, superior = Util.calcular_quartil(reactivas)
                    reactivas = Util.reemplazar(reactivas, inferior, q2, superior)

                    # Agrega cada lectura depurada a la base de datos.
                    for i in range(len(activas)):
                        self.base_datos.insertar_potencias_dep(
                            fechas[i],
                            hora,
                            alimentador,
                            dia,
                            activas[i],
                            reactivas[i],
                        )

        # Confirma los cambios en la base de datos.
        self.base_datos.commit()

    def depurar_tensiones(self):
        """Realiza un análisis intercuartil sobre las registros de tensiones de un mes,
        ya cargadas en la base de datos. Se agrupan los registros de un día,
        alimentador y hora específica, para cada grupo, los registros que se encuentren
        fuera del rango intercuartil son reemplazados por la media.
        """
        # Asigna la lista de todas las subestaciones.
        subestaciones = self.base_datos.obtener_subestaciones()

        # Itera por cada día de la semana, cada alimentador y cada hora.
        for dia in range(7):
            for subestacion in subestaciones:
                for hora in self.horas:
                    # Se recogen los registros del rango actual desde la base de datos.
                    fechas, tensiones = self.base_datos.obtener_tension(
                        self.year, self.mes, hora, subestacion, dia
                    )

                    # Si no se encontraron registros para el rango actual, salta a la
                    # siguiente iteración.
                    if len(fechas) == 0:
                        continue

                    # Calcula el rango intercuartil para las potencias activas y
                    # reemplaza las lecturas fuera del rango.
                    inferior, q2, superior = Util.calcular_quartil(tensiones)
                    tensiones = Util.reemplazar(tensiones, inferior, q2, superior)

                    # Agrega cada lectura depurada a la base de datos.
                    for i in range(len(tensiones)):
                        self.base_datos.insertar_tension_dep(
                            fechas[i],
                            hora,
                            subestacion,
                            dia,
                            tensiones[i],
                        )

        # Confirma los cambios en la base de datos.
        self.base_datos.commit()


class Util:
    """Clase con funciones estáticas para transformar entradas y/o realizar
    cálculos.
    """

    @staticmethod
    def parse_subestacion_potencia(subestacion):
        """Extrae, por medio de una expresión regular, el número de subestación a
        partir del texto en su cabecera en la hoja de potencia. Si la entrada es
        `'SUBEST 23'`, la salida será `23`; si la entrada es `'SUB 2 NOMBRE'`, la
        salida será `2`.
        """
        encontrado = re.findall(r"\d+", subestacion)

        if not encontrado:
            raise Exception(
                f"Error: No se encuentra el número de subestacion en '{subestacion}'."
            )

        return int(encontrado[0])

    @staticmethod
    def parse_subestacion_tension(subestacion):
        """Extrae, el número de subestación a partir del texto en su cabecera en la
        hoja de tensión.
        """
        if isinstance(subestacion, int):
            # Si la entrada es entera, devuelve la misma entrada.
            return subestacion

        # Si la entrada no es entera, será string; devuelve el número cortando el
        # prefijo "SE_" del nombre.
        return int(subestacion[3:])

    @staticmethod
    def parse_alimentador(alimentador):
        """Extrae, el código de alimentador a partir del texto en su cabecera en la
        hoja de potencia.
        """
        try:
            # Intenta convertir a entero la entrada.
            int_alimentador = int(alimentador)
        except:
            # Si no se puede convertir, se trata de un código de alimentador especial;
            # devuelve el nombre en mayúsculas sin números, sin espacios y sin ":".
            return re.sub(r"[0-9\s:]", "", alimentador.upper())

        # Si se consiguió convertir, devuelve el código formatéandolo con hasta cuatro
        # cifras.
        return f"{int_alimentador:04}"

    @staticmethod
    def parse_tension(tension):
        """Transforma el texto en la cabecera de tension nominal a un número de tensión
        en miles.
        """
        try:
            tension = tension.lower()
        except:
            raise Exception(
                f"Error: No se puede inferir el valor de tensión a partir de {tension}"
                + " indique si es V o kV."
            )

        # Busca el número entero o decimal en la entrada.
        encontrado = re.findall(r"\d+\.\d+|\d+", tension)

        if not encontrado:
            return tension

        # Convierte a decimal el número en string encontrado.
        valor = float(encontrado[0])

        if "kv" in tension:
            # Si existía 'kv' en la entrada, el decimal ya está en miles; devuelve el
            # valor.
            return valor

        # Si no existía 'kv' en la entrada, devuelve el valor en miles.
        return valor / 1000

    @staticmethod
    def calcular_quartil(datos):
        """Calcula los límites inferior y superior de una lista de datos, basado en el
        índice intercuartil.
        """
        # Calcula los quartiles.
        q1 = np.quantile(datos, 0.25)
        q2 = np.quantile(datos, 0.5)
        q3 = np.quantile(datos, 0.75)

        # Calcula el ínice intercuartil.
        iqr = q3 - q1

        # Define los límites inferior y superior
        superior = q3 + 1.5 * iqr
        inferior = q1 - 1.5 * iqr

        # Devuelve los límites y la media (quartil 2).
        return inferior, q2, superior

    @staticmethod
    def generar_horas(minutos=15):
        """Crea una lista de horas (empezando a las 00:00) en un día con una separación
        de `minutos`.
        """
        hora_actual = datetime.datetime(1, 1, 1)

        horas = []
        for _ in range(int(24 * 60 / minutos)):
            horas.append(hora_actual.time())
            hora_actual += datetime.timedelta(minutes=minutos)

        return horas

    @staticmethod
    def reemplazar(datos, inferior, media, superior):
        """Modifica los valores de una lista de datos que no se encuentren en entre
        `inferior` y `superior`, estos son reemplazados por `media`.
        """
        for i in range(len(datos)):
            if datos[i] < inferior or datos[i] > superior:
                datos[i] = media

        return datos


selector = Selector()
t1 = time.time()
CargaMes(
    selector.year,
    selector.mes,
    selector.mes_int,
    selector.directorio,
    selector.hoja_potencia,
    selector.hoja_tension,
)
t2 = time.time()

print(f"Proceso completado exitosamente en {(t2-t1)/60:.2f} m")
