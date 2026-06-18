import os
import shutil
import json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook

BASE_DIR = r"C:\Automatizaciones\Nomina"
PLANTILLA = os.path.join(BASE_DIR, "plantilla", "NOMINA.xlsx")
SALIDAS_DIR = os.path.join(BASE_DIR, "salidas")
TRAZA_DIR = os.path.join(BASE_DIR, "trazabilidad")

# Diccionario global para almacenar datos de empleados con sus horas
datos_empleados = {}

# Variable global para almacenar la ruta del último archivo creado
ultimo_archivo_creado = None

# Variables globales para modo corrección
modo_correccion = False
archivo_actual = None
valores_originales = {}  # Para almacenar valores originales antes de corrección

# Estructura global para almacenar novedades de comisiones
novedades_comisiones = {}

# Estructura global para calendario nocturnos
calendario_nocturnos = {}

# Estructura global para horas calculadas de nocturnos
horas_nocturnos_calculadas = {}

# Mapeo de tipos de comisión a columnas Excel (solo tipos reales en plantilla)
MAPEO_COMISIONES = {
    "Comisión contratación y R+R": "AI",
    "Comisión créditos": "AJ",
    "Comisión SINDRI": "AK",
    "Comisión venta de oro y traslado": "AL",
    "Comisión ROSETT": "AM",
    "Bonificación por reemplazo": "AN",
    "Bonificación variable por desplazamiento jefes de zona": "AO",
    "Bonificación variable N° joyerías jefes de zona": "AP",
    "Bonificación bunker": "AQ"
}

# Mapeo de comisiones externas pendientes de configuración (documentación)
MAPEO_COMISIONES_EXTERNAS_PENDIENTES = {
    "Comisión por rentabilidad": "PENDIENTE",
    "Comisión R. inventario": "PENDIENTE",
    "Comisión lingotes": "PENDIENTE",
    "Comisión LUXE": "PENDIENTE",
    "Clientes nuevos": "PENDIENTE",
    "Recaudo Sistecredito": "PENDIENTE",
    "Operaciones Efecty": "PENDIENTE",
    "Contratación Art.": "PENDIENTE",
    "Comisión por contratación": "PENDIENTE"
}

# Mantener compatibilidad con código existente
comisiones_mapping = MAPEO_COMISIONES


def centrar_ventana(ventana, ancho, alto):
    """Centra una ventana en la pantalla con el ancho y alto especificados."""
    ventana.update_idletasks()
    ancho_pantalla = ventana.winfo_screenwidth()
    alto_pantalla = ventana.winfo_screenheight()
    x = (ancho_pantalla - ancho) // 2
    y = (alto_pantalla - alto) // 2
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")


def obtener_hoja_por_quincena(quincena):
    """Retorna el nombre de la hoja según la quincena seleccionada."""
    if quincena == "Primera Quincena":
        return "Nómina 01-15"
    elif quincena == "Segunda Quincena":
        return "Nómina 16-30"
    else:
        return None


def leer_empleados(ruta_excel, quincena):
    """Lee los empleados desde el archivo Excel según la quincena.
    
    Retorna una lista de diccionarios con:
    - fila: número de fila en Excel
    - nombre: nombre del empleado
    """
    nombre_hoja = obtener_hoja_por_quincena(quincena)
    if not nombre_hoja:
        return []
    
    try:
        wb = load_workbook(ruta_excel, data_only=True)
        hoja = wb[nombre_hoja]
    except Exception as e:
        return []
    
    empleados = []
    fila_actual = 18
    filas_vacias_consecutivas = 0
    
    while filas_vacias_consecutivas < 5:
        celda = hoja[f'E{fila_actual}']
        valor = celda.value
        
        if valor and str(valor).strip():
            empleados.append({
                "fila": fila_actual,
                "nombre": str(valor).strip()
            })
            filas_vacias_consecutivas = 0
        else:
            filas_vacias_consecutivas += 1
        
        fila_actual += 1
    
    wb.close()
    return empleados


def leer_empleados_con_novedades(ruta_excel, quincena):
    """Lee los empleados y sus novedades actuales desde el archivo Excel según la quincena.
    
    Retorna una lista de diccionarios con:
    - fila: número de fila en Excel
    - nombre: nombre del empleado
    - hed, hen, hrdf, rn, rnd, hedf, hendf: valores actuales de novedades
    """
    nombre_hoja = obtener_hoja_por_quincena(quincena)
    if not nombre_hoja:
        return []
    
    try:
        wb = load_workbook(ruta_excel, data_only=True)
        hoja = wb[nombre_hoja]
    except Exception as e:
        return []
    
    empleados = []
    fila_actual = 18
    filas_vacias_consecutivas = 0
    
    # Mapeo de columnas de novedades
    columnas_novedades = {
        "hed": "U",
        "hen": "W",
        "hrdf": "Y",
        "rn": "AA",
        "rnd": "AC",
        "hedf": "AE",
        "hendf": "AG"
    }
    
    while filas_vacias_consecutivas < 5:
        celda_nombre = hoja[f'E{fila_actual}']
        valor_nombre = celda_nombre.value
        
        if valor_nombre and str(valor_nombre).strip():
            # Leer novedades
            novedades = {}
            for campo, columna in columnas_novedades.items():
                celda = hoja[f"{columna}{fila_actual}"]
                valor = celda.value
                novedades[campo] = valor if valor is not None else 0
            
            empleados.append({
                "fila": fila_actual,
                "nombre": str(valor_nombre).strip(),
                **novedades
            })
            filas_vacias_consecutivas = 0
        else:
            filas_vacias_consecutivas += 1
        
        fila_actual += 1
    
    wb.close()
    return empleados


def crear_nomina():
    global ultimo_archivo_creado, archivo_actual
    anio = entry_anio.get().strip()
    mes = combo_mes.get().strip()
    quincena = combo_quincena.get().strip()
    fecha_inicio = fecha_inicio_entry.get()
    fecha_fin = fecha_fin_entry.get()

    if not anio or not mes or not quincena:
        messagebox.showerror("Error", "Debes completar año, mes y quincena.")
        return

    if not os.path.exists(PLANTILLA):
        messagebox.showerror("Error", f"No se encontro la plantilla:\n{PLANTILLA}")
        return

    os.makedirs(SALIDAS_DIR, exist_ok=True)
    os.makedirs(TRAZA_DIR, exist_ok=True)

    nombre_archivo = f"NOMINA_{quincena.upper().replace(' ', '_')}_{mes.upper()}_{anio}.xlsx"
    ruta_salida = os.path.join(SALIDAS_DIR, nombre_archivo)

    if os.path.exists(ruta_salida):
        respuesta = messagebox.askyesno(
            "Archivo existente",
            "Ya existe una nomina con ese nombre.\n¿Deseas reemplazarla?"
        )
        if not respuesta:
            return

    shutil.copy2(PLANTILLA, ruta_salida)
    
    # Guardar la ruta del último archivo creado
    ultimo_archivo_creado = ruta_salida
    archivo_actual = ruta_salida
    
    print(f"[crear_nomina] Archivo creado: {archivo_actual}")

    trazabilidad = {
        "fecha_creacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "plantilla_origen": PLANTILLA,
        "archivo_generado": ruta_salida,
        "anio": anio,
        "mes": mes,
        "quincena": quincena,
        "periodo_novedades": {
            "fecha_inicio": fecha_inicio,
            "fecha_fin": fecha_fin
        }
    }

    nombre_traza = f"trazabilidad_{quincena.upper().replace(' ', '_')}_{mes.upper()}_{anio}.json"
    ruta_traza = os.path.join(TRAZA_DIR, nombre_traza)

    with open(ruta_traza, "w", encoding="utf-8") as f:
        json.dump(trazabilidad, f, indent=4, ensure_ascii=False)

    messagebox.showinfo(
        "Proceso exitoso",
        f"Nomina creada correctamente:\n\n{ruta_salida}\n\nTrazabilidad:\n{ruta_traza}"
    )


def cargar_empleados():
    """Carga y muestra los empleados desde la plantilla según la quincena seleccionada."""
    global datos_empleados, modo_correccion, archivo_actual, valores_originales
    quincena = combo_quincena.get().strip()
    
    if not quincena:
        messagebox.showerror("Error", "Debes seleccionar una quincena.")
        return
    
    if modo_correccion:
        # Modo corrección: cargar desde archivo existente
        if not archivo_actual or not os.path.exists(archivo_actual):
            messagebox.showerror("Error", "No hay archivo seleccionado para corrección.")
            return
        
        try:
            empleados = leer_empleados_con_novedades(archivo_actual, quincena)
            
            if not empleados:
                messagebox.showwarning("Advertencia", "No se encontraron empleados en la hoja seleccionada.")
                return
            
            # Limpiar tabla existente y diccionario
            for item in tree_empleados.get_children():
                tree_empleados.delete(item)
            datos_empleados.clear()
            valores_originales.clear()
            
            # Llenar tabla con empleados y sus novedades actuales
            for emp in empleados:
                fila_excel = emp["fila"]
                datos_empleados[fila_excel] = {
                    "fila": fila_excel,
                    "nombre": emp["nombre"],
                    "hed": emp["hed"],
                    "hen": emp["hen"],
                    "hrdf": emp["hrdf"],
                    "rn": emp["rn"],
                    "rnd": emp["rnd"],
                    "hedf": emp["hedf"],
                    "hendf": emp["hendf"]
                }
                # Guardar valores originales para trazabilidad
                valores_originales[fila_excel] = {
                    "hed": emp["hed"],
                    "hen": emp["hen"],
                    "hrdf": emp["hrdf"],
                    "rn": emp["rn"],
                    "rnd": emp["rnd"],
                    "hedf": emp["hedf"],
                    "hendf": emp["hendf"]
                }
                tree_empleados.insert("", "end", values=(
                    fila_excel,
                    emp["nombre"],
                    emp["hed"], emp["hen"], emp["hrdf"], emp["rn"], emp["rnd"], emp["hedf"], emp["hendf"]
                ))
            
            messagebox.showinfo("Carga exitosa", f"Se cargaron {len(empleados)} empleados con sus novedades actuales.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar empleados:\n{str(e)}")
    else:
        # Modo normal: cargar desde plantilla
        if not os.path.exists(PLANTILLA):
            messagebox.showerror("Error", f"No se encontro la plantilla:\n{PLANTILLA}")
            return
        
        try:
            empleados = leer_empleados(PLANTILLA, quincena)
            
            if not empleados:
                messagebox.showwarning("Advertencia", "No se encontraron empleados en la hoja seleccionada.")
                return
            
            # Limpiar tabla existente y diccionario
            for item in tree_empleados.get_children():
                tree_empleados.delete(item)
            datos_empleados.clear()
            valores_originales.clear()
            
            # Llenar tabla con empleados y inicializar horas en 0
            for emp in empleados:
                fila_excel = emp["fila"]
                datos_empleados[fila_excel] = {
                    "fila": fila_excel,
                    "nombre": emp["nombre"],
                    "hed": 0,
                    "hen": 0,
                    "hrdf": 0,
                    "rn": 0,
                    "rnd": 0,
                    "hedf": 0,
                    "hendf": 0
                }
                tree_empleados.insert("", "end", values=(
                    fila_excel,
                    emp["nombre"],
                    0, 0, 0, 0, 0, 0, 0
                ))
            
            messagebox.showinfo("Carga exitosa", f"Se cargaron {len(empleados)} empleados.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar empleados:\n{str(e)}")


def validar_numero(valor):
    """Valida que el valor sea numérico, aceptando punto o coma como decimal."""
    if not valor or valor.strip() == "":
        return 0.0
    valor = valor.strip().replace(",", ".")
    try:
        return float(valor)
    except ValueError:
        return None


def generar_nombre_correccion(ruta_original):
    """Genera un nombre de archivo de corrección con incremento.
    
    Si existe archivo_original_CORRECCION_01.xlsx, usa 02, etc.
    """
    base_dir = os.path.dirname(ruta_original)
    nombre_base = os.path.splitext(os.path.basename(ruta_original))[0]
    
    contador = 1
    while True:
        nombre_correccion = f"{nombre_base}_CORRECCION_{contador:02d}.xlsx"
        ruta_correccion = os.path.join(base_dir, nombre_correccion)
        if not os.path.exists(ruta_correccion):
            return ruta_correccion
        contador += 1


def abrir_modulo_nocturnos():
    """Abre el módulo de nocturnos en una ventana secundaria."""
    global calendario_nocturnos, horas_nocturnos_calculadas
    
    if not datos_empleados:
        messagebox.showerror("Error", "Primero debes cargar los empleados.")
        return
    
    # Crear ventana secundaria
    ventana_nocturnos = tk.Toplevel(ventana)
    ventana_nocturnos.title("Módulo de Nocturnos")
    ventana_nocturnos.resizable(True, True)
    ventana_nocturnos.grab_set()
    centrar_ventana(ventana_nocturnos, 900, 700)
    
    # Frame principal para botones (fuera del scroll)
    frame_botones_principal = ttk.Frame(ventana_nocturnos, padding=10)
    frame_botones_principal.pack(side="bottom", fill="x")
    
    # Canvas y scrollbar para contenido scrollable
    canvas = tk.Canvas(ventana_nocturnos)
    scrollbar = ttk.Scrollbar(ventana_nocturnos, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    canvas.pack(side="top", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    frame_nocturnos = ttk.Frame(scrollable_frame, padding=20)
    frame_nocturnos.pack(fill="both", expand=True)
    
    # Sección de selección de nocturnos
    frame_seleccion = ttk.LabelFrame(frame_nocturnos, text="Selección de Nocturnos")
    frame_seleccion.pack(fill="x", pady=10)
    
    # Obtener lista de empleados
    lista_empleados = [datos["nombre"] for datos in datos_empleados.values()]
    
    # Nocturno A
    ttk.Label(frame_seleccion, text="Nocturno A:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5, padx=5)
    combo_nocturno_a = ttk.Combobox(frame_seleccion, values=lista_empleados, state="readonly")
    combo_nocturno_a.grid(row=0, column=1, sticky="ew", pady=5, padx=5)
    
    # Nocturno B
    ttk.Label(frame_seleccion, text="Nocturno B:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5, padx=5)
    combo_nocturno_b = ttk.Combobox(frame_seleccion, values=lista_empleados, state="readonly")
    combo_nocturno_b.grid(row=1, column=1, sticky="ew", pady=5, padx=5)
    
    frame_seleccion.columnconfigure(1, weight=1)
    
    # Validación para impedir seleccionar mismo empleado
    def validar_nocturno_a(event):
        if combo_nocturno_a.get() == combo_nocturno_b.get():
            combo_nocturno_a.set("")
            messagebox.showwarning("Advertencia", "No puedes seleccionar el mismo empleado en ambos nocturnos.")
    
    def validar_nocturno_b(event):
        if combo_nocturno_b.get() == combo_nocturno_a.get():
            combo_nocturno_b.set("")
            messagebox.showwarning("Advertencia", "No puedes seleccionar el mismo empleado en ambos nocturnos.")
    
    combo_nocturno_a.bind("<<ComboboxSelected>>", validar_nocturno_a)
    combo_nocturno_b.bind("<<ComboboxSelected>>", validar_nocturno_b)
    
    # Sección de calendario
    frame_calendario = ttk.LabelFrame(frame_nocturnos, text="Calendario Mensual")
    frame_calendario.pack(fill="both", expand=True, pady=10)
    
    # Obtener año y mes
    anio = entry_anio.get().strip()
    mes = combo_mes.get().strip()
    
    # Mapeo de mes a número
    mes_a_numero = {
        "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,
        "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8,
        "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
    }
    
    try:
        numero_mes = mes_a_numero.get(mes, 1)
        anio_int = int(anio)
    except:
        messagebox.showerror("Error", "Año o mes no válidos.")
        ventana_nocturnos.destroy()
        return
    
    # Calcular días del mes
    import calendar
    dias_mes = calendar.monthrange(anio_int, numero_mes)[1]
    
    # Crear grid del calendario
    dias_semana = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
    
    # Encabezados de días
    for i, dia in enumerate(dias_semana):
        ttk.Label(frame_calendario, text=dia, font=("Arial", 10, "bold")).grid(row=0, column=i, padx=2, pady=2)
    
    # Diccionario para almacenar widgets de cada día
    widgets_dia = {}
    
    # Crear celdas para cada día
    dia_actual = 1
    fila = 1
    while dia_actual <= dias_mes:
        for col in range(7):
            if dia_actual > dias_mes:
                break
            
            # Frame para el día
            frame_dia = ttk.Frame(frame_calendario, borderwidth=1, relief="solid")
            frame_dia.grid(row=fila, column=col, padx=2, pady=2, sticky="nsew")
            
            # Número del día
            ttk.Label(frame_dia, text=str(dia_actual), font=("Arial", 10, "bold")).pack(anchor="w")
            
            # Combobox para selección (Libre, A, B)
            combo_dia = ttk.Combobox(frame_dia, values=["Libre", "A", "B"], state="readonly", width=5)
            combo_dia.set("Libre")
            combo_dia.pack(anchor="w")
            
            # Checkboxes
            var_festivo = tk.BooleanVar()
            var_descanso = tk.BooleanVar()
            
            check_festivo = ttk.Checkbutton(frame_dia, text="Festivo", variable=var_festivo)
            check_festivo.pack(anchor="w")
            
            check_descanso = ttk.Checkbutton(frame_dia, text="Descanso", variable=var_descanso)
            check_descanso.pack(anchor="w")
            
            # Almacenar widgets
            widgets_dia[dia_actual] = {
                "combo": combo_dia,
                "festivo": var_festivo,
                "descanso": var_descanso
            }
            
            dia_actual += 1
        fila += 1
    
    # Configurar pesos del grid
    for i in range(7):
        frame_calendario.columnconfigure(i, weight=1)
    
    # Función para guardar calendario
    def guardar_calendario():
        global calendario_nocturnos
        calendario_nocturnos = {}
        
        nocturno_a = combo_nocturno_a.get()
        nocturno_b = combo_nocturno_b.get()
        
        if not nocturno_a and not nocturno_b:
            messagebox.showwarning("Advertencia", "Debes seleccionar al menos un nocturno.")
            return
        
        for dia, widgets in widgets_dia.items():
            seleccion = widgets["combo"].get()
            festivo = widgets["festivo"].get()
            descanso = widgets["descanso"].get()
            
            if seleccion != "Libre":
                fecha = f"{anio_int}-{numero_mes:02d}-{dia:02d}"
                
                if seleccion == "A":
                    empleado = nocturno_a
                    etiqueta = "A"
                elif seleccion == "B":
                    empleado = nocturno_b
                    etiqueta = "B"
                else:
                    continue
                
                calendario_nocturnos[fecha] = {
                    "fecha": fecha,
                    "empleado_asignado": empleado,
                    "etiqueta": etiqueta,
                    "festivo": festivo,
                    "descanso_trabajado": descanso
                }
        
        messagebox.showinfo("Éxito", f"Calendario guardado con {len(calendario_nocturnos)} días asignados.")
    
    # Función para calcular resumen
    def calcular_resumen():
        if not calendario_nocturnos:
            messagebox.showwarning("Advertencia", "Primero debes guardar el calendario.")
            return
        
        # Calcular resumen por empleado
        resumen = {}
        
        for fecha, datos in calendario_nocturnos.items():
            empleado = datos["empleado_asignado"]
            
            if empleado not in resumen:
                resumen[empleado] = {
                    "dias_trabajados": 0,
                    "festivos_trabajados": 0,
                    "descansos_trabajados": 0
                }
            
            resumen[empleado]["dias_trabajados"] += 1
            if datos["festivo"]:
                resumen[empleado]["festivos_trabajados"] += 1
            if datos["descanso_trabajado"]:
                resumen[empleado]["descansos_trabajados"] += 1
        
        # Mostrar resumen en ventana
        ventana_resumen = tk.Toplevel(ventana_nocturnos)
        ventana_resumen.title("Resumen de Nocturnos")
        ventana_resumen.resizable(False, False)
        ventana_resumen.grab_set()
        centrar_ventana(ventana_resumen, 500, 400)
        
        frame_resumen = ttk.Frame(ventana_resumen, padding=20)
        frame_resumen.pack(fill="both", expand=True)
        
        ttk.Label(frame_resumen, text="Resumen de Nocturnos", font=("Arial", 14, "bold")).pack(pady=10)
        
        # Tabla de resumen
        frame_tabla = ttk.Frame(frame_resumen)
        frame_tabla.pack(fill="both", expand=True, pady=10)
        
        tree_resumen = ttk.Treeview(frame_tabla, columns=("empleado", "dias", "festivos", "descansos"), show="headings", height=8)
        tree_resumen.heading("empleado", text="Empleado")
        tree_resumen.heading("dias", text="Días trabajados")
        tree_resumen.heading("festivos", text="Festivos trabajados")
        tree_resumen.heading("descansos", text="Descansos trabajados")
        
        tree_resumen.column("empleado", width=200)
        tree_resumen.column("dias", width=100)
        tree_resumen.column("festivos", width=100)
        tree_resumen.column("descansos", width=100)
        
        for empleado, datos in resumen.items():
            tree_resumen.insert("", "end", values=(
                empleado,
                datos["dias_trabajados"],
                datos["festivos_trabajados"],
                datos["descansos_trabajados"]
            ))
        
        tree_resumen.pack(fill="both", expand=True)
        
        ttk.Button(frame_resumen, text="Cerrar", command=ventana_resumen.destroy).pack(pady=10)
    
    # Función para calcular horas nocturnos
    def calcular_horas_nocturnos():
        global horas_nocturnos_calculadas
        
        if not calendario_nocturnos:
            messagebox.showwarning("Advertencia", "Primero debes guardar el calendario.")
            return
        
        horas_nocturnos_calculadas = {}
        
        # Inicializar horas para cada empleado
        for fecha, datos in calendario_nocturnos.items():
            empleado = datos["empleado_asignado"]
            if empleado not in horas_nocturnos_calculadas:
                horas_nocturnos_calculadas[empleado] = {
                    "hed": 0,
                    "hen": 0,
                    "hrdf": 0,
                    "rn": 0,
                    "rnd": 0,
                    "hedf": 0,
                    "hendf": 0,
                    "festivos_aplicados": []
                }
        
        # Obtener día de la semana para cada fecha
        import datetime
        fechas_por_empleado = {}
        for fecha, datos in calendario_nocturnos.items():
            empleado = datos["empleado_asignado"]
            fecha_obj = datetime.datetime.strptime(fecha, "%Y-%m-%d")
            dia_semana = fecha_obj.weekday()  # 0=Lunes, 6=Domingo
            es_festivo = datos["festivo"]
            
            if empleado not in fechas_por_empleado:
                fechas_por_empleado[empleado] = []
            fechas_por_empleado[empleado].append((fecha, dia_semana, es_festivo))
        
        # Detectar bloques viernes-sábado-domingo (solo si no son festivos)
        bloques_fds = {}
        for empleado, fechas in fechas_por_empleado.items():
            fechas.sort()  # Ordenar por fecha
            i = 0
            while i < len(fechas):
                fecha, dia, es_festivo = fechas[i]
                
                # Si es festivo, no puede ser parte de bloque FDS
                if es_festivo:
                    i += 1
                    continue
                
                # Si es viernes (4), verificar si sigue sábado (5) y domingo (6)
                if dia == 4:  # Viernes
                    # Verificar si hay sábado y domingo consecutivos no festivos
                    bloque = [fecha]
                    j = i + 1
                    dias_esperados = [5, 6]  # Sábado, Domingo
                    
                    for dia_esperado in dias_esperados:
                        if j < len(fechas):
                            fecha_sig, dia_sig, es_festivo_sig = fechas[j]
                            if dia_sig == dia_esperado and not es_festivo_sig:
                                # Verificar que sea el día siguiente
                                fecha_obj = datetime.datetime.strptime(fecha, "%Y-%m-%d")
                                fecha_sig_obj = datetime.datetime.strptime(fecha_sig, "%Y-%m-%d")
                                if (fecha_sig_obj - fecha_obj).days == 1:
                                    bloque.append(fecha_sig)
                                    fecha = fecha_sig
                                    j += 1
                                else:
                                    break
                            else:
                                break
                        else:
                            break
                    
                    # Si es un bloque completo de 3 días
                    if len(bloque) == 3:
                        if empleado not in bloques_fds:
                            bloques_fds[empleado] = []
                        bloques_fds[empleado].append(bloque)
                        i = j  # Saltar los días del bloque
                    else:
                        i += 1
                else:
                    i += 1
        
        # Calcular horas
        for empleado, fechas in fechas_por_empleado.items():
            dias_en_bloque = set()
            # Marcar días que están en bloques FDS
            if empleado in bloques_fds:
                for bloque in bloques_fds[empleado]:
                    for fecha in bloque:
                        dias_en_bloque.add(fecha)
            
            for fecha, dia, es_festivo in fechas:
                # Si está en bloque FDS, se procesa aparte
                if fecha in dias_en_bloque:
                    continue
                
                # Si es festivo, aplicar regla de festivo trabajado
                if es_festivo:
                    horas_nocturnos_calculadas[empleado]["hed"] += 2.5
                    horas_nocturnos_calculadas[empleado]["hen"] += 4.5
                    horas_nocturnos_calculadas[empleado]["hrdf"] += 10.5
                    horas_nocturnos_calculadas[empleado]["hendf"] += 5
                    horas_nocturnos_calculadas[empleado]["festivos_aplicados"].append(fecha)
                    continue
                
                # Lunes a jueves (0-3) no festivos
                if dia <= 3:
                    horas_nocturnos_calculadas[empleado]["hed"] += 2
                    horas_nocturnos_calculadas[empleado]["rn"] += 9.5
        
        # Procesar bloques FDS
        for empleado, bloques in bloques_fds.items():
            for bloque in bloques:
                horas_nocturnos_calculadas[empleado]["rn"] += 29
                horas_nocturnos_calculadas[empleado]["hedf"] += 7.5
                horas_nocturnos_calculadas[empleado]["hendf"] += 0.5
        
        messagebox.showinfo("Éxito", f"Horas nocturnos calculadas para {len(horas_nocturnos_calculadas)} empleados.")
    
    # Función para mostrar resumen de horas calculadas
    def mostrar_resumen_horas():
        if not horas_nocturnos_calculadas:
            messagebox.showwarning("Advertencia", "Primero debes calcular las horas nocturnos.")
            return
        
        ventana_resumen = tk.Toplevel(ventana_nocturnos)
        ventana_resumen.title("Resumen de Horas Nocturnos")
        ventana_resumen.resizable(False, False)
        ventana_resumen.grab_set()
        centrar_ventana(ventana_resumen, 800, 500)
        
        frame_resumen = ttk.Frame(ventana_resumen, padding=20)
        frame_resumen.pack(fill="both", expand=True)
        
        ttk.Label(frame_resumen, text="Resumen de Horas Nocturnos", font=("Arial", 14, "bold")).pack(pady=10)
        
        # Tabla de resumen
        frame_tabla = ttk.Frame(frame_resumen)
        frame_tabla.pack(fill="both", expand=True, pady=10)
        
        tree_resumen = ttk.Treeview(frame_tabla, columns=("empleado", "hed", "hen", "hrdf", "rn", "rnd", "hedf", "hendf", "festivos"), show="headings", height=8)
        tree_resumen.heading("empleado", text="Empleado")
        tree_resumen.heading("hed", text="HED")
        tree_resumen.heading("hen", text="HEN")
        tree_resumen.heading("hrdf", text="HRDF")
        tree_resumen.heading("rn", text="RN")
        tree_resumen.heading("rnd", text="RND")
        tree_resumen.heading("hedf", text="HEDF")
        tree_resumen.heading("hendf", text="HENDF")
        tree_resumen.heading("festivos", text="Festivos aplicados")
        
        tree_resumen.column("empleado", width=150)
        tree_resumen.column("hed", width=60)
        tree_resumen.column("hen", width=60)
        tree_resumen.column("hrdf", width=60)
        tree_resumen.column("rn", width=60)
        tree_resumen.column("rnd", width=60)
        tree_resumen.column("hedf", width=60)
        tree_resumen.column("hendf", width=60)
        tree_resumen.column("festivos", width=150)
        
        for empleado, horas in horas_nocturnos_calculadas.items():
            festivos_text = ", ".join(horas["festivos_aplicados"]) if horas["festivos_aplicados"] else ""
            tree_resumen.insert("", "end", values=(
                empleado,
                horas["hed"],
                horas["hen"],
                horas["hrdf"],
                horas["rn"],
                horas["rnd"],
                horas["hedf"],
                horas["hendf"],
                festivos_text
            ))
        
        tree_resumen.pack(fill="both", expand=True)
        
        ttk.Button(frame_resumen, text="Cerrar", command=ventana_resumen.destroy).pack(pady=10)
    
    # Función para aplicar horas a tabla principal
    def aplicar_horas_tabla():
        global datos_empleados
        
        if not horas_nocturnos_calculadas:
            messagebox.showwarning("Advertencia", "Primero debes calcular las horas nocturnos.")
            return
        
        # Buscar fila de cada empleado y sumar horas
        for empleado, horas_calc in horas_nocturnos_calculadas.items():
            for fila, datos in datos_empleados.items():
                if datos["nombre"] == empleado:
                    # Sumar a valores existentes
                    datos["hed"] += horas_calc["hed"]
                    datos["hen"] += horas_calc["hen"]
                    datos["hrdf"] += horas_calc["hrdf"]
                    datos["rn"] += horas_calc["rn"]
                    datos["rnd"] += horas_calc["rnd"]
                    datos["hedf"] += horas_calc["hedf"]
                    datos["hendf"] += horas_calc["hendf"]
                    break
        
        # Actualizar tabla
        for item in tree_empleados.get_children():
            valores = tree_empleados.item(item, "values")
            fila_excel = int(valores[0])
            if fila_excel in datos_empleados:
                datos = datos_empleados[fila_excel]
                tree_empleados.item(item, values=(
                    fila_excel,
                    datos["nombre"],
                    datos["hed"],
                    datos["hen"],
                    datos["hrdf"],
                    datos["rn"],
                    datos["rnd"],
                    datos["hedf"],
                    datos["hendf"]
                ))
        
        messagebox.showinfo("Éxito", f"Horas nocturnos aplicadas a la tabla principal para {len(horas_nocturnos_calculadas)} empleados.")
    
    # Botones en frame principal (fuera del scroll)
    ttk.Button(frame_botones_principal, text="Guardar calendario", command=guardar_calendario).pack(side="left", padx=5)
    ttk.Button(frame_botones_principal, text="Calcular resumen nocturnos", command=calcular_resumen).pack(side="left", padx=5)
    ttk.Button(frame_botones_principal, text="Calcular horas nocturnos", command=calcular_horas_nocturnos).pack(side="left", padx=5)
    ttk.Button(frame_botones_principal, text="Ver resumen horas", command=mostrar_resumen_horas).pack(side="left", padx=5)
    ttk.Button(frame_botones_principal, text="Aplicar horas a tabla principal", command=aplicar_horas_tabla).pack(side="left", padx=5)
    ttk.Button(frame_botones_principal, text="Cerrar", command=ventana_nocturnos.destroy).pack(side="left", padx=5)


def mostrar_resumen_validacion(tipo_guardado):
    """Muestra ventana modal de resumen y validación antes de guardar.
    
    Args:
        tipo_guardado: 'horas' o 'comisiones'
    
    Returns:
        True si el usuario confirma, False si cancela
    """
    # Recopilar datos
    anio = entry_anio.get().strip()
    mes = combo_mes.get().strip()
    quincena = combo_quincena.get().strip()
    periodo = f"{quincena} de {mes} {anio}"
    
    # Validaciones y advertencias
    advertencias = []
    empleados_modificados = []
    total_horas = 0
    total_comisiones = 0
    
    if tipo_guardado == 'horas':
        # Analizar horas
        for fila, datos in datos_empleados.items():
            horas_totales = datos['hed'] + datos['hen'] + datos['hrdf'] + datos['rn'] + datos['rnd'] + datos['hedf'] + datos['hendf']
            if horas_totales > 0:
                empleados_modificados.append(datos['nombre'])
                total_horas += horas_totales
                if horas_totales > 40:
                    advertencias.append(f"{datos['nombre']}: {horas_totales} horas (supera 40)")
        
        # Analizar comisiones si existen
        for tipo_comision, empleados_comision in novedades_comisiones.items():
            for fila, valor in empleados_comision.items():
                total_comisiones += valor
                if valor > 5000000:
                    nombre = datos_empleados.get(fila, {}).get('nombre', 'Desconocido')
                    advertencias.append(f"{nombre}: Comisión {tipo_comision} de ${valor:,.2f} (supera $5,000,000)")
    else:  # comisiones
        # Analizar comisiones
        for tipo_comision, empleados_comision in novedades_comisiones.items():
            for fila, valor in empleados_comision.items():
                nombre = datos_empleados.get(fila, {}).get('nombre', 'Desconocido')
                if nombre not in empleados_modificados:
                    empleados_modificados.append(nombre)
                total_comisiones += valor
                if valor > 5000000:
                    advertencias.append(f"{nombre}: Comisión {tipo_comision} de ${valor:,.2f} (supera $5,000,000)")
        
        # Analizar horas si existen
        for fila, datos in datos_empleados.items():
            horas_totales = datos['hed'] + datos['hen'] + datos['hrdf'] + datos['rn'] + datos['rnd'] + datos['hedf'] + datos['hendf']
            if horas_totales > 0:
                total_horas += horas_totales
                if horas_totales > 40:
                    advertencias.append(f"{datos['nombre']}: {horas_totales} horas (supera 40)")
    
    # Validación: no hay empleados modificados
    if not empleados_modificados:
        messagebox.showwarning("Advertencia", "No hay empleados con novedades para guardar.")
        return False
    
    # Crear ventana modal
    ventana_resumen = tk.Toplevel(ventana)
    ventana_resumen.title("Resumen y Validación")
    ventana_resumen.resizable(False, False)
    ventana_resumen.geometry("600x500")
    ventana_resumen.grab_set()
    centrar_ventana(ventana_resumen, 600, 500)
    
    frame_resumen = ttk.Frame(ventana_resumen, padding=20)
    frame_resumen.pack(fill="both", expand=True)
    
    # Título
    ttk.Label(frame_resumen, text="Resumen de Novedades", font=("Arial", 14, "bold")).pack(pady=10)
    
    # Periodo
    ttk.Label(frame_resumen, text=f"Período: {periodo}", font=("Arial", 10)).pack(anchor="w", pady=5)
    
    # Empleados modificados
    ttk.Label(frame_resumen, text=f"Empleados con novedades: {len(empleados_modificados)}", font=("Arial", 10)).pack(anchor="w", pady=5)
    
    # Horas totales
    if total_horas > 0:
        ttk.Label(frame_resumen, text=f"Total horas: {total_horas}", font=("Arial", 10)).pack(anchor="w", pady=5)
    
    # Comisiones totales
    if total_comisiones > 0:
        ttk.Label(frame_resumen, text=f"Total comisiones: ${total_comisiones:,.2f}", font=("Arial", 10)).pack(anchor="w", pady=5)
    
    # Lista de empleados
    frame_lista = ttk.LabelFrame(frame_resumen, text="Empleados con novedades")
    frame_lista.pack(fill="both", expand=True, pady=10)
    
    lista_empleados = tk.Listbox(frame_lista, height=8)
    scrollbar_lista = ttk.Scrollbar(frame_lista, orient="vertical", command=lista_empleados.yview)
    lista_empleados.configure(yscrollcommand=scrollbar_lista.set)
    
    for emp in empleados_modificados:
        lista_empleados.insert("end", emp)
    
    lista_empleados.pack(side="left", fill="both", expand=True)
    scrollbar_lista.pack(side="right", fill="y")
    
    # Advertencias
    if advertencias:
        frame_advertencias = ttk.LabelFrame(frame_resumen, text="Advertencias")
        frame_advertencias.pack(fill="x", pady=10)
        
        lista_advertencias = tk.Listbox(frame_advertencias, height=5, bg="#fff3cd")
        scrollbar_adv = ttk.Scrollbar(frame_advertencias, orient="vertical", command=lista_advertencias.yview)
        lista_advertencias.configure(yscrollcommand=scrollbar_adv.set)
        
        for adv in advertencias:
            lista_advertencias.insert("end", adv)
        
        lista_advertencias.pack(side="left", fill="both", expand=True)
        scrollbar_adv.pack(side="right", fill="y")
    
    # Variable para resultado
    resultado = [False]
    
    def confirmar():
        resultado[0] = True
        ventana_resumen.destroy()
    
    def cancelar():
        resultado[0] = False
        ventana_resumen.destroy()
    
    # Botones
    frame_botones = ttk.Frame(frame_resumen)
    frame_botones.pack(pady=15)
    
    ttk.Button(frame_botones, text="Confirmar y Guardar", command=confirmar).pack(side="left", padx=5)
    ttk.Button(frame_botones, text="Cancelar", command=cancelar).pack(side="left", padx=5)
    
    # Esperar a que se cierre la ventana
    ventana_resumen.wait_window()
    
    return resultado[0]


def corregir_nomina_existente():
    """Abre filedialog para seleccionar un archivo existente y entra en modo corrección."""
    global modo_correccion, archivo_actual, ultimo_archivo_creado
    
    from tkinter import filedialog
    
    ruta_archivo = filedialog.askopenfilename(
        title="Seleccionar nómina existente",
        initialdir=SALIDAS_DIR,
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    
    if not ruta_archivo:
        return
    
    # Verificar que sea un archivo .xlsx
    if not ruta_archivo.lower().endswith(".xlsx"):
        messagebox.showerror("Error", "Debes seleccionar un archivo .xlsx")
        return
    
    # Verificar que esté en el directorio de salidas usando comparación robusta de rutas
    salidas_abs = os.path.normcase(os.path.abspath(SALIDAS_DIR))
    archivo_abs = os.path.normcase(os.path.abspath(ruta_archivo))
    
    if os.path.commonpath([salidas_abs, archivo_abs]) != salidas_abs:
        messagebox.showerror("Error", "El archivo debe estar en el directorio de salidas")
        return
    
    # Activar modo corrección
    modo_correccion = True
    archivo_actual = ruta_archivo
    ultimo_archivo_creado = ruta_archivo
    
    messagebox.showinfo(
        "Modo corrección activado",
        f"Archivo seleccionado:\n{ruta_archivo}\n\n"
        "Ahora puedes cargar empleados y editar sus novedades.\n"
        "Al guardar, se creará una copia corregida sin modificar el original."
    )


def editar_empleado(event):
    """Abre ventana secundaria para editar las horas del empleado seleccionado."""
    seleccion = tree_empleados.selection()
    if not seleccion:
        return
    
    item = seleccion[0]
    valores = tree_empleados.item(item, "values")
    fila_excel = int(valores[0])
    
    if fila_excel not in datos_empleados:
        messagebox.showerror("Error", "Empleado no encontrado en datos.")
        return
    
    datos = datos_empleados[fila_excel]
    
    # Crear ventana secundaria
    ventana_edicion = tk.Toplevel(ventana)
    ventana_edicion.title(f"Editar horas - {datos['nombre']}")
    ventana_edicion.resizable(False, False)
    ventana_edicion.grab_set()
    centrar_ventana(ventana_edicion, 450, 400)
    
    frame_edicion = ttk.Frame(ventana_edicion, padding=20)
    frame_edicion.pack(fill="both", expand=True)
    
    # Información del empleado
    frame_info = ttk.Frame(frame_edicion)
    frame_info.pack(fill="x", pady=(0, 10))
    ttk.Label(frame_info, text=f"Empleado: {datos['nombre']}", font=("Arial", 10, "bold")).pack(anchor="w")
    ttk.Label(frame_info, text=f"Fila Excel: {fila_excel}").pack(anchor="w")
    
    # Frame para los campos de edición
    frame_campos = ttk.Frame(frame_edicion)
    frame_campos.pack(fill="both", expand=True)
    
    campos = [
        ("HED (Horas Extras Diurnas)", "hed"),
        ("HEN (Horas Extras Nocturnas)", "hen"),
        ("HRDF (Horas Recargo Dominical/Festivo)", "hrdf"),
        ("RN (Recargos Nocturnos)", "rn"),
        ("RND (Recargo Nocturno Dominical)", "rnd"),
        ("HEDF (Horas Extras Dominicales/Festivas)", "hedf"),
        ("HENDF (Horas Extras Nocturnas Dominicales/Festivas)", "hendf")
    ]
    
    entries = {}
    for i, (label, key) in enumerate(campos):
        ttk.Label(frame_campos, text=label).grid(row=i, column=0, sticky="w", pady=3, padx=5)
        entry = ttk.Entry(frame_campos)
        entry.insert(0, str(datos[key]))
        entry.grid(row=i, column=1, sticky="ew", pady=3, padx=5)
        entries[key] = entry
    
    frame_campos.columnconfigure(1, weight=1)
    
    def guardar_cambios():
        try:
            nuevos_valores = {}
            for key, entry in entries.items():
                valor = entry.get().strip()
                if valor:
                    valor = valor.replace(",", ".")
                    try:
                        nuevos_valores[key] = float(valor)
                    except ValueError:
                        messagebox.showerror("Error", f"Valor inválido en {key}. Use números con punto o coma.")
                        return
                else:
                    nuevos_valores[key] = 0.0
            
            # Actualizar diccionario global
            datos_empleados[fila_excel].update(nuevos_valores)
            
            # Actualizar tabla
            tree_empleados.item(item, values=(
                fila_excel,
                datos["nombre"],
                nuevos_valores["hed"],
                nuevos_valores["hen"],
                nuevos_valores["hrdf"],
                nuevos_valores["rn"],
                nuevos_valores["rnd"],
                nuevos_valores["hedf"],
                nuevos_valores["hendf"]
            ))
            
            messagebox.showinfo("Éxito", "Horas actualizadas correctamente.")
            ventana_edicion.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar cambios:\n{str(e)}")
    
    def cancelar():
        ventana_edicion.destroy()
    
    frame_botones = ttk.Frame(frame_edicion)
    frame_botones.pack(pady=15)
    
    ttk.Button(frame_botones, text="Guardar cambios", command=guardar_cambios).pack(side="left", padx=5)
    ttk.Button(frame_botones, text="Cancelar", command=cancelar).pack(side="left", padx=5)


def guardar_horas_excel():
    """Guarda las horas de los empleados en el archivo Excel generado."""
    global archivo_actual, modo_correccion, valores_originales
    
    if not archivo_actual:
        messagebox.showerror("Error", "Primero debes crear o seleccionar una nómina.")
        return
    
    print(f"[guardar_horas_excel] Archivo actual: {archivo_actual}")
    
    if not os.path.exists(archivo_actual):
        messagebox.showerror("Error", f"El archivo no existe:\n{archivo_actual}")
        return
    
    quincena = combo_quincena.get().strip()
    if not quincena:
        messagebox.showerror("Error", "Debes seleccionar una quincena.")
        return
    
    nombre_hoja = obtener_hoja_por_quincena(quincena)
    if not nombre_hoja:
        messagebox.showerror("Error", "Quincena no válida.")
        return
    
    # Mostrar resumen y validación antes de guardar
    if not mostrar_resumen_validacion('horas'):
        return
    
    # Mapeo de campos a columnas Excel
    columnas_mapping = {
        "hed": "U",
        "hen": "W",
        "hrdf": "Y",
        "rn": "AA",
        "rnd": "AC",
        "hedf": "AE",
        "hendf": "AG"
    }
    
    # Determinar archivo de trabajo
    if modo_correccion:
        archivo_origen = archivo_actual
        archivo_destino = generar_nombre_correccion(archivo_origen)
    else:
        archivo_origen = archivo_actual
        archivo_destino = archivo_actual
    
    print(f"[guardar_horas_excel] Archivo origen: {archivo_origen}")
    print(f"[guardar_horas_excel] Archivo destino: {archivo_destino}")
    print(f"[guardar_horas_excel] Modo corrección: {modo_correccion}")
    
    try:
        from openpyxl import load_workbook
        print(f"[guardar_horas_excel] Abriendo archivo: {archivo_origen}")
        wb = load_workbook(archivo_origen)
        
        if nombre_hoja not in wb.sheetnames:
            messagebox.showerror("Error", f"La hoja '{nombre_hoja}' no existe en el archivo.")
            wb.close()
            return
        
        hoja = wb[nombre_hoja]
        
        # Lista para trazabilidad
        trazabilidad_horas = []
        
        # Recorrer filas del Treeview
        for item in tree_empleados.get_children():
            valores = tree_empleados.item(item, "values")
            fila_excel = int(valores[0])
            nombre = valores[1]
            hed = valores[2]
            hen = valores[3]
            hrdf = valores[4]
            rn = valores[5]
            rnd = valores[6]
            hedf = valores[7]
            hendf = valores[8]
            
            # Obtener datos del diccionario global
            if fila_excel not in datos_empleados:
                continue
            
            datos = datos_empleados[fila_excel]
            
            # Escribir valores en las columnas correspondientes
            for campo, columna in columnas_mapping.items():
                valor = datos[campo]
                celda = hoja[f"{columna}{fila_excel}"]
                celda.value = valor
                
                # Agregar a trazabilidad
                if modo_correccion and fila_excel in valores_originales:
                    valor_anterior = valores_originales[fila_excel][campo]
                    trazabilidad_horas.append({
                        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "archivo_origen": archivo_origen,
                        "archivo_corregido": archivo_destino,
                        "hoja": nombre_hoja,
                        "empleado": nombre,
                        "fila": fila_excel,
                        "columna_excel": columna,
                        "campo": campo,
                        "valor_anterior": valor_anterior,
                        "valor_nuevo": valor
                    })
                else:
                    trazabilidad_horas.append({
                        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "archivo_modificado": archivo_destino,
                        "hoja": nombre_hoja,
                        "empleado": nombre,
                        "fila": fila_excel,
                        "columna_excel": columna,
                        "campo": campo,
                        "valor_guardado": valor
                    })
        
        # Guardar el archivo
        print(f"[guardar_horas_excel] Guardando archivo: {archivo_destino}")
        wb.save(archivo_destino)
        print(f"[guardar_horas_excel] Guardado correctamente")
        wb.close()
        
        # Actualizar variables globales si es modo corrección
        if modo_correccion:
            archivo_actual = archivo_destino
            ultimo_archivo_creado = archivo_destino
        
        # Crear archivo de trazabilidad para las horas
        if trazabilidad_horas:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if modo_correccion:
                nombre_traza_horas = f"trazabilidad_correccion_{timestamp}.json"
            else:
                nombre_traza_horas = f"trazabilidad_horas_{timestamp}.json"
            ruta_traza_horas = os.path.join(TRAZA_DIR, nombre_traza_horas)
            
            with open(ruta_traza_horas, "w", encoding="utf-8") as f:
                json.dump(trazabilidad_horas, f, indent=4, ensure_ascii=False)
        
        mensaje = f"Horas guardadas correctamente en:\n{archivo_destino}\n\nTrazabilidad:\n{ruta_traza_horas if trazabilidad_horas else 'Sin cambios'}"
        if modo_correccion:
            mensaje = f"Archivo de corrección creado:\n{archivo_destino}\n\nOriginal no modificado:\n{archivo_origen}\n\nTrazabilidad:\n{ruta_traza_horas if trazabilidad_horas else 'Sin cambios'}"
        
        messagebox.showinfo("Éxito", f"Excel actualizado correctamente:\n{archivo_destino}\n\n{mensaje}")
        
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar horas en Excel:\n{str(e)}")


def abrir_modulo_comisiones():
    """Abre el módulo de comisiones en una ventana secundaria."""
    if not datos_empleados:
        messagebox.showerror("Error", "Primero debes cargar los empleados.")
        return
    
    # Crear ventana secundaria
    ventana_comisiones = tk.Toplevel(ventana)
    ventana_comisiones.title("Módulo de Comisiones")
    ventana_comisiones.resizable(True, True)
    ventana_comisiones.grab_set()
    centrar_ventana(ventana_comisiones, 850, 700)
    
    # Frame principal para botones (fuera del scroll)
    frame_botones_principal = ttk.Frame(ventana_comisiones, padding=10)
    frame_botones_principal.pack(side="bottom", fill="x")
    
    # Canvas y scrollbar para contenido scrollable
    canvas = tk.Canvas(ventana_comisiones)
    scrollbar = ttk.Scrollbar(ventana_comisiones, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    canvas.pack(side="top", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    frame_comisiones = ttk.Frame(scrollable_frame, padding=20)
    frame_comisiones.pack(fill="both", expand=True)
    
    # Tipo de comisión
    ttk.Label(frame_comisiones, text="Tipo de comisión:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
    combo_tipo_comision = ttk.Combobox(frame_comisiones, values=list(comisiones_mapping.keys()), state="readonly")
    combo_tipo_comision.grid(row=0, column=1, sticky="ew", pady=5)
    combo_tipo_comision.set(list(comisiones_mapping.keys())[0])
    
    # Valor total
    ttk.Label(frame_comisiones, text="Valor total de comisión:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
    entry_valor_total = ttk.Entry(frame_comisiones)
    entry_valor_total.grid(row=1, column=1, sticky="ew", pady=5)
    
    # Repartir en partes iguales
    ttk.Label(frame_comisiones, text="Repartir en partes iguales:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky="w", pady=5)
    combo_partes_iguales = ttk.Combobox(frame_comisiones, values=["Sí", "No"], state="readonly")
    combo_partes_iguales.grid(row=2, column=1, sticky="ew", pady=5)
    combo_partes_iguales.set("Sí")
    
    frame_comisiones.columnconfigure(1, weight=1)
    
    # Frame para lista de empleados
    frame_empleados = ttk.LabelFrame(frame_comisiones, text="Seleccionar empleados")
    frame_empleados.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=10)
    
    # Treeview para empleados con checkbox
    tree_empleados_comision = ttk.Treeview(frame_empleados, columns=("seleccion", "fila", "nombre"), show="headings", height=10)
    tree_empleados_comision.heading("seleccion", text="✓")
    tree_empleados_comision.heading("fila", text="Fila")
    tree_empleados_comision.heading("nombre", text="Nombre")
    tree_empleados_comision.column("seleccion", width=40, anchor="center")
    tree_empleados_comision.column("fila", width=60, anchor="center")
    tree_empleados_comision.column("nombre", width=300, anchor="w")
    
    scrollbar_emp = ttk.Scrollbar(frame_empleados, orient="vertical", command=tree_empleados_comision.yview)
    tree_empleados_comision.configure(yscrollcommand=scrollbar_emp.set)
    
    tree_empleados_comision.pack(side="left", fill="both", expand=True)
    scrollbar_emp.pack(side="right", fill="y")
    
    # Llenar treeview con empleados
    empleados_seleccionados = {}
    for fila, datos in datos_empleados.items():
        item_id = tree_empleados_comision.insert("", "end", values=(" ", fila, datos["nombre"]))
        empleados_seleccionados[item_id] = {"fila": fila, "nombre": datos["nombre"], "seleccionado": False}
    
    # Evento para toggle selección
    def toggle_seleccion(event):
        item = tree_empleados_comision.selection()
        if item:
            item = item[0]
            current = tree_empleados_comision.item(item, "values")
            if current[0] == " ":
                tree_empleados_comision.item(item, values=("✓", current[1], current[2]))
                empleados_seleccionados[item]["seleccionado"] = True
            else:
                tree_empleados_comision.item(item, values=(" ", current[1], current[2]))
                empleados_seleccionados[item]["seleccionado"] = False
    
    tree_empleados_comision.bind("<Double-1>", toggle_seleccion)
    
    # Frame para distribución manual (inicialmente oculto)
    frame_distribucion_manual = ttk.Frame(frame_comisiones)
    
    # Frame para preview de distribución igual
    frame_preview = ttk.Frame(frame_comisiones)
    
    def actualizar_interfaz():
        partes_iguales = combo_partes_iguales.get()
        if partes_iguales == "No":
            frame_distribucion_manual.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=10)
            frame_preview.grid_forget()
            # Ocultar columna de check en modo manual
            tree_empleados_comision.column("seleccion", width=0, stretch=False)
            tree_empleados_comision.heading("seleccion", text="")
        else:
            frame_distribucion_manual.grid_forget()
            frame_preview.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=10)
            # Mostrar columna de check en modo igual
            tree_empleados_comision.column("seleccion", width=40, stretch=False)
            tree_empleados_comision.heading("seleccion", text="✓")
    
    combo_partes_iguales.bind("<<ComboboxSelected>>", lambda e: actualizar_interfaz())
    
    # Función para calcular distribución igual
    def calcular_distribucion_igual():
        try:
            valor_total = entry_valor_total.get().strip().replace(",", ".")
            valor_total = float(valor_total)
        except ValueError:
            messagebox.showerror("Error", "Valor total inválido.")
            return None
        
        seleccionados = [emp for emp in empleados_seleccionados.values() if emp["seleccionado"]]
        if not seleccionados:
            messagebox.showerror("Error", "Debes seleccionar al menos un empleado.")
            return None
        
        valor_por_empleado = valor_total / len(seleccionados)
        return {emp["fila"]: valor_por_empleado for emp in seleccionados}
    
    # Variables para distribución manual
    entries_manual = {}
    
    # Construir interfaz de distribución manual
    ttk.Label(frame_distribucion_manual, text="Asignar valor manual por empleado:", font=("Arial", 10, "bold")).pack(pady=5)
    
    frame_manual_campos = ttk.Frame(frame_distribucion_manual)
    frame_manual_campos.pack(fill="both", expand=True)
    
    for i, emp in enumerate(sorted(empleados_seleccionados.values(), key=lambda x: x["fila"])):
        ttk.Label(frame_manual_campos, text=emp["nombre"]).grid(row=i, column=0, sticky="w", pady=2, padx=5)
        entry = ttk.Entry(frame_manual_campos, width=15)
        entry.insert(0, "0")
        entry.grid(row=i, column=1, sticky="w", pady=2, padx=5)
        entries_manual[emp["fila"]] = entry
    
    # Función para guardar comisiones en memoria
    def guardar_comisiones_memoria():
        tipo_comision = combo_tipo_comision.get()
        partes_iguales = combo_partes_iguales.get()
        
        try:
            valor_total = entry_valor_total.get().strip().replace(",", ".")
            valor_total = float(valor_total)
        except ValueError:
            messagebox.showerror("Error", "Valor total inválido.")
            return
        
        if partes_iguales == "Sí":
            seleccionados = [emp for emp in empleados_seleccionados.values() if emp["seleccionado"]]
            if not seleccionados:
                messagebox.showerror("Error", "Debes seleccionar al menos un empleado.")
                return
            
            distribucion = calcular_distribucion_igual()
            if not distribucion:
                return
        else:
            # Distribución manual - ignorar checks, usar todos los empleados
            distribucion = {}
            suma_asignada = 0
            for fila, datos in datos_empleados.items():
                valor = entries_manual[fila].get().strip().replace(",", ".")
                try:
                    valor = float(valor)
                except ValueError:
                    messagebox.showerror("Error", f"Valor inválido para empleado {datos['nombre']}.")
                    return
                
                # Solo incluir empleados con valor mayor a 0
                if valor > 0:
                    distribucion[fila] = valor
                    suma_asignada += valor
            
            if not distribucion:
                messagebox.showerror("Error", "Debes asignar un valor mayor a 0 al menos a un empleado.")
                return
            
            if abs(suma_asignada - valor_total) > 0.01:
                diferencia = valor_total - suma_asignada
                if diferencia > 0:
                    messagebox.showerror("Error", f"Falta asignar ${diferencia:.2f}. La suma actual es ${suma_asignada:.2f} y el total es ${valor_total:.2f}.")
                else:
                    messagebox.showerror("Error", f"Sobra ${abs(diferencia):.2f}. La suma actual es ${suma_asignada:.2f} y el total es ${valor_total:.2f}.")
                return
        
        # Guardar en estructura global
        if tipo_comision not in novedades_comisiones:
            novedades_comisiones[tipo_comision] = {}
        
        for fila, valor in distribucion.items():
            novedades_comisiones[tipo_comision][fila] = valor
        
        messagebox.showinfo("Éxito", f"Comisiones guardadas en memoria para {len(distribucion)} empleados.")
    
    # Botones en frame principal (fuera del scroll)
    ttk.Button(frame_botones_principal, text="Guardar en memoria", command=guardar_comisiones_memoria).pack(side="left", padx=5)
    ttk.Button(frame_botones_principal, text="Cerrar", command=ventana_comisiones.destroy).pack(side="left", padx=5)
    
    # Inicializar interfaz
    actualizar_interfaz()


def guardar_comisiones_excel():
    """Guarda las comisiones de los empleados en el archivo Excel generado."""
    global archivo_actual
    
    if not archivo_actual:
        messagebox.showerror("Error", "Primero debes crear o seleccionar una nómina.")
        return
    
    print(f"[guardar_comisiones_excel] Archivo actual: {archivo_actual}")
    
    if not os.path.exists(archivo_actual):
        messagebox.showerror("Error", f"El archivo no existe:\n{archivo_actual}")
        return
    
    if not novedades_comisiones:
        messagebox.showerror("Error", "No hay comisiones guardadas en memoria. Primero usa el módulo de comisiones.")
        return
    
    quincena = combo_quincena.get().strip()
    if not quincena:
        messagebox.showerror("Error", "Debes seleccionar una quincena.")
        return
    
    # Mostrar resumen y validación antes de guardar
    if not mostrar_resumen_validacion('comisiones'):
        return
    
    nombre_hoja = obtener_hoja_por_quincena(quincena)
    if not nombre_hoja:
        messagebox.showerror("Error", "Quincena no válida.")
        return
    
    try:
        from openpyxl import load_workbook
        print(f"[guardar_comisiones_excel] Archivo a modificar: {archivo_actual}")
        print(f"[guardar_comisiones_excel] Abriendo archivo: {archivo_actual}")
        wb = load_workbook(archivo_actual)
        
        if nombre_hoja not in wb.sheetnames:
            messagebox.showerror("Error", f"La hoja '{nombre_hoja}' no existe en el archivo.")
            wb.close()
            return
        
        hoja = wb[nombre_hoja]
        
        # Lista para trazabilidad
        trazabilidad_comisiones = []
        
        # Recorrer tipos de comisión
        for tipo_comision, empleados_comision in novedades_comisiones.items():
            columna_excel = comisiones_mapping.get(tipo_comision)
            if not columna_excel:
                continue
            
            # Verificar si la columna está pendiente de configuración
            if columna_excel == "PENDIENTE":
                messagebox.showerror(
                    "Error",
                    f"No se ha configurado la columna Excel para el tipo de comisión:\n'{tipo_comision}'\n\nPor favor configure la columna en el diccionario MAPEO_COMISIONES antes de guardar."
                )
                wb.close()
                return
            
            # Recorrer empleados para este tipo de comisión
            for fila, valor in empleados_comision.items():
                # Obtener nombre del empleado
                if fila in datos_empleados:
                    nombre = datos_empleados[fila]["nombre"]
                else:
                    nombre = "Desconocido"
                
                # Escribir valor en la celda correspondiente
                celda = hoja[f"{columna_excel}{fila}"]
                celda.value = valor
                
                # Agregar a trazabilidad
                trazabilidad_comisiones.append({
                    "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "archivo_modificado": archivo_actual,
                    "hoja": nombre_hoja,
                    "empleado": nombre,
                    "fila": fila,
                    "tipo_comision": tipo_comision,
                    "columna_excel": columna_excel,
                    "valor_guardado": valor
                })
        
        # Guardar el archivo
        print(f"[guardar_comisiones_excel] Guardando archivo: {archivo_actual}")
        wb.save(archivo_actual)
        print(f"[guardar_comisiones_excel] Guardado correctamente")
        wb.close()
        
        # Crear archivo de trazabilidad para las comisiones
        if trazabilidad_comisiones:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_traza_comisiones = f"trazabilidad_comisiones_{timestamp}.json"
            ruta_traza_comisiones = os.path.join(TRAZA_DIR, nombre_traza_comisiones)
            
            with open(ruta_traza_comisiones, "w", encoding="utf-8") as f:
                json.dump(trazabilidad_comisiones, f, indent=4, ensure_ascii=False)
        
        messagebox.showinfo(
            "Éxito",
            f"Excel actualizado correctamente:\n{archivo_actual}\n\nComisiones guardadas correctamente en:\n{archivo_actual}\n\nTrazabilidad:\n{ruta_traza_comisiones if trazabilidad_comisiones else 'Sin cambios'}"
        )
        
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar comisiones en Excel:\n{str(e)}")


ventana = tk.Tk()
ventana.title("GAMAN - Generador de Nomina V1")
ventana.resizable(True, True)
centrar_ventana(ventana, 900, 700)

frame = ttk.Frame(ventana, padding=20)
frame.pack(fill="both", expand=True)

titulo = ttk.Label(frame, text="Generador de Nomina V1", font=("Arial", 16, "bold"))
titulo.pack(pady=10)

form = ttk.Frame(frame)
form.pack(pady=10, fill="x")

ttk.Label(form, text="Año:").grid(row=0, column=0, sticky="w", pady=5)
entry_anio = ttk.Entry(form)
entry_anio.insert(0, "2026")
entry_anio.grid(row=0, column=1, sticky="ew", pady=5)

ttk.Label(form, text="Mes:").grid(row=1, column=0, sticky="w", pady=5)
combo_mes = ttk.Combobox(form, values=[
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
], state="readonly")
combo_mes.set("Junio")
combo_mes.grid(row=1, column=1, sticky="ew", pady=5)

ttk.Label(form, text="Quincena:").grid(row=2, column=0, sticky="w", pady=5)
combo_quincena = ttk.Combobox(form, values=[
    "Primera Quincena",
    "Segunda Quincena"
], state="readonly")
combo_quincena.set("Primera Quincena")
combo_quincena.grid(row=2, column=1, sticky="ew", pady=5)

ttk.Label(form, text="Inicio novedades:").grid(row=3, column=0, sticky="w", pady=5)
fecha_inicio_entry = DateEntry(form, date_pattern="dd/mm/yyyy")
fecha_inicio_entry.grid(row=3, column=1, sticky="ew", pady=5)

ttk.Label(form, text="Fin novedades:").grid(row=4, column=0, sticky="w", pady=5)
fecha_fin_entry = DateEntry(form, date_pattern="dd/mm/yyyy")
fecha_fin_entry.grid(row=4, column=1, sticky="ew", pady=5)

form.columnconfigure(1, weight=1)

# Frame para organizar botones
frame_botones = ttk.Frame(frame)
frame_botones.pack(pady=10, fill="x")

# Primera fila de botones
frame_botones_fila1 = ttk.Frame(frame_botones)
frame_botones_fila1.pack(fill="x", pady=2)

btn_corregir = ttk.Button(frame_botones_fila1, text="Corregir nómina existente", command=corregir_nomina_existente)
btn_corregir.pack(side="left", padx=5)

btn_crear = ttk.Button(frame_botones_fila1, text="Crear nomina desde plantilla", command=crear_nomina)
btn_crear.pack(side="left", padx=5)

btn_cargar = ttk.Button(frame_botones_fila1, text="Cargar empleados", command=cargar_empleados)
btn_cargar.pack(side="left", padx=5)

btn_guardar_horas = ttk.Button(frame_botones_fila1, text="Guardar horas en Excel", command=guardar_horas_excel)
btn_guardar_horas.pack(side="left", padx=5)

# Segunda fila de botones
frame_botones_fila2 = ttk.Frame(frame_botones)
frame_botones_fila2.pack(fill="x", pady=2)

btn_modulo_comisiones = ttk.Button(frame_botones_fila2, text="Módulo de comisiones", command=abrir_modulo_comisiones)
btn_modulo_comisiones.pack(side="left", padx=5)

btn_guardar_comisiones = ttk.Button(frame_botones_fila2, text="Guardar comisiones en Excel", command=guardar_comisiones_excel)
btn_guardar_comisiones.pack(side="left", padx=5)

btn_modulo_nocturnos = ttk.Button(frame_botones_fila2, text="Módulo nocturnos", command=abrir_modulo_nocturnos)
btn_modulo_nocturnos.pack(side="left", padx=5)

# Tabla para mostrar empleados
frame_tabla = ttk.LabelFrame(frame, text="Empleados detectados")
frame_tabla.pack(pady=10, fill="both", expand=True)

tree_empleados = ttk.Treeview(frame_tabla, columns=("fila", "nombre", "hed", "hen", "hrdf", "rn", "rnd", "hedf", "hendf"), show="headings", height=12)
tree_empleados.heading("fila", text="Fila Excel")
tree_empleados.heading("nombre", text="Nombre")
tree_empleados.heading("hed", text="HED")
tree_empleados.heading("hen", text="HEN")
tree_empleados.heading("hrdf", text="HRDF")
tree_empleados.heading("rn", text="RN")
tree_empleados.heading("rnd", text="RND")
tree_empleados.heading("hedf", text="HEDF")
tree_empleados.heading("hendf", text="HENDF")
tree_empleados.column("fila", width=80, anchor="center")
tree_empleados.column("nombre", width=260, anchor="w")
tree_empleados.column("hed", width=70, anchor="center")
tree_empleados.column("hen", width=70, anchor="center")
tree_empleados.column("hrdf", width=70, anchor="center")
tree_empleados.column("rn", width=70, anchor="center")
tree_empleados.column("rnd", width=70, anchor="center")
tree_empleados.column("hedf", width=70, anchor="center")
tree_empleados.column("hendf", width=70, anchor="center")

scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical", command=tree_empleados.yview)
tree_empleados.configure(yscrollcommand=scrollbar.set)

tree_empleados.bind("<Double-1>", editar_empleado)
tree_empleados.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

ventana.mainloop()