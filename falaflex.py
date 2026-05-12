import pdfplumber
import pandas as pd
import os
import re
import sys 
import tkinter as tk
from tkinter import filedialog

# Declaracion Variables
nombreExcel = "Distribución R.M"
segAbreviado = "Seguimiento" 
nombreCliente = "Nombre Cliente"
tlfCliente = "Telefono Cliente"
dirCliente = "Direccion" 
refCliente = "Referencia"
comunaCliente = "Comuna"
vehiculoRuta = "Vehiculo"
ocExcel = "OC" 
skuExcel = "SKU"
prodExcel = "PRODUCTO"
unidadesExcel = "Unidades"
bultosExcel = "Bultos"          
datos_extraidos = []

# Declaracion ruta archivo
root = tk.Tk()
root.withdraw()

print("Abriendo ventana para seleccion de PDFs: ")
archivosPDF = filedialog.askopenfilenames(
    title="Selecciona PDFs de Etiquetas TEN SERIES",
    filetypes=[("Archivos PDF", "*.pdf")]
)

if not archivosPDF:
    print("No seleccionaste nada, adios ")
    sys.exit()

print(f"Has seleccionado {len(archivosPDF)} archivo(s)")

print("Abriendo ventana para seleccionar excel maestro. ")
rutaPlantilla = filedialog.askopenfilename(
    title="Selecciona el Excel Maestro",
    filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")]
)

if not rutaPlantilla:
    sys.exit()

nombre_base = input("\nEscribe el nombre para el archivo final (sin .xlsx): ").strip().replace('.xlsx', '')
nombre_final = f"{nombre_base}.xlsx"
print(f"\nProcesando etiquetas...")

for indice, rutaActual in enumerate(archivosPDF):
    nombreArchivo = os.path.basename(rutaActual)
    print(f"[{indice + 1}/{len(archivosPDF)}] Leyendo: {nombreArchivo} ...")

    with pdfplumber.open(rutaActual) as pdf:
        for i, pagina in enumerate(pdf.pages):
            
            texto = pagina.extract_text(x_tolerance=2, y_tolerance=2)
            if not texto: continue

            # --- LÓGICA DE EXTRACCIÓN (MODIFICADA PARA OC) ---

            # 1. BUSCAR LA OC (N- de orden)
            # Buscamos la frase "N- de orden:" y capturamos los números que siguen
            match_oc = re.search(r'N-\s*de\s*orden:?\s*(\d+)', texto, re.IGNORECASE)
            
            oc_encontrada = "No encontrada"
            
            if match_oc:
                oc_encontrada = match_oc.group(1)
            else:
                # Si falla, intentamos buscar un número de 10 dígitos que empiece con 3 (Formato Falabella)
                match_fallback = re.search(r'\b3\d{9}\b', texto)
                if match_fallback:
                    oc_encontrada = match_fallback.group(0)

            # 2. EXTRACCIÓN DE BULTOS (Para el orden interno)
            match_bulto = re.search(r'BULTO\(S\):\s*(\d+)\s*de\s*(\d+)', texto)
            if match_bulto:
                bulto_actual = match_bulto.group(1) 
                seg_interno = f"Bulto {bulto_actual}" 
            else:
                seg_interno = "Bulto 1"

            # Guardamos usando la OC como llave
            datos_etiqueta = {
                'OC_CRUCE': oc_encontrada,
                'Seg Interno': seg_interno,
            }
            datos_extraidos.append(datos_etiqueta)

# --- SECCIÓN DE CRUCE CON EXCEL (NUEVA LÓGICA COMENTADA) ---

if datos_extraidos:
    print(f"\nLeyendo Plantilla Maestra '{nombreExcel}'...")
    df_pdf = pd.DataFrame(datos_extraidos)

    # Limpiamos la OC extraída del PDF (quitamos espacios)
    df_pdf['OC_CRUCE'] = df_pdf['OC_CRUCE'].astype(str).str.strip()
    
    # Ordenamos el PDF por OC y luego por Bulto para mantener el orden de las etiquetas
    df_pdf = df_pdf.sort_values(by=['OC_CRUCE', 'Seg Interno'])
    
    # Generamos un ID de posición (0, 1, 2...) para poder cruzar bulto a bulto
    df_pdf['posicion_item'] = df_pdf.groupby('OC_CRUCE').cumcount()

    # Leemos el Excel Maestro
    df_master = pd.read_excel(rutaPlantilla, sheet_name=nombreExcel, dtype=str)
    df_master.columns = df_master.columns.str.strip() 

    # Verificamos que la columna OC exista en el Excel
    if ocExcel not in df_master.columns:
        print(f"ERROR: No encontré la columna '{ocExcel}' en el Excel.")
        sys.exit()

    # Limpiamos la columna OC del Excel (quitamos espacios y decimales .0)
    df_master[ocExcel] = df_master[ocExcel].astype(str).str.strip()
    df_master[ocExcel] = df_master[ocExcel].str.replace(r'\.0$', '', regex=True)

    # Creamos la Dirección Final uniendo Calle + Comuna
    if dirCliente in df_master.columns:
        df_master['Direccion Final'] = df_master[dirCliente].fillna('').astype(str) + f"\nReferencia:" +df_master[refCliente].fillna('').astype(str) + f"\n" + df_master[comunaCliente].fillna('').astype(str)
    else:
        df_master['Direccion Final'] = df_master.get(dirCliente, "")

    # --- EXPANSIÓN DE FILAS (MULTIBULTO) ---
    # Si no existe columna Bultos, asumimos 1
    if bultosExcel not in df_master.columns:
        df_master[bultosExcel] = 1
    else:
        df_master[bultosExcel] = df_master[bultosExcel].replace('', '1').fillna('1').astype(float).astype(int)

    # Aquí duplicamos las filas del Excel según la cantidad de bultos
    df_master_expandido = df_master.loc[df_master.index.repeat(df_master[bultosExcel])].copy()
    
    # Numeramos los bultos (1, 2, 3...)
    df_master_expandido['contador_bulto'] = df_master_expandido.groupby(level=0).cumcount() + 1
    
    # Agregamos (1/3) al nombre del producto si son varios bultos
    mask_multi = df_master_expandido[bultosExcel] > 1
    df_master_expandido.loc[mask_multi, prodExcel] = (
        df_master_expandido.loc[mask_multi, prodExcel].astype(str) + " (" + 
        df_master_expandido.loc[mask_multi, 'contador_bulto'].astype(str) + "/" + 
        df_master_expandido.loc[mask_multi, bultosExcel].astype(str) + ")"
    )

    # Generamos el ID de posición en el Excel para que coincida con el PDF
    df_master_expandido['posicion_item'] = df_master_expandido.groupby(ocExcel).cumcount()

    # --- MERGE (EL CRUCE FINAL) ---
    columnas_a_traer = [
        ocExcel, segAbreviado, skuExcel, prodExcel, unidadesExcel, 
        nombreCliente, tlfCliente, 'Direccion Final', vehiculoRuta, 
        'posicion_item'
    ]
    # Filtramos solo columnas que existan para evitar errores
    cols_utiles = [c for c in columnas_a_traer if c in df_master_expandido.columns]
    
    # Cruzamos usando la OC y la Posición
    df_final = pd.merge(
        df_pdf, 
        df_master_expandido[cols_utiles], 
        left_on=['OC_CRUCE', 'posicion_item'], 
        right_on=[ocExcel, 'posicion_item'], 
        how='left'
    )

    # --- LIMPIEZA Y EXPORTACIÓN ---
    columnas_finales = {
        'OC_CRUCE': 'OC',
        segAbreviado: 'Seguimiento', # Traemos el seguimiento del Excel
        'Seg Interno': 'Seg Interno',
        nombreCliente: 'Destinatario',
        'Direccion Final': 'Direccion',
        tlfCliente: 'Telefono',
        skuExcel: 'SKU',
        prodExcel: 'PRODUCTO',
        unidadesExcel: 'Unidades',
        vehiculoRuta: 'Vehículo'
    }

    df_export = pd.DataFrame()
    for col_origen, col_destino in columnas_finales.items():
        if col_origen in df_final.columns:
            df_export[col_destino] = df_final[col_origen]
        else:
            df_export[col_destino] = "" 

    if os.path.exists(nombre_final):
        print(f"Agregando datos a '{nombre_final}'...")
        try:
            df_existente = pd.read_excel(nombre_final, dtype=str)
            df_consolidado = pd.concat([df_existente, df_export], ignore_index=True)
            df_consolidado.to_excel(nombre_final, index=False)
        except:
             df_export.to_excel(nombre_final, index=False)
    else:
        print(f"Creando '{nombre_final}'...")
        df_export.to_excel(nombre_final, index=False)

    print(f"¡Listo! Procesadas: {len(df_pdf)}")
    input("Presiona ENTER para salir...")

else:
    print("No se encontró información válida en los PDFs.")
    input("Enter para salir...")