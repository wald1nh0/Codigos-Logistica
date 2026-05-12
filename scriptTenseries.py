import pdfplumber
import pandas as pd
import os
import re
import sys 
import tkinter as tk
from tkinter import filedialog

# ==========================================
#   CONFIGURACIÓN EXCEL MAESTRO
# ==========================================
NOMBRE_HOJA_EXCEL  = "Distribución R.M"

# COLUMNAS EN TU EXCEL (ORIGEN)
COL_TRACKING_EXCEL = "Seguimiento" 
COL_NOMBRE_CLIENTE = "Nombre Cliente"
COL_TELEFONO       = "Telefono Cliente"
COL_DIRECCION      = "Direccion" 
COL_REFERENCIA     = "Referencia"
COL_COMUNA         = "Comuna"
COL_VEHICULO       = "Vehículo"   # <--- NUEVO: Columna del transporte

# Columnas de producto
COL_OC_EXCEL       = "OC" 
COL_SKU_EXCEL      = "SKU"
COL_PRODUCTO_EXCEL = "PRODUCTO"
colUnidadesExcel   = "Unidades"
COL_CANT_BULTOS    = "Bultos" 

# ==========================================

datos_extraidos = []

root = tk.Tk()
root.withdraw()

print(">>> Abriendo ventana para seleccionar los PDFs TEN SERIES...")
archivos_pdf = filedialog.askopenfilenames(
    title="Selecciona PDFs de Etiquetas TEN SERIES",
    filetypes=[("Archivos PDF", "*.pdf")]
)

if not archivos_pdf:
    print("No seleccionaste nada. Adiós.")
    sys.exit()

print(f"Has seleccionado {len(archivos_pdf)} archivo(s).")

print("\n>>> Abriendo ventana para seleccionar la PLANILLA MAESTRA...")
ruta_plantilla = filedialog.askopenfilename(
    title="Selecciona el Excel Maestro",
    filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")]
)
if not ruta_plantilla:
    sys.exit()

nombre_base = input("\nEscribe el nombre para el archivo final (sin .xlsx): ").strip().replace('.xlsx', '')
nombre_final = f"{nombre_base}.xlsx"
print(f"\nProcesando etiquetas...")

# ==========================================
#   PROCESAMIENTO PDF
# ==========================================

for indice, ruta_actual in enumerate(archivos_pdf):
    nombre_archivo = os.path.basename(ruta_actual)
    print(f"[{indice + 1}/{len(archivos_pdf)}] Leyendo: {nombre_archivo} ...")

    with pdfplumber.open(ruta_actual) as pdf:
        for i, pagina in enumerate(pdf.pages):
            
            texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
            if not texto_pagina: continue

            # 1. SEGUIMIENTO
            match_id = re.search(r'(\d{4}-\d{8}-\d{4})', texto_pagina)
            id_full = match_id.group(1) if match_id else "No encontrado"
        

            # 2. DATOS QR
            match_envio = re.search(r'ENVÍO:\s*"?(\d+)', texto_pagina)
            shp = match_envio.group(1) if match_envio else ""

            match_control = re.search(r'CONTROL:\s*"?(\d+)', texto_pagina)
            code = match_control.group(1) if match_control else ""
            
            lbc = id_full
            context = "tenseries-cl"
            type_qr = "zmv1"

            contenido_qr = f'{{"type":"{type_qr}","context":"{context}","shp":"{shp}","lbc":"{lbc}","code":"{code}"}}'
            
            match_bulto = re.search(r'(Bulto\s*[a-zA-Z0-9\-]*)', texto_pagina)
            seg_interno = match_bulto.group(1) if match_bulto else "No encontrado"

            datos_etiqueta = {
                'Seguimiento': id_full,  
                'Seg Interno': seg_interno,     
                'Contenido qr': contenido_qr 
            }
            datos_extraidos.append(datos_etiqueta)

# ==========================================
#   CRUCE CON EXCEL
# ==========================================

if datos_extraidos:
    print(f"\nLeyendo Plantilla Maestra '{NOMBRE_HOJA_EXCEL}'...")
    df_pdf = pd.DataFrame(datos_extraidos)

    # Ordenamos PDF
    df_pdf['Seguimiento'] = df_pdf['Seguimiento'].astype(str).str.strip()
    df_pdf = df_pdf.sort_values(by=['Seguimiento', 'Seg Interno'])
    df_pdf['posicion_item'] = df_pdf.groupby('Seguimiento').cumcount()

    # Leemos Excel
    df_master = pd.read_excel(ruta_plantilla, sheet_name=NOMBRE_HOJA_EXCEL, dtype=str)
    df_master.columns = df_master.columns.str.strip() 
    
    if COL_TRACKING_EXCEL not in df_master.columns:
        print(f"ERROR: No encontré la columna '{COL_TRACKING_EXCEL}' en el Excel.")
        sys.exit()

    df_master[COL_TRACKING_EXCEL] = df_master[COL_TRACKING_EXCEL].astype(str).str.strip()
    df_master[COL_TRACKING_EXCEL] = df_master[COL_TRACKING_EXCEL].str.replace(r'\.0$', '', regex=True)

    # --- DIRECCION FINAL ---
    if COL_DIRECCION in df_master.columns and COL_REFERENCIA in df_master.columns:
        df_master['Direccion Final'] = df_master[COL_DIRECCION].fillna('') + f"\nReferencia: "+ df_master[COL_REFERENCIA].fillna('') + f"\n" + df_master[COL_COMUNA].fillna('') 
    elif COL_REFERENCIA not in df_master.columns:
        df_master['Direccion Final'] = df_master[COL_DIRECCION].fillna('') + f"\n" + df_master[COL_COMUNA].fillna('')
    else:
        df_master['Direccion Final'] = "No encontrada"

    # --- EXPANSIÓN BULTOS ---
    if COL_CANT_BULTOS not in df_master.columns:
        print("AVISO: No vi columna 'Bultos', asumiendo 1.")
        df_master[COL_CANT_BULTOS] = 1
    else:
        df_master[COL_CANT_BULTOS] = df_master[COL_CANT_BULTOS].replace('', '1').fillna('1').astype(float).astype(int)

    df_master_expandido = df_master.loc[df_master.index.repeat(df_master[COL_CANT_BULTOS])].copy()
    df_master_expandido['contador_bulto'] = df_master_expandido.groupby(level=0).cumcount() + 1
    
    mask_multi = df_master_expandido[COL_CANT_BULTOS] > 1
    df_master_expandido.loc[mask_multi, COL_PRODUCTO_EXCEL] = (
        df_master_expandido.loc[mask_multi, COL_PRODUCTO_EXCEL].astype(str) + " (" + 
        df_master_expandido.loc[mask_multi, 'contador_bulto'].astype(str) + "/" + 
        df_master_expandido.loc[mask_multi, COL_CANT_BULTOS].astype(str) + ")"
    )

    df_master_expandido['posicion_item'] = df_master_expandido.groupby(COL_TRACKING_EXCEL).cumcount()

    # --- MERGE ---
    # AGREGAMOS 'COL_VEHICULO' A LA LISTA DE COSAS ÚTILES A TRAER
    columnas_a_traer = [
        COL_TRACKING_EXCEL, COL_OC_EXCEL, COL_SKU_EXCEL, COL_PRODUCTO_EXCEL, 
        colUnidadesExcel, COL_NOMBRE_CLIENTE, COL_TELEFONO, 'Direccion Final', 
        COL_VEHICULO, 'posicion_item' # <--- AQUÍ ESTÁ EL VEHÍCULO
    ]
    cols_utiles = [c for c in columnas_a_traer if c in df_master_expandido.columns]
    
    df_final = pd.merge(
        df_pdf, 
        df_master_expandido[cols_utiles], 
        left_on=['Seguimiento', 'posicion_item'], 
        right_on=[COL_TRACKING_EXCEL, 'posicion_item'], 
        how='left'
    )

    # --- ORDENAR Y FILTRAR COLUMNAS FINALES ---
    columnas_deseadas = {
        'Seguimiento': 'Seguimiento',
        'Seg Interno': 'Seg Interno', 
        'Contenido qr': 'Contenido qr',
        COL_NOMBRE_CLIENTE: 'Destinatario',
        'Direccion Final': 'Direccion',
        COL_TELEFONO: 'Telefono',
        COL_OC_EXCEL: 'OC',
        COL_SKU_EXCEL: 'SKU',
        COL_PRODUCTO_EXCEL: 'PRODUCTO',
        colUnidadesExcel: 'Unidades',
        COL_VEHICULO: 'Vehiculo'  # <--- AGREGADO AL FINAL
    }

    df_export = pd.DataFrame()
    for col_origen, col_destino in columnas_deseadas.items():
        if col_origen in df_final.columns:
            df_export[col_destino] = df_final[col_origen]
        else:
            df_export[col_destino] = "" 

    # Guardar
    if os.path.exists(nombre_final):
        print(f"Agregando a '{nombre_final}'...")
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