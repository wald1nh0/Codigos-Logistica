import pdfplumber
import pandas as pd
import os
import re
import sys 
import tkinter as tk
from tkinter import filedialog

#nombre hoja excel
NOMBRE_HOJA_EXCEL  = "EDITABLE" 

#titulos plantilla
COL_TRACKING_EXCEL = "SEG"
COL_OC_EXCEL       = "OC" 
COL_SKU_EXCEL      = "SKU"
COL_PRODUCTO_EXCEL = "PRODUCTO"
colUnidadesExcel   = "Unidades"
COL_CANT_BULTOS    = "Bultos" 

#cajas
caja_TipoEnvio = None
caja_SegInterno = None
caja_SegAbreviado = None
caja_Direccion = None
caja_destinatario = None
caja_rut = None
caja_telefono = None
caja_segInternoWalmart = None
cajaNroBulto = None
punto = None
texto_SegInterno_OjoHumano = None

#variables
datos_extraidos = []
cont = 0

# Inicializamos tkinter y ocultamos la ventana principal fea
root = tk.Tk()
root.withdraw()

print(">>> Abriendo ventana para seleccionar LOS PDFs de Etiquetas...")
# CAMBIO 1: Usamos 'askopenfilenames' (con 's' al final) para permitir selección múltiple
archivos_pdf = filedialog.askopenfilenames(
    title="Selecciona UNO o VARIOS PDFs de Etiquetas",
    filetypes=[("Archivos PDF", "*.pdf")]
)

if not archivos_pdf:
    print("No seleccionaste ningún PDF. Cerrando programa.")
    sys.exit()

# Mostramos cuántos archivos se seleccionaron
print(f"Has seleccionado {len(archivos_pdf)} archivo(s).")

print("\n>>> Abriendo ventana para seleccionar la Plantilla Maestra...")
ruta_plantilla = filedialog.askopenfilename(
    title="Selecciona el Excel Plantilla Maestro",
    filetypes=[("Archivos Excel", "*.xlsx *.xls *.xlsm")]
)
if not ruta_plantilla:
    print("No seleccionaste la Plantilla. Cerrando programa.")
    sys.exit()
print(f"Plantilla Seleccionada: {ruta_plantilla}")

# El nombre del archivo final lo pedimos por texto porque es para CREAR uno nuevo
nombre_base = input("\nEscribe el nombre para el archivo final (sin .xlsx): ").strip().replace('.xlsx', '')
nombre_final = f"{nombre_base}.xlsx"

#seleccion de tipo de etiqueta
print("\n--- TIPO DE ETIQUETA ---")
# Nota: Asumimos que TODOS los PDFs seleccionados son del MISMO TIPO (ej: todos Starken o todos Bluex)
while cont != 1:
    tipo_etiqueta = input(f"'S' Starken | 'B' Bluex | 'Z' ZipNova | 'P' Paris| 'WS' Walmart Starken | 'W' Walmart interno | 'PB' Paris Bluex | 'R' Ripley: ").strip().replace("'","").upper()

    if tipo_etiqueta == "S": 
        caja_TipoEnvio = (48, 246, 84, 258)                 
        caja_SegInterno = (132, 204, 264, 222)              
        caja_SegAbreviado = (72, 390, 192, 430)
        caja_Direccion = (84 , 294, 276, 318)
        caja_destinatario = (90, 318, 276, 326)
        caja_rut = (90, 326, 114, 336)
        caja_telefono = (120, 326, 168, 336)
        cajaNroBulto = (294,294,348,324)
        cont +=1

    elif tipo_etiqueta == "B": 
        caja_TipoEnvio = (156,222,318,276)                 
        caja_SegInterno = (186,198,414,216)          
        caja_SegAbreviado = (216,507,378,534)               
        caja_Direccion = (42,564,522,600)
        caja_destinatario = (42,546,504,564)
        caja_telefono = (42,618,180,636)
        cajaNroBulto = (474,704,558,726)
        cont +=1

    elif tipo_etiqueta == "Z": 
        caja_TipoEnvio = (96,282,180,324)                 
        caja_SegInterno = (96,254,240,276)
        caja_SegAbreviado = (60,126,162,135)                 
        caja_Direccion = (18,153,180,160)
        caja_destinatario = (18,143,156,150)
        caja_telefono = (18,142,300,152)
        cajaNroBulto = (150,15,192,42)
        cont +=1
        
    elif tipo_etiqueta == "WS": 
        caja_TipoEnvio = (16,35,42,44)                 
        caja_SegInterno = (84,117,197,123)
        caja_SegAbreviado = (93,43,152,51)                 
        caja_Direccion = (10,124,246,138)
        caja_destinatario = (10,159,114,174)
        caja_telefono = (0, 159,168,174)
        caja_segInternoWalmart = (78,280,180,292)
        punto = "STARKEN"
        cont +=1

    elif tipo_etiqueta == "W": 
        caja_TipoEnvio = (12,46,54,56) 
        caja_SegAbreviado = (84,114,168,123)
        caja_SegInterno = (78,276,174,288)                
        caja_Direccion = (0,153,252,168)
        caja_destinatario = (0,132,252,144)
        caja_telefono = (0,132,252,144)
        punto = "ENVIAME"
        cont +=1

    elif tipo_etiqueta == "P": 
        caja_TipoEnvio = (12,33,42,45)
        caja_SegInterno =(72,117,186,124) 
        caja_SegAbreviado = (90,42,156,51)                 
        caja_Direccion = (36,124,252,138)
        caja_destinatario = (6,166,66,174)
        caja_telefono = (0,164,252,174)
        cont +=1

    elif tipo_etiqueta == "PB": 
        caja_TipoEnvio = (54,110,116,124)
        caja_SegInterno =(24,90,162,112) 
        caja_SegAbreviado = (90,4,180,24)                 
        caja_Direccion = (33,183,252,189)
        caja_destinatario = (0,162,166,168)
        caja_telefono = (24,174,252,183)
        cajaNroBulto = (180,191,246,200)
        cont +=1
    
    elif tipo_etiqueta == "R": 
        caja_TipoEnvio = (126, 270, 204, 292)                 
        caja_SegInterno = (144, 252, 312, 264)              
        caja_SegAbreviado = (372, 146, 428, 159)
        caja_Direccion = (60 , 375, 432, 385)
        caja_destinatario = (60, 360, 156, 373)
        caja_telefono = (258, 367, 312, 374)
        cajaNroBulto = (300,406,414,419)
        cont +=1

    else:
        print("Opción no válida.")
        cont = 0 

print(f"\nProcesando PDFs...")

# CAMBIO 2: Iteramos sobre la lista de archivos seleccionados
for indice, ruta_actual in enumerate(archivos_pdf):
    
    # Extraemos solo el nombre del archivo para mostrarlo (sin toda la ruta C:/users/...)
    nombre_archivo = os.path.basename(ruta_actual)
    print(f"[{indice + 1}/{len(archivos_pdf)}] Leyendo: {nombre_archivo} ...")

    # lectura etiquetas (Ahora usamos 'ruta_actual' en vez de 'archivo_pdf')
    with pdfplumber.open(ruta_actual) as pdf:
        for i, pagina in enumerate(pdf.pages):
            
            def extraer_seguro(pagina, caja, x_tol=2):
                if caja is None: return None
                return pagina.crop(caja).extract_text(x_tolerance=x_tol)

            texto_TipoEnvio = extraer_seguro(pagina, caja_TipoEnvio)

            if tipo_etiqueta == "S" or tipo_etiqueta == "P":
                texto_SegInterno = f">8{extraer_seguro(pagina,caja_SegInterno)}"
                texto_SegInterno_OjoHumano = extraer_seguro(pagina,caja_SegInterno)
            else:
                texto_SegInterno = extraer_seguro(pagina, caja_SegInterno)
            texto_SegAbreviado = extraer_seguro(pagina, caja_SegAbreviado)
            texto_Direccion = extraer_seguro(pagina, caja_Direccion)
            texto_destinatario = extraer_seguro(pagina, caja_destinatario)
            texto_rut = extraer_seguro(pagina, caja_rut)
            texto_tlf = extraer_seguro(pagina, caja_telefono)
            texto_SIW = extraer_seguro(pagina,caja_segInternoWalmart)
            textoBulto = extraer_seguro(pagina,cajaNroBulto) if cajaNroBulto else None

            val_tlf = texto_tlf.replace(f'\n','').replace('-','').replace(':','') if texto_tlf else ""
            val_seg_abr = texto_SegAbreviado.replace(f'\n','').replace(".","").strip() if texto_SegAbreviado else "No encontrado"
            val_seg_iw = texto_SIW.replace(f'\n', '').replace(".","").strip() if texto_SIW else "No encontrado"

            datos_etiqueta = {
                'Tipo Envio': re.sub(r'[0-9]','', texto_TipoEnvio ).replace(f'\n', ' ') if texto_TipoEnvio else "No encontrado",
                'Seg Interno': re.sub(r'[a-zA-ZáéíóúÁÉÍÓÚñÑ]','', texto_SegInterno).replace(f'\n','').replace('-','').replace(' ','') if texto_SegInterno else "No encontrado",
                'seg Humano': texto_SegInterno_OjoHumano,
                'Seguimiento': re.sub(r'[a-zA-ZáéíóúÁÉÍÓÚñÑ]','', val_seg_abr).replace(":",""), 
                'Direccion': texto_Direccion.replace(f'\n','').replace('DIRECCION:','').replace('OBSERVACION','') if texto_Direccion else "No encontrado",
                'Destinatario': texto_destinatario.replace('ENVIAR A:','').replace('DESTINATARIO','') if texto_destinatario else "No encontrado",
                'Rut': texto_rut.replace(f'\n','') if texto_rut else "No aplica",
                'Telefono': re.sub(r'[a-zA-ZáéíóúÁÉÍÓÚñÑ]','', val_tlf).replace(".","") if val_tlf else "No encontrado",
                'Seg Walmart':val_seg_iw,
                'Bulto': textoBulto.replace("00"," ").replace(f'\n',': ')if textoBulto else "no encontrado",
                'punto' : punto if punto else "No encontrado",
            }
            
            datos_extraidos.append(datos_etiqueta)

# cruce de datos
if datos_extraidos:
    print(f"\nLeyendo hoja '{NOMBRE_HOJA_EXCEL}' de la Plantilla Maestra...")
    
    df_pdf = pd.DataFrame(datos_extraidos)
    
    # 1. PREPARACIÓN PDF (ORDENAMIENTO PARA 1-A-1)
    df_pdf['Seguimiento'] = df_pdf['Seguimiento'].astype(str).str.strip()
    
    # Ordenamos por Seguimiento y Seg Interno
    df_pdf = df_pdf.sort_values(by=['Seguimiento', 'Seg Interno'])
    
    # Generamos ID de posición (0, 1, 2...) para cada grupo de seguimiento
    df_pdf['posicion_item'] = df_pdf.groupby('Seguimiento').cumcount()
    
    # -----------------------------------------------------
    
    # 2. Leer Excel
    df_master = pd.read_excel(ruta_plantilla, sheet_name=NOMBRE_HOJA_EXCEL, dtype=str)
    df_master.columns = df_master.columns.str.strip() 

    # Verificación
    if COL_TRACKING_EXCEL not in df_master.columns:
        print(f"ERROR: No encontré la columna '{COL_TRACKING_EXCEL}' en el Excel.")
        print(f"Columnas detectadas: {df_master.columns.tolist()}")
        sys.exit()

    # Limpieza Excel
    print("Normalizando datos para el cruce...")
    df_master[COL_TRACKING_EXCEL] = df_master[COL_TRACKING_EXCEL].astype(str).str.strip()
    df_master[COL_TRACKING_EXCEL] = df_master[COL_TRACKING_EXCEL].str.replace(r'\.0$', '', regex=True)

    # --- 3. LÓGICA DE EXPANSIÓN Y ETIQUETADO DE BULTOS ---
    if COL_CANT_BULTOS not in df_master.columns:
        print(f"AVISO: No existe columna '{COL_CANT_BULTOS}'. Asumiendo 1 bulto por producto.")
        df_master[COL_CANT_BULTOS] = 1
    else:
        df_master[COL_CANT_BULTOS] = df_master[COL_CANT_BULTOS].replace('', '1').fillna('1').astype(float).astype(int)

    # Expandimos
    df_master_expandido = df_master.loc[df_master.index.repeat(df_master[COL_CANT_BULTOS])].copy()

    # Generamos el contador (1, 2, 3...)
    df_master_expandido['contador_bulto'] = df_master_expandido.groupby(level=0).cumcount() + 1
    
    # Etiquetamos el producto (1/3)
    mask_multibulto = df_master_expandido[COL_CANT_BULTOS] > 1
    
    df_master_expandido.loc[mask_multibulto, COL_PRODUCTO_EXCEL] = (
        df_master_expandido.loc[mask_multibulto, COL_PRODUCTO_EXCEL].astype(str) + " (" + 
        df_master_expandido.loc[mask_multibulto, 'contador_bulto'].astype(str) + "/" + 
        df_master_expandido.loc[mask_multibulto, COL_CANT_BULTOS].astype(str) + ")"
    )

    # Generamos ID de posición para el Excel
    df_master_expandido['posicion_item'] = df_master_expandido.groupby(COL_TRACKING_EXCEL).cumcount()

    # -----------------------------------------------------

    # 4. Filtrar columnas Excel
    cols_disponibles = [c for c in [COL_TRACKING_EXCEL, COL_OC_EXCEL, COL_SKU_EXCEL, COL_PRODUCTO_EXCEL, colUnidadesExcel, 'posicion_item'] if c in df_master_expandido.columns]
    df_master_limpio = df_master_expandido[cols_disponibles].copy()

    # 5. MERGE
    df_final_merge = pd.merge(
        df_pdf, 
        df_master_limpio, 
        left_on=['Seguimiento', 'posicion_item'], 
        right_on=[COL_TRACKING_EXCEL, 'posicion_item'], 
        how='left'
    )

    if 'posicion_item' in df_final_merge.columns:
        df_final_merge.drop(columns=['posicion_item'], inplace=True)
    
    if COL_TRACKING_EXCEL + "_y" in df_final_merge.columns:
         df_final_merge.drop(columns=[COL_TRACKING_EXCEL + "_y"], inplace=True)

    # 6. Guardar
    if os.path.exists(nombre_final):
        print(f"El archivo '{nombre_final}' ya existe. Agregando datos...")
        df_existente = pd.read_excel(nombre_final, dtype=str)
        df_consolidado = pd.concat([df_existente, df_final_merge], ignore_index=True)
        df_consolidado.to_excel(nombre_final, index=False)
    else:
        print(f"Creando archivo nuevo '{nombre_final}'...")
        df_final_merge.to_excel(nombre_final, index=False)

    print(f"¡Éxito! Total etiquetas procesadas: {len(df_pdf)}")
    input("\nPresiona ENTER para salir...")

else:
    print("No se extrajeron datos de ningún PDF.")
    input("\nPresiona ENTER para salir...")
    