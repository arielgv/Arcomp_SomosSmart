import os
import pdfplumber
import re
import pandas as pd
from colorama import Fore, Style, init
import time
import shutil
from datetime import datetime
from pathlib import Path

init(autoreset=True)

def crear_estructuras_directorios():
    carpeta_salida = "1 - Salida de deposito"
    carpeta_reentrada = "2 - Re-entrada a deposito"
    carpeta_procesados = "3 - Archivos procesados"
    
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)
        print(f"{Fore.GREEN}Carpeta '{carpeta_salida}' creada.")
    
    if not os.path.exists(carpeta_reentrada):
        os.makedirs(carpeta_reentrada)
        print(f"{Fore.GREEN}Carpeta '{carpeta_reentrada}' creada.")
    
    if not os.path.exists(carpeta_procesados):
        os.makedirs(carpeta_procesados)
        print(f"{Fore.GREEN}Carpeta '{carpeta_procesados}' creada.")
    
    return carpeta_salida, carpeta_reentrada, carpeta_procesados

def listar_pdf_directorio(directorio):
    archivos_pdf = []
    if os.path.exists(directorio):
        for archivo in os.listdir(directorio):
            if archivo.lower().endswith('.pdf'):
                archivos_pdf.append(os.path.join(directorio, archivo))
    
    return sorted(archivos_pdf)

def mostrar_archivos_pdf(titulo, archivos):
    print(f"\n{Fore.CYAN}{titulo}")
    print(f"{Fore.CYAN}{'-' * len(titulo)}")
    
    if not archivos:
        print(f"{Fore.YELLOW}  No hay archivos seleccionados")
    else:
        for i, archivo in enumerate(archivos):
            print(f"{Fore.GREEN}  {i+1}. {os.path.basename(archivo)}")

def seleccionar_archivos(todos_archivos, titulo):
    archivos_seleccionados = []
    
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        
        print(f"\n{Fore.CYAN}{'='*80}")
        print(f"{Fore.CYAN}{titulo:^80}")
        print(f"{Fore.CYAN}{'='*80}")
        
        archivos_disponibles = [a for a in todos_archivos if a not in archivos_seleccionados]
        print(f"\n{Fore.WHITE}Archivos PDF disponibles:")
        if archivos_disponibles:
            for i, archivo in enumerate(archivos_disponibles):
                print(f"{Fore.WHITE}  {i+1}. {os.path.basename(archivo)}")
        else:
            print(f"{Fore.YELLOW}  No hay más archivos disponibles")
        
        mostrar_archivos_pdf("Archivos seleccionados", archivos_seleccionados)
        
        print(f"\n{Fore.YELLOW}Opciones:")
        print(f"{Fore.WHITE}  1. Añadir archivo")
        print(f"{Fore.WHITE}  2. Eliminar archivo de la selección")
        print(f"{Fore.WHITE}  3. Continuar con los archivos seleccionados")
        print(f"{Fore.WHITE}  4. Seleccionar todos los archivos disponibles")
        print(f"{Fore.WHITE}  5. Volver atrás / Cancelar")
        
        opcion = input(f"\n{Fore.GREEN}Ingrese su opción (1-5): ")
        
        if opcion == '1' and archivos_disponibles:
            try:
                idx = int(input(f"{Fore.GREEN}Ingrese el número del archivo a añadir (1-{len(archivos_disponibles)}): "))
                if 1 <= idx <= len(archivos_disponibles):
                    archivos_seleccionados.append(archivos_disponibles[idx-1])
                    print(f"{Fore.GREEN}Archivo añadido.")
                    time.sleep(1)
                else:
                    print(f"{Fore.RED}Número fuera de rango.")
                    time.sleep(1)
            except ValueError:
                print(f"{Fore.RED}Por favor, ingrese un número válido.")
                time.sleep(1)
        
        elif opcion == '2' and archivos_seleccionados:
            try:
                idx = int(input(f"{Fore.GREEN}Ingrese el número del archivo a eliminar (1-{len(archivos_seleccionados)}): "))
                if 1 <= idx <= len(archivos_seleccionados):
                    eliminado = archivos_seleccionados.pop(idx-1)
                    print(f"{Fore.GREEN}Archivo '{os.path.basename(eliminado)}' eliminado de la selección.")
                    time.sleep(1)
                else:
                    print(f"{Fore.RED}Número fuera de rango.")
                    time.sleep(1)
            except ValueError:
                print(f"{Fore.RED}Por favor, ingrese un número válido.")
                time.sleep(1)
        
        elif opcion == '3':
            return archivos_seleccionados
        
        elif opcion == '4' and archivos_disponibles:
            archivos_seleccionados.extend(archivos_disponibles)
            print(f"{Fore.GREEN}Todos los archivos disponibles han sido seleccionados.")
            time.sleep(1)
        
        elif opcion == '5':
            return None
        
        else:
            print(f"{Fore.RED}Opción no válida o no hay archivos para esta acción.")
            time.sleep(1)

def extract_data_from_pdf(pdf_path):
    info_documento = {
        'numero': 'Desconocido',
        'fecha': 'Desconocido',
        'origen': 'Desconocido',
        'destino': 'Desconocido',
        'usuario': 'Desconocido'
    }
    
    productos = {}
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            print(f"{Fore.WHITE}Procesando PDF: {os.path.basename(pdf_path)}")
            
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                
                match_numero = re.search(r'Nº:\s*([^\n]+)', texto)
                match_fecha = re.search(r'FECHA:\s*([^\n]+)', texto)
                match_origen = re.search(r'DEPÓSITO\s+ORIGEN:\s*([^\n]+)', texto)
                match_destino = re.search(r'DEPÓSITO\s+DESTINO:\s*([^\n]+)', texto)
                match_usuario = re.search(r'USUARIO:\s*([^\n]+)', texto)
                
                if match_numero:
                    info_documento['numero'] = match_numero.group(1).strip()
                if match_fecha:
                    info_documento['fecha'] = match_fecha.group(1).strip()
                if match_origen:
                    info_documento['origen'] = match_origen.group(1).strip()
                if match_destino:
                    info_documento['destino'] = match_destino.group(1).strip()
                if match_usuario:
                    info_documento['usuario'] = match_usuario.group(1).strip()
                
                tablas = pagina.extract_tables()
                productos_encontrados = False
                
                for tabla in tablas:
                    es_tabla_productos = False
                    indice_encabezado = None
                    
                    for idx, fila in enumerate(tabla):
                        if fila and len(fila) >= 3:
                            col1 = str(fila[0] or '').strip().upper()
                            col2 = str(fila[1] or '').strip().upper()
                            col3 = str(fila[2] or '').strip().upper()
                            
                            if ('COD' in col1 or 'CÓD' in col1) and 'ITEM' in col2 and 'CANTIDAD' in col3:
                                es_tabla_productos = True
                                indice_encabezado = idx
                                break
                    
                    if not es_tabla_productos:
                        filas_producto = []
                        for idx, fila in enumerate(tabla):
                            if fila and len(fila) >= 3 and fila[0] and fila[1] and fila[2]:
                                try:
                                    codigo = str(fila[0]).strip()
                                    if codigo.isdigit():
                                        filas_producto.append(idx)
                                except:
                                    pass
                        
                        if len(filas_producto) >= 3:
                            es_tabla_productos = True
                            indice_encabezado = filas_producto[0] - 1 if filas_producto[0] > 0 else 0
                    
                    if es_tabla_productos:
                        inicio = indice_encabezado + 1 if indice_encabezado is not None else 0
                        
                        for fila in tabla[inicio:]:
                            if fila and len(fila) >= 3 and fila[0] and fila[1]:
                                try:
                                    codigo = str(fila[0]).strip()
                                    nombre = str(fila[1]).strip()
                                    cantidad_str = str(fila[2]).strip() if fila[2] else '0'
                                    
                                    try:
                                        cantidad = int(float(cantidad_str))
                                    except ValueError:
                                        cantidad = 0
                                    
                                    if codigo and nombre and codigo.isdigit():
                                        productos[codigo] = (nombre, cantidad)
                                        productos_encontrados = True
                                except Exception as e:
                                    print(f"{Fore.RED}Error al procesar fila: {e}")
                
                if not productos_encontrados:
                    match_seccion = re.search(r'CÓD\.?\s*ITEM\s+ITEM\s+CANTIDAD(.*?)(?=\n\n|\Z)', texto, re.DOTALL)
                    
                    if match_seccion:
                        seccion_productos = match_seccion.group(1)
                        lineas = seccion_productos.strip().split('\n')
                        
                        for linea in lineas:
                            match_producto = re.match(r'^\s*(\d+)\s+(.*?)\s+(\d+)\s*$', linea)
                            
                            if match_producto:
                                codigo = match_producto.group(1).strip()
                                nombre = match_producto.group(2).strip()
                                cantidad_str = match_producto.group(3).strip()
                                
                                try:
                                    cantidad = int(float(cantidad_str))
                                except ValueError:
                                    cantidad = 0
                                
                                productos[codigo] = (nombre, cantidad)
                                productos_encontrados = True
                    
                    if not productos_encontrados:
                        lineas = texto.split('\n')
                        
                        for linea in lineas:
                            match_producto = re.match(r'^\s*(\d+)\s+(.*?)\s+(\d+)\s*$', linea)
                            
                            if match_producto:
                                codigo = match_producto.group(1).strip()
                                nombre = match_producto.group(2).strip()
                                cantidad_str = match_producto.group(3).strip()
                                
                                try:
                                    cantidad = int(float(cantidad_str))
                                except ValueError:
                                    cantidad = 0
                                
                                productos[codigo] = (nombre, cantidad)
        
        print(f"{Fore.GREEN}Se extrajeron {len(productos)} productos del PDF {os.path.basename(pdf_path)}")
        
    except Exception as e:
        print(f"{Fore.RED}Error al procesar el PDF {pdf_path}: {str(e)}")
    
    return info_documento, productos

def procesar_multiples_pdfs(rutas_pdf):
    todos_los_productos = {}
    
    for ruta_pdf in rutas_pdf:
        try:
            _, productos = extract_data_from_pdf(ruta_pdf)
            
            for codigo, (nombre, cantidad) in productos.items():
                if codigo in todos_los_productos:
                    nombre_actual = todos_los_productos[codigo][0]
                    cantidad_actual = todos_los_productos[codigo][1]
                    todos_los_productos[codigo] = (nombre_actual, cantidad_actual + cantidad)
                else:
                    todos_los_productos[codigo] = (nombre, cantidad)
        except Exception as e:
            print(f"{Fore.RED}Error al procesar {ruta_pdf}: {e}")
    
    return todos_los_productos

def comparar_productos(productos_salida, productos_regreso):
    no_devueltos = {}
    
    for codigo, (nombre, cantidad_salida) in productos_salida.items():
        cantidad_regreso = 0
        
        if codigo in productos_regreso:
            cantidad_regreso = productos_regreso[codigo][1]
        
        diferencia = cantidad_salida - cantidad_regreso
        
        if diferencia > 0:
            no_devueltos[codigo] = (nombre, diferencia)
    
    return no_devueltos

def generar_reporte(productos_salida, productos_regreso, productos_instalados, ruta_salida, formato_salida='excel'):
    todos_codigos = set(list(productos_salida.keys()) + list(productos_regreso.keys()))
    
    datos_reporte = []
    
    for codigo in sorted(todos_codigos):
        nombre = "Desconocido"
        cantidad_salida = 0
        cantidad_regreso = 0
        cantidad_instalada = 0
        
        if codigo in productos_salida:
            nombre = productos_salida[codigo][0]
            cantidad_salida = productos_salida[codigo][1]
        
        if codigo in productos_regreso:
            if nombre == "Desconocido":
                nombre = productos_regreso[codigo][0]
            cantidad_regreso = productos_regreso[codigo][1]
        
        if codigo in productos_instalados:
            cantidad_instalada = productos_instalados[codigo][1]
        
        if cantidad_salida > 0 or cantidad_regreso > 0:
            datos_reporte.append({
                'Código': codigo,
                'Producto': nombre,
                'Cantidad Salida': cantidad_salida,
                'Cantidad Regreso': cantidad_regreso,
                'Cantidad Instalada': cantidad_instalada,
                'Estado': "Instalado" if cantidad_instalada > 0 else "Regresado"
            })
    
    df_reporte = pd.DataFrame(datos_reporte)
    
    if formato_salida.lower() == 'csv':
        df_reporte.to_csv(ruta_salida, index=False, encoding='utf-8-sig')
    else:
        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            df_reporte.to_excel(writer, sheet_name='Todos los Productos', index=False)
            
            df_instalados = df_reporte[df_reporte['Cantidad Instalada'] > 0].copy()
            if not df_instalados.empty:
                df_instalados.to_excel(writer, sheet_name='Productos Instalados', index=False)
            
            df_salida = df_reporte[df_reporte['Cantidad Salida'] > 0].copy()
            if not df_salida.empty:
                df_salida.to_excel(writer, sheet_name='Productos Salida', index=False)
            
            df_regreso = df_reporte[df_reporte['Cantidad Regreso'] > 0].copy()
            if not df_regreso.empty:
                df_regreso.to_excel(writer, sheet_name='Productos Regreso', index=False)
    
    print(f"{Fore.GREEN}Reporte guardado en: {ruta_salida}")
    
    return df_reporte

def mostrar_resumen(productos_salida, productos_regreso, productos_instalados):
    print(f"\n{Fore.CYAN}{'='*80}")
    print(f"{Fore.CYAN}{'RESUMEN DE LA COMPARACIÓN':^80}")
    print(f"{Fore.CYAN}{'='*80}")
    
    print(f"{Fore.WHITE}Total de productos que salieron: {len(productos_salida)}")
    print(f"{Fore.WHITE}Total de productos que regresaron: {len(productos_regreso)}")
    print(f"{Fore.GREEN}Total de productos instalados: {len(productos_instalados)}")
    
    if productos_instalados:
        print(f"\n{Fore.CYAN}{'='*80}")
        print(f"{Fore.CYAN}{'PRODUCTOS INSTALADOS EN CASA DEL CLIENTE':^80}")
        print(f"{Fore.CYAN}{'='*80}")
        
        datos_instalados = []
        for codigo, (nombre, cantidad) in productos_instalados.items():
            datos_instalados.append({
                'Código': codigo,
                'Producto': nombre,
                'Cantidad Instalada': cantidad
            })
        
        df_instalados = pd.DataFrame(datos_instalados)
        df_instalados = df_instalados.sort_values('Código')
        
        print(f"{Fore.WHITE}{df_instalados.to_string(index=False)}")
    
    print(f"{Fore.CYAN}{'='*80}")

def mover_archivos_procesados(archivos, carpeta_procesados):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta_destino = os.path.join(carpeta_procesados, f"Procesado_{timestamp}")
    
    try:
        os.makedirs(carpeta_destino)
        
        for archivo in archivos:
            nombre_archivo = os.path.basename(archivo)
            shutil.copy2(archivo, os.path.join(carpeta_destino, nombre_archivo))
            print(f"{Fore.GREEN}Archivo '{nombre_archivo}' copiado a carpeta de procesados.")
        
        return True
    except Exception as e:
        print(f"{Fore.RED}Error al mover archivos procesados: {str(e)}")
        return False

def main():
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print(f"{Fore.CYAN}{'='*80}")
    print(f"{Fore.CYAN}{'COMPARADOR DE TRANSFERENCIAS DE DEPÓSITO':^80}")
    print(f"{Fore.CYAN}{'='*80}")
    
    carpeta_salida, carpeta_reentrada, carpeta_procesados = crear_estructuras_directorios()
    
    pdfs_carpeta_salida = listar_pdf_directorio(carpeta_salida)
    pdfs_carpeta_reentrada = listar_pdf_directorio(carpeta_reentrada)
    
    if not pdfs_carpeta_salida:
        print(f"{Fore.YELLOW}No se encontraron archivos PDF en la carpeta '{carpeta_salida}'.")
        print(f"{Fore.YELLOW}Por favor, coloque los PDFs de productos que SALEN del depósito en esta carpeta.")
        return
    
    print(f"\n{Fore.GREEN}Se encontraron {len(pdfs_carpeta_salida)} archivos PDF en la carpeta '{carpeta_salida}':")
    for i, pdf in enumerate(pdfs_carpeta_salida):
        print(f"{Fore.WHITE}{i+1}. {os.path.basename(pdf)}")
    
    print(f"\n{Fore.YELLOW}Seleccione los PDFs de productos que SALEN del depósito...")
    time.sleep(1)
    
    pdfs_salida = seleccionar_archivos(pdfs_carpeta_salida, "SELECCIÓN DE PDFs - PRODUCTOS QUE SALEN DEL DEPÓSITO")
    if pdfs_salida is None:
        print(f"{Fore.YELLOW}Operación cancelada por el usuario.")
        return
    
    if not pdfs_salida:
        print(f"{Fore.RED}Debe seleccionar al menos un PDF de salida.")
        return
    
    if not pdfs_carpeta_reentrada:
        print(f"{Fore.YELLOW}No se encontraron archivos PDF en la carpeta '{carpeta_reentrada}'.")
        continuar = input(f"{Fore.YELLOW}¿Desea continuar sin archivos de re-entrada? (s/n): ")
        if continuar.lower() != 's':
            print(f"{Fore.YELLOW}Operación cancelada. Coloque los PDFs de productos que REGRESAN al depósito en la carpeta '{carpeta_reentrada}'.")
            return
        pdfs_regreso = []
    else:
        print(f"\n{Fore.GREEN}Se encontraron {len(pdfs_carpeta_reentrada)} archivos PDF en la carpeta '{carpeta_reentrada}':")
        for i, pdf in enumerate(pdfs_carpeta_reentrada):
            print(f"{Fore.WHITE}{i+1}. {os.path.basename(pdf)}")
        
        print(f"\n{Fore.YELLOW}Seleccione los PDFs de productos que REGRESAN al depósito...")
        time.sleep(1)
        
        pdfs_regreso = seleccionar_archivos(pdfs_carpeta_reentrada, "SELECCIÓN DE PDFs - PRODUCTOS QUE REGRESAN AL DEPÓSITO")
        if pdfs_regreso is None:
            print(f"{Fore.YELLOW}Operación cancelada por el usuario.")
            return
    
    os.system('cls' if os.name == 'nt' else 'clear')
    
    print(f"{Fore.CYAN}{'='*80}")
    print(f"{Fore.CYAN}{'ARCHIVOS SELECCIONADOS PARA COMPARACIÓN':^80}")
    print(f"{Fore.CYAN}{'='*80}")
    
    mostrar_archivos_pdf("PDFs DE PRODUCTOS QUE SALEN DEL DEPÓSITO", pdfs_salida)
    mostrar_archivos_pdf("PDFs DE PRODUCTOS QUE REGRESAN AL DEPÓSITO", pdfs_regreso)
    
    print(f"\n{Fore.YELLOW}¿Desea proceder con la comparación de estos archivos?")
    print(f"{Fore.WHITE}1. Sí, proceder con la comparación")
    print(f"{Fore.WHITE}2. No, volver a seleccionar archivos")
    print(f"{Fore.WHITE}3. Cancelar y salir")
    
    opcion = input(f"\n{Fore.GREEN}Ingrese su opción (1-3): ")
    
    if opcion == '1':
        print(f"\n{Fore.YELLOW}Seleccione el formato del reporte:")
        print(f"{Fore.WHITE}1. Excel (recomendado)")
        print(f"{Fore.WHITE}2. CSV")
        
        formato = input(f"\n{Fore.GREEN}Ingrese su opción (1-2): ")
        formato_salida = 'excel' if formato != '2' else 'csv'
        
        extension = 'xlsx' if formato_salida == 'excel' else 'csv'
        nombre_defecto = f"reporte_productos_instalados.{extension}"
        
        nombre_archivo = input(f"\n{Fore.GREEN}Ingrese nombre del archivo de salida [{nombre_defecto}]: ")
        if not nombre_archivo:
            nombre_archivo = nombre_defecto
        elif not nombre_archivo.endswith(f".{extension}"):
            nombre_archivo = f"{nombre_archivo}.{extension}"
        
        print(f"\n{Fore.CYAN}{'='*80}")
        print(f"{Fore.CYAN}{'PROCESANDO ARCHIVOS':^80}")
        print(f"{Fore.CYAN}{'='*80}")
        
        print(f"\n{Fore.MAGENTA}PROCESANDO PRODUCTOS QUE SALEN DEL DEPÓSITO:")
        productos_salida = procesar_multiples_pdfs(pdfs_salida)
        print(f"{Fore.GREEN}Total de productos que salen: {len(productos_salida)}")
        
        productos_regreso = {}
        if pdfs_regreso:
            print(f"\n{Fore.MAGENTA}PROCESANDO PRODUCTOS QUE REGRESAN AL DEPÓSITO:")
            productos_regreso = procesar_multiples_pdfs(pdfs_regreso)
            print(f"{Fore.GREEN}Total de productos que regresan: {len(productos_regreso)}")
        else:
            print(f"\n{Fore.YELLOW}No hay archivos de regreso para procesar.")
        
        print(f"\n{Fore.MAGENTA}COMPARANDO PRODUCTOS:")
        productos_instalados = comparar_productos(productos_salida, productos_regreso)
        print(f"{Fore.GREEN}Total de productos instalados: {len(productos_instalados)}")
        
        mostrar_resumen(productos_salida, productos_regreso, productos_instalados)
        
        print(f"\n{Fore.MAGENTA}GENERANDO REPORTE DETALLADO:")
        generar_reporte(productos_salida, productos_regreso, productos_instalados, nombre_archivo, formato_salida)
        
        archivos_procesados = pdfs_salida + pdfs_regreso
        if archivos_procesados:
            mover = input(f"\n{Fore.YELLOW}¿Desea mover los archivos procesados a la carpeta '{carpeta_procesados}'? (s/n): ")
            if mover.lower() == 's':
                exito = mover_archivos_procesados(archivos_procesados, carpeta_procesados)
                if exito:
                    print(f"{Fore.GREEN}Archivos movidos correctamente.")
        
        print(f"\n{Fore.GREEN}Proceso completado con éxito.")
    
    elif opcion == '2':
        main()
    
    else:
        print(f"{Fore.YELLOW}Operación cancelada por el usuario.")

if __name__ == "__main__":
    main()

# INSTRUCCIONES DE USO:
# 1. Coloque los PDFs de SALIDA en la carpeta "1 - Salida de deposito"
# 2. Coloque los PDFs de REGRESO en la carpeta "2 - Re-entrada a deposito"
# 3. Ejecute el programa: python comparador_depositos.py
# 4. Seleccione los archivos a procesar usando los menús interactivos
# 5. Elija el formato y nombre del reporte
# 6. Revise el reporte generado para ver qué productos quedaron instalados
# 7. Opcionalmente, los archivos procesados se pueden mover a "3 - Archivos procesados"