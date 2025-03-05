# Comparador de Transferencias de Deposito

Este programa permite comparar PDFs de transferencias de productos entre depositos para identificar que elementos quedaron instalados, generando reportes detallados en Excel o CSV.

## Requisitos previos

### Instalación de Python (solo primera vez)

1. **Descargar Python**:
   - Visite [python.org](https://www.python.org/downloads/)
   - Descargue Python 3.8 o superior
   - **Importante**: Al instalar, marque "Add Python to PATH"

2. **Verificar instalación**:
   - Abra una terminal o línea de comandos
   - Escriba: `python --version`
   - Debería mostrar la versión instalada (ej. "Python 3.10.4")

## Configuración inicial (solo primera vez)

1. **Descargar el programa**:
   - Descargue y extraiga el ZIP en una ubicación de fácil acceso

2. **CREAR CARPETAS**:
   - El programa esta configurado para crearlas en su primer uso, pero es muy recomendable crear previamente estas tres carpetas en la ubicacion que se encuentra main.py:
      - carpeta 1: "1 - Salida de deposito", sin comillas doble
         (ahí se colocarán los archivos pdf relacionados)
      - carpeta 2: "2 - Re-entrada a deposito", sin comillas doble
         (ahí se colocarán los archivos pdf relacionados)
      - carpeta 3: "3 - Archivos procesados", sin comillas doble.
         (carpeta de uso interno del programa)


2. **Preparar el entorno**:
   - Abra una terminal en la carpeta del programa
   - Cree un entorno virtual:
     ```
     python -m venv venv
     ```
   - Active el entorno virtual:
     - Windows: `venv\Scripts\activate`
     - Mac/Linux: `source venv/bin/activate`
   - Instale las dependencias:
     ```
     pip install -r requirements.txt
     ```

## Uso diario del programa

### 1. Preparación de archivos

- Coloque los PDFs de salida en la carpeta: `1 - Salida de deposito`
- Coloque los PDFs de regreso en la carpeta: `2 - Re-entrada a deposito`

### 2. Ejecutar el programa

- Abra una terminal en la carpeta del programa
- Active el entorno virtual:
  - Windows: `venv\Scripts\activate`
  - Mac/Linux: `source venv/bin/activate`
- Ejecute el programa:
  ```
  python main.py
  ```

### 3. Uso del programa

1. **Seleccione los archivos** siguiendo las instrucciones en pantalla
2. **Elija formato** del reporte (Excel recomendado)
3. **Revise el informe** generado para ver qué productos quedaron instalados
4. **Opcionalmente**, permita que el programa mueva copias de los archivos procesados a la carpeta de históricos

## Estructura de carpetas

El programa crea y utiliza tres carpetas principales:

- **1 - Salida de deposito**: Contiene PDFs de productos que salen
- **2 - Re-entrada a deposito**: Contiene PDFs de productos que regresan
- **3 - Archivos procesados**: Almacena copias de PDFs ya procesados

## Solución de problemas comunes

- **Error al iniciar**: Verifique que activó el entorno virtual
- **Error "No module found"**: Ejecute nuevamente `pip install -r requirements.txt`
- **No encuentra archivos**: Confirme que los PDFs están en las carpetas correctas
- **Error al procesar PDFs**: Confirme que los PDFs tienen el mismo formato que los ejemplos

Para cualquier consulta adicional, se puede comunicar a avillafane@cloudacio.com
