# %%
import os
import fnmatch
import pandas as pd
import numpy as np
from datetime import datetime
import shutil
import sys
 
def generar_excel_by_df(df, nombre_base):
    """
    Genera un archivo Excel a partir de un DataFrame, añadiendo la fecha y hora actual al nombre del archivo.

    :param df: DataFrame a exportar.
    :param nombre_base: Nombre base del archivo (sin extensión).
    :return: Ruta completa del archivo generado.
    """
    try:
        # Obtener la fecha y hora actuales en formato 'YYYYMMDD_HHMMSS'
        fecha_hora = datetime.now().strftime("%Y-%m-%dT%H-%M-%S")
        
        # Construir el nombre del archivo
        nombre_archivo = f"{nombre_base}_{fecha_hora}.xlsx"
        
        # Exportar el DataFrame a Excel
        df.to_excel(nombre_archivo, index=False, engine="openpyxl")
        
        print(f"Archivo Excel generado exitosamente: {nombre_archivo}")
        return nombre_archivo
    except Exception as e:
        print(f"Error al generar el archivo Excel: {e}")
        return None

def fusionar_archivos_excel(lista_archivos, hoja=None, nombre_salida="archivo_fusionado.xlsx"):
    """
    Fusiona múltiples archivos Excel en un solo archivo, permitiendo especificar una hoja de cada archivo.
    
    :param lista_archivos: Lista de rutas de los archivos Excel a fusionar.
    :param hoja: Nombre de la hoja a leer de cada archivo. Si es None, se usará la primera hoja.
    :param nombre_salida: Nombre del archivo de salida fusionado.
    :return: Nombre del archivo fusionado, o una cadena vacía si no se pudieron procesar archivos.
    """
    dataframes = []

    for archivo in lista_archivos:
        # Verificar si el archivo existe y tiene la extensión correcta
        if os.path.isfile(archivo) and archivo.endswith(".xlsx"):
            try:
                # Leer el archivo con la hoja especificada
                df = None
                if hoja is None:
                    df = pd.read_excel(archivo, engine="openpyxl")
                else:
                    df = pd.read_excel(archivo, engine="openpyxl", sheet_name=hoja)
                

                if df is not None:
                   dataframes.append(df)
            except ValueError:
                print(f"La hoja '{hoja}' no existe en el archivo {archivo}.")
            except Exception as e:
                print(f"Error al leer el archivo {archivo}: {e}")
        else:
            print(f"Archivo no válido o no encontrado: {archivo}")

    if dataframes:
        # Concatenar los DataFrames si hay al menos uno válido
        df_fusionado = pd.concat(dataframes, ignore_index=True)
        # Especificar el motor openpyxl al guardar
        df_fusionado.to_excel(nombre_salida, index=False, engine="openpyxl")
        return nombre_salida
    else:
        print("No se encontraron archivos válidos para fusionar.")
        return ""

def listar_archivos_excel_por_cadena(directorio: str, cadena: str):
    archivos_excel = []
    patron = f"*{cadena}*.xlsx"
    
    for archivo in os.listdir(directorio):
        if fnmatch.fnmatch(archivo, patron):
            archivos_excel.append(archivo)

    return archivos_excel

def borrar_archivos(lista_archivos):
    for archivo in lista_archivos:
        if os.path.isfile(archivo):
            try:
                os.remove(archivo)
                print(f"Archivo eliminado: {archivo}")
            except Exception as e:
                print(f"Error al eliminar el archivo {archivo}: {e}")
        else:
            print(f"Archivo no encontrado: {archivo}")

def crear_dataframe_desde_archivo(archivo: str, columnas: list, hoja: str = None):

    try:
        # Leer el archivo completo usando pandas con la hoja especificada
        df = None
        if hoja is None:
             df = pd.read_excel(archivo, engine='openpyxl')
        else:
             df = pd.read_excel(archivo, engine='openpyxl', sheet_name=hoja)

        # Filtrar solo las columnas deseadas
        df_filtrado = df[columnas]

        return df_filtrado
    except FileNotFoundError:
        print(f"El archivo {archivo} no fue encontrado.")
    except KeyError as e:
        print(f"Una o más columnas no se encuentran en el archivo: {e}")
    except ValueError:
        print(f"La hoja '{hoja}' no existe en el archivo {archivo}.")
    except Exception as e:
        print(f"Se produjo un error al procesar el archivo: {e}")

def eliminar_columnas_df(dataframe, columnas):
    """
    Elimina una lista de columnas de un DataFrame.

    :param dataframe: DataFrame de pandas del que se desean eliminar las columnas.
    :param columnas: Lista de nombres de columnas a eliminar.
    :return: DataFrame con las columnas eliminadas.
    """
    try:
        # Verificar cuáles columnas existen en el DataFrame
        columnas_existentes = [col for col in columnas if col in dataframe.columns]
        columnas_no_existentes = [col for col in columnas if col not in dataframe.columns]

        if columnas_existentes:
            dataframe = dataframe.drop(columns=columnas_existentes)
            print(f"Las columnas eliminadas exitosamente: {columnas_existentes}")
        
        if columnas_no_existentes:
            print(f"Las siguientes columnas no existen en el DataFrame: {columnas_no_existentes}")
        
        return dataframe
    except Exception as e:
        print(f"Se produjo un error al intentar eliminar las columnas: {e}")
        return dataframe
    
def filtrar_columnas_df(dataframe, columnas):
    """
    Devuelve un DataFrame que contiene solo las columnas especificadas.

    :param dataframe: DataFrame de pandas del que se desea conservar las columnas.
    :param columnas: Lista de nombres de columnas a conservar.
    :return: DataFrame con solo las columnas especificadas.
    """
    try:
        # Verificar cuáles columnas existen en el DataFrame
        columnas_existentes = [col for col in columnas if col in dataframe.columns]
        columnas_no_existentes = [col for col in columnas if col not in dataframe.columns]

        if columnas_no_existentes:
            print(f"Las siguientes columnas no existen en el DataFrame: {columnas_no_existentes}")

        # Seleccionar solo las columnas existentes
        dataframe = dataframe[columnas_existentes]

        return dataframe
    except Exception as e:
        print(f"Se produjo un error al intentar mantener las columnas: {e}")
        return dataframe
    
def reemplazar_ceros_con_nan(dataframe, columnas):
    """
    Reemplaza ceros en las columnas especificadas de un DataFrame con NaN.

    :param dataframe: DataFrame en el que se procesarán las columnas.
    :param columnas: Lista de nombres de columnas donde se reemplazarán los ceros por NaN.
    :return: DataFrame con los ceros reemplazados por NaN en las columnas especificadas.
    """
    try:
        # Validar que las columnas existan en el DataFrame
        columnas_validas = [col for col in columnas if col in dataframe.columns]

        # Reemplazar ceros por NaN en las columnas válidas
        dataframe[columnas_validas] = dataframe[columnas_validas].replace(0, np.nan)

        print(f"Ceros reemplazados por NaN en las columnas: {columnas_validas}")
        return dataframe
    except Exception as e:
        print(f"Error al reemplazar ceros por NaN: {e}")
        return dataframe
    
def crear_carpeta(base_nombre_carpeta, ruta_base="."):
    """
    Crea una carpeta con un nombre que incluye la fecha y hora actual al final.
    
    :param base_nombre_carpeta: Nombre base para la carpeta.
    :param ruta_base: Ruta donde se creará la carpeta. Por defecto, en el directorio actual.
    :return: Ruta completa de la carpeta creada.
    """
    try:
        # Obtener la fecha y hora actuales en formato 'YYYY-MM-DDTHH-MM-SS'
        fecha_hora = datetime.now().strftime("%Y-%m-%dT%H-%M-%S")
        
        # Construir el nombre completo de la carpeta
        nombre_completo_carpeta = f"{base_nombre_carpeta}_{fecha_hora}"
        ruta_completa_carpeta = os.path.join(ruta_base, nombre_completo_carpeta)
        
        # Crear la carpeta
        os.makedirs(ruta_completa_carpeta, exist_ok=True)
        print(f"Carpeta creada: {ruta_completa_carpeta}")
        return ruta_completa_carpeta
    except Exception as e:
        print(f"Error al crear la carpeta: {e}")
        return None

def mover_archivos_a_carpeta(lista_archivos, carpeta_destino):
    """
    Mueve una lista de archivos a una carpeta destino.
    
    :param lista_archivos: Lista con las rutas de los archivos a mover.
    :param carpeta_destino: Ruta de la carpeta destino.
    :return: None
    """
    try:
        # Crear la carpeta destino si no existe
        carpeta_destino = crear_carpeta(carpeta_destino)
        
        for archivo in lista_archivos:
            if os.path.isfile(archivo):  # Verificar que el archivo existe
                destino = os.path.join(carpeta_destino, os.path.basename(archivo))  # Ruta destino
                shutil.move(archivo, destino)  # Mover el archivo
                print(f"Archivo movido: {archivo} -> {destino}")
            else:
                print(f"El archivo no existe: {archivo}")
    except Exception as e:
        print(f"Error al mover archivos: {e}")

def validar_archivos(lista_archivos):
    """
    Valida que una lista de archivos no contenga valores vacíos.
    
    :param lista_archivos: Lista de rutas de archivos.
    :return: True si todos los archivos son válidos, False si hay archivos faltantes.
    """
    # Filtrar archivos vacíos
    archivos_faltantes = [archivo for archivo in lista_archivos if archivo == ""]
    
    if archivos_faltantes:
        print("Error: Faltan archivos necesarios para el proceso del reporte.")
        sys.exit(1)  # Rompe la ejecución con un código de error
        return False
    
    print("Todos los archivos son válidos.")
    return True



def crearDataframeExistenciaFinal(dfExistencias):
    #Crea copia de DF de existencias para agrupación de existencias globales
    dfExistenciasGrp = dfExistencias.copy(deep=True)
    dfExistenciasGrp = filtrar_columnas_df(dfExistenciasGrp, ["ProdConcat", "Existencia"])
    dfExistenciasGrpGlobal = dfExistenciasGrp.groupby('ProdConcat').agg({
            'Existencia': 'sum'
        }).reset_index()


    # Realiza un PIVOT para los almacenes y calculos para agrupar los productos por ProdConcat
    # Pivotar los datos para que cada Almacen tenga su propia columna
    pivot = dfExistencias.pivot_table(
        index="ProdConcat", 
        columns="Almacen", 
        values="Existencia", 
        aggfunc="sum"  # Sumamos si hay duplicados
    )

    # Renombrar las columnas del pivote agregando un prefijo
    pivot = pivot.rename(columns=lambda col: f"Existencias en {col}")

    # Resetear el índice del pivot
    pivot.reset_index(inplace=True)

    # Agrupar para obtener el primer valor de 'Nombre' y 'TipoProducto'
    metadata = dfExistencias.groupby('ProdConcat').agg({
        'Nombre': 'first',
        'TipoProducto': 'first',
        'Modelo': 'first',
        'Marca': 'first',
        "Publico En General": 'mean'
    }).reset_index()
    
    # Combinar el pivot con los datos adicionales
    dfExistenciasAlmacenes = pd.merge(metadata, pivot, on="ProdConcat", how="left")

    # #USAR dfExistenciasAlmacenes y dfExistenciasGrpGlobal para obtener datos limpios
    dfExistenciasFinal =  pd.merge(dfExistenciasAlmacenes, dfExistenciasGrpGlobal, on="ProdConcat", how="inner")
    return dfExistenciasFinal


def creaReporteExistenciaConcentrada(dfExistenciasFinal):
    dfConcentradoExistencias = dfExistenciasFinal.copy(deep=True)
    dfConcentradoExistencias = eliminar_columnas_df(dfConcentradoExistencias, ["ProdConcat", "TipoProducto"])

    # Reemplazar NaN con un valor predeterminado antes de agrupar
    dfConcentradoExistencias[["Marca", "Modelo", "Nombre"]] = dfConcentradoExistencias[["Marca", "Modelo", "Nombre"]].fillna("Desconocido")

    dfConcentradoExistencias = dfConcentradoExistencias.groupby(["Marca", "Modelo", "Nombre"]).agg({
        'Existencias en Central Cell 20 de noviembre': 'sum',
        'Existencias en Central Cell Almacén general': 'sum',
        'Existencias en Central Cell Abastos': 'sum',
        'Existencias en Central Cell Fortín': 'sum',
        'Existencias en Central Cell Labotienda': 'sum',
        'Existencias en Central Cell Nuño del Mercado': 'sum',
        'Existencias en Central Cell Plaza Bella': 'sum',
        'Existencias en Central Cell Plaza Bonn': 'sum',
        'Existencias en Central Cell Reforma': 'sum',
        'Existencias en Central Cell Revistería': 'sum',
        'Existencias en Central Cell Violetas': 'sum',
        'Existencia': 'sum'
    }).reset_index()
        # Renombrar columnas del DataFrame
    dfPiezasConsumidas.rename(columns={
        "Nombre": "Clasificación"
    }, inplace=True)


    generar_excel_by_df(dfConcentradoExistencias, "BI-CONCENTRADO-EXISTENCIAS-BY-MODELO-MARCA")



def creaDataFrameExistenciasComprasFinal(dfExistenciasFinal, dfCompras):
    dfComprasAdjusted = dfCompras.copy(deep=True)
    dfComprasAdjusted = eliminar_columnas_df(dfComprasAdjusted, ["Almacen"])
    #Agrupa los datos de compras para limpiar la muestra
    # Paso 1: Transformar la columna Fecha para que solo contenga la fecha sin la hora
    dfComprasAdjusted["Fecha"] = pd.to_datetime(dfComprasAdjusted["Fecha"]).dt.date
    # Paso 2: Filtrar los registros con la fecha más reciente por producto
    # Ordenar el DataFrame por Producto y Fecha en orden descendente
    dfComprasAdjusted = dfComprasAdjusted.sort_values(by=["Producto", "Fecha"], ascending=[True, False])
    # Mantener solo el registro más reciente para cada Producto
    dfFiltradoCompras = dfComprasAdjusted.drop_duplicates(subset="Producto", keep="first")
    dfFiltradoCompras.rename(columns={'Producto': 'ProdConcat'}, inplace=True)
    #CASI LO FINAL
    dfExistenciasComprasFinal =  pd.merge(dfExistenciasFinal, dfFiltradoCompras, on="ProdConcat", how="left")
    dfExistenciasComprasFinal.rename(columns={'Existencia': 'Existencia Global', 'Fecha':'Última Fecha Compra', 'Costo':'Precio Compra', 'Cantidad':'Cantidad Comprada Ultimo Mov'}, inplace=True)
    # Verificar si 'Publico En General' es un DataFrame y corregir
    if isinstance(dfExistenciasComprasFinal["Publico En General"], pd.DataFrame):
        # Si es un DataFrame, tomar la primera columna válida (ajustar según necesidad)
        dfExistenciasComprasFinal["Publico En General"] = dfExistenciasComprasFinal["Publico En General"].iloc[:, 0]
    # Lista actual de columnas en el DataFrame
    columnas_actuales = dfExistenciasComprasFinal.columns.tolist()
    # Crear un nuevo orden, asegurando que no se dupliquen columnas
    columnas_nuevo_orden = []
    for col in columnas_actuales:
        if col != "Publico En General" and col != "Cantidad Comprada Ultimo Mov":  # Evitar duplicar la columna en su posición original
            columnas_nuevo_orden.append(col)
        if col == "Precio Compra":  # Insertar "Publico En General" después de "Precio Compra"
            columnas_nuevo_orden.append("Publico En General")
        if col == "Última Fecha Compra":
            columnas_nuevo_orden.append("Cantidad Comprada Ultimo Mov")
    # Reorganizar las columnas del DataFrame
    dfExistenciasComprasFinal = dfExistenciasComprasFinal[columnas_nuevo_orden]
    dfExistenciasComprasFinal['Precio Compra'] = pd.to_numeric(dfExistenciasComprasFinal['Precio Compra'], errors='coerce')
    IVA = .16
    CIEN = 100
    # Verificar que las columnas sean numéricas y manejar NaN
    if pd.api.types.is_numeric_dtype(dfExistenciasComprasFinal["Publico En General"]) and pd.api.types.is_numeric_dtype(dfExistenciasComprasFinal["Precio Compra"]):
        # Rellenar NaN con 0 para evitar errores durante la resta
        dfExistenciasComprasFinal["Publico En General"] = dfExistenciasComprasFinal["Publico En General"].fillna(0)
        dfExistenciasComprasFinal["Precio Compra"] = dfExistenciasComprasFinal["Precio Compra"].fillna(0)
        # Crear la columna 'Utilidad Bruta'
        dfExistenciasComprasFinal['Costo'] = dfExistenciasComprasFinal["Precio Compra"]+(dfExistenciasComprasFinal["Precio Compra"]*IVA)
        # Crear la columna 'Utilidad Bruta'
        dfExistenciasComprasFinal['Utilidad'] = ((dfExistenciasComprasFinal["Publico En General"]-dfExistenciasComprasFinal['Costo'])/dfExistenciasComprasFinal['Publico En General']) * CIEN
        print("Columna 'Utilidad' y 'Costo' creada exitosamente.")
    else:
        print("Error: Las columnas 'Publico En General' y 'Precio Compra' deben ser numéricas.")
    # Lista actual de columnas en el DataFrame
    columnas_actuales = dfExistenciasComprasFinal.columns.tolist()
    # Crear un nuevo orden, asegurando que no se dupliquen columnas
    columnas_nuevo_orden = []
    for col in columnas_actuales:
        if col != "Costo":  # Evitar duplicar la columna en su posición original
            columnas_nuevo_orden.append(col)
        if col == "Precio Compra": 
            columnas_nuevo_orden.append("Costo")
    # Reorganizar las columnas del DataFrame
    dfExistenciasComprasFinal = dfExistenciasComprasFinal[columnas_nuevo_orden]
    return dfExistenciasComprasFinal


def creaDataFrameVentasFinal(dfVentas, dfPiezasConsumidas):
    # Combinar dfVentas y dfPiezasConsumidas
    dfVentas = pd.concat([dfVentas, dfPiezasConsumidas], ignore_index=True)
    dfVentas["Cantidad"] = dfVentas["Cantidad"].fillna(0)
    # Agrupar por 'Almacen' y 'ProdConcat', sumando las cantidades
    dfVentas = dfVentas.groupby(["Almacen", "ProdConcat"], as_index=False).agg({"Cantidad": "sum"})


    dfVentasTotales = dfVentas.copy(deep=True)

    # Paso 1: Agrupar por 'ProdConcat' y 'Almacen' para sumar las cantidades
    dfVentasAgrupado = dfVentas.groupby(["ProdConcat", "Almacen"]).agg({"Cantidad": "sum"}).reset_index()

    # Paso 2: Pivoteo por 'ProdConcat' y 'Almacen'
    pivotVentas = dfVentasAgrupado.pivot_table(
        index="ProdConcat", 
        columns="Almacen", 
        values="Cantidad", 
        aggfunc="sum"  # Ya no debería ser necesario, pero lo dejamos por seguridad
    )

    # Paso 3: Renombrar las columnas del pivote agregando un prefijo
    pivotVentas = pivotVentas.rename(columns=lambda col: f"Ventas de {col}")

    # Paso 4: Convertir el índice del pivote a una columna para un DataFrame plano
    dfVentasFinal = pivotVentas.reset_index()


    dfVentasTotales = eliminar_columnas_df(dfVentasTotales, ["Almacen"])

    dfVentasTotalesAgp = dfVentasTotales.groupby(["ProdConcat"]).agg({"Cantidad": "sum"}).reset_index()

    # Merge de dfVentasFinal y dfVentasTotalesAgp por 'ProdConcat'
    dfVentasFinalMerged = pd.merge(
        dfVentasFinal, 
        dfVentasTotalesAgp, 
        on="ProdConcat",  # Clave común
        how="inner"       # Tipo de merge (inner join)
    )

    dfVentasFinalMerged.rename(columns={ 'Cantidad':'Ventas Totales'}, inplace=True)
    return dfVentasFinalMerged

def creaReporteExistenciasComprasVentasCC(dfExistenciasComprasFinal, dfVentasFinalMerged):
    # Merge de dfExistenciasComprasFinal y dfVentasFinalMerged por 'ProdConcat'
    dfResultadoFinalBIData = pd.merge(
        dfExistenciasComprasFinal,
        dfVentasFinalMerged,
        on="ProdConcat",  # Clave común
        how="left"       # Tipo de merge (inner join)
    )

    generar_excel_by_df(dfResultadoFinalBIData, "BI-EXISTENCIAS-COMPRAS-VENTAS-CC")


def convertir_columna_uppercase(df, columna="ProdConcat"):
    """
    Convierte todos los valores de una columna de un DataFrame a mayúsculas.

    :param df: DataFrame que contiene la columna a transformar.
    :param columna: Nombre de la columna que se desea convertir a mayúsculas. Por defecto, 'ProdConcat'.
    :return: DataFrame con la columna transformada.
    """
    try:
        if columna not in df.columns:
            raise ValueError(f"La columna '{columna}' no existe en el DataFrame.")

        # Convertir la columna a mayúsculas
        df[columna] = df[columna].str.upper()

        return df
    except Exception as e:
        print(f"Error al convertir la columna '{columna}' a mayúsculas: {e}")
        return df



# %%
print("####################################################")
print("Iniciando análisis de datos...")

directorio = "./"  # Directorio actual

#Archivos de existencias de productos
archivosExistencias = "Existencia"
archivosExitenciasMap = listar_archivos_excel_por_cadena(directorio, archivosExistencias)

#Archivos de datos de compras
archivosCompras = "Excel_Movimientos"
archivosComprasMap = listar_archivos_excel_por_cadena(directorio, archivosCompras)

#Archivos de datos de ventas
archivosVentas = "Analisis de Ventas por Tickets"
archivosVentasMap = listar_archivos_excel_por_cadena(directorio, archivosVentas)

#Archivos de datos de ventas
archivosPiezasConsumidas = "Excel_Reparaciones_Refacciones_Consumidas"
archivosPiezasConsumidasMap = listar_archivos_excel_por_cadena(directorio, archivosPiezasConsumidas)

archivosTrabajados = archivosExitenciasMap+ archivosComprasMap+ archivosVentasMap + archivosPiezasConsumidasMap

#Fusión de archvios clasificados por reportes
existenciasCC = fusionar_archivos_excel(lista_archivos=archivosExitenciasMap, nombre_salida="ExistenciasCC.xlsx")
comprasCC = fusionar_archivos_excel(lista_archivos=archivosComprasMap, hoja="Detalle de movimientos", nombre_salida="ComprasCC.xlsx")
ventasCC = fusionar_archivos_excel(lista_archivos=archivosVentasMap, nombre_salida="VentasCC.xlsx")
piezasConsumidasCC = fusionar_archivos_excel(lista_archivos=archivosPiezasConsumidasMap, nombre_salida="PiezasConsumidasCC.xlsx")


validar_archivos([existenciasCC, comprasCC, ventasCC, piezasConsumidasCC]) 

#Qué columnas ocupamos de cada paquete de archivos
columnasExistencias = ["Almacen", "ProdConcat", "Existencia", "Nombre", "TipoProducto", "Marca", "Modelo", "Publico En General"]
columnasCompras = ["Almacen", "Fecha", "Producto", "Costo", "Cantidad"]
columnasVentas = ["Almacen", "ProdConcat", "Cantidad"]
columnasPiezasConsumidas = ["Almacén Salida Reparación", "Producto", "Cantidad"]

#Generación de dataframe de existencias y ajustes por valores numéricos
dfExistencias = crear_dataframe_desde_archivo(existenciasCC, columnasExistencias)
dfExistencias = reemplazar_ceros_con_nan(dfExistencias, ["Existencia"])

#Generación de dataframe de compras
dfCompras = crear_dataframe_desde_archivo(comprasCC, columnasCompras)

#Generación de dataframe de ventas
dfVentas = crear_dataframe_desde_archivo(ventasCC, columnasVentas)

#Generación de dataframe de piezas consumidas 
dfPiezasConsumidas = crear_dataframe_desde_archivo(piezasConsumidasCC, columnasPiezasConsumidas)
# Renombrar columnas del DataFrame de piezas consumidas
dfPiezasConsumidas.rename(columns={
    "Almacén Salida Reparación": "Almacen",
    "Producto": "ProdConcat"
}, inplace=True)

# Eliminar filas duplicadas considerando todas las columnas
dfPiezasConsumidas.drop_duplicates(inplace=True)


dfVentas = convertir_columna_uppercase(dfVentas, "ProdConcat")
dfExistencias = convertir_columna_uppercase(dfExistencias, "ProdConcat")
dfPiezasConsumidas = convertir_columna_uppercase(dfPiezasConsumidas, "ProdConcat")
dfCompras = convertir_columna_uppercase(dfCompras, "Producto")

# Crea un un Dataframe que contenga los valores de existencias 
# por almacen en forma de columnas y en otra la existencia global
dfExistenciasFinal = crearDataframeExistenciaFinal(dfExistencias)

# Genera el primer reporte que dará como resultado el acumulado 
# de existencias de Productos dividido por MARCA-MODELO-CATEGORÍA 
# por sucursal y globalmente
creaReporteExistenciaConcentrada(dfExistenciasFinal)

# Crea un DataFrame que contiene las existencias de productos por almacén 
# (en columnas) y una columna con la existencia global total
# a su vez, quedan agrupada la ultima compra hecha, junto con la fecha para cada uno de los productos
dfExistenciasComprasFinal = creaDataFrameExistenciasComprasFinal(dfExistenciasFinal, dfCompras)
generar_excel_by_df(dfExistenciasComprasFinal, "BI-EXISTENCIA-CC")

# Fusiona los DataFrames de ventas y piezas consumidas, consolidando las 
# cantidades de productos vendidos por almacén y obteniendo un DataFrame 
# con el detalle completo de ventas
dfVentasFinalMerged = creaDataFrameVentasFinal(dfVentas, dfPiezasConsumidas)
generar_excel_by_df(dfVentasFinalMerged, "BI-VENTAS-CC")


# Crea un reporte final que integra existencias, compras y ventas, 
# mostrando el desglose de productos por almacén, acumulados y ventas 
# globales, facilitando el análisis comparativo
creaReporteExistenciasComprasVentasCC(dfExistenciasComprasFinal, dfVentasFinalMerged)


# Reagrupar archivos y nuevos 
archivosCompilados = listar_archivos_excel_por_cadena(directorio, "CC")
archivosBI = listar_archivos_excel_por_cadena(directorio, "BI-")

archivosTrabajados = archivosTrabajados + archivosCompilados + archivosBI
print(archivosTrabajados)
mover_archivos_a_carpeta(archivosTrabajados, "BI-DATA-CC")


# Cerrar la ventana de la terminal
os.system("TASKKILL /F /IM cmd.exe")

