import pandas as pd
import streamlit as st
from collections import defaultdict
from datetime import datetime, timedelta
import difflib
import re
import numpy as np
import io


##############
# Armado CSV #
##############

#Estos son codigos de motivos de ausencia que detallan para que si
#se declara en la planilla que una persona hizo horas extras en un 
#día bajo alguno de estos codigos no se descuenten las mismas.
codigos_ausencias_no_descontables = (set([1,4,5,6,7,8,9,10,11,12,13,
                                          14,15,16,17,18,21,22,25,26,
                                          30,31,32,33,34,35,36,37,38,39,
                                          41,42,43,44,48,49,50,57,58,
                                          59,62,65,70,82,83,84,85,89,
                                          90,91,100,102,103,104,110,
                                          111,120,121,130,131,140,141,
                                          500,501,502,504,505,506,601,
                                          602,603,772,773,774,777,780,
                                          781,783,784,785,788,791,796,
                                          797,798]))

def encontrar_indice_oficina(df):
    for i, value in df.iloc[:, 0].items():
        if isinstance(value, str) and value.strip().startswith("Oficina:"):
            return i
    return None


def indexar_hojas_excel(planilla_csv):
    
    """
    Lee un archivo Excel (.xls o .xlsx) y analiza hoja por hoja
    para detectar si contiene la cadena 'PLANILLA DE LIQUIDACION DE HORAS EXTRAS'.

    Errores que puede levantar:
        -de lectura de archivo u hoja de excel
        -que el excel tenga 1 hoja o más de 2 

    Retorna:
        {
            "es_planilla": List[Boolean],
            "nombres_hojas": {
                "hoja_planilla": String,
                "hoja_csv": String
            }
        }
    """
    # Detecta tipo de archivo automáticamente
    try:
        excel_file = pd.ExcelFile(planilla_csv)
    except Exception as e:
        raise ValueError(f"Error al abrir el archivo: {e}")
    

    nombres_hojas = excel_file.sheet_names
    es_planilla = []

    if len(nombres_hojas) > 2:
        st.error("No se pudo procesar el archivo. La planilla de horas extras cargada tiene más de dos hojas.")
        st.stop()

    # Iterar por cada hoja
    for hoja in nombres_hojas:
        try:
            df = pd.read_excel(planilla_csv, sheet_name=hoja, header=None, dtype=str)
            # Es planilla si tiene más de 30 columnas
            tiene_forma_planilla = df.shape[1] > 30
            es_planilla.append(tiene_forma_planilla)
        except Exception as e:
            st.error(f"⚠️ Error leyendo hoja '{hoja}': {e}. Reportar con el equipo de desarrollo.")
            st.stop()

    # Clasificar hojas
    hoja_planilla = nombres_hojas[0] if es_planilla[0] else nombres_hojas[1]
    hoja_csv = nombres_hojas[0] if not es_planilla[0] else nombres_hojas[1]

    return {
        "es_planilla": es_planilla,
        "nombres_hojas": {
            "hoja_planilla": hoja_planilla,
            "hoja_csv": hoja_csv
        }
    }

def buscar_tabla(df,keyword,nombre_tabla=""):
    '''
    El df es una lectura de dataframe de un archivo excel o csv que está muy mal formateado.
    Por ejemplo: la tabla del mismo empiece en otra fila distinta de la primera, que tenga anotaciones
    que no correspondan en otras celdas, etc.
    Asume que la keyword que es el encabezado de la tabla está en la primera columna del df
    Devuelve el indice idx_header de donde está el encabezado de la tabla.
    '''
    idx_header = None
    
    keyword = keyword.lower()

    if keyword in df.columns[0].lower(): #Si ya está el encabezado en la fila 0, que pandas lea como encabezado de DataFrame
        return -1

    #Si la tabla empieza en otra fila distinta de la de 1
    for i in range(df.shape[0]):
        cell = str(df.iat[i, 0]).lower()
        if keyword in cell: # LIKE '%keyword%'
            idx_header = i
            break
        if idx_header is not None:
            break

    if idx_header is None:
        st.error(
            f"❌ No se encontró la tabla para el archivo {nombre_tabla}\n\n"
            f"Se esperaba encontrar un encabezado que contenga la palabra '{keyword}' "
            "en alguna celda del archivo."
        )
        st.stop()

    return idx_header

def transformar_hhee_a_csv(df: pd.DataFrame):
    '''
    Arma el csv que se carga al sistema
    Se ejecuta una vez que se compararon las ausencias con la planilla de hhee.
    '''
    # Columnas que representan días/horas (todas menos legajo, nombre, tipo_hora)
    non_day = ["legajo", "nombre", "tipo_hora"]
    day_cols = [c for c in df.columns if c not in non_day]

    # 1) Sumar por legajo, nombre y tipo_hora todos los días
    #    -> queda un DataFrame con la suma por tipo_hora para ese legajo/nombre (aún en columnas de día)
    grouped = df.groupby(["legajo", "nombre", "tipo_hora"])[day_cols].sum()

    # 2) Colapsar las columnas de día a un único total por cada (legajo,nombre,tipo_hora)
    grouped_total = grouped.sum(axis=1).reset_index(name="horas")

    grouped_total["horas"] = np.ceil(grouped_total["horas"])
    # 3) Pivotear para que cada tipo_hora quede en su propia columna
    summary = grouped_total.pivot_table(
        index=["legajo", "nombre"],
        columns="tipo_hora",
        values="horas",
        fill_value=0
    ).reset_index()
  
    # 4) Renombrar las columnas según tu nomenclatura solicitada
    # Identificar los valores únicos en orden de aparición
    unique_types = list(df['tipo_hora'].dropna().unique())
    unique_types = unique_types[0:3]
    # Asegurar que tengamos exactamente 3 tipos
    if len(unique_types) < 3:
        st.error("Advertencia: se esperaban exactamente 3 tipos de hora. Se detectaron menos")
        st.stop()

    # Mapeo universal según orden
    mapping = {
        unique_types[0]: 'horas_normales',
        unique_types[1]: 'horas_50',
        unique_types[2]: 'horas_100'
    }

    summary = summary.rename(columns=mapping)
  
    # 5) Asegurarse de que existan las 3 columnas esperadas
    for col in ["horas_normales", "horas_50", "horas_100"]:
        if col not in summary.columns:
            summary[col] = 0

    # 6) Orden final de columnas: legajo, horas_normales, horas_50, horas_100, nombre
    summary_final = summary[["legajo", "horas_normales", "horas_50", "horas_100", "nombre"]]
    summary_final.insert(1, "columna(0)", 0)

    numeric_cols = ["horas_normales", "horas_50", "horas_100"]
    for col in numeric_cols:
        summary_final[col] = (
            summary_final[col]
                .astype(str)          # por si vienen como object/float/string
                .str.replace(",", ".", regex=False)  # reemplaza coma decimal si aparece
        )
        summary_final[col] = pd.to_numeric(summary_final[col], errors="coerce").fillna(0)
      
    # Lo transformamos a CSV
    return summary_final
    #csv = summary_final.to_csv(index=False).encode('latin1')
    #return csv

def cambiar_fechas(df):
    '''
    Recibe el dataframe de ausencias
    Lo que hace es cambiar las ausencias de forma tal que no se reemplace en
    la tabla las fechas de dia_inicia y dia_fin por el numero de día que correspondería al
    mes anterior (si es que la fecha es del mes pasado).
    '''
    df["dia_inicio"] = pd.to_datetime(df["dia_inicio"],format="%d/%m/%Y")
    df["dia_fin"] = pd.to_datetime(df["dia_fin"],format="%d/%m/%Y")

    hoy = datetime.today()
    hoy = hoy.replace(hour=0, minute=0, second=0, microsecond=0)

    # Determinar el mes anterior
    primer_dia_mes_anterior = (hoy.replace(day=1) - timedelta(days=1)).replace(day=1)
    ultimo_dia_mes_anterior = hoy.replace(day=1) - timedelta(days=1)

    # Función para acotar el rango al mes anterior (i.e. si es anterior al mes pasado
    # se inicializa en el primer día del mes anterior, análogo a si es un mes posterior).
    def acotar_al_mes_anterior(row):
        nuevo_inicio = row["dia_inicio"]
        nuevo_fin = row["dia_fin"]
        if row["dia_inicio"] < primer_dia_mes_anterior:
            nuevo_inicio = primer_dia_mes_anterior
        if row["dia_fin"] > ultimo_dia_mes_anterior:
            nuevo_fin = ultimo_dia_mes_anterior
        if nuevo_inicio > nuevo_fin:
            return pd.Series([primer_dia_mes_anterior,ultimo_dia_mes_anterior])  # rangos fuera del mes anterior
        return pd.Series([nuevo_inicio, nuevo_fin])

    df[["dia_inicio", "dia_fin"]] = df.apply(acotar_al_mes_anterior, axis=1)

    # Extraemos solo los días
    df["dia_inicio"] = df["dia_inicio"].dt.day
    df["dia_fin"] = df["dia_fin"].dt.day

def transformar_ausencias_a_dict(ausencias) -> dict:
    '''
    A partir de las ausencias se arma un diccionario:
    dict[legajo] = { "nombre": string, "dias": [int] }
    donde dias es una lista de numeros de los dias en 
    que esa persona estuvo ausente.
    '''
    df_raw = pd.read_excel(ausencias)

    oficina = None
    empleado = None
    legajo = None

    rows = []
 
    for _, row in df_raw.iterrows():
        primera_col = str(row[0]).strip() if pd.notna(row[0]) else ""

        # Detectar inicio de un bloque por Oficina
        if primera_col.startswith("Oficina :"):
            oficina = primera_col.replace("Oficina :", "").strip()
            empleado = None
            legajo = None
            continue

        # Detectar empleado
        if primera_col.startswith("Empleado:"):
            match = re.search(r"Empleado:\s*(.*?)\s*Legajo:\s*0*([0-9]+)", primera_col)
            if match:
                empleado = match.group(1).strip()
                legajo = match.group(2).strip()
            continue

        # Filas de ausencias (requieren oficina + empleado + fechas)
        if oficina and empleado and pd.notna(row[0]) and pd.notna(row[1]):
            primer_dia = row[0]
            ultimo_dia = row[1]
            motivo_raw = row[4] if len(row) > 4 else None
            nro_motivo = (
                motivo_raw.split("-")[0].strip()
                if isinstance(motivo_raw, str) and "-" in motivo_raw
                else None
            )

            rows.append([oficina, legajo, empleado, primer_dia, ultimo_dia, nro_motivo])
    
    df = pd.DataFrame(rows, columns=["oficina","legajo", "empleado", "dia_inicio", "dia_fin", "nro_motivo"])
    df["legajo"] = df["legajo"].astype(str).str.lstrip("0")
    
    cambiar_fechas(df)
    df["nro_motivo"] = df["nro_motivo"].astype(int)
    df = df[~df["nro_motivo"].isin(codigos_ausencias_no_descontables)]

    st.write(f"Esta es la planilla de ausencias")
    st.write(df)

    legajo_dict = defaultdict(lambda: {"nombre": None, "dias": []})
    
    for _, row in df.iterrows():
        legajo = str(row["legajo"])
        nombre = row["empleado"]
        oficina = row["oficina"]
        if int(row["dia_fin"]) == 30:
            dias = list(range(int(row["dia_inicio"]), int(row["dia_fin"]) + 2))
        else:
            dias = list(range(int(row["dia_inicio"]), int(row["dia_fin"]) + 1))

        legajo_dict[legajo]["nombre"] = nombre
        legajo_dict[legajo]["oficina"] = oficina
        legajo_dict[legajo]["dias"].extend(dias)
    
    # opcional: eliminar duplicados y ordenar
    for v in legajo_dict.values():
        v["dias"] = sorted(set(v["dias"]))

    return dict(legajo_dict)

def normalizar_planilla_hhee(planilla_hhee):
    '''
    Para la planilla de horas extras
    '''
    df = planilla_hhee
    df = df[df.columns[:34]]

    idx = buscar_tabla(df,"Legajo","Planilla de horas extras")
    if idx != -1: # si la tabla no tiene como encabezados los que queremos
        df.columns = df.iloc[idx]
        df = df.drop(idx)
      
    df = df.reset_index(drop=True)
    df = df.dropna(how="all")
    # Rellenar legajos vacíos con último valor válido
    df["Legajo"] = df["Legajo"].ffill(limit=2)
    df["Apellido y Nombre"] = df["Apellido y Nombre"].ffill(limit=2)
    # Quitar filas donde "Legajo" está vacío
    df = df[df["Legajo"].notna()]

    # Identificar columnas de días → las primeras 31 después de las 3 iniciales
    day_cols = df.columns[3:34]
  
    # forzar los valores a numeric
    df.iloc[:, 3:34] = df.iloc[:, 3:34].apply(pd.to_numeric, errors='coerce')

    # Renombrar las fechas por números 1–31
    df.rename(columns={day_cols[i]: i+1 for i in range(len(day_cols))}, inplace=True)

    df = df.rename(columns={
        df.columns[0]: "legajo",
        df.columns[1]: "nombre",
        df.columns[2]: "tipo_hora"
    }) 
    
    df["legajo"] = (
        df["legajo"]
        .apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x))
        .str.replace(r"[.,\s]", "", regex=True)
    )
    df = df[df["legajo"].astype(str).str.isdigit()]
    df["legajo"] = df["legajo"].astype(str)

    #Como las columnas que quedan que podrían tener na son de dias y hrs extras, se ponen en cero
    df = df.fillna(0)

    #Quitar aquellos espacios donde legajo quedó en 0
    df = df[df["legajo"] != '0']

    return df

def limpiar_nombre(nombre):
    nombre = nombre.upper() # mayúsculas
    nombre = re.sub(r"['’]", "", nombre) # quitar comas y apóstrofes
    nombre = re.sub(r",\s", " ", nombre)
    nombre = re.sub(r",", " ", nombre)
    reemplazos = {"Á": "A", 
                  "É": "E",
                  "Í": "I", 
                  "Ó": "O",
                  "Ú": "U",
                  "Ü": "U"}
    patron = re.compile("|".join(reemplazos.keys()))
    nombre = patron.sub(lambda m: reemplazos[m.group()], nombre) # reemplazar tildes
    nombre = nombre.strip() # sacar espacios adicionales
    nombre = re.sub(' +',' ',nombre) # idem 
    return nombre.split(' ')

def son_similares(nombre1, nombre2, umbral=0.8):
    ratio = difflib.SequenceMatcher(None, nombre1, nombre2).ratio()
    return ratio >= umbral

def imprimir_dict(lista,dicc):
    s = ''
    for legajo in lista:
        s += "- " + f"Legajo {legajo}, nombre sistema:{dicc[legajo]["nombre"]}, nombre en planilla: {dicc[legajo]["nombre"]}"  + "\n"
    st.markdown(s)

def reportar_inconsistencias(ausencias_ofi,planilla_hhee):
    '''
    Recibe los diccionarios ausencias_ofi y hhee_ofi.
    Para cada día donde se ausentaron e hicieron horas extras,
    segun el tipo de ausencia, se pone en 0 la hora extra en planilla_hhee
    Dejandolo listo para exportar a csv
    '''
    hhee = planilla_hhee
    inconsistencias_ausencias = []
    nombres_distintos = []
    #faltaría filtrar antes si hay legajos repetidos!!!!!!
    legajos_planilla = set(hhee["legajo"].unique().tolist())
    for legajo in ausencias_ofi.keys():
        legajo = str(legajo)
        if legajo in legajos_planilla:
            for dia in ausencias_ofi[legajo]["dias"]:
                # Para ese legajo si en algun día que estuvo ausente tiene horas extras, ponerlas en 0
                if (hhee.loc[hhee["legajo"] == legajo, dia] > 0).any():
                    inconsistencias_ausencias.append(f"Para {legajo} tenemos inconstencia el dia {dia}")
                    hhee.loc[hhee["legajo"] == legajo, dia] = 0

    st.write(f"Esta es la planilla de hhee después de ver inconsistencias")
    st.write(hhee)
    
    if len(inconsistencias_ausencias) > 0:
        st.write("Estos legajos tienen horas extras en un día que estuvieron ausentes")
        s = ""
        for x in inconsistencias_ausencias:
            s += "- " + x + "\n"
        st.markdown(s)         

######################
# Calcular variación #
######################  
# 
#Recibe dos (o más) planillas de horas extras y liquidación de dos meses distintos, el actual y el anterior

#------- Funciones auxiliares -----------
def esta_en_columna(df: pd.DataFrame, columna: str, valor):
  """
  Docstring for esta_en_columna
  
  :param df: dataFrame que tenga la columna que queremos ver
  :param columna: Nombre de la columna en donde se quiere ver si existe el valor
  :param valor: valor puede ser un string o un número que queremos ver si pertenece a la columna dada.
  """
  return df[columna].isin([valor]).any()

def tiene_guion(valor) -> bool:  
    # Verifica si el valor es una cadena y si contiene un guion
    if isinstance(valor, str) and ' - ' in valor:
        return True
    return False

def tipoDeFila(fila) -> int:
  '''
  Si es una persona devuelve un 0
  Si es un tipo de hora devuelve un 1
  Si es una oficina devuelve un 2
  '''
  tipo_fila = -1

  if tiene_guion(fila["Muni"]):

    palabra_dividida = fila['Muni'].split("-")

    primer_codigo = palabra_dividida[0].strip()

    nombre_value = fila['Nombre']

    if not pd.isna(nombre_value):
      tipo_fila = 0
    elif primer_codigo == "MU":
      tipo_fila = 1
    else:
      tipo_fila = 2

  else:
    tipo_fila = 0

  return tipo_fila

def obtener_tipo_de_hora(fila) -> int:
  '''
  Precondición: tipoDeFila(fila) == 1
  '''
  palabra_dividida = fila['Muni'].split("-")

  codigo_hora = palabra_dividida[1].strip()

  return codigo_hora

def obtener_oficina(fila) -> str:
  '''
  Precondición: tipoDeFila(fila) == 2
  '''
  if 'Muni' not in fila or pd.isna(fila['Muni']):
    return ""
  else:
    codigo_oficina = fila['Muni'].split("-")[0]
    return codigo_oficina.strip()

 #--------- Limpieza de datos originales ---------------

def limpiar(data: pd.DataFrame) -> defaultdict:
  """
  Arregla el dataSet para que sea de tipo Nombre, Legajo, Horas_86, Horas_87, Horas_89, valor_86,valor_87, valor_89.
  
  :param data: archivo en formato .xls pasado a dataFrame
  :return: Diccionario con legajo como clave y una lista de valores de la forma [nombre, oficina, legajo, cant_horas_86, valor_horas_86, cant_horas_87...]
  """
  d = defaultdict(list)

  cant_filas = data.shape[0]
  oficina = ""
  tipo_hora_extra = ""

  for i in range(0, cant_filas):

    fila = data.iloc[i]
    tipo_de_fila = tipoDeFila(fila)

    if tipo_de_fila == 1:

      tipo_hora_extra = obtener_tipo_de_hora(fila)

    elif tipo_de_fila == 2:

      oficina = obtener_oficina(fila)

    elif tipo_de_fila == 0:

      legajo = fila['Legajo']
      cant_horas = fila['Cant horas'] # Keep as is, will convert if not NaN
      valor_por_hora = fila['Valor por hora'] # Keep as is, will convert if not NaN
      valor = fila['Valor total'] # Keep as is, will convert if not NaN
      nombre = fila['Nombre']

      # A veces estos valores pueden ser nulos y no aportan significado
      # a la hora de calcular variaciones mes a mes (según nos comentaron
      # en Asistencias), por eso se ignoran
      if np.isnan(cant_horas) or np.isnan(valor_por_hora):
        continue

      # Convert to int/float after checking for NaN
      cant_horas = int(cant_horas)
      valor = float(valor)

      if legajo in d:

        if tipo_hora_extra == "A0786":

          d[legajo][0][3] = int(cant_horas)
          d[legajo][0][4] = float(valor)

        elif tipo_hora_extra == "A0787":

          d[legajo][0][5] = int(cant_horas)
          d[legajo][0][6] = float(valor)

        else:

          d[legajo][0][7] = int(cant_horas)
          d[legajo][0][8] = float(valor)

      else:

        if tipo_hora_extra == "A0786":

          valores = [nombre, oficina, legajo, int(cant_horas), float(valor), 0, 0, 0, 0]
          d[legajo].append(valores)

        elif tipo_hora_extra == "A0787":

          valores = [nombre, oficina, legajo, 0, 0 ,int(cant_horas), float(valor), 0, 0]
          d[legajo].append(valores)

        else:

          valores = [nombre, oficina, legajo, 0, 0, 0, 0, int(cant_horas), float(valor)]
          d[legajo].append(valores)

  return d

#-------- En caso de tener dos liquidaciones -------------

def agregar_liquidacion_extra(d: defaultdict, data: pd.DataFrame) -> None:
  """
  Precondición: tener más de una liquidación en un mes
  d (defaultDict): Diccionario creado con limpiar(data)
  data (DataFrame): DataFrame del archivo .xls correspondiente a la liquidación extra del mes
  """

  cant_filas = data.shape[0]
  tipo_hora_extra = ""
  oficina = ""


  for i in range(0, cant_filas):

    fila = data.iloc[i]

    if tipoDeFila(fila) == 1:

      tipo_hora_extra = obtener_tipo_de_hora(fila)

    elif tipoDeFila(fila) == 2:

      oficina = obtener_oficina(fila)

    elif tipoDeFila(fila) == 0:

      legajo = fila['Legajo']
      cant_horas = fila["Cant horas"]
      valor = fila["Valor total"]
      nombre = fila["Nombre"]
      valor_por_hora = fila['Valor por hora']
      
      #Fijarse si ese leajo es parte del diccionario del primer excel procesado

      if np.isnan(cant_horas) or np.isnan(valor_por_hora):
        continue

      # Convert to int/float after checking for NaN
      cant_horas = int(cant_horas)
      valor = float(valor)

      if legajo in d:
        if tipo_hora_extra == "A0786":
          
          d[legajo][0][3] = d[legajo][0][3] + cant_horas
          d[legajo][0][4] = d[legajo][0][4] + valor

        elif tipo_hora_extra == "A0787":
          
          d[legajo][0][5] = d[legajo][0][5] + cant_horas
          d[legajo][0][6] = d[legajo][0][6] + valor

        elif tipo_hora_extra == "A0789":
          
          d[legajo][0][7] = d[legajo][0][7] + cant_horas
          d[legajo][0][8] = d[legajo][0][8] + valor

      else:

          if tipo_hora_extra == "A0786":

            valores = [nombre, oficina, legajo, cant_horas, valor, 0, 0, 0, 0]
            d[legajo].append(valores)

          elif tipo_hora_extra == "A0787":

            valores = [nombre, oficina, legajo, 0, 0 ,cant_horas, valor, 0, 0]
            d[legajo].append(valores)

          else:

            valores = [nombre, oficina, legajo, 0, 0, 0, 0, cant_horas, valor]
            d[legajo].append(valores)

#------------- Creación de dataset --------------------

def armar_data_set(d: defaultdict) -> pd.DataFrame:
  """
  Pasar de un diccionario con legajo como clave, y una lista de valores a un dataFrame
  
  :param d: Diccionario con clave "legajo" y valores, los datos recolectados 
  :return: Dataframe resumen del diccionario
  """

  dataSetLimpio = pd.DataFrame(columns = ["Nombre", "Oficina", "Legajo",
                                        "HorasExtra_86", "Valor_86",
                                        "HorasExtra_87", "Valor_87",
                                        "HorasExtra_89", "Valor_89"])

  legajos = []
  nombres = []
  oficinas = []
  horas_86 = []
  horas_87 = []
  horas_89 = []
  valor_86 = []
  valor_87 = []
  valor_89 = []

  for key in d.keys():

    valores = d[key][0]

    legajos.append(key)
    nombres.append(valores[0])
    oficinas.append(valores[1])
    horas_86.append(valores[3])
    horas_87.append(valores[5])
    horas_89.append(valores[7])
    valor_86.append(valores[4])
    valor_87.append(valores[6])
    valor_89.append(valores[8])

  dataSetLimpio["Nombre"] = nombres
  dataSetLimpio["Oficina"] = oficinas
  dataSetLimpio["Legajo"] = legajos
  dataSetLimpio["HorasExtra_86"] = horas_86
  dataSetLimpio["HorasExtra_87"] = horas_87
  dataSetLimpio["HorasExtra_89"] = horas_89
  dataSetLimpio["Valor_86"] = valor_86
  dataSetLimpio['Valor_87'] = valor_87
  dataSetLimpio['Valor_89'] = valor_89

  return dataSetLimpio

def agregar_total(df_anterior: pd.DataFrame, df_actual:pd.DataFrame) -> None:
  """
  Agrega las columnas de totales, donde para su respectiva columna se suma la respectiva columna del dataframe anterior y del actual
  
  :param df_anterior: dataframe del mes anterior
  :param df_actual: dataframe del mes actual
  """

  df_anterior["Total horas"] = df_anterior["HorasExtra_86"] + df_anterior["HorasExtra_87"] + df_anterior["HorasExtra_89"]
  df_actual["Total horas"] = df_actual["HorasExtra_86"] + df_actual["HorasExtra_87"] + df_actual["HorasExtra_89"]

  df_anterior["Total valor"] = df_anterior["Valor_86"] + df_anterior["Valor_87"] + df_anterior["Valor_89"]
  df_actual["Total valor"] = df_actual["Valor_86"] + df_actual["Valor_87"] + df_actual["Valor_89"]


# ----------- Archivos necesarios ------------------
# Son los archivos que van a ser devueltos en formato .xlsx

def unir_oficinas(df_anterior: pd.DataFrame, df_actual: pd.DataFrame) -> pd.DataFrame:
  df_area = pd.DataFrame(columns = ["Oficina", "Dif. horas extras normales", "Dif. liquidado normales",
                                  "Dif. horas extras al 50", "Dif. liquidado al 50",
                                  "Dif.horas extras al 100", "Dif. liquidado al 100","Dif. porcentual",
                                  "Dif. porcentual ponderado"])

  oficinas = []
  dif_horas_86 = []
  dif_valor_86 = []
  dif_horas_87 = []
  dif_valor_87 = []
  dif_horas_89 = []
  dif_valor_89 = []
  dif_horas_total = []
  dif_valor_total = []
  dif_porcentual_horas = []
  dif_porcentual_ponderado = []

  oficinas_mes_ant = df_anterior["Oficina"].unique()
  oficinas_mes_act = df_actual["Oficina"].unique()

  oficinas_unicas = np.union1d(oficinas_mes_ant, oficinas_mes_act)

  for i,oficina in enumerate(oficinas_unicas):

    if (esta_en_columna(df_anterior,"Oficina", oficina)) and (not esta_en_columna(df_actual,"Oficina", oficina)):

      # La oficina solo se encuentra en un archivo .xls
      df_oficinas_ant = df_anterior[df_anterior["Oficina"] == oficina]
      cant_oficina_ant = df_oficinas_ant.shape[0]

      cant_oficina_act = 0

      valor_86_ant = df_oficinas_ant["Valor_86"].sum()
      valor_86_act = 0

      valor_87_ant = df_oficinas_ant["Valor_87"].sum()
      valor_87_act = 0

      valor_89_ant = df_oficinas_ant["Valor_89"].sum()
      valor_89_act = 0

      horas_86_ant = df_oficinas_ant["HorasExtra_86"].sum()
      horas_86_act = 0

      horas_87_ant = df_oficinas_ant["HorasExtra_87"].sum()
      horas_87_act = 0

      horas_89_ant = df_oficinas_ant["HorasExtra_89"].sum()
      horas_89_act = 0

      #Total
      horas_total_ant = df_oficinas_ant["Total horas"].sum()
      horas_total_act = 0
      valor_total_ant = df_oficinas_ant["Total valor"].sum()
      valor_total_act = 0

    elif (not esta_en_columna(df_anterior,"Oficina", oficina)) and (esta_en_columna(df_actual,"Oficina", oficina)):

      # La oficina solo se encuentra en un archivo .xls
      df_oficinas_act = df_actual[df_actual["Oficina"] == oficina]
      cant_oficina_act = df_oficinas_act.shape[0]
      cant_oficina_ant = 0

      valor_86_ant = 0
      valor_86_act = df_oficinas_act["Valor_86"].sum()

      valor_87_ant = 0
      valor_87_act = df_oficinas_act["Valor_87"].sum()

      valor_89_ant = 0
      valor_89_act = df_oficinas_act["Valor_89"].sum()

      horas_86_ant = 0
      horas_86_act = df_oficinas_act["HorasExtra_86"].sum()

      horas_87_ant = 0
      horas_87_act = df_oficinas_act["HorasExtra_87"].sum()

      horas_89_ant = 0
      horas_89_act = df_oficinas_act["HorasExtra_89"].sum()

      #Total
      horas_total_ant = 0
      horas_total_act = df_oficinas_act["Total horas"].sum()
      valor_total_ant = 0
      valor_total_act = df_oficinas_act["Total valor"].sum()

    elif esta_en_columna(df_anterior, "Oficina", oficina) and esta_en_columna(df_actual, "Oficina", oficina):
      # La oficina se encuentra en ambos archivos .xls

      df_oficinas_act = df_actual[df_actual["Oficina"] == oficina]
      cant_oficina_act = df_oficinas_act.shape[0]

      df_oficinas_ant = df_anterior[df_anterior["Oficina"] == oficina]
      cant_oficina_ant = df_oficinas_ant.shape[0]

      valor_86_ant = df_oficinas_ant["Valor_86"].sum()
      valor_86_act = df_oficinas_act["Valor_86"].sum()

      valor_87_ant = df_oficinas_ant["Valor_87"].sum()
      valor_87_act = df_oficinas_act["Valor_87"].sum()

      valor_89_ant = df_oficinas_ant["Valor_89"].sum()
      valor_89_act = df_oficinas_act["Valor_89"].sum()

      horas_86_ant = df_oficinas_ant["HorasExtra_86"].sum()
      horas_86_act = df_oficinas_act["HorasExtra_86"].sum()

      horas_87_ant = df_oficinas_ant["HorasExtra_87"].sum()
      horas_87_act = df_oficinas_act["HorasExtra_87"].sum()

      horas_89_ant = df_oficinas_ant["HorasExtra_89"].sum()
      horas_89_act = df_oficinas_act["HorasExtra_89"].sum()

      #Total
      horas_total_ant = df_oficinas_ant["Total horas"].sum()
      horas_total_act = df_oficinas_act["Total horas"].sum()
      valor_total_ant = df_oficinas_ant["Total valor"].sum()
      valor_total_act = df_oficinas_act["Total valor"].sum()

    oficinas.append(oficina)
    #Tipo de hora: 86
    dif_horas_86.append(horas_86_act - horas_86_ant)
    dif_valor_86.append(valor_86_act - valor_86_ant)
    #Tipo de hora:87
    dif_horas_87.append(horas_87_act - horas_87_ant)
    dif_valor_87.append(valor_87_act - valor_87_ant)
    #Tipo de hora:89
    dif_horas_89.append(horas_89_act - horas_89_ant)
    dif_valor_89.append(valor_89_act - valor_89_ant)
    #Total
    dif_horas_total.append(horas_total_act - horas_total_ant)
    dif_valor_total.append(valor_total_act - valor_total_ant)
    #Dif. porcentual
    if horas_total_ant == 0:
      dif_porcentual_horas.append(np.nan)
    else:
      dif_porcentual_horas.append(horas_total_act/horas_total_ant - 1)
    #Dif. porcentual ponderado
    if horas_total_ant == 0:
      dif_porcentual_ponderado.append(np.nan)
    elif horas_total_act == 0:
      dif_porcentual_ponderado.append(np.nan)
    else:
      dif_porcentual_ponderado.append((horas_total_act/cant_oficina_act - horas_total_ant/cant_oficina_ant)/(horas_total_act/cant_oficina_act))

  df_area["Oficina"] = oficinas
  df_area["Dif. horas extras normales"] = dif_horas_86
  df_area["Dif. liquidado normales"] = dif_valor_86
  df_area["Dif. horas extras al 50"] = dif_horas_87
  df_area["Dif. liquidado al 50"] = dif_valor_87
  df_area["Dif.horas extras al 100"] = dif_horas_89
  df_area["Dif. liquidado al 100"] = dif_valor_89
  df_area["Dif. horas total"] = dif_horas_total
  df_area["Dif. valor total"] = dif_valor_total
  df_area["Dif. porcentual"] = dif_porcentual_horas
  df_area["Dif. porcentual ponderado"] = dif_porcentual_ponderado

  return df_area

def unir_personas(df_1, df_2):

  df_personas = pd.DataFrame(columns = ["Legajo", "Nombre", "Oficina","Dif. horas extras", "Dif. liquidado"])

  dif_horas = []
  dif_valor = []
  nombres = []
  legajos = []
  oficinas = []

  legajos_mes_1 = df_1["Legajo"].unique()
  legajos_mes_2 = df_2["Legajo"].unique()

  # Filtra valores NaN y se consideran legajos como valor numérico
  legajos_mes_1 = legajos_mes_1[~pd.isna(legajos_mes_1)]
  legajos_mes_2 = legajos_mes_2[~pd.isna(legajos_mes_2)]

  legajos_unicos = np.union1d(legajos_mes_1, legajos_mes_2)

  for legajo in legajos_unicos:
    legajos.append(legajo)

    #Una persona puede no estar en ambos meses
    if esta_en_columna(df_1, "Legajo", legajo) and esta_en_columna(df_2, "Legajo", legajo):
      #La persona hizo horas extras ambos meses
      fila_persona_1 = df_1[df_1["Legajo"] == legajo].iloc[0]
      fila_persona_2 = df_2[df_2["Legajo"] == legajo].iloc[0]
      oficina_persona = fila_persona_1["Oficina"]

      dif_horas.append(fila_persona_2["Total horas"] - fila_persona_1["Total horas"])

      dif_valor.append(fila_persona_2["Total valor"] - fila_persona_1["Total valor"])

      nombres.append(fila_persona_1["Nombre"])

    elif esta_en_columna(df_1, "Legajo", legajo) and (not esta_en_columna(df_2, "Legajo", legajo)):
      #La persona hizo horas extras solo el mes 1
      fila_persona_1 = df_1[df_1["Legajo"] == legajo].iloc[0]
      oficina_persona = fila_persona_1["Oficina"]

      dif_horas.append(-fila_persona_1["Total horas"])

      dif_valor.append(-fila_persona_1["Total valor"])
      nombres.append(fila_persona_1["Nombre"])

    else:
      # La persona hizo horas extras solo el mes 2
      fila_persona_2 = df_2[df_2["Legajo"] == legajo].iloc[0]
      oficina_persona = fila_persona_2["Oficina"]

      dif_horas.append(fila_persona_2["Total horas"])

      dif_valor.append(fila_persona_2["Total valor"])

      nombres.append(fila_persona_2["Nombre"])
    oficinas.append(oficina_persona)

  df_personas["Legajo"] = legajos
  df_personas["Nombre"] = nombres
  df_personas["Dif. horas extras"] = dif_horas
  df_personas["Dif. liquidado"] = dif_valor
  df_personas["Oficina"] = oficinas

  return df_personas

def resumen_oficinas(df):
  """
  Solo para el mes actual
  df (pd.DataFrame): dataFrame del mes actual
  """

  oficinas_unicas = df["Oficina"].unique()

  df_area_total = pd.DataFrame(columns = ["Oficina", "Total horas normales", "Total liquidado normales",
                                          "Total horas al 50", "Total liquidado al 50", "Total horas al 100",
                                          "Total liquidado al 100"])

  oficinas = []
  total_hora_86 = []
  total_hora_87 = []
  total_hora_89 = []

  total_valor_86 = []
  total_valor_87 = []
  total_valor_89 = []

  for oficina in oficinas_unicas:

    oficinas.append(oficina) 

    df_oficina = df[df["Oficina"] == oficina]

    hora_86 = df_oficina["HorasExtra_86"].sum()
    hora_87 = df_oficina["HorasExtra_87"].sum()
    hora_89 = df_oficina["HorasExtra_89"].sum()

    valor_86 = df_oficina["Valor_86"].sum()
    valor_87 = df_oficina["Valor_87"].sum()
    valor_89 = df_oficina["Valor_89"].sum()

    total_hora_86.append(hora_86)
    total_hora_87.append(hora_87)
    total_hora_89.append(hora_89)

    total_valor_86.append(valor_86)
    total_valor_87.append(valor_87)
    total_valor_89.append(valor_89)

  df_area_total["Oficina"] = oficinas
  df_area_total["Total horas normales"] = total_hora_86
  df_area_total["Total liquidado normales"] = total_valor_86
  df_area_total["Total horas al 50"] = total_hora_87
  df_area_total["Total liquidado al 50"] = total_valor_87
  df_area_total["Total horas al 100"] = total_hora_89
  df_area_total["Total liquidado al 100"] = total_valor_89

  return df_area_total  

#############################
#         STREAMLIT         #
#############################
st.title("Asistencia's Assistant 🤖")

tab1, tab2 = st.tabs(["Armar CSV","Variación intermensual"])

with tab1:
    st.subheader("📝Comparación con ausencias y armado del CSV")
    planilla_csv = st.file_uploader("Subí la planilla de horas extras")
    ausencias = st.file_uploader("Subí la planilla de ausencias")
    nombre_archivo = st.text_input("Escribí el nombre del archivo csv que querés generar")
    if planilla_csv and ausencias and nombre_archivo:
        #para planilla_csv hay que indexar las hojas:
        resumen = indexar_hojas_excel(planilla_csv)
        #trabajamos con planilla_hhee
        nombres_hojas = resumen["nombres_hojas"]
        planilla_hhee = pd.read_excel(planilla_csv, sheet_name = nombres_hojas["hoja_planilla"], engine = "calamine")
        #comparamos con ausencias
        ausencias_ofi = transformar_ausencias_a_dict(ausencias)    
        planilla = normalizar_planilla_hhee(planilla_hhee)
        #transformamos planilla_hhee en csv resumen planilla
        resumen_planilla_antes = transformar_hhee_a_csv(planilla)
        reportar_inconsistencias(ausencias_ofi,planilla)
        st.write(f"Este es el resumen antes de modificarlo")
        st.write(resumen_planilla_antes)
        resumen_planilla = transformar_hhee_a_csv(planilla)
        st.write(f"Este es el resumen después de modificarlo")
        st.write(resumen_planilla)

        csv = resumen_planilla.to_csv(index=False).encode('latin1')
        st.download_button(
            label="Descargar CSV",
            data=csv,
            file_name=f"{nombre_archivo}.csv",
            mime="text/csv",
            key='download_csv_no_index'
        )

with tab2:

    ahora = datetime.now()

    # Obtener el primer día del mes actual
    primer_dia_mes = ahora.replace(day=1)

    # Restar un día para ir al último día del mes anterior
    mes_anterior_fecha = primer_dia_mes - timedelta(days=1)

    # Formatear como "YYYY-MM"
    mes_anterior_str = mes_anterior_fecha.strftime("%Y-%m")

    st.subheader("📊 Variación intermensual")
    st.markdown("📂 Subí los archivos correspondientes al _mes anterior_. En caso de tener dos liquidaciones, subí ambos juntos.")

    archivos_1 = st.file_uploader("", type=["xls"], key="archivo1", accept_multiple_files=True)
    cant_mes_anterior = len(archivos_1)

    #st.write(f"Archivos seleccionados: {cant_mes_anterior}")

    st.markdown("📂 Subí los archivos correspondientes al _mes actual_. En caso de tener dos liquidaciones, sube ambos juntos.")

    archivos_2 = st.file_uploader("", type=["xls"], key="archivo2",accept_multiple_files=True)
    cant_mes_actual = len(archivos_2)

    #st.write(f"Archivos seleccionados: {cant_mes_actual}")

    # --- Cuando ambos archivos son subidos ---
    if archivos_1 and archivos_2:
        
        st.success("Archivos cargados correctamente.")

        dfs_1 = []
        dfs_2 = []

        for archivo_1 in archivos_1:
            df_1 = pd.read_excel(archivo_1, engine='xlrd')
            dfs_1.append(df_1)

        for archivo_2 in archivos_2:
            df_2 = pd.read_excel(archivo_2,engine='xlrd')
            dfs_2.append(df_2)
        
        for i,df in enumerate(dfs_1):
            df.columns =  ["Muni", "Legajo", "Nombre", "Liq","Base",
                    "Cant horas","Valor por hora","Saporte",
                    "Fecha","Valor total"]
            
            if i == 0:
                mes_anterior = limpiar(df)
            else:
                agregar_liquidacion_extra(mes_anterior, df)

        for j,df in enumerate(dfs_2):
            df.columns = ["Muni", "Legajo", "Nombre", "Liq","Base",
                    "Cant horas","Valor por hora","Saporte",
                    "Fecha","Valor total"]
            
            if j == 0:
                mes_actual = limpiar(df)
            else:
                agregar_liquidacion_extra(mes_actual, df)

        dataSetLimpio_mes1 = armar_data_set(mes_anterior)
        dataSetLimpio_mes2 = armar_data_set(mes_actual)

        agregar_total(dataSetLimpio_mes1,dataSetLimpio_mes2)

        df_area = unir_oficinas(dataSetLimpio_mes1,dataSetLimpio_mes2)

        df_personas = unir_personas(dataSetLimpio_mes1, dataSetLimpio_mes2)

        df_area_total = resumen_oficinas(dataSetLimpio_mes2)

        output1 = io.BytesIO()
        output2 = io.BytesIO()
        output3 = io.BytesIO()

        df_area.to_excel(output1, index=False)
        df_personas.to_excel(output2, index=False)
        df_area_total.to_excel(output3, index = False)

        nombre_archivo_1 = f"Dif. horas extras por oficina_{mes_anterior_str}.xlsx"
        nombre_archivo_2 = f"Dif. horas extras por persona_{mes_anterior_str}.xlsx"
        nombre_archivo_3 = f"Resumen horas extras mes actual_{mes_anterior_str}.xlsx"

        st.download_button(
            label="📄 Descargar planilla de diferencias de horas extras por oficina",
            data=output1.getvalue(),
            file_name=nombre_archivo_1,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="📄 Descargar planilla de diferencias de horas extras por persona",
            data=output2.getvalue(),
            file_name=nombre_archivo_2,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
        label="📄 Descargar resumen por oficina para el mes actual",
        data=output3.getvalue(),
        file_name=nombre_archivo_3,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
