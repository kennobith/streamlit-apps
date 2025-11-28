import pandas as pd
import streamlit as st
from collections import defaultdict
from datetime import datetime, timedelta
import difflib
import re

#Estos son codigos de motivos de ausencia que detallan para que si
#se declara en la planilla que una persona hizo horas extras en un 
#día bajo alguno de estos codigos no se descuenten las mismas.
codigos_ausencias_no_descontables = (set([1,4,5,6,7,8,9,10,11,12,13,
                                          14,15,16,17,18,21,22,25,26,
                                          30,31,32,33,34,35,36,37,39,
                                          41,42,43,44,48,49,50,57,58,
                                          59,62,65,70,82,83,84,85,89,
                                          90,91,104,110,111,120,121,130,
                                          131,140,141,500,501,502,504,
                                          505,506,601,602,603,772,773,
                                          774,777,780,781,783,784,785,
                                          788,791,796,797,798]))

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
    df = df[df["nro_motivo"].isin(codigos_ausencias_no_descontables)]

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
    
def obtener_dict_personas_hhee(df):
    '''
    Creo que no va mas
    A partir de la planilla de horas extras se arma un diccionario:
    dict[legajo] = { "nombre": string, "dias": [int] }
    donde dias es una lista de los dias del mes anterior que ese legajo
    hizo horas extras
    '''
    #quizás todo esto no aca falta después si ya tenemos normalizar_planilla_hhee
    df = df[df.columns[:34]]

    idx = buscar_tabla(df,"Legajo","Planilla de horas extras")
    df.columns = df.iloc[idx]
    df = df.drop(idx)
    df = df.reset_index(drop=True)
    df = df.dropna(how="all")
    # Rellenar legajos vacíos con último valor válido
    df["Legajo"] = df["Legajo"].replace("", pd.NA).ffill()

    # Quitar filas donde "Legajo" está vacío
    df = df[df["Legajo"].notna()]

    # Identificar columnas de días → las primeras 31 después de las 3 iniciales
    day_cols = df.columns[3:34]

    # Renombrar las fechas por números 1–31
    df.rename(columns={day_cols[i]: i+1 for i in range(len(day_cols))}, inplace=True)

    # Agrupar por legajo (cada grupo tiene 3 filas: N, 50, 100)
    result = {}

    for legajo, group in df.groupby("Legajo"):
        nombre = group["Apellido y Nombre"].iloc[0]
        
        # Sumar horas por día (N + 50 + 100)
        sum_days = group[list(range(1, 32))].sum()
        
        # Días con horas extras
        dias = [int(d) for d, v in sum_days.items() if v > 0]
        
        result[str(legajo)] = {
            "nombre": nombre,
            "dias": dias
        }

    return result

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
    #if len(nombres_distintos) > 0:
    #    st.write("Estos legajos parecen no coincidir en nombres")
    #    imprimir_dict(nombres_distintos,ausencias_ofi)
    
    if len(inconsistencias_ausencias) > 0:
        st.write("Estos legajos tienen horas extras en un día que estuvieron ausentes")
        s = ""
        for x in inconsistencias_ausencias:
            s += "- " + x + "\n"
        st.markdown(s)
    
    #if len(inconsistencias_ausencias) > 0:
    #   st.stop()
               

#############################
#         STREAMLIT         #
#############################
st.title("Asistencia's Assistant 🤖")
planilla_csv = st.file_uploader("subí la planilla de horas extras")
ausencias = st.file_uploader("subí la planilla de ausencias")
if planilla_csv and ausencias:
    #para planilla_csv hay que indexar las hojas:
    resumen = indexar_hojas_excel(planilla_csv)
    #trabajamos con planilla_hhee
    nombres_hojas = resumen["nombres_hojas"]
    planilla_hhee = pd.read_excel(planilla_csv, sheet_name = nombres_hojas["hoja_planilla"], engine = "calamine")
    #comparamos con ausencias
    ausencias_ofi = transformar_ausencias_a_dict(ausencias)    
    #hhee_ofi = obtener_dict_personas_hhee(planilla_hhee)
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
        label="Download CSV without Index",
        data=csv,
        file_name="my_data_no_index.csv",
        mime="text/csv",
        key='download_csv_no_index'
    )
    









