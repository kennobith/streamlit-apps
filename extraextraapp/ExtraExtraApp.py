import pandas as pd
import re
import streamlit as st
import io
import difflib

# FUNCIONES AUXILIARES SCRIPT
def type_cast_to_integer(df,nombres_col):
    for nombre_col in nombres_col:
        df[nombre_col] = pd.to_numeric(df[nombre_col],downcast='integer')
    return df

def type_cast_to_string(df,nombres_col):
    for nombre_col in nombres_col:
        df[nombre_col] = df[nombre_col].astype(str)
    return df

def flatten(oficinas,clave):
    lista = [
        x 
        for i in range(len(oficinas))
        for x in oficinas[i][clave]
        ]
    
    return lista

def ordenar_por_legajo_y_dict(df):
    df = df.sort_values(by = 'legajo')
    df= df.set_index('legajo').T.to_dict('list')
    return df

def dict_a_dataframe(diccionario, columnas):
    df = pd.DataFrame.from_dict(diccionario, orient='index', columns=columnas)
    df.reset_index(inplace=True)       # vuelve el índice (legajo) a columna
    df.rename(columns={'index': 'legajo'}, inplace=True)
    return df

def limpiar_csv(archivo):
    df = pd.read_csv(archivo, encoding="latin1",skip_blank_lines=True)

    ultima_columna  = df.columns[-1]
    primera_columna = df.columns[0]
    
    # Quitar ultima columna si todos los elementos son nulos.
    if df[ultima_columna].isnull().all(): 
        df = df.drop(columns=[ultima_columna])

    # Quitar filas que tengan legajos nulos.
    df = df[df.iloc[:, 0].notna()]

    # Typecast columna de legajos dependiendo si es o no string.
    if df[primera_columna].dtype == object and isinstance(df.iloc[0][primera_columna], str):
        df[primera_columna] = df[primera_columna].apply(lambda line: "".join(filter(lambda ch: ch not in " ?.!/;:,", line)))
    else:
        df = type_cast_to_integer(df,[primera_columna])
        df = type_cast_to_string(df,[primera_columna])

    # A las celdas vacías les ponemos cero
    df = df.fillna(0)

    return df

def expand_column(col, prefix):
        return pd.DataFrame(col.tolist(), 
                            index=col.index, 
                            columns=[f"Nombre en {prefix}", f"Oficina en {prefix}", f"H.E. normales en {prefix}", f"H.E. al 50 en {prefix}", f"H.E. al 100 en {prefix}"])

def esta_en_oficinas(resultados,legajo,oficinas):
    return resultados[legajo][1].strip() in oficinas

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

def son_iguales(nombre1, nombre2, umbral=0.8):
    ratio = difflib.SequenceMatcher(None, nombre1, nombre2).ratio()
    return ratio >= umbral

# FUNCIONES PRINCIPALES
def procesar_novedades_sistema(novedades_sistema):
    novedades_sistema = pd.read_excel(novedades_sistema, engine="xlrd")

    # Dividimos columnas que tienen doble información.
    novedades_sistema[['legajo','nro_cargo']] = novedades_sistema['LEGAJO'].str.split('-',n=1, expand=True)
    novedades_sistema[['año','oficina']] = novedades_sistema['OFICINA'].str.split('-',n=1,expand=True)

    # Cambiamos nombres de columnas.
    novedades_sistema['nombre_completo'] = novedades_sistema['APELLIDO Y NOMBRE']
    novedades_sistema['valor_hora_extra'] = novedades_sistema['VALOR']
    novedades_sistema['tipo_hora_extra'] = novedades_sistema['DESCRIPCIÓN']
    novedades_sistema['oficina'] = novedades_sistema['oficina'].str.strip()

    # Armamos df_con lo que nos interesa.
    df = pd.DataFrame(
                    {
                        'legajo': novedades_sistema['legajo'],
                        'nombre_completo': novedades_sistema['nombre_completo'],
                        'oficina': novedades_sistema['oficina'],
                        'tipo_hora_extra': novedades_sistema['tipo_hora_extra'],
                        'valor_hora_extra': novedades_sistema['valor_hora_extra'],
                    }
    )

    # Reemplazar valores de data frame.
    mapeo = {f'@HRSEXTR{i}': i for i in range(1,4)}
    df.replace(mapeo,inplace=True)

    # Pivotear tabla.
    df = pd.pivot_table(
        df,
        index = ['legajo','nombre_completo','oficina'],
        columns = ['tipo_hora_extra'],
        values = ['valor_hora_extra'],
        fill_value = 0
    )   

    # Aplanar columnas.
    df.columns = [f'{col[0]}_{col[1]}' for col in df.columns]
    df = df.reset_index()

    # Type casting columnas.
    columnas_a_integrar = ['legajo'] + [f'valor_hora_extra_{i}' for i in range(1,4)]
    df = type_cast_to_integer(df,columnas_a_integrar)
    df['legajo'] = df['legajo'].apply(str)

    # Ordenar por legajo y convertir a dict.
    resultados_sistema = ordenar_por_legajo_y_dict(df)

    return resultados_sistema

def procesar_csvs_oficinas(archivos):
    oficinas = []

    # Procesar cada csv.
    for archivo in archivos:
        if archivo.name.endswith(".csv"): 
            df_reportado = limpiar_csv(archivo)

            ofi = archivo.name.strip(".csv")

            nombres_columnas = df_reportado.columns.tolist()
            oficinas.append({'nro_ofi': ofi,
                             'tam_ofi': len(df_reportado),
                             'legajos': df_reportado[nombres_columnas[0]].tolist(),
                             'nombres': df_reportado[nombres_columnas[5]].tolist(),
                             'hs_tip1': df_reportado[nombres_columnas[2]].tolist(),
                             'hs_tip2': df_reportado[nombres_columnas[3]].tolist(),
                             'hs_tip3': df_reportado[nombres_columnas[4]].tolist()})
    # Ponemos en listas todos los atributos de cada diccionario para armar el df_reportado

    # Armar lista que te da todos los numeros de oficinas en el orden en el que está en oficina
    # Si oficinas[0] = diccionario de la 310 con 3 empleados
    # Si oficinas[1] = diccionario de la 311 con 2 empleados
    # => oficinas_todas = [310,310,310,311,311]
    oficinas_todas = [
            oficinas[i]['nro_ofi']
            for i in range(len(oficinas))
            for _ in range(oficinas[i]['tam_ofi'])
        ]
    
    legajos = flatten(oficinas,'legajos')
    nombres = flatten(oficinas,'nombres')
    hs_tip1 = flatten(oficinas,'hs_tip1')
    hs_tip2 = flatten(oficinas,'hs_tip2')
    hs_tip3 = flatten(oficinas,'hs_tip3')

    # Armar resultados_reporte
    df = pd.DataFrame(
        {   
            'legajo': legajos,
            'nombre_completo': [nombre.upper() for nombre in nombres],
            'oficinas': oficinas_todas,
            'valor_hora_extra_1': hs_tip1,
            'valor_hora_extra_2': hs_tip2,
            'valor_hora_extra_3': hs_tip3
        }
    )

    resultados_reporte = ordenar_por_legajo_y_dict(df)
    return resultados_reporte

def comparar_y_armar_df(resultados_sistema,resultados_reporte,oficinas):

    no_coinciden = {} # Legajos de quienes no coinciden lo reportado y lo cargado en sistema.

    no_reportados = [] # Legajos de quienes estan en sistema pero no fueron reportados.
    no_estan_en_sistema = [] # Legajos de quienes fueron reportados pero no cargados en sistema.

    # Por cada legajo reportado ver si está en sistema:
    for legajo in resultados_reporte.keys():
        if legajo not in resultados_sistema.keys():
            # Aquellos a quienes se reportan horas extras nulas no van a aparecer en la planilla del sistema.
            if resultados_reporte[legajo][2:5] != [0,0,0]:
                no_estan_en_sistema.append(f'Legajo: {legajo} - Archivo: {resultados_reporte[legajo][1]}')

    # Por cada legajo en sistema, ver si está en reporte
    for legajo in resultados_sistema.keys():
        # Comparar
        if legajo in resultados_reporte.keys():
            if resultados_sistema[legajo][2:5] != resultados_reporte[legajo][2:5]:
                no_coinciden[legajo] = {'sistema':resultados_sistema[legajo],'reporte': resultados_reporte[legajo]}
        # Si no esta en el reporte, ver si...
        else:
            # la oficna del mismo está en una de las oficinas que se ingresaron
            if oficinas != [1,1,1] and esta_en_oficinas(resultados_sistema,legajo,oficinas):
                no_reportados.append(f'Legajo: {legajo} - Oficina: {resultados_sistema[legajo][1]}')
            # si pidieron todas las oficinas, informalos siempre
            elif oficinas == [1,1,1]:
                no_reportados.append(f'Legajo: {legajo} - Oficina: {resultados_sistema[legajo][1]}')

    df = pd.DataFrame(no_coinciden.values(),index=no_coinciden.keys())

    # Si hay coincidencias, devolver los que no estén en sistema o no estén reportados
    # Si no hay coincidencias, armar dataframe
    df_final = None
    if not df.empty:
        df_sistema_expandido = expand_column(df['sistema'], 'sistema')
        df_reporte_expandido = expand_column(df['reporte'], 'reporte')
        df_final = pd.concat([df_sistema_expandido, df_reporte_expandido], axis=1)
        
    return df_final, no_estan_en_sistema, no_reportados

def comparar_nombres(resultados_sistema,resultados_reporte):
    columnas = ['nombre','oficina','hr_extr1','hr_extr2','hr_extr3']
    df_s = dict_a_dataframe(resultados_sistema,columnas)
    df_r = dict_a_dataframe(resultados_reporte,columnas)
    df_s = df_s[['legajo','nombre']]
    df_r = df_r[['legajo','nombre','oficina']]
    df = pd.merge(df_s,df_r,on='legajo',how='outer')
    df = df.dropna() # quitar donde no este reportado o no esté cargado
    no_coinciden = {}
    personas = ordenar_por_legajo_y_dict(df)
    for legajo,nombres in personas.items():
        nombre_s = limpiar_nombre(nombres[0])
        nombre_r = limpiar_nombre(nombres[1])
        archivo = nombres[2]
        coincidencias = 0
        for palabra1 in nombre_s:
            for palabra2 in nombre_r:
                if son_iguales(palabra1,palabra2):
                    coincidencias +=1
        if coincidencias < 2:
            no_coinciden[legajo] = [nombre_s,nombre_r,archivo]
    return no_coinciden

# FUNCIONES AUXILIARES PAGINA
def procesar_oficinas(oficinas):
    res = []

    if len(oficinas) == 0:
        return None
    
    if oficinas.strip().lower() == 'todo':
        return [1,1,1]

    oficinas = oficinas.split(",")
    for ofi in oficinas:
        if len(ofi.split("-")) > 1: # Si es un rango de oficinas, ej: 310-312 = 310,311,312
            rango = ofi.split("-")
            for k in range(int(rango[0]),int(rango[1])+1):
                res.append(k)
        else:
            res.append(ofi)

    # Convertir todo a string
    for i in range(0,len(res)):
        if type(res[i]) is int:
            res[i] = str(res[i])

    return res

def imprimir_lista(lista):
    s = ''
    for i in lista:
        s += "- " + f"{i}" + "\n"
    st.markdown(s)

def imprimir_no_coinciden(dict):
    s = ''
    for key, value in dict.items():
        nombre1 = " ".join(value[0])
        nombre2 = " ".join(value[1])
        archivo = value[2]
        st.write(f"+ Legajo {key}, en sistema {nombre1}, en reporte {nombre2} del archivo {archivo}")
    st.markdown(s)

 
# PAGINA
st.title('Extra! Extra! 🗞️')

st.header('Procedimiento')
with st.expander('Paso 1️⃣: Descargá el archivo de novedades'):
    st.markdown('''
                - Entrar a M@JOR e ir a Informes > Informes de empleados > Empleados por novedad
                - Elegir partición MU
                - Seleccionar Novedades vigentes en el año actual y el mes anterior
                - Elegir variables desde @HRSEXTR1 a @HRSEXTR3
                - Establecer restricciones > Ejecutar
                - ⚠️**Importante**⚠️: exportarlo en el formato "Excel 5.0 (XLS) Tabular" y confirmar "Column headings"
                '''
                )
    
archivos = None
oficinas = None
with st.expander('Paso 2️⃣: Subí todos los archivos, tanto los csvs como el de novedades descargado del sistema'):
    archivos = st.file_uploader('Subí aca abajo los archivos arrastrando o seleccionando en \'Browse files\'',accept_multiple_files=True)
    st.write("Ingresá las oficinas en un listado con comas, si querés indicar rangos de oficinas separalas por un guion. No uses espacios entre cada uno.")
    st.write("Por ejemplo si ingresás '100-102,200,310' es que querés procesar las oficinas 100, 101, 102, 200 y 310")
    st.write("Si escribís la palabra 'TODO' vas a procesar considerando todas las oficinas (aviso: seguramente aparezcan muchas personas no reportadas pero que sí figuran en sistema)")
    oficinas = st.text_area("Escribí las oficinas o 'todo' abajo, y presioná Ctrl+Enter")
oficinas = procesar_oficinas(oficinas)

novedades = None
with st.expander('Paso 3️⃣: Procesar los datos y ver los resultados'):
    if st.button("Procesar") and archivos:
        # Hallar archivo de novedades
        for archivo in archivos:
            if archivo.name.endswith('.xls'): 
                novedades = archivo
                break

        if novedades is None:
            st.error('No subiste el archivo de novedades, hacelo en el paso 2.', icon = '🚨')
        # Procesar
        else:
            resultados_sistema = procesar_novedades_sistema(novedades)
            resultados_reporte = procesar_csvs_oficinas(archivos)
            df,no_estan_en_sistema,no_reportados = comparar_y_armar_df(resultados_sistema,resultados_reporte,oficinas)

            with st.expander('Ver resultados'):
                if len(no_estan_en_sistema) > 0:
                    st.write("1) Estos legajos fueron reportados pero no cargados en el sistema.")
                    with st.expander("Ver más"):
                        imprimir_lista(no_estan_en_sistema)
                else: 
                    st.write("1) Todos los legajos reportados están cargados al sistema.")
                
                if len(no_reportados) > 0:
                    st.write("2) Estos legajos no fueron reportados por las oficinas pero están cargados en el sistema.")
                    with st.expander("Ver más"):
                        imprimir_lista(no_reportados)
                else:  
                    st.write("2) Todos los legajos de las oficinas dadas están reportados.")

                buffer = io.BytesIO()
                if df is not None:
                    st.write("3) Se encontraron las siguientes inconsistencias:")
                    st.write(df)
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df.to_excel(writer, sheet_name='inconsistencias_hrs_extra', index=False)
                    buffer.seek(0)
                    st.download_button(
                        label="Descargar resultados",
                        data=buffer,
                        file_name="inconsistencias_hrs_extra.xlsx",
                        mime="application/vnd.ms-excel",
                        icon=":material/download:",
                    )
                else:
                    st.write("3) No se encontraron inconsistencias entre lo reportado y el sistema.")

                nombres_no_coinciden = comparar_nombres(resultados_sistema,resultados_reporte)

                if len(nombres_no_coinciden) > 0:
                    st.write('Los siguientes nombres no coinciden:')
                    imprimir_no_coinciden(nombres_no_coinciden)




