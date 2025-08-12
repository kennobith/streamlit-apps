import pandas as pd
import os
import streamlit as st
import io
import streamlit.components.v1 as components

def procesar_novedades_sistema(datos):

    datos = pd.read_excel(datos, engine="xlrd")

    # dividimos columnas que tienen doble información
    datos[['legajo','nro_cargo']] = datos['LEGAJO'].str.split('-',n=1, expand=True)
    datos[['año','oficina']] = datos['OFICINA'].str.split('-',n=1,expand=True)

    # cambiamos nombres de columnas
    datos['nombre_completo'] = datos['APELLIDO Y NOMBRE']
    datos['valor_hora_extra'] = datos['VALOR']
    datos['tipo_hora_extra'] = datos['DESCRIPCIÓN']
    datos['oficina'] = datos['oficina'].str.strip()

    # armamos df_con lo que nos interesa
    # df_s == df_sistema
    df_s = pd.DataFrame(
                    {
                        'legajo': datos['legajo'],
                        'nombre_completo': datos['nombre_completo'],
                        'oficina': datos['oficina'],
                        'tipo_hora_extra': datos['tipo_hora_extra'],
                        'valor_hora_extra': datos['valor_hora_extra'],
                    }
    )

    df_s.replace(to_replace=['@HRSEXTR1','@HRSEXTR2','@HRSEXTR3'],value=[1,2,3],inplace=True)

    df_sistema = pd.pivot_table(
        df_s,
        index = ['legajo','nombre_completo','oficina'],
        columns = ['tipo_hora_extra'],
        values = ['valor_hora_extra'],
        fill_value = 0
    )   

    # Aplanar columnas
    df_sistema.columns = [f'{col[0]}_{col[1]}' for col in df_sistema.columns]
    df_sistema = df_sistema.reset_index()
    df_sistema['legajo'] = pd.to_numeric(df_sistema['legajo'])
    df_sistema['valor_hora_extra_1'] = pd.to_numeric(df_sistema['valor_hora_extra_1'],downcast='integer')
    df_sistema['valor_hora_extra_2'] = pd.to_numeric(df_sistema['valor_hora_extra_2'],downcast='integer')
    df_sistema['valor_hora_extra_3'] = pd.to_numeric(df_sistema['valor_hora_extra_3'],downcast='integer')
    df_sistema = df_sistema.sort_values(by = 'legajo')


    # lo convertimos a un dict
    resultados_sistema = df_sistema.set_index('legajo').T.to_dict('list')
    return resultados_sistema

# esto nos va a servir para aplanar unas listas
def flatten(oficinas,clave):
    lista = [
        x 
        for i in range(len(oficinas))
        for x in oficinas[i][clave]
        ]
    
    return lista

def procesar_csvs_oficinas(archivos):
    oficinas = []

    for archivo in archivos:
        if archivo.name.endswith(".csv"): 

            df_reportado = pd.read_csv(archivo, encoding="latin1",skip_blank_lines=True)

            nombre_ultima_columna  = df_reportado.columns[-1]
            nombre_primera_columna = df_reportado.columns[0]
            if df_reportado[nombre_ultima_columna].isnull().all(): # si en la ultima columna todos los elementos son nulos
                df_reportado = df_reportado.drop(columns=[nombre_ultima_columna])
            #quitar filas nulas de tipo ,0,,,,
            df_reportado = df_reportado[~(df_reportado.drop(df_reportado.columns[1], axis=1).isna().all(axis=1))]
            df_reportado[nombre_primera_columna] = df_reportado[nombre_primera_columna].astype(str).str.replace(' ', '', regex=False)
            df_reportado = df_reportado.fillna(0)
            ofi = archivo.name.strip(".csv")
            columnas_nombres = df_reportado.columns.tolist()
            oficinas.append({'nro_ofi': ofi,
                             'tam_ofi': len(df_reportado),
                             'legajos': df_reportado[columnas_nombres[0]].tolist(),
                             'nombres': df_reportado[columnas_nombres[5]].tolist(),
                             'hs_tip1': df_reportado[columnas_nombres[2]].tolist(),
                             'hs_tip2': df_reportado[columnas_nombres[3]].tolist(),
                             'hs_tip3': df_reportado[columnas_nombres[4]].tolist()})

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
    df_r = pd.DataFrame(
        {   
            'legajo': legajos,
            'nombre_completo': [nombre.upper() for nombre in nombres],
            'oficinas': oficinas_todas,
            'valor_hora_extra_1': hs_tip1,
            'valor_hora_extra_2': hs_tip2,
            'valor_hora_extra_3': hs_tip3
        }
    )

    df_r = df_r.sort_values('legajo')

    resultados_reporte = df_r.set_index('legajo').T.to_dict('list')

    return resultados_reporte

def expand_column(col, prefix):
        return pd.DataFrame(col.tolist(), 
                            index=col.index, 
                            columns=[f"Nombre en {prefix}", f"Oficina en {prefix}", f"H.E. normales en {prefix}", f"H.E. al 50 en {prefix}", f"H.E. al 100 en {prefix}"])

def no_fue_reportado_y_esta_en_oficinas(resultados_sistema,legajo,oficinas):
    return resultados_sistema[legajo][1] in oficinas

def comparar_y_armar_df(resultados_sistema,resultados_reporte,oficinas):

    no_reportados = []
    no_coinciden = {} #legajos de quienes no coinciden lo reportado y lo cargado en sistema
    no_estan_en_sistema = [] #legajos de quienes fueron reportados pero no cargados en sistema

    # está en reporte pero no en sistema:
    for legajo in resultados_reporte.keys():
        if legajo in resultados_sistema.keys():
            continue
        else:
            no_estan_en_sistema.append(f'Legajo: {legajo} - Oficina: {resultados_reporte[legajo][1]}')

    # hacer comparacion
    for legajo in resultados_sistema.keys():
        # si el legajo está en reporte
        if legajo in resultados_reporte.keys():
            if resultados_sistema[legajo][2:5] == resultados_reporte[legajo][2:5]:
                continue
            # si no coinciden
            else:
                no_coinciden[legajo] = {'sistema':resultados_sistema[legajo],'reporte': resultados_reporte[legajo]}
        # si no esta en el reporte
        else:
            #si esta en las oficinas dadas
            if no_fue_reportado_y_esta_en_oficinas(resultados_sistema,legajo,oficinas):
                no_reportados.append(f'Legajo: {legajo} - Oficina: {resultados_sistema[legajo][1]}')

    df = pd.DataFrame(no_coinciden.values(),index=no_coinciden.keys())

    if df.empty:
        return None,no_estan_en_sistema,no_reportados
    else:
        df_sistema_expandido = expand_column(df['sistema'], 'sistema')
        df_reporte_expandido = expand_column(df['reporte'], 'reporte')

        df_final = pd.concat([df_sistema_expandido, df_reporte_expandido], axis=1)
        return df_final,no_estan_en_sistema,no_reportados

#################################
def procesar_oficinas(oficinas):
    oficinas = oficinas.split("\n")
    return [ofi.strip() for ofi in oficinas]

def imprimir_lista(lista):
    s = ''
    for i in lista:
        s += "- " + f"{i}" + "\n"
    st.markdown(s)

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
novedades = None
oficinas = None
with st.expander('Paso 2️⃣: Subí todos los archivos, tanto los csvs como el de novedades descargado del sistema'):

    archivos = st.file_uploader('Subí aca abajo los archivos arrastrando o seleccionando en \'Browse files\'',accept_multiple_files=True)

    oficinas = st.text_area(
        "Escribí las oficinas en una lista, es decir en cada línea va un número de oficina, "
        "todas las que incluyan los csvs a procesar. \n Cuando termines, presioná Ctrl + Enter"
    )

oficinas = procesar_oficinas(oficinas)

with st.expander('Paso 3️⃣: Procesar los datos y ver los resultados'):

    if st.button("Procesar"):

        if archivos:

            for archivo in archivos:
                if archivo.name.endswith('.xls'): 
                    novedades = archivo
                    break

            if novedades is None:
                st.error('No subiste el archivo de novedades, hacelo en el paso 2.', icon = '🚨')
            if oficinas is None:
                st.error('No escribiste ninguna oficina, hacelo en el paso 2.', icon = '🚨')
            else:
                resultados_sistema = procesar_novedades_sistema(novedades)
                
                resultados_reporte = procesar_csvs_oficinas(archivos)
                
                df,no_estan_en_sistema,no_reportados = comparar_y_armar_df(resultados_sistema,resultados_reporte,oficinas)

                with st.expander('Ver resultados'):
                    if len(no_estan_en_sistema) > 0:
                        st.write("1) Estos legajos fueron reportados pero no cargados en el sistema.")
                        with st.expander("Ver mas"):
                            imprimir_lista(no_estan_en_sistema)
                    else: 
                        st.write("1) Todos los legajos reportados están cargados al sistema.")
                    
                    if len(no_reportados) > 0:
                        st.write("2) Estos legajos no fueron reportados por las oficinas pero están cargados en el sistema.")
                        with st.expander("Ver mas"):
                            imprimir_lista(no_reportados)
                    else: 
                        st.write("2) Todos los legajos de las oficinas dadas están reportados.")

                    buffer = io.BytesIO()
                    if df is not None and not df.empty:
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

                        
    
hvar = """
    <script>
        var elements = window.parent.document.querySelectorAll('.streamlit-expanderHeader');
        elements[0].style.color = 'rgba(83, 36, 118, 1)';
        elements[0].style.fontFamily = 'Didot';
        elements[0].style.fontSize = 'x-large';
        elements[0].style.fontWeight = 'bold';
    </script>
"""


components.html(hvar, height=0, width=0)



