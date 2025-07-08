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
            df_reportado = pd.read_csv(archivo)
            ofi = archivo.name.strip(".csv")
            oficinas.append({'nro_ofi': ofi,
                            'tam_ofi': len(df_reportado),
                            'legajos': df_reportado['LEGAJO'].tolist(),
                            'nombres': df_reportado['NOMBRE Y APELLIDO '].tolist(),
                            'hs_tip1': df_reportado['HORAS NORMALES'].tolist(),
                            'hs_tip2': df_reportado['HORAS 50'].tolist(),
                            'hs_tip3': df_reportado['HORAS 100'].tolist()})

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

def comparar_y_armar_df(resultados_sistema,resultados_reporte):

    no_reportados = [] #legajos de quienes están cargados en sistema pero no fueron reportados
    no_coinciden = {} #legajos de quienes no coinciden lo reportado y lo cargado en sistema
    no_cargados = [] #legajos de quienes fueron reportados pero no cargados en sistema
    coinciden = {}

    # está en reporte pero no en sistema:
    for legajo in resultados_reporte.keys():
        if legajo in resultados_sistema.keys():
            continue
        else:
            no_cargados.append(legajo)

    # hacer comparacion
    for legajo in resultados_sistema.keys():
        # si el legajo no está en reporte
        if legajo not in resultados_reporte.keys():
            no_reportados.append(legajo)
        # si el legajo está en reporte
        else:
            # si coinciden
            if resultados_sistema[legajo][2:5] == resultados_reporte[legajo][2:5]:
                coinciden[legajo] = {'sistema':resultados_sistema[legajo],'reporte': resultados_reporte[legajo]}
            # si no coinciden
            else:
                no_coinciden[legajo] = {'sistema':resultados_sistema[legajo],'reporte': resultados_reporte[legajo]}

    df = pd.DataFrame(no_coinciden.values(),index=no_coinciden.keys())


    def expand_column(col, prefix):
        return pd.DataFrame(col.tolist(), 
                            index=col.index, 
                            columns=[f"Nombre en {prefix}", f"Oficina en {prefix}", f"H.E. normales en {prefix}", f"H.E. al 50 en {prefix}", f"H.E. al 100 en {prefix}"])

    df_sistema_expandido = expand_column(df["sistema"], "sistema")
    df_reporte_expandido = expand_column(df["reporte"], "reporte")

    df_final = pd.concat([df_sistema_expandido, df_reporte_expandido], axis=1)
    return df_final

#################################

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

with st.expander('Paso 2️⃣: Subí todos los archivos, tanto los csvs como el de novedades descargado del sistema'):

    archivos = st.file_uploader('Subí aca abajo los archivos arrastrando o seleccionando en \'Browse files\'',accept_multiple_files=True)
    novedades = None
    if archivos:

        for archivo in archivos:
            if archivo.name.endswith('.xls'): 
                novedades = archivo
                break

        resultados_sistema = procesar_novedades_sistema(novedades)
        
        resultados_reporte = procesar_csvs_oficinas(archivos)
        
        df = comparar_y_armar_df(resultados_sistema,resultados_reporte)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='incongruencias_hrs_extra')
        buffer.seek(0)
        with st.expander('Ver resultados'):
            st.dataframe(df)
            
        st.download_button(
            label="Descargar resultados",
            data=buffer,
            file_name="incongruencias_hrs_extra.xlsx",
            mime="application/vnd.ms-excel",
            icon=":material/download:",
        )
    
    
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
