import numpy as np
import pandas as pd
import streamlit as st
import re



#--------- LECTURA de archivos ---------------------
#PRODUCTIVIDADES -> EXCEL, DECRETO -> CSV

def lectura_archivo_prod(archivo_prod:str):
    '''
    Convierte en DataFrame el excel subido del sistema. Ordena los legajos por orden numérico.
    
    :param archivo_prod: Nombre del archivo .xlsx
    :type archivo_prod: Str
    :return DataFrame. 
    '''
    df_prod = pd.read_excel(archivo_prod)
    df_prod.columns = ["Legajo","Cargo","Apellido y Nombre","Nula1","Leyenda","Inicio","Nula2","Fin","Cantidad","Base","Importe","Porcentaje","Indicativo"]

    return df_prod

def lectura_archivo_dec(archivo_dec: str,decreto, nombre_original):
    '''
    Convierte en DataFrame los archivos subidos con extensión .csv. Si al convertirlo tiene más de 4 columnas, se toman las primeras cuatro.
    
    :param archivo_dec: String. Nombre del archivo .csv
    :type archivo_dec: Str.
    :return DataFrame.
    '''
    df_decreto = pd.read_csv(archivo_dec,header=None)

    if df_decreto.shape[1] > 4: #Si el archivo csv tiene más de 4 columnas, me quedo con las primeras 4

        df_decreto = df_decreto.iloc[:,:4]
    
    df_decreto.columns = ["Legajo", "Nula", "Nula2", "Importe"] #Renombro las columnas
    df_decreto = df_decreto.dropna() # Elimino las filas con algún Nan
    df_decreto["Legajo"] = df_decreto["Legajo"].astype('Int64') #Cambio tipo del Legajo para que tipe con df_prod
    df_decreto["Decreto"] = decreto #Agrego la columna de decreto
    df_decreto["Nombre original"] = nombre_original #Agrego la columna de decreto del nombre original
    df_decreto.sort_values(by = "Legajo") #Ordeno según legajo
    df_decreto["Decreto"] = df_decreto["Decreto"].astype(str)
    df_decreto['Importe'] = df_decreto['Importe'].astype(float).round(2)

    return df_decreto

#--------- LIMPIEZA DE LA COLUMNA DECRETOS -----------

def limpieza_decreto(df: pd.DataFrame) -> dict:
    '''
    Modifica la columna "Leyenda" del dataFrame correspondiente al archivo subido del sistema para que el decreto quede de la forma num/num. Y le agrega
    al dataFrame una columna que sea el nombre original del decreto. Para los decretos donde no se pueda sacar apropiadamente el decreto de la forma num/num, ponemos
    como nombre original : ""
    
    :param df: dataFrame correspondiente a la concatenación de archivos csv
    :type df: pd.DataFrame
    :return dict: diccionario con un decreto, como clave y un conjunto asociado correspondiente a los decretos sin limpiar
    '''

    #TODO sacar el diccionario, al final no lo uso
    patron = r"DTO\.?\s*(\d+/\d+)"

    decretos = []
    cant_prod = df.shape[0]

    for i in range(cant_prod):

        leyenda = df.iloc[i]["Leyenda"]
        
        if pd.isna(leyenda):
            # + 1 por el index de python, 1 por el encabezado de Excel
            decretos.append("")
            #st.write("La fila ", i + 2, " no tiene leyenda detallada." ) TODO me queda ver como imprimir bien esto, podria agregarle una columna de 
            #que archivo fue sacado
            
        else:
            match = re.search(patron, leyenda)
            if match: #Si cuando hacemos el split nos queda de longitud mayor o igual a dos, indexamos en 1 para obtenerlo,
                # en cas contrario, decimos que no podemos extraer el decreto

                decreto_prod = match.group(1)
                decretos.append(leyenda)
                df.loc[i,"Leyenda"] = decreto_prod
                
                if decreto_prod not in dicc_decretos:
                    dicc_decretos[decreto_prod] = set()

                dicc_decretos[decreto_prod].add(leyenda)

            else:
                decretos.append("")

    df.sort_values(by="Legajo")
    df["Leyenda"] = df["Leyenda"].astype(str)
    df['Importe'] = df['Importe'].astype(float).round(2) #Redondeo a dos decimales así puede matchear con mejor exactitud

    df["Nombre original"] = decretos


    

#--------- OBTENGO nombre del decreto, según nombre del archivo-----

def obtener_decreto(nombre_archivo: str) -> str:
    '''
    Dado el nombre del archivo, le quita la extensión .csv. Si el decreto está separado por "-", lo convierte a la forma num/num.
    
    :param nombre_archivo: Description
    :type nombre_archivo: str
    :return: Devuelve  el nombre del archivo sin la extensión.
    :rtype: str
    '''
    decreto = nombre_archivo.split(".")[0] #Con esto saco la extension .csv
    decreto = decreto.split(" ")
    decreto = decreto[0].split("-")[0] + "/" +decreto[0].split("-")[1] #Lo renombro a tipo num/año

    return decreto

#--------- FUNCION PRINCIPAL -------------------------

def comparar(
    df_origen: pd.DataFrame,
    df_destino: pd.DataFrame,
    legajos: list,
    importes: list,
    decretos_originales: list,
    decretos: list,
    decreto: str
) -> None:

    comparacion = df_origen.merge(
        df_destino[["Legajo", "Importe"]],
        on=["Legajo", "Importe"],
        how="left",
        indicator=True
    )

    diferencias = comparacion[comparacion["_merge"] == "left_only"]

    for _, row in diferencias.iterrows():
        legajos.append(row["Legajo"])
        importes.append(row["Importe"])
        decretos_originales.append(row["Nombre original"])
        decretos.append(decreto)




#--------- STREAMLIT -------------------------------

st.title("📝 PRODUCTIVIDADES")

st.divider()

tab1,tab2,tab3 = st.tabs(["Subir archivos", "Sin comparar","Resultados"])

with tab1:

    #Si dan clic al siguiente checkbox la comparación se hará bajo la suposición de que todos los decretos que se encuentran en el excel tienen 
    #un archivo csv corrrespondiente a ese decreto, en caso contrario avisa que ningún csv correspondiente al decreto fue subido
    agree = st.checkbox("Comparar todos los decretos del archivo excel de productividades")

    st.markdown("Subir los archivos de productividades arrojados por el sistema")

    archivos_excel = st.file_uploader("Seleccionar archivo", type = "xls",key = "productividades",accept_multiple_files=True)
    #Acepta multiples, concatenarlos en ese caso (asumimos que las columnas y los nombres son iguales)
    st.markdown("Subir los archivos .csv que se quieren comparar")

    archivos_csv = st.file_uploader("Seleccionar archivo", type = "csv", key = "decreto",accept_multiple_files=True)
    #Acepta multiples, concatenarlos en ese caso(las columnas y sus nombres son iguales porque se 
    #procesan todos en la misma función)

#-------- LECTURA Y LIMPIEZA de los archivos --------

with tab2:

    #bien, tengo los archivos de excel y csv, primero voy a leer los de excel y concatenarlos

    dicc_decretos = dict() #Variable global para guardar en un diccionario el numero de decreto y todas las variantes
    #que aparezcan en el archivo excel de productividades

    if archivos_excel and archivos_csv:

        dfs_excel = []
        dfs_csv = []
        
        #Como se aceptan múltiples archivos, los concatenamos todos en un archivo final, mismo con los archivos csv
        for archivo_excel in archivos_excel:

            df_excel_i = lectura_archivo_prod(archivo_excel)
            limpieza_decreto(df_excel_i)
            dfs_excel.append(df_excel_i)
        
        df_excel_final = pd.concat(dfs_excel,ignore_index = True)
        df_excel_sin_procesar = df_excel_final[df_excel_final["Nombre original"] == ""] #Filtramos el dataFrame por los decretos que no pudimos obtener
        st.write("Esta es la lista de productividades que no pudo ser procesada debido al formato de las leyendas")
        df_excel_sin_procesar = df_excel_sin_procesar.drop(columns = ["Nula1","Inicio","Nula2","Fin","Cantidad","Base","Porcentaje","Indicativo","Nombre original"])
        st.write(df_excel_sin_procesar)
        
        df_excel_final = df_excel_final[df_excel_final["Nombre original"] != ""]


        for archivo_csv in archivos_csv:

            #Acá hago la lectura del decreto y del nombre original del decreto
            decreto_original = archivo_csv.name
            decreto_split = decreto_original.split(".")
            decreto_original = decreto_split[0] #Este es el nombre original del arhivo, le sacamos la extensión .csv
            decreto_original_split = decreto_original.split(" ")#Este es el numero de la forma num-num
            decreto_con_guion = decreto_original_split[0] #Me quedo con el numero de decreto nada más
            decreto_con_guion = decreto_con_guion.split("-")
            decreto = decreto_con_guion[0] + "/" + decreto_con_guion[1]
            

            df_csv_i = lectura_archivo_dec(archivo_csv,decreto,decreto_original)
            dfs_csv.append(df_csv_i)

        df_csv_final = pd.concat(dfs_csv, ignore_index = True)
        #st.write(df_csv_final)

        dfs_dif_excel = []
        dfs_dif_csv = []
        diferencias = set()
        legajos_comp_excel = [] #Lista de legajos con la comparación de excel a csv
        importes_comp_excel = [] #Lista de importes con la comparación de excel a csv
        legajos_comp_csv = [] #Lista de legajos con la comparación de csv a excel
        importes_comp_csv = [] #Lista de importes con la comparación de csv a excel
        decreto_original_excel = []
        decreto_original_csv = []
        decretos_comp_excel = []
        decretos_comp_csv = []
        no_existe_csv = []


        if agree:
            #Si se piden procesar todos los decretos procedemos de la siguiente manera:
            #obtenemos todos los decretos únicos del excel e iteramos sobre ellos para  hacer la comparación de excel a csv y de csv a excel

            decretos_unicos_excel = df_excel_final["Leyenda"].unique()

            for decreto_excel in decretos_unicos_excel:
            
                df_csv_decreto = df_csv_final[df_csv_final["Decreto"] == decreto_excel]
                df_excel_decreto = df_excel_final[df_excel_final["Leyenda"] == decreto_excel]

                if df_csv_decreto.shape[0] == 0:
                    no_existe_csv.append(decreto_excel)
                    #st.write(f"No fue cargado ningún csv correspondiente al decreto {decreto_excel}")

                else:

                    comparar(df_excel_decreto, df_csv_decreto,legajos_comp_excel, importes_comp_excel,decreto_original_excel,decretos_comp_excel,decreto_excel) #Comparación Excel a CSV

                    comparar(df_csv_decreto,df_excel_decreto,legajos_comp_csv,importes_comp_csv,decreto_original_csv,decretos_comp_csv,decreto_excel) #Comparación CSV a Excel

            df_diferencias_excel = pd.DataFrame({"Legajo": legajos_comp_excel, "Importe": importes_comp_excel,"Decreto":decretos_comp_excel,"Decreto original":decreto_original_excel}) 
            df_diferencias_csv = pd.DataFrame({"Legajo": legajos_comp_csv, "Importe": importes_comp_csv,"Decreto":decretos_comp_csv,"Decreto original":decreto_original_csv})

            st.write("Estos son los legajos e importes que no pudieron ser matcheados debido a que su CSV no fue subido:")

            df_no_subidos = df_excel_final[df_excel_final["Leyenda"].isin(no_existe_csv)]
            df_no_subidos = df_no_subidos.drop(columns = ["Nula1","Inicio","Nula2","Fin","Cantidad","Base","Porcentaje","Indicativo","Nombre original"])

            st.write(df_no_subidos)
                    

                    #df_diferencias_csv = pd.DataFrame({"Legajo": legajos_comp_csv, "Importe" : importes_comp_csv,"Decreto original":decreto_original_csv})

                    #df_diferencias_excel = pd.DataFrame({"Legajo": legajos_comp_excel, "Importe" : importes_comp_excel,"Decreto original":decreto_original_excel})

        else:

            decretos_unicos_csv = df_csv_final["Decreto"].unique()

            for decreto_csv in decretos_unicos_csv:
                
                df_csv_decreto = df_csv_final[df_csv_final["Decreto"] == decreto_csv]
                df_excel_decreto = df_excel_final[df_excel_final["Leyenda"] == decreto_csv]

                comparar(df_excel_decreto, df_csv_decreto,legajos_comp_excel, importes_comp_excel,decreto_original_excel,decretos_comp_excel,decreto_csv)

                comparar(df_csv_decreto,df_excel_decreto,legajos_comp_csv,importes_comp_csv,decreto_original_csv,decretos_comp_csv,decreto_csv)
                
            df_diferencias_excel = pd.DataFrame({"Legajo": legajos_comp_excel, "Importe": importes_comp_excel,"Decreto":decretos_comp_excel,"Decreto original":decreto_original_excel}) 
            df_diferencias_csv = pd.DataFrame({"Legajo": legajos_comp_csv, "Importe": importes_comp_csv,"Decreto":decretos_comp_csv,"Decreto original":decreto_original_csv})

        

        df_diferencias_excel.columns = ["Legajo","Importe","Decreto","Leyenda original"]
        df_diferencias_csv.columns = ["Legajo","Importe","Decreto","Nombre archivo original"]

        with tab3:

            diferencias = pd.concat([df_diferencias_excel["Decreto"], df_diferencias_csv["Decreto"]]).unique().tolist()

            if(len(diferencias)>0):

                tabs = st.tabs(diferencias)

                for i in range(len(diferencias)):

                    with tabs[i]:

                        df_diferencias_excel_dec = df_diferencias_excel[df_diferencias_excel["Decreto"] == diferencias[i]]
                        df_diferencias_csv_dec = df_diferencias_csv[df_diferencias_csv["Decreto"] == diferencias[i]]

                        if df_diferencias_excel_dec.shape[0] != 0:
                            st.write("Los siguientes importes de la planilla del sistema no fueron encontrados en ninguno de los CSVs subidos: ")
                            st.write(df_diferencias_excel_dec)
                        
                        if df_diferencias_csv_dec.shape[0] != 0:
                            st.write("Los siguientes importes de los CSVs subidos no fueron encontrados en ninguna de las planillas del sistema subidas: ")
                            st.write(df_diferencias_csv_dec)

            else:
                st.markdown("No se encontraron diferencias")



       
