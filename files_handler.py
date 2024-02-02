import openpyxl
import pandas as pd

def read_xlsx_file(uploaded_file):
    # Read the file into a pandas DataFrame
    # Read the file into a pandas DataFrame
    df = pd.read_excel(uploaded_file, engine='openpyxl')

    # Get the file name without the extension and convert it to lowercase
    file_name = uploaded_file.name.rsplit('.', 1)[0].lower()

    # Return a dictionary where the key is the file name and the value is the DataFrame
    return {f"{file_name}_df": df}


def filter_dataframe_vida(df, year, riesgo, vida=False):
    # Initialize an empty list to store the selected DataFrames
    selected_dfs = []

    # Flags to track whether "TOTAL:" and "Contrato:" are found
    total_found = False
    contrato_found = False

    # Iterate through the DataFrame
    for i in range(len(df)):
        # Check conditions for each row
        # Qué secciones calcular: en este caso se excluyen VIDA y CAUCION
        if (
            df.iloc[i, 2] == "TOTAL:" and
            df.iloc[i, 3] == year and
            (isinstance(df.iloc[i, 7], str) and (df.iloc[i, 7]).split()[0] == "VIDA") and
            (isinstance(df.iloc[i, 7], str) and not (df.iloc[i, 7]).split()[0] == "CAUCION") and
            df.iloc[i, 4] == riesgo
        ):
            total_found = True
            contrato_found = False
        elif df.iloc[i, 2] == "Contrato:":
            contrato_found = True
            total_found = False
    
        # Select rows between "TOTAL:" and "Contrato:"
        if total_found and not contrato_found:
            selected_dfs.append(df.iloc[i])

    # Check if there are selected DataFrames
    if selected_dfs:
        # Concatenate the list of selected DataFrames into a single DataFrame
        result_df = pd.concat(selected_dfs, axis=1).T

        # Filter out rows that start with specific words in the first column
        result_df = result_df[~result_df.iloc[:, 0].astype(str).str.startswith(('Póliza', 'Operador:'))]

        # Filter out rows that in the third column start with specific words
        result_df = result_df[~result_df.iloc[:, 2].astype(str).str.startswith(('REGIONAL', 'Listado', 'Emisiones', 'Anulaciones'))]

        # Remove columns with all NaN values
        result_df = result_df.dropna(axis=1, how='all')

        # Change the column names for letters. Example: "Unamed: 0" to "A"
        result_df.columns = [chr(65 + i) for i in range(len(result_df.columns))]

        return result_df
    else:
        print("No matching rows found.")
        return None
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_filtered_df = filter_dataframe_vida(df, 2023, 'EXCEDENTE')
if your_filtered_df is not None:
    print(your_filtered_df)
"""


def filter_dataframe_patrimoniales(df, year, riesgo, vida=False):
    # Initialize an empty list to store the selected DataFrames
    selected_dfs = []

    # Flags to track whether "TOTAL:" and "Contrato:" are found
    total_found = False
    contrato_found = False

    # Iterate through the DataFrame
    for i in range(len(df)):
        # Check conditions for each row
        # Qué secciones calcular: en este caso se excluyen VIDA y CAUCION
        if (
            df.iloc[i, 2] == "TOTAL:" and
            df.iloc[i, 3] == year and
            (isinstance(df.iloc[i, 7], str) and not (df.iloc[i, 7]).split()[0] == "VIDA") and
            (isinstance(df.iloc[i, 7], str) and not (df.iloc[i, 7]).split()[0] == "CAUCION") and
            df.iloc[i, 4] == riesgo
        ):
            total_found = True
            contrato_found = False
        elif df.iloc[i, 2] == "Contrato:":
            contrato_found = True
            total_found = False
    
        # Select rows between "TOTAL:" and "Contrato:"
        if total_found and not contrato_found:
            selected_dfs.append(df.iloc[i])

    # Check if there are selected DataFrames
    if selected_dfs:
        # Concatenate the list of selected DataFrames into a single DataFrame
        result_df = pd.concat(selected_dfs, axis=1).T

        # Filter out rows that start with specific words in the first column
        result_df = result_df[~result_df.iloc[:, 0].astype(str).str.startswith(('Póliza', 'Operador:'))]

        # Filter out rows that in the third column start with specific words
        result_df = result_df[~result_df.iloc[:, 2].astype(str).str.startswith(('REGIONAL', 'Listado', 'Emisiones', 'Anulaciones'))]

        # Remove columns with all NaN values
        result_df = result_df.dropna(axis=1, how='all')

        # Change the column names for letters. Example: "Unamed: 0" to "A"
        result_df.columns = [chr(65 + i) for i in range(len(result_df.columns))]

        return result_df
    else:
        print("No matching rows found.")
        return None
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_filtered_df = filter_dataframe_patrimoniales(df, 2023, 'EXCEDENTE')
if your_filtered_df is not None:
    print(your_filtered_df)
"""


def filter_recuperos_patrimoniales(df, year, riesgo, vida=False):
    # Initialize an empty list to store the selected DataFrames
    selected_dfs = []

    # Flags to track whether "TOTAL:" and "Contrato:" are found
    total_found = False
    contrato_found = False

    # Iterate through the DataFrame
    for i in range(len(df)):
        # Check conditions for each row
        # Qué secciones calcular: en este caso se excluyen VIDA y CAUCION
        if (
            df.iloc[i, 0] == "Total Contrato:" and
            df.iloc[i, 2] == year and
            not (df.iloc[i, 9]).split()[0] == "VIDA" and
            not (df.iloc[i, 9]).split()[0] == "CAUCION" and
            df.iloc[i, 3] == riesgo
        ):
            total_found = True
            contrato_found = False
        elif df.iloc[i, 0] == "Contrato:" or df.iloc[i, 1] == "Resumen":
            contrato_found = True
            total_found = False
    
        # Select rows between "TOTAL:" and "Contrato:"
        if total_found and not contrato_found:
            selected_dfs.append(df.iloc[i])

    # Check if there are selected DataFrames
    if selected_dfs:
        # Concatenate the list of selected DataFrames into a single DataFrame
        result_df = pd.concat(selected_dfs, axis=1).T

        # Filter out rows that start with specific words in the first column
        result_df = result_df[~result_df.iloc[:, 0].astype(str).str.startswith(('Póliza', 'Operador:'))]

        # Filter out rows that in the third column start with specific words
        result_df = result_df[~result_df.iloc[:, 2].astype(str).str.startswith(('REGIONAL', 'Listado', 'Emisiones', 'Anulaciones'))]

        # Remove columns with all NaN values
        result_df = result_df.dropna(axis=1, how='all')

        # Change the column names for letters. Example: "Unamed: 0" to "A"
        result_df.columns = [chr(65 + i) for i in range(len(result_df.columns))]

        return result_df
    else:
        print("No matching rows found.")
        return None
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_filtered_df = filter_recuperos_patrimoniales(df, 2023, 'EXCEDENTE')
if your_filtered_df is not None:
    print(your_filtered_df)
"""


def filter_recuperos_vida(df, year, riesgo, vida=False):
    # Initialize an empty list to store the selected DataFrames
    selected_dfs = []

    # Flags to track whether "TOTAL:" and "Contrato:" are found
    total_found = False
    contrato_found = False

    # Iterate through the DataFrame
    for i in range(len(df)):
        # Check conditions for each row
        # Qué secciones calcular: en este caso se excluyen VIDA y CAUCION
        if (
            df.iloc[i, 0] == "Total Contrato:" and
            df.iloc[i, 2] == year and
            (df.iloc[i, 9]).split()[0] == "VIDA" and
            not (df.iloc[i, 9]).split()[0] == "CAUCION" and
            df.iloc[i, 3] == riesgo
        ):
            total_found = True
            contrato_found = False
        elif df.iloc[i, 0] == "Contrato:" or df.iloc[i, 1] == "Resumen":
            contrato_found = True
            total_found = False
    
        # Select rows between "TOTAL:" and "Contrato:"
        if total_found and not contrato_found:
            selected_dfs.append(df.iloc[i])

    # Check if there are selected DataFrames
    if selected_dfs:
        # Concatenate the list of selected DataFrames into a single DataFrame
        result_df = pd.concat(selected_dfs, axis=1).T

        # Filter out rows that start with specific words in the first column
        result_df = result_df[~result_df.iloc[:, 0].astype(str).str.startswith(('Póliza', 'Operador:'))]

        # Filter out rows that in the third column start with specific words
        result_df = result_df[~result_df.iloc[:, 2].astype(str).str.startswith(('REGIONAL', 'Listado', 'Emisiones', 'Anulaciones'))]

        # Remove columns with all NaN values
        result_df = result_df.dropna(axis=1, how='all')

        # Change the column names for letters. Example: "Unamed: 0" to "A"
        result_df.columns = [chr(65 + i) for i in range(len(result_df.columns))]

        return result_df
    else:
        print("No matching rows found.")
        return None
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_filtered_df = filter_recuperos_patrimoniales(df, 2023, 'EXCEDENTE')
if your_filtered_df is not None:
    print(your_filtered_df)
"""


def remove_totals(df):
    try:
        # Remove all the rows that start with "TOTAL:" in the second column
        result_df = df[~df.iloc[:, 1].astype(str).str.startswith('TOTAL:')]
        
        # Remove all empty rows
        result_df = result_df.dropna(how='all')

        return result_df
    except Exception as e:
        print("No matches found.")
        return None

"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
your_processed_df = process_dataframe(your_dataframe)
print(your_processed_df)
"""


def remove_totals_recuperos(df):
    try:
        # Remove all the rows that start with "Total Contrato:" in the first column
        result_df = df[~df.iloc[:, 0].astype(str).str.startswith('Total Contrato:')]
        
        # Remove all empty rows
        result_df = result_df.dropna(how='all')

        return result_df
    except Exception as e:
        print("No matches found.")
        return None
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
your_processed_df = remove_totals_recuperos(your_dataframe)
print(your_processed_df)
"""


def calculate_sums(df):
    try:
        # Suma los valores numéricos de las columnas H, la columna L y la columna M,
        # devuelve los valores y los almacena en variables "prima", "importe_comision" y "a_favor" respectivamente
        prima = pd.to_numeric(df['H'], errors='coerce').fillna(0).sum()
        importe_comision = pd.to_numeric(df['L'], errors='coerce').fillna(0).sum()
        a_favor = pd.to_numeric(df['M'], errors='coerce').fillna(0).sum()

        # Return the calculated values as a dictionary
        return {
            'prima': prima,
            'importe_comision': importe_comision,
            'a_favor': a_favor
        }
    except Exception as e:
        print("No matches found. Values set to 0")
        return {
            'prima': 0,
            'importe_comision': 0,
            'a_favor': 0
        }


def calculate_sums_recuperos(df):
    try:
        # Suma los valores numéricos de las columnas J, la columna M y la columna O,
        # devuelve los valores y los almacena en variables "indemnizacion", "gastos" y "importe_total" respectivamente
        indemnizacion = pd.to_numeric(df['J'], errors='coerce').fillna(0).sum()
        gastos = pd.to_numeric(df['M'], errors='coerce').fillna(0).sum()
        importe_total = pd.to_numeric(df['O'], errors='coerce').fillna(0).sum()

        # Return the calculated values as a dictionary
        return {
            'indemnizacion': indemnizacion,
            'gastos': gastos,
            'importe_total': importe_total
        }
    except Exception as e:
        print("No matches found. Values set to 0")
        return {
            'indemnizacion': 0,
            'gastos': 0,
            'importe_total': 0
        }
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
your_processed_df = process_dataframe(your_dataframe)
prima_value, importe_comision_value, a_favor_value = calculate_sums(your_processed_df)
print(prima_value, importe_comision_value, a_favor_value)
"""


def process_and_sum_vida(df, year, riesgo):
    # Step 1: Filter the DataFrame
    filtered_df = filter_dataframe_vida(df, year, riesgo)

    # Step 2: Generate reaseguros dictionary
    reaseguros_dict = create_reaseguros_dict(filtered_df)
    
    # Step 3: Remove totals from the filtered DataFrame
    processed_df = remove_totals(filtered_df)
    
    # Step 4: Calculate the sums
    sums_result = calculate_sums(processed_df)
    
    return sums_result, reaseguros_dict
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_sums_result = process_and_sum_vida(df, 2023, 'EXCEDENTE')
print("Prima:", your_sums_result[0])
print("Importe Comision:", your_sums_result[1])
print("A Favor:", your_sums_result[2])
"""


def process_and_sum_patrimoniales(df, year, riesgo):
    # Step 1: Filter the DataFrame
    filtered_df = filter_dataframe_patrimoniales(df, year, riesgo)

    # Step 2: Generate reaseguros dictionary
    reaseguros_dict = create_reaseguros_dict(filtered_df)
    
    # Step 2: Remove totals from the filtered DataFrame
    processed_df = remove_totals(filtered_df)
    
    # Step 3: Calculate the sums
    sums_result = calculate_sums(processed_df)
    
    return sums_result, reaseguros_dict
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_sums_result = process_and_sum_patrimoniales(df, 2023, 'EXCEDENTE')
print("Prima:", your_sums_result[0])
print("Importe Comision:", your_sums_result[1])
print("A Favor:", your_sums_result[2])
"""


def process_sum_recuperos_patrimoniales(df, year, riesgo):
    # Step 1: Filter the DataFrame
    filtered_df = filter_recuperos_patrimoniales(df, year, riesgo)

    # Step 2: Generate reaseguros dictionary
    reaseguros_dict = create_reaseguros_dict_recuperos(filtered_df)
    
    # Step 3: Remove totals from the filtered DataFrame
    processed_df = remove_totals_recuperos(filtered_df)
    
    # Step 4: Calculate the sums
    sums_result = calculate_sums_recuperos(processed_df)
    
    return sums_result, reaseguros_dict

"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_sums_result = process_sum_recuperos_patrimoniales(df, 2023, 'EXCEDENTE')
print("Indemnizacion:", your_sums_result[0])
print("Gastos:", your_sums_result[1])
print("Importe Total:", your_sums_result[2])
"""


def process_sum_recuperos_vida(df, year, riesgo):
    # Step 1: Filter the DataFrame
    filtered_df = filter_recuperos_vida(df, year, riesgo)

    # Step 2: Generate reaseguros dictionary
    reaseguros_dict = create_reaseguros_dict_recuperos(filtered_df)
    
    # Step 3: Remove totals from the filtered DataFrame
    processed_df = remove_totals_recuperos(filtered_df)
    
    # Step 4: Calculate the sums
    sums_result = calculate_sums_recuperos(processed_df)
    
    return sums_result, reaseguros_dict
"""
# Example usage
# Replace 'your_dataframe' with the actual DataFrame variable name
# Replace 2022 and 'EXCEDENTE' with the desired values for year and riesgo
your_sums_result = process_sum_recuperos_vida(df, 2023, 'EXCEDENTE')
print("Indemnizacion:", your_sums_result[0])
print("Gastos:", your_sums_result[1])
print("Importe Total:", your_sums_result[2])
"""

# Deprecated
def create_reaseguros_dict_2(df):
    if df is None:
        print("Error: df is None")
        return None
    
    # Elimina las filas que tienen la palabra "Reasegurador" en la primera columna
    filtered_df = df[df.iloc[:, 0] != 'Reasegurador']

    # Obtén los valores únicos de la primera columna
    unique_values_A = filtered_df.iloc[:, 0].unique()

    # Crea un diccionario con valores de la primera columna como keys y valores de la quinta columna como values
    result_dict = {}
    for value in unique_values_A:
        # Verifica si la fila no está vacía antes de acceder a la quinta columna
        if not filtered_df.loc[filtered_df.iloc[:, 0] == value, filtered_df.columns[5]].empty:
            result_dict[value] = filtered_df.loc[filtered_df.iloc[:, 0] == value, filtered_df.columns[5]].values[0]

    return result_dict

def create_reaseguros_dict(df):
    if df is None:
        print("Error: df is None")
        return None
    # Elimina las filas que tienen la palabra "Reasegurador" en la primera columna
    filtered_df = df[df.iloc[:, 0] != 'Reasegurador']

    # Obtén los valores únicos de la primera columna
    unique_values_A = filtered_df.iloc[:, 0].unique()

    # Crea un diccionario con valores de la primera columna como keys y sumas de las columnas 8, 12, 13 como values
    result_dict = {}
    for value in unique_values_A:
        if pd.notna(value):  # Check if the value is not null or NaN
            participacion = filtered_df.loc[filtered_df.iloc[:, 0] == value, filtered_df.columns[5]].values[0]
            prima = filtered_df.loc[filtered_df.iloc[:, 0] == value, filtered_df.columns[7]].sum()
            importe_comision = filtered_df.loc[filtered_df.iloc[:, 0] == value, filtered_df.columns[11]].sum()
            a_favor = filtered_df.loc[filtered_df.iloc[:, 0] == value, filtered_df.columns[12]].sum()
            result_dict[value] = {'participacion': participacion,'prima': prima, 'importe_comision': importe_comision, 'a_favor': a_favor}

    return result_dict

"""
Output example:
 dict = {
'MS AMLIN AG': 	{ 'participacion': 1087343123,
        'prima': 1087343123,
		'importe_comision': 337076291, 
		'a_favor': 750266832}, 

'SCOR REINSURANCE COMPANY': 	{'participacion': 1087343123,
                'prima': 434937249, 
				'importe_comision': 108734290, 
				'a_favor': 326202959}, 

'KOREAN REINSURANCE COMPANY': 	{'participacion': 1087343123,
                'prima': 326202936, 
				'importe_comision': 110908972, 
				'a_favor': 215293964}, 

'MAPFRE RE COMPAÑIA DE REASEGUROS S.A.': {'participacion': 1087343123,
                    'prima': 217468624, 
					'importe_comision': 69589944, 
					'a_favor': 147878680}, 

'REASEGURADORA PATRIA S.A.': 	{'participacion': 1087343123,
                'prima': 108734311, 
				'importe_comision': 30445601, 
				'a_favor': 78288710}
    }
"""

def create_reaseguros_dict_recuperos(df):
    if df is None:
        print("Error: df is None")
        return None
    # Elimina las filas que tienen la palabra "Reasegurador" en la primera columna
    filtered_df = df[df.iloc[:, 4] != 'Reasegurador']

    # Obtén los valores únicos de la primera columna
    unique_values_A = filtered_df.iloc[:, 4].unique()


    result_dict = {}
    for value in unique_values_A:
        if pd.notna(value):  # Check if the value is not null or NaN
            participacion = filtered_df.loc[filtered_df.iloc[:, 4] == value, filtered_df.columns[6]].values[0]
            indemnizacion = filtered_df.loc[filtered_df.iloc[:, 4] == value, filtered_df.columns[9]].sum()
            importe_gastos = filtered_df.loc[filtered_df.iloc[:, 4] == value, filtered_df.columns[12]].sum()
            total = filtered_df.loc[filtered_df.iloc[:, 4] == value, filtered_df.columns[14]].sum()
            result_dict[value] = {'participacion': participacion, 'indemnizacion': indemnizacion, 'importe_gastos': importe_gastos, 'total': total}
    
    return result_dict

"""
Output example:
dict = {
'REASEGURADORA PATRIA S.A.': 
            {'participacion': 1087343123, 'indemnizacion': 31904870, 'importe_gastos': 806737, 'total': 32711607}, 
'MAPFRE RE COMPAÑIA DE REASEGUROS S.A.': 
            {'participacion': 1087343123, 'indemnizacion': 63809451, 'importe_gastos': 1613404, 'total': 65422855}, 
'SCOR REINSURANCE COMPANY': 
            {'participacion': 1087343123, 'indemnizacion': 127619052, 'importe_gastos': 3226735, 'total': 130845787}, 
'KOREAN REINSURANCE COMPANY': 
            {'participacion': 1087343123, 'indemnizacion': 95714326, 'importe_gastos': 2420141, 'total': 98134467}, 
'MS AMLIN AG': 
            {'participacion': 1087343123, 'indemnizacion': 319047413, 'importe_gastos': 8066947, 'total': 327114360}
        }
"""


def calculate_table_values(resumen_dic, tasa):
    # Desempacar los valores del diccionario
    prima_qs = resumen_dic['prima_qs']
    prima_exc = resumen_dic['prima_exc']
    prima_anulada_qs = resumen_dic['prima_anulada_qs']
    prima_anulada_exc = resumen_dic['prima_anulada_exc']
    comisiones_qs = resumen_dic['comisiones_qs']
    comisiones_exc = resumen_dic['comisiones_exc']
    comisiones_anuladas_qs = resumen_dic['comisiones_anulacion_qs']
    comisiones_anuladas_exc = resumen_dic['comisiones_anulacion_exc']
    siniestros_qs = resumen_dic['siniestros_qs']
    siniestros_exc = resumen_dic['siniestros_exc']
    tasa = tasa

    # Calculate the values
    primas_cedidas = prima_qs + prima_exc
    primas_anuladas = prima_anulada_qs + prima_anulada_exc
    comisiones = (comisiones_qs - comisiones_anuladas_qs) + (comisiones_exc - comisiones_anuladas_exc)
    siniestros_pagados = siniestros_qs + siniestros_exc
    impuestos = ((primas_cedidas - primas_anuladas) - ((comisiones_qs - comisiones_anuladas_qs)+(comisiones_exc - comisiones_anuladas_exc)))*tasa
    balance_a_favor_debe = primas_anuladas + comisiones + siniestros_pagados + impuestos
    balance_a_favor_haber = primas_cedidas

    # Return the calculated values as a dictionary
    return {
        'primas_cedidas': primas_cedidas,
        'primas_anuladas': primas_anuladas,
        'comisiones': comisiones,
        'siniestros_pagados': siniestros_pagados,
        'impuestos': impuestos,
        'balance_a_favor_debe': balance_a_favor_debe,
        'balance_a_favor_haber': balance_a_favor_haber,
        'tasa': tasa
    }


def calculate_resumen_values(emitidos_qs, anulados_qs, recuperos_qs, emitidos_exc, anulados_exc, recuperos_exc):
    prima_qs = emitidos_qs['prima']
    prima_exc = emitidos_exc['prima']
    prima_anulada_qs = anulados_qs['prima']
    prima_anulada_exc = anulados_exc['prima']
    comisiones_qs = emitidos_qs['importe_comision'] - anulados_qs['importe_comision']
    comisiones_exc = emitidos_exc['importe_comision'] - anulados_exc['importe_comision']
    siniestros_qs = recuperos_qs['indemnizacion']
    siniestros_exc = recuperos_exc['indemnizacion']

    return {
        'prima_qs': prima_qs,
        'prima_exc': prima_exc,
        'prima_anulada_qs': prima_anulada_qs,
        'prima_anulada_exc': prima_anulada_exc,
        'comisiones_qs': comisiones_qs,
        'comisiones_exc': comisiones_exc,
        'siniestros_qs': siniestros_qs,
        'siniestros_exc': siniestros_exc
    }

# Esta función primero crea un conjunto de todos los reaseguradores únicos presentes en los cinco diccionarios de entrada. 
# Luego, para cada reasegurador, recopila los valores correspondientes de cada diccionario de entrada y realiza los cálculos
# necesarios para completar la tabla. Si un reasegurador no está presente en un diccionario en particular, se utilizan ceros 
# como valores predeterminados.
def generate_invoice_dict(dict_emitida_qs, dict_emitida_exc, dict_anulada_qs, dict_anulada_exc, dict_recupero_qs, dict_recupero_exc, tasa):
    # Initialize None dictionaries to empty dictionaries
    dict_emitida_qs = dict_emitida_qs or {}
    dict_emitida_exc = dict_emitida_exc or {}
    dict_anulada_qs = dict_anulada_qs or {}
    dict_anulada_exc = dict_anulada_exc or {}
    dict_recupero_qs = dict_recupero_qs or {}
    dict_recupero_exc = dict_recupero_exc or {}

    reaseguradores = set(list(dict_emitida_qs.keys()) + list(dict_emitida_exc.keys()) + list(dict_anulada_qs.keys()) + list(dict_anulada_exc.keys()) + list(dict_recupero_qs.keys()) + list(dict_recupero_exc.keys()))
    
    result_dict = {}
    for reasegurador in reaseguradores:
        prima_emitida_qs = dict_emitida_qs.get(reasegurador, {}).get('prima', 0)
        comisiones_qs = dict_emitida_qs.get(reasegurador, {}).get('importe_comision', 0)
        prima_anulada_qs = dict_anulada_qs.get(reasegurador, {}).get('prima', 0)
        comisiones_anuladas_qs = dict_anulada_qs.get(reasegurador, {}).get('importe_comision', 0)
        prima_emitida_exc = dict_emitida_exc.get(reasegurador, {}).get('prima', 0)
        comision_exc = dict_emitida_exc.get(reasegurador, {}).get('importe_comision', 0)
        prima_anulada_exc = dict_anulada_exc.get(reasegurador, {}).get('prima', 0)
        comision_anulada_exc = dict_anulada_exc.get(reasegurador, {}).get('importe_comision', 0)
        recupero_qs_total = dict_recupero_qs.get(reasegurador, {}).get('total', 0)
        recupero_exc_total = dict_recupero_exc.get(reasegurador, {}).get('total', 0)
        participacion = (
            dict_emitida_qs.get(reasegurador, {}).get('participacion') or
            dict_emitida_exc.get(reasegurador, {}).get('participacion') or
            dict_anulada_qs.get(reasegurador, {}).get('participacion') or
            dict_anulada_exc.get(reasegurador, {}).get('participacion') or
            dict_recupero_qs.get(reasegurador, {}).get('participacion') or
            dict_recupero_exc.get(reasegurador, {}).get('participacion') or
            0
        )

        prima_neta_qs = prima_emitida_qs - prima_anulada_qs
        prima_neta_exc = prima_emitida_exc - prima_anulada_exc
        comision_neta_qs = comisiones_qs - comisiones_anuladas_qs
        comision_neta_exc = comision_exc - comision_anulada_exc
        iva_qs = (prima_neta_qs - comision_neta_qs) * tasa
        menos_iva_qs = prima_neta_qs - comision_neta_qs - iva_qs
        iva_exc = (prima_neta_exc - comision_neta_exc) * tasa
        menos_iva_exc = prima_neta_exc - comision_neta_exc - iva_exc
        final_qs = recupero_qs_total - menos_iva_qs
        final_exc = recupero_exc_total - menos_iva_exc

        result_dict[reasegurador] = {
            'prima_emitida_qs': prima_emitida_qs,
            'comisiones_qs': comisiones_qs,
            'prima_anulada_qs': prima_anulada_qs,
            'comisiones_anuladas_qs': comisiones_anuladas_qs,
            'prima_emitida_exc': prima_emitida_exc,
            'comision_exc': comision_exc,
            'prima_anulada_exc': prima_anulada_exc,
            'comision_anulada_exc': comision_anulada_exc,
            'recupero_qs_total': recupero_qs_total,
            'recupero_exc_total': recupero_exc_total,
            'prima_neta_qs': prima_neta_qs,
            'prima_neta_exc': prima_neta_exc,
            'comision_neta_qs': comision_neta_qs,
            'comision_neta_exc': comision_neta_exc,
            'iva_qs': iva_qs,
            'menos_iva_qs': menos_iva_qs,
            'iva_exc': iva_exc,
            'menos_iva_exc': menos_iva_exc,
            'final_qs': int(final_qs),
            'final_exc': int(final_exc),
            'participacion': participacion
        }
    return result_dict