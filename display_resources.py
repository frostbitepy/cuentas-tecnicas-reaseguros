import pandas as pd
import streamlit as st
import openpyxl
from files_handler import (process_and_sum_patrimoniales, process_sum_recuperos_patrimoniales,
                           calculate_table_values, process_and_sum_vida, process_sum_recuperos_vida,
                           generate_invoice_dict, process_and_sum_caucion, process_sum_recuperos_caucion)



# Recibe los archivos subidos y los convierte en DataFrames
def convert_to_dataframe(uploaded_files):
    for uploaded_file in uploaded_files:
        # Convertir archivo a DataFrame
        if uploaded_file.name == "Emisiones.xlsx" or uploaded_file.name == "emisiones.xls":
            emisiones_df = pd.read_excel(uploaded_file)
        elif uploaded_file.name == "Anulaciones.xlsx" or uploaded_file.name == "anulaciones.xls":
            anulaciones_df = pd.read_excel(uploaded_file)
        elif uploaded_file.name == "Recuperos.xlsx" or uploaded_file.name == "recuperos.xls":
            recuperos_df = pd.read_excel(uploaded_file)
        else:
            st.error("No se pudo cargar el archivo")

    return emisiones_df, anulaciones_df, recuperos_df
        

# Generar los resumenes de Patrimoniales
def generate_resumen(year, emisiones_df, anulaciones_df, recuperos_df):

    sums_result_emisiones_excedente,reaseguros_dict_exc = process_and_sum_patrimoniales(emisiones_df, year,'EXCEDENTE')
    sums_result_emisiones_cuota_parte,reaseguros_dict_qs = process_and_sum_patrimoniales(emisiones_df, year,'CUOTA PARTE')
    sums_result_anulaciones_excedente,reaseguros_dict_aexc = process_and_sum_patrimoniales(anulaciones_df, year,'EXCEDENTE')
    sums_result_anulaciones_cuota_parte,reaseguros_dict_aqs = process_and_sum_patrimoniales(anulaciones_df, year, 'CUOTA PARTE')
    sums_result_recuperos_excedente,reaseguros_dict_rexc = process_sum_recuperos_patrimoniales(recuperos_df, year, 'EXCEDENTE')
    sums_result_recuperos_cuota_parte,reaseguros_dict_rqs = process_sum_recuperos_patrimoniales(recuperos_df, year, 'CUOTA PARTE')
    
    tasa = 0.045

    prima_qs=sums_result_emisiones_cuota_parte['prima']
    prima_exc=sums_result_emisiones_excedente['prima']
    prima_anulada_qs=sums_result_anulaciones_cuota_parte['prima']
    prima_anulada_exc=sums_result_anulaciones_excedente['prima']
    comisiones_qs=sums_result_emisiones_cuota_parte['importe_comision']
    comisiones_exc=sums_result_emisiones_excedente['importe_comision']
    comisiones_anulacion_qs=sums_result_anulaciones_cuota_parte['importe_comision']
    comisiones_anulacion_exc=sums_result_anulaciones_excedente['importe_comision']
    siniestros_qs=sums_result_recuperos_cuota_parte['importe_total']
    siniestros_exc=sums_result_recuperos_excedente['importe_total']

    resumen_dict = {
        'prima_qs': prima_qs,
        'prima_exc': prima_exc,
        'prima_anulada_qs': prima_anulada_qs,
        'prima_anulada_exc': prima_anulada_exc,
        'comisiones_qs': comisiones_qs,
        'comisiones_exc': comisiones_exc,
        'comisiones_anulacion_qs': comisiones_anulacion_qs,
        'comisiones_anulacion_exc': comisiones_anulacion_exc,
        'siniestros_qs': siniestros_qs,
        'siniestros_exc': siniestros_exc
    }

    table_values_dict = calculate_table_values(resumen_dict, 0.045)

    # Data
    data_resumen = {
    'Vigencia/Contrato': [
        'Prima cedida en el periodo (QS)',
        'Prima cedida en el periodo (EXC)',
        'Prima anulada en el periodo (QS)',
        'Prima anulada en el periodo (EXC)',
        'Comisiones Emitidas (QS)',
        'Comisiones Emitidas (EXC)',
        'Comisiones Anuladas (QS)',
        'Comisiones Anuladas (EXC)',
        'Siniestros pagados en el periodo QS',
        'Siniestros pagados en el periodo EXC'
    ],
    'Monto' : [
        prima_qs,
        prima_exc,
        prima_anulada_qs,
        prima_anulada_exc,
        comisiones_qs,
        comisiones_exc,
        comisiones_anulacion_qs,
        comisiones_anulacion_exc,
        siniestros_qs,
        siniestros_exc
    ]
    }

    data_table_values = {
        'CONCEPTO': [
            'Primas cedidas',
            'Primas anuladas',
            'Comisiones',
            'Siniestros pagados',
            'Impuesto 4,5%',
            'Balance Saldo a favor del Reasegurador'
        ],
        'DEBE': [
            0,  # Debe asignar el valor correspondiente
            table_values_dict['primas_anuladas'],  # Debe asignar el valor correspondiente
            table_values_dict['comisiones'],  # Debe asignar el valor correspondiente
            table_values_dict['siniestros_pagados'],  # Debe asignar el valor correspondiente
            table_values_dict['impuestos'],  # Debe asignar el valor correspondiente
            table_values_dict['balance_a_favor_debe']  # Valor proporcionado
        ],
        'HABER': [
            table_values_dict['primas_cedidas'],  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            table_values_dict['balance_a_favor_haber']  # Valor proporcionado
        ]
    }

    balance_saldo = table_values_dict['balance_a_favor_haber'] - table_values_dict['balance_a_favor_debe']

    
    invoice_dict = generate_invoice_dict(reaseguros_dict_qs, reaseguros_dict_exc, reaseguros_dict_aqs, reaseguros_dict_aexc, reaseguros_dict_rqs, reaseguros_dict_rexc, 0.045)

    # Crear DataFrame
    resumen_df = pd.DataFrame(data_resumen)
    table_values_df = pd.DataFrame(data_table_values)
    reaseguradores_values_df = pd.DataFrame(generate_reaseguradores_data(invoice_dict))
    invoice_df = pd.DataFrame(invoice_dict)

    # Mostrar DataFrame
    return resumen_df,table_values_df, reaseguradores_values_df, invoice_df


# Generar los resumenes de Vida
def generate_resumen_vida(year, emisiones_df, anulaciones_df, recuperos_df):

    sums_result_emisiones_excedente,reaseguros_dict_exc = process_and_sum_vida(emisiones_df, year,'EXCEDENTE')
    sums_result_emisiones_cuota_parte,reaseguros_dict_qs = process_and_sum_vida(emisiones_df, year,'CUOTA PARTE')
    sums_result_anulaciones_excedente,reaseguros_dict_aexc = process_and_sum_vida(anulaciones_df, year,'EXCEDENTE')
    sums_result_anulaciones_cuota_parte,reaseguros_dict_aqs = process_and_sum_vida(anulaciones_df, year, 'CUOTA PARTE')
    sums_result_recuperos_excedente,reaseguros_dict_rexc = process_sum_recuperos_vida(recuperos_df, year, 'EXCEDENTE')
    sums_result_recuperos_cuota_parte,reaseguros_dict_rqs = process_sum_recuperos_vida(recuperos_df, year, 'CUOTA PARTE')
    
    tasa = 0.045

    prima_qs=sums_result_emisiones_cuota_parte['prima']
    prima_exc=sums_result_emisiones_excedente['prima']
    prima_anulada_qs=sums_result_anulaciones_cuota_parte['prima']
    prima_anulada_exc=sums_result_anulaciones_excedente['prima']
    comisiones_qs=sums_result_emisiones_cuota_parte['importe_comision']
    comisiones_exc=sums_result_emisiones_excedente['importe_comision']
    comisiones_anulacion_qs=sums_result_anulaciones_cuota_parte['importe_comision']
    comisiones_anulacion_exc=sums_result_anulaciones_excedente['importe_comision']
    siniestros_qs=sums_result_recuperos_cuota_parte['importe_total']
    siniestros_exc=sums_result_recuperos_excedente['importe_total']

    resumen_dict = {
        'prima_qs': prima_qs,
        'prima_exc': prima_exc,
        'prima_anulada_qs': prima_anulada_qs,
        'prima_anulada_exc': prima_anulada_exc,
        'comisiones_qs': comisiones_qs,
        'comisiones_exc': comisiones_exc,
        'comisiones_anulacion_qs': comisiones_anulacion_qs,
        'comisiones_anulacion_exc': comisiones_anulacion_exc,
        'siniestros_qs': siniestros_qs,
        'siniestros_exc': siniestros_exc
    }

    table_values_dict = calculate_table_values(resumen_dict, 0.045)

    # Data
    data_resumen = {
    'Vigencia/Contrato': [
        'Prima cedida en el periodo (QS)',
        'Prima cedida en el periodo (EXC)',
        'Prima anulada en el periodo (QS)',
        'Prima anulada en el periodo (EXC)',
        'Comisiones Emitidas (QS)',
        'Comisiones Emitidas (EXC)',
        'Comisiones Anuladas (QS)',
        'Comisiones Anuladas (EXC)',
        'Siniestros pagados en el periodo QS',
        'Siniestros pagados en el periodo EXC'
    ],
    'Monto' : [
        prima_qs,
        prima_exc,
        prima_anulada_qs,
        prima_anulada_exc,
        comisiones_qs,
        comisiones_exc,
        comisiones_anulacion_qs,
        comisiones_anulacion_exc,
        siniestros_qs,
        siniestros_exc
    ]
    }

    data_table_values = {
        'CONCEPTO': [
            'Primas cedidas',
            'Primas anuladas',
            'Comisiones',
            'Siniestros pagados',
            'Impuesto 4,5%',
            'Balance Saldo a favor del Reasegurador'
        ],
        'DEBE': [
            0,  # Debe asignar el valor correspondiente
            table_values_dict['primas_anuladas'],  # Debe asignar el valor correspondiente
            table_values_dict['comisiones'],  # Debe asignar el valor correspondiente
            table_values_dict['siniestros_pagados'],  # Debe asignar el valor correspondiente
            table_values_dict['impuestos'],  # Debe asignar el valor correspondiente
            table_values_dict['balance_a_favor_debe']  # Valor proporcionado
        ],
        'HABER': [
            table_values_dict['primas_cedidas'],  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            table_values_dict['balance_a_favor_haber']  # Valor proporcionado
        ]
    }

    balance_saldo = table_values_dict['balance_a_favor_haber'] - table_values_dict['balance_a_favor_debe']

    invoice_dict = generate_invoice_dict(reaseguros_dict_qs, reaseguros_dict_exc, reaseguros_dict_aqs, reaseguros_dict_aexc, reaseguros_dict_rqs, reaseguros_dict_rexc, 0.045)

    # Crear DataFrame
    resumen_df = pd.DataFrame(data_resumen)
    table_values_df = pd.DataFrame(data_table_values)
    reaseguradores_values_df = pd.DataFrame(generate_reaseguradores_data(invoice_dict))
    invoice_df = pd.DataFrame(invoice_dict)

    # Mostrar DataFrame
    return resumen_df,table_values_df, reaseguradores_values_df, invoice_df


# Generar los resumenes de Caucion
def generate_resumen_caucion(year, emisiones_df, anulaciones_df, recuperos_df):

    sums_result_emisiones_excedente,reaseguros_dict_exc = process_and_sum_caucion(emisiones_df, year,'EXCEDENTE')
    sums_result_emisiones_cuota_parte,reaseguros_dict_qs = process_and_sum_caucion(emisiones_df, year,'CUOTA PARTE')
    sums_result_anulaciones_excedente,reaseguros_dict_aexc = process_and_sum_caucion(anulaciones_df, year,'EXCEDENTE')
    sums_result_anulaciones_cuota_parte,reaseguros_dict_aqs = process_and_sum_caucion(anulaciones_df, year, 'CUOTA PARTE')
    sums_result_recuperos_excedente,reaseguros_dict_rexc = process_sum_recuperos_caucion(recuperos_df, year, 'EXCEDENTE')
    sums_result_recuperos_cuota_parte,reaseguros_dict_rqs = process_sum_recuperos_caucion(recuperos_df, year, 'CUOTA PARTE')
    
    tasa = 0.045

    prima_qs=sums_result_emisiones_cuota_parte['prima']
    prima_exc=sums_result_emisiones_excedente['prima']
    prima_anulada_qs=sums_result_anulaciones_cuota_parte['prima']
    prima_anulada_exc=sums_result_anulaciones_excedente['prima']
    comisiones_qs=sums_result_emisiones_cuota_parte['importe_comision']
    comisiones_exc=sums_result_emisiones_excedente['importe_comision']
    comisiones_anulacion_qs=sums_result_anulaciones_cuota_parte['importe_comision']
    comisiones_anulacion_exc=sums_result_anulaciones_excedente['importe_comision']
    siniestros_qs=sums_result_recuperos_cuota_parte['importe_total']
    siniestros_exc=sums_result_recuperos_excedente['importe_total']

    resumen_dict = {
        'prima_qs': prima_qs,
        'prima_exc': prima_exc,
        'prima_anulada_qs': prima_anulada_qs,
        'prima_anulada_exc': prima_anulada_exc,
        'comisiones_qs': comisiones_qs,
        'comisiones_exc': comisiones_exc,
        'comisiones_anulacion_qs': comisiones_anulacion_qs,
        'comisiones_anulacion_exc': comisiones_anulacion_exc,
        'siniestros_qs': siniestros_qs,
        'siniestros_exc': siniestros_exc
    }

    table_values_dict = calculate_table_values(resumen_dict, 0.045)

    # Data
    data_resumen = {
    'Vigencia/Contrato': [
        'Prima cedida en el periodo (QS)',
        'Prima cedida en el periodo (EXC)',
        'Prima anulada en el periodo (QS)',
        'Prima anulada en el periodo (EXC)',
        'Comisiones Emitidas (QS)',
        'Comisiones Emitidas (EXC)',
        'Comisiones Anuladas (QS)',
        'Comisiones Anuladas (EXC)',
        'Siniestros pagados en el periodo QS',
        'Siniestros pagados en el periodo EXC'
    ],
    'Monto' : [
        prima_qs,
        prima_exc,
        prima_anulada_qs,
        prima_anulada_exc,
        comisiones_qs,
        comisiones_exc,
        comisiones_anulacion_qs,
        comisiones_anulacion_exc,
        siniestros_qs,
        siniestros_exc
    ]
    }

    data_table_values = {
        'CONCEPTO': [
            'Primas cedidas',
            'Primas anuladas',
            'Comisiones',
            'Siniestros pagados',
            'Impuesto 4,5%',
            'Balance Saldo a favor del Reasegurador'
        ],
        'DEBE': [
            0,  # Debe asignar el valor correspondiente
            table_values_dict['primas_anuladas'],  # Debe asignar el valor correspondiente
            table_values_dict['comisiones'],  # Debe asignar el valor correspondiente
            table_values_dict['siniestros_pagados'],  # Debe asignar el valor correspondiente
            table_values_dict['impuestos'],  # Debe asignar el valor correspondiente
            table_values_dict['balance_a_favor_debe']  # Valor proporcionado
        ],
        'HABER': [
            table_values_dict['primas_cedidas'],  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            0,  # Valor proporcionado
            table_values_dict['balance_a_favor_haber']  # Valor proporcionado
        ]
    }

    balance_saldo = table_values_dict['balance_a_favor_haber'] - table_values_dict['balance_a_favor_debe']

    invoice_dict = generate_invoice_dict(reaseguros_dict_qs, reaseguros_dict_exc, reaseguros_dict_aqs, reaseguros_dict_aexc, reaseguros_dict_rqs, reaseguros_dict_rexc, 0.045)

    # Crear DataFrame
    resumen_df = pd.DataFrame(data_resumen)
    table_values_df = pd.DataFrame(data_table_values)
    reaseguradores_values_df = pd.DataFrame(generate_reaseguradores_data(invoice_dict))
    invoice_df = pd.DataFrame(invoice_dict)

    # Mostrar DataFrame
    return resumen_df,table_values_df, reaseguradores_values_df, invoice_df


# Deprecated
def generate_table(dict):
    # Datos proporcionados
    data = {
        'CONCEPTO': [
            'Primas cedidas',
            'Primas anuladas',
            'Comisiones',
            'Siniestros pagados',
            'Impuesto 4,5%',
            'Balance Saldo a favor del Reasegurador'
        ],
        'DEBE': [
            None,  # Debe asignar el valor correspondiente
            None,  # Debe asignar el valor correspondiente
            None,  # Debe asignar el valor correspondiente
            None,  # Debe asignar el valor correspondiente
            None,  # Debe asignar el valor correspondiente
            3229117315  # Valor proporcionado
        ],
        'HABER': [
            4973894666,  # Valor proporcionado
            416154325,  # Valor proporcionado
            1407700446,  # Valor proporcionado
            1263510749,  # Valor proporcionado
            141751795,  # Valor proporcionado
            4973894666  # Valor proporcionado
        ]
    }

    # Crear DataFrame
    table_df = pd.DataFrame(data)

    # Mostrar DataFrame
    return table_df
###########
# Deprecated
def generate_reaseguradores_resumen(dict):
    # Data
    data = {
        'Reasegurador': [
            'MS AMLIN AG',
            'SCOR REINSURANCE COMPANY',
            'KOREAN REINSURANCE CORPORATION',
            'MAPFRE RE COMPAÑÍA DE REASEGUROS S.A.',
            'REASEGURADORA PATRIA S.A.'
        ],
        'Participación': [
             '50%', 
             '20%', 
             '15%', 
             '10%', 
             '5%'],
        'Monto': [
             0, 
             0, 
             0, 
             0, 
             0]
    }

    # Create DataFrame
    df_reaseguradores = pd.DataFrame(data)

    return df_reaseguradores
############

def generate_cuenta_tecnica(dict):
    impuestos = ((dict['Prima cedida en el periodo (EXC)'] + dict['Prima cedida en el periodo (QS)']) - (dict['Prima anulada en el periodo (EXC)'] + dict['Prima anulada en el periodo (QS)'] +
                (dict['Comisiones Emitidas (EXC)'] - dict['Comisiones Anuladas (EXC)']) + (dict['Comisiones Emitidas (QS)'] - dict['Comisiones Anuladas (QS)'])))*0.045
    subtotal_debe = (dict['Prima anulada en el periodo (EXC)']+dict['Prima anulada en el periodo (QS)']+dict['Comisiones Emitidas (EXC)']+dict['Comisiones Emitidas (QS)']-dict['Comisiones Anuladas (EXC)'] - 
                    dict['Comisiones Anuladas (QS)'] + impuestos + dict['Siniestros pagados en el periodo QS']+dict['Siniestros pagados en el periodo EXC'])
    subtotal_haber = (dict['Prima cedida en el periodo (EXC)']+dict['Prima cedida en el periodo (QS)'])
    saldo_estado_cuenta = subtotal_debe - subtotal_haber
    total_debe = subtotal_debe 
    total_haber = subtotal_haber + saldo_estado_cuenta


    data = {
        'CONCEPTO':[
            'PRIMAS CEDIDAS EXCEDENTE',
            'PRIMAS CEDIDAS CUOTA PARTE',
            'PRIMAS CEDIDAS ANULADAS EXCEDENTE',
            'PRIMAS CEDIDAS ANULADAS CUOTA PARTE',
            'COMISION EMITIDA EXCEDENTE DEL PERIODO',
            'COMISION EMITIDA CUOTA PARTE DEL PERIODO',
            'COMISION ANULADA EXCEDENTE DEL PERIODO',
            'COMISION ANULADA CUOTA PARTE DEL PERIODO',
            'IMPUESTO DE LEY - 4,5% DEL PERIODO DE CESION',
            'SINIESTROS PAGADOS DEL PERIODO DE CESION EXCEDENTE',
            'SINIESTROS PAGADOS DEL PERIODO DE CESION CUOTA PARTE',
            'SUBTOTALES',
            'SALDO DEL ESTADO DE CUENTA',
            'TOTAL'
        ],
        'DEBE':[
            0,
            0,
            dict['Prima anulada en el periodo (EXC)'],
            dict['Prima anulada en el periodo (QS)'],
            dict['Comisiones Emitidas (EXC)'],
            dict['Comisiones Emitidas (QS)'],
            dict['Comisiones Anuladas (EXC)'],
            dict['Comisiones Anuladas (QS)'],
            int(impuestos), # IMPUESTO DE LEY - 4,5% DEL PERIODO DE CESION
            dict['Siniestros pagados en el periodo EXC'],
            dict['Siniestros pagados en el periodo QS'],
            int(subtotal_debe), # SUBTOTALES
            0,
            int(total_debe) # TOTAL
        ],
        'HABER':[
            dict['Prima cedida en el periodo (EXC)'],
            dict['Prima cedida en el periodo (QS)'],
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            int(subtotal_haber), # SUBTOTALES
            int(saldo_estado_cuenta), # SALDO DEL ESTADO DE CUENTA
            int(total_haber) # TOTAL
        ]
    }

    # Create dataframe
    df_cuenta_tecnica = pd.DataFrame(data)

    return df_cuenta_tecnica

# Deprecated
def generate_reaseguradores_data_depr(input_dict, valor):
    reaseguradores = list(input_dict.keys())
    participaciones = [str(value) + '%' for value in input_dict.values()]
    montos = [value * valor / 100 for value in input_dict.values()]

    data_reaseguradores = {
        'Reasegurador': reaseguradores,
        'Participación': participaciones,
        'Monto': montos
    }
    return data_reaseguradores
#############

def generate_reaseguradores_data(input_dict):
    reaseguradores = list(input_dict.keys())
    participacion = [str(input_dict[key]['participacion']) + '%' for key in input_dict.keys()]
    monto_qs = [input_dict[key]['final_qs'] for key in input_dict.keys()]
    monto_exc = [input_dict[key]['final_exc'] for key in input_dict.keys()]

    data_reaseguradores = {
        'Reasegurador': reaseguradores,
        'Participación': participacion,
        'Monto QS': monto_qs,
        'Monto EXC': monto_exc
    }
    return data_reaseguradores


def sum_dataframe_values(resumen_df_container):
    sums = {}
    for df in resumen_df_container:
        unique_rows = df['Vigencia/Contrato'].unique()
        for row in unique_rows:
            if row in sums:
                sums[row] += df[df['Vigencia/Contrato'] == row]['Monto'].sum()
            else:
                sums[row] = df[df['Vigencia/Contrato'] == row]['Monto'].sum()
    #result_df = pd.DataFrame(list(sums.items()), columns=['Vigencia/Contrato', 'Monto'])
    return sums


