import pandas as pd
import streamlit as st
from files_handler import (filter_dataframe_patrimoniales, filter_dataframe_vida, 
                            filter_dataframe_emisiones, filter_dataframe_anulaciones, 
                            filter_dataframe_recuperos, filter_dataframe_reaseguradores, 
                            filter_dataframe_reaseguradores_resumen, filter_dataframe_table, 
                            filter_dataframe_resumen)


# No esto seguro de si deba pasar el año como parámetro
def generate_resumen(year, dict):
    data = {
        'Vigencia/Contrato {year}': "",
        'Prima cedida en el periodo (QS)': "",
        'Prima cedida en el periodo (EXC)': "",
        'Prima anulada en el periodo (QS)': "",  # Valores de ejemplo
        'Prima anulada en el periodo (EXC)': "",  # Valores de ejemplo
        'Comisiones (QS)': "",  # Valores de ejemplo
        'Comisiones (EXC)': "",  # Valores de ejemplo
        'Siniestros pagados en el periodo QS': "",  # Valores de ejemplo
        'Siniestros pagados en el periodo EXC': ""  # Valores de ejemplo
    }
    # Crear DataFrame
    resumen_df = pd.DataFrame(data)

    # Mostrar DataFrame
    return resumen_df


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