import streamlit as st
import pandas as pd
import locale
import time
import openpyxl
from display_resources import (convert_to_dataframe, generate_resumen)




def main():

    """Sidebar elements"""
    st.sidebar.title("Filtros")

    st.sidebar.subheader("Sección")
    # Using object notation
    section_fitler = st.sidebar.selectbox(
        '',
        ('Incendios', 'Vida'),
        placeholder="Elige una opción",
        )
    

    #Center elements
    st.title("Elaboración de Cuentas Técnicas")

    # Upload files
    uploaded_files = st.file_uploader("Subir planillas de Excel", accept_multiple_files=True, type=["xlsx"], help="Subir las planillas excel de Emisiones, Anulaciones y Recuperos una tras otra")
    
    # Inicializar las variables
    emisiones_df = anulaciones_df = recuperos_df = resumen_df = None

    # Convertir archivos a DataFrames
    if st.button('Generar resúmenes'):
        if uploaded_files:
            progress_text = "Operación en progreso. Aguarde un momento."
            my_bar = st.progress(0)
            st.text(progress_text)
            time.sleep(1)
            my_bar.progress(10)

            emisiones_df, anulaciones_df, recuperos_df = convert_to_dataframe(uploaded_files)
            my_bar.progress(33)
            st.success('Archivos convertidos a DataFrames correctamente.')
            my_bar.progress(66)

            resumen_df_2020,table_df_2020,reaseguradores_df_2020=generate_resumen(2020, emisiones_df, anulaciones_df, recuperos_df)
            st.success('Resumen 2020 generado correctamente.')
            my_bar.progress(70)

            resumen_df_2021,table_df_2021,reaseguradores_df_2021=generate_resumen(2021, emisiones_df, anulaciones_df, recuperos_df)
            st.success('Resumen 2021 generado correctamente.')
            my_bar.progress(80)

            resumen_df_2022,table_df_2022,reaseguradores_df_2022=generate_resumen(2022, emisiones_df, anulaciones_df, recuperos_df)
            st.success('Resumen 2022 generado correctamente.')
            my_bar.progress(90)

            resumen_df_2023,table_df_2023,reaseguradores_df_2023=generate_resumen(2023, emisiones_df, anulaciones_df, recuperos_df)
            st.success('Resumen 2023 generado correctamente.')
            my_bar.progress(100)
            time.sleep(1)
            my_bar.empty()


        # Mostrar resumen
        if resumen_df_2020 is not None and table_df_2020 is not None and reaseguradores_df_2020 is not None:
            st.subheader("Resumen periodo 2019-2020")
            st.dataframe(resumen_df_2020, hide_index=True)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2020, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2020, hide_index=True, use_container_width=True)

        if resumen_df_2021 is not None and table_df_2021 is not None and reaseguradores_df_2021 is not None:
            st.subheader("Resumen periodo 2020-2021")
            st.dataframe(resumen_df_2021, hide_index=True)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2021, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2021, hide_index=True, use_container_width=True)

        if resumen_df_2022 is not None and table_df_2022 is not None and reaseguradores_df_2022 is not None:
            st.subheader("Resumen periodo 2021-2022")
            st.dataframe(resumen_df_2022, hide_index=True)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2022, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2022, hide_index=True, use_container_width=True)

        if resumen_df_2023 is not None and table_df_2023 is not None and reaseguradores_df_2023 is not None:
            st.subheader("Resumen periodo 2022-2023")
            st.dataframe(resumen_df_2023, hide_index=True)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2023, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2023, hide_index=True, use_container_width=True)
    

if __name__ == "__main__":
    main()