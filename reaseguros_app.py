import streamlit as st
import pandas as pd
import locale
import time
import base64
import io
import xlsxwriter
from display_resources import (convert_to_dataframe, generate_resumen,
                               sum_dataframe_values, generate_cuenta_tecnica,
                               generate_resumen_vida, generate_resumen_caucion)


@st.cache_data
def convert_dict_to_excel(data_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        cell_format = workbook.add_format({'text_wrap': False})

        for key, value_list in data_dict.items():
            sheet_name = str(key)

            for i, item in enumerate(value_list):
                sheet_name_item = sheet_name
                
                item.to_excel(writer, sheet_name=sheet_name_item, index=False, startrow=i * (len(item) + 10), header=True)
                
                for j, column in enumerate(item.columns):
                    max_len = max(item[column].astype(str).apply(len).max(), len(column))
                    writer.sheets[sheet_name_item].set_column(j, j, max_len + 2, cell_format)

    processed_data = output.getvalue()
    return processed_data


def main():

    """Sidebar elements"""
    st.sidebar.title("Filtros")

    st.sidebar.subheader("Sección")
    # Using object notation
    section_filter = st.sidebar.selectbox(
        '',
        ('Incendios', 'Vida', 'Caución'),
        placeholder="Elige una opción",
        )
    

    #Center elements
    st.title("Estados de Cuentas Trimestrales")

    # Upload files
    uploaded_files = st.file_uploader("Subir planillas de Excel", accept_multiple_files=True, type=["xlsx"], help="Subir las planillas excel de Emisiones, Anulaciones y Recuperos una tras otra")
    
    # Inicializar las variables
    emisiones_df = anulaciones_df = recuperos_df = resumen_df = None
    resumen_df_container = []
    resumen_dict_container = {}

    # Convertir archivos a DataFrames
    if st.button('Generar resúmenes'):
        if uploaded_files:
            progress_text = "Operación en progreso. Aguarde un momento."
            my_bar = st.progress(0)
            st.text(progress_text)
            time.sleep(1)
            my_bar.progress(10)
            if section_filter == 'Incendios':
                emisiones_df, anulaciones_df, recuperos_df = convert_to_dataframe(uploaded_files)
                my_bar.progress(33)
                st.success('Archivos convertidos a DataFrames correctamente.')
                my_bar.progress(66)

                resumen_df_2020,table_df_2020,reaseguradores_df_2020,invoice_df_2020=generate_resumen(2020, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2020 generado correctamente.')
                my_bar.progress(70)

                resumen_df_2021,table_df_2021,reaseguradores_df_2021,invoice_df_2021=generate_resumen(2021, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2021 generado correctamente.')
                my_bar.progress(80)

                resumen_df_2022,table_df_2022,reaseguradores_df_2022,invoice_df_2022=generate_resumen(2022, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2022 generado correctamente.')
                my_bar.progress(90)

                resumen_df_2023,table_df_2023,reaseguradores_df_2023,invoice_df_2023=generate_resumen(2023, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2023 generado correctamente.')
                my_bar.progress(95)

                resumen_df_2024,table_df_2024,reaseguradores_df_2024,invoice_df_2024=generate_resumen(2024, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2024 generado correctamente.')
                my_bar.progress(100)

                time.sleep(1)
                my_bar.empty()

            elif section_filter == 'Vida':
                emisiones_df, anulaciones_df, recuperos_df = convert_to_dataframe(uploaded_files)
                my_bar.progress(33)
                st.success('Archivos convertidos a DataFrames correctamente.')
                my_bar.progress(66)

                resumen_df_2020,table_df_2020,reaseguradores_df_2020,invoice_df_2020=generate_resumen_vida(2020, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2020 generado correctamente.')
                my_bar.progress(70)

                resumen_df_2021,table_df_2021,reaseguradores_df_2021,invoice_df_2021=generate_resumen_vida(2021, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2021 generado correctamente.')
                my_bar.progress(80)

                resumen_df_2022,table_df_2022,reaseguradores_df_2022,invoice_df_2022=generate_resumen_vida(2022, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2022 generado correctamente.')
                my_bar.progress(90)

                resumen_df_2023,table_df_2023,reaseguradores_df_2023,invoice_df_2023=generate_resumen_vida(2023, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2023 generado correctamente.')
                my_bar.progress(95)

                resumen_df_2024,table_df_2024,reaseguradores_df_2024,invoice_df_2024=generate_resumen_vida(2024, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2024 generado correctamente.')
                my_bar.progress(100)

                time.sleep(1)
                my_bar.empty()

            elif section_filter == 'Caución':
                emisiones_df, anulaciones_df, recuperos_df = convert_to_dataframe(uploaded_files)
                my_bar.progress(33)
                st.success('Archivos convertidos a DataFrames correctamente.')
                my_bar.progress(66)

                resumen_df_2020,table_df_2020,reaseguradores_df_2020,invoice_df_2020=generate_resumen_caucion(2020, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2020 generado correctamente.')
                my_bar.progress(70)

                resumen_df_2021,table_df_2021,reaseguradores_df_2021,invoice_df_2021=generate_resumen_caucion(2021, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2021 generado correctamente.')
                my_bar.progress(80)

                resumen_df_2022,table_df_2022,reaseguradores_df_2022,invoice_df_2022=generate_resumen_caucion(2022, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2022 generado correctamente.')
                my_bar.progress(90)

                resumen_df_2023,table_df_2023,reaseguradores_df_2023,invoice_df_2023=generate_resumen_caucion(2023, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2023 generado correctamente.')
                my_bar.progress(95)

                resumen_df_2024,table_df_2024,reaseguradores_df_2024,invoice_df_2024=generate_resumen_caucion(2024, emisiones_df, anulaciones_df, recuperos_df)
                #st.success('Resumen 2024 generado correctamente.')
                my_bar.progress(100)

                time.sleep(1)
                my_bar.empty()


            
        else:
            st.warning('Debe cargar las planillas correspondientes a Emisiones, Anulaciones y Recuperos', icon="⚠️")
            

        
        # Mostrar resumen
        if (resumen_df_2020 is not None and not resumen_df_2020.empty) and \
            (table_df_2020 is not None and not table_df_2020.empty) and \
            (reaseguradores_df_2020 is not None and not reaseguradores_df_2020.empty):
            st.subheader("Resumen periodo 2019-2020")
            st.dataframe(resumen_df_2020, hide_index=True, use_container_width=True)
            resumen_df_container.append(resumen_df_2020)
            resumen_dict_container["2020"] = []
            resumen_dict_container["2020"].append(resumen_df_2020)
            resumen_dict_container["2020"].append(table_df_2020)
            resumen_dict_container["2020"].append(reaseguradores_df_2020)
            resumen_dict_container["2020"].append(invoice_df_2020)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2020, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2020, hide_index=True, use_container_width=True)
            #st.dataframe(invoice_df_2020, hide_index=True, use_container_width=True)

        if (resumen_df_2021 is not None and not resumen_df_2021.empty) and \
            (table_df_2021 is not None and not table_df_2021.empty) and \
            (reaseguradores_df_2021 is not None and not reaseguradores_df_2021.empty):
            st.subheader("Resumen periodo 2020-2021")
            st.dataframe(resumen_df_2021, hide_index=True, use_container_width=True)
            resumen_df_container.append(resumen_df_2021)
            resumen_dict_container["2021"] = []
            resumen_dict_container["2021"].append(resumen_df_2021)
            resumen_dict_container["2021"].append(table_df_2021)
            resumen_dict_container["2021"].append(reaseguradores_df_2021)
            resumen_dict_container["2021"].append(invoice_df_2021)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2021, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2021, hide_index=True, use_container_width=True)
            #st.dataframe(invoice_df_2021, hide_index=True, use_container_width=True)

        if (resumen_df_2022 is not None and not resumen_df_2022.empty) and \
            (table_df_2022 is not None and not table_df_2022.empty) and \
            (reaseguradores_df_2022 is not None and not reaseguradores_df_2022.empty):
            st.subheader("Resumen periodo 2021-2022")
            st.dataframe(resumen_df_2022, hide_index=True, use_container_width=True)
            resumen_df_container.append(resumen_df_2022)
            resumen_dict_container["2022"] = []
            resumen_dict_container["2022"].append(resumen_df_2022)
            resumen_dict_container["2022"].append(table_df_2022)
            resumen_dict_container["2022"].append(reaseguradores_df_2022)
            resumen_dict_container["2022"].append(invoice_df_2022)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2022, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2022, hide_index=True, use_container_width=True)
            #st.dataframe(invoice_df_2022, hide_index=True, use_container_width=True)

        if (resumen_df_2023 is not None and not resumen_df_2023.empty) and \
            (table_df_2023 is not None and not table_df_2023.empty) and \
            (reaseguradores_df_2023 is not None and not reaseguradores_df_2023.empty):
            st.subheader("Resumen periodo 2022-2023")
            st.dataframe(resumen_df_2023, hide_index=True, use_container_width=True)
            resumen_df_container.append(resumen_df_2023)
            resumen_dict_container["2023"] = []
            resumen_dict_container["2023"].append(resumen_df_2023)
            resumen_dict_container["2023"].append(table_df_2023)
            resumen_dict_container["2023"].append(reaseguradores_df_2023)
            resumen_dict_container["2023"].append(invoice_df_2023)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2023, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2023, hide_index=True, use_container_width=True)
            #st.dataframe(invoice_df_2023, hide_index=True, use_container_width=True)

        if (resumen_df_2024 is not None and not resumen_df_2024.empty) and \
                (table_df_2024 is not None and not table_df_2024.empty) and \
                (reaseguradores_df_2024 is not None and not reaseguradores_df_2024.empty):
            st.subheader("Resumen periodo 2023-2024")
            st.dataframe(resumen_df_2024, hide_index=True, use_container_width=True)
            resumen_df_container.append(resumen_df_2024)
            resumen_dict_container["2024"] = []
            resumen_dict_container["2024"].append(resumen_df_2024)
            resumen_dict_container["2024"].append(table_df_2024)
            resumen_dict_container["2024"].append(reaseguradores_df_2024)
            resumen_dict_container["2024"].append(invoice_df_2024)
            #st.subheader("Tabla de valores")
            st.dataframe(table_df_2024, hide_index=True, use_container_width=True)
            st.dataframe(reaseguradores_df_2024, hide_index=True, use_container_width=True)
            #st.dataframe(invoice_df_2024, hide_index=True, use_container_width=True)

        # Mostrar estado de cuenta
        if resumen_df_container:    
            st.subheader("Estado de Cuenta Trimestral")
            sums_dict = sum_dataframe_values(resumen_df_container)
            cuenta_tecnica_df = (generate_cuenta_tecnica(sums_dict))
            resumen_dict_container["cuenta_tecnica"] = []
            resumen_dict_container["cuenta_tecnica"].append(cuenta_tecnica_df)
            st.dataframe(cuenta_tecnica_df, hide_index=True, use_container_width=True)


        xlsx_file = convert_dict_to_excel(resumen_dict_container)


        st.download_button(
            label="Download Excel File",
            data=xlsx_file,
            file_name='resumen.xlsx',
            mime='text/xlsx',
        )

    
          

if __name__ == "__main__":
    main()