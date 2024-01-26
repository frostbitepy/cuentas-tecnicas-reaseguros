import streamlit as st
import pandas as pd
import locale




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
    


if __name__ == "__main__":
    main()