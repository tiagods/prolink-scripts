import streamlit as st
import pandas as pd
from io import BytesIO

def processar_arquivo(file):
    df = pd.read_excel(file)
    return df

def main():
    st.title("Processador de Arquivos Excel")
    
    uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")
    
    if uploaded_file is not None:
        df = processar_arquivo(uploaded_file)
        st.write(df)

if __name__ == "__main__":
    main()
