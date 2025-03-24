import streamlit as st
import pandas as pd
from io import BytesIO

# Função para processar o Excel
def processar_excel(arquivo):
    df = pd.read_excel(arquivo)
    
    # Simples exemplo de processamento (tudo maiúsculo)
    df = df.applymap(lambda x: str(x).upper() if isinstance(x, str) else x)

    # Salvar em memória
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado")
    output.seek(0)
    
    return output

# Interface Streamlit
st.title("Processador de Arquivos Excel")

# Upload do arquivo
arquivo = st.file_uploader("Envie um arquivo Excel", type=["xls", "xlsx"])

if arquivo:
    st.write("Arquivo carregado com sucesso!")
    
    # Processar o arquivo
    output = processar_excel(arquivo)
    
    # Botão de download
    st.download_button(
        label="Baixar Arquivo Processado",
        data=output,
        file_name="resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
        # Botão de download
    st.download_button(
        label="Baixar Arquivo Processado 2",
        data=output,
        file_name="resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )