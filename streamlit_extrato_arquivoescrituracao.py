import streamlit as st
import pandas as pd
from io import BytesIO
import re
from datetime import datetime
import xlrd

errorReport = []

def handle_file_upload(file, file_type):
    """Handle file upload event and log it"""
    if file is not None:
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.info(f"Arquivo {file_type} carregado em {current_time}")
        return True
    return False

def carregar_arquivo(file):
        df = pd.read_excel(file, header=None)
        return df

def processar_arquivo(nome, df, nome_colunas):
    if df.empty:
        raise ValueError(f"O arquivo {nome} está vazio.")
    # Procura pela linha que contém todos os nomes das colunas
    linha_colunas = -1
    for i, linha in df.iterrows():
        if all(nome in linha.values for nome in nome_colunas):
            linha_colunas = i
            break

    if linha_colunas == -1:
        erro = f"Nenhuma linha contendo os nomes das colunas foi encontrada no arquivo {nome}."
        raise ValueError(erro)

    msg = f"A linha {linha_colunas} contém os nomes das colunas no arquivo {nome}."
    st.write(msg)
    return linha_colunas

def gerar_extrato( nome,  df_modificado):
    extrato = []
    for index, row in df_modificado.iterrows():
        texto = row['Notas Fiscais']
        data = row['Data']
        valor = row['Valor (R$)']
        
        texto_limpo = re.sub(r"\s*(\d)\s+(\d)\s*", r"\1\2", texto, flags=re.IGNORECASE)  # Remove espaços entre números
        texto_limpo = re.sub(r"NFS-?e(\d+)", r"NFS-e \1", texto_limpo, flags=re.IGNORECASE)  # Adiciona espaço após "NFS-e" quando seguido por números
        texto_limpo = re.sub(r"\s+", " ", texto_limpo).strip()        # Limpa espaços extras
        
        padrao = r"nfs-?e\s+?(\d+)"
        numeros_notas = re.findall(padrao, texto_limpo, flags=re.IGNORECASE)
        lista_inteiros = [int(x) for x in numeros_notas if x is not None]  # Ignora valores None
        
        if data == '':
            erro = f"{nome}: Nenhuma data encontrada para o lançamento {index}"
            errorReport.append(erro)
            continue

        dataStr = datetime.strftime(data, '%m/%d/%Y')

        if valor == '':
            erro = f"{nome}: Nenhuma valor encontrado para o lançamento {index}"
            errorReport.append(erro)
            continue

        if len(lista_inteiros) == 0:
            erro = f"{nome}: Nenhuma nota fiscal encontrada para o lançamento {index}, {texto}, {texto_limpo}"
            errorReport.append(erro)
            continue
        
        extrato.append({
                'index': index,
                'data': dataStr,
                'lançamento': row['Lançamento'],
                'valor': valor,
                'notas': lista_inteiros,
                'nfs': "NF N " + str(lista_inteiros[0]) + "".join([", " + str(num) for num in lista_inteiros[1:]])
            })
    return extrato
def processar_notas_fiscais():
    global uploaded_extrato, uploaded_servicos, df_servicos_filtrado, extrato
    # Filtra as notas fiscais do extrato
    uploaded_extrato = st.file_uploader("Escolha o arquivo de Extrato", type=['xlsx', 'xls'])
    if handle_file_upload(uploaded_extrato, "extrato"):
        try:
            df = carregar_arquivo(uploaded_extrato)
            nomes_colunas_extrato = ["Data", "Lançamento", "Valor (R$)", "Saldo (R$)"]
            linha_colunas = processar_arquivo("Extrato", df, nomes_colunas_extrato)

            df_com_cabecalho = pd.read_excel(uploaded_extrato, header=linha_colunas)
            colunas_para_remover = [0]
            df_extrato = df_com_cabecalho.drop(df_com_cabecalho.columns[colunas_para_remover], axis=1)
            novo_nome = "Notas Fiscais"
            df_extrato.rename(columns={df_extrato.columns[-1]: novo_nome}, inplace=True)
            extrato = gerar_extrato("Acompanhamento de Serviços", df_extrato)
            st.success(f"Arquivo de extrato processado com sucesso! {len(extrato)} notas fiscais encontradas.")
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {str(e)}")
            raise e

    uploaded_servicos = st.file_uploader("Escolha o arquivo de Acompanhamento de Serviços", type=['xlsx', 'xls'])
    if handle_file_upload(uploaded_servicos, "Acompanhamento de Serviços"):
        try:
            df = carregar_arquivo(uploaded_servicos)
            nomes_colunas_servicos = ["Código", "Data", "Nota", "Série", "Espécie", "Código", "Cliente", "AC.", "UF", "Valor Contábil", "Tipo", "Base Cálculo", "Alíq.", "Valor", "Isentas", "Outras" ]
            linha_colunas_servicos = processar_arquivo("Acompanhamento de Serviços", df, nomes_colunas_servicos)

            nomes_colunas_servicos_2 = ["Nota", "Cliente"]
            result = pd.read_excel(uploaded_servicos, header=linha_colunas_servicos, usecols=nomes_colunas_servicos_2)
            df_servicos_filtrado = result.copy()
            df_servicos_filtrado = df_servicos_filtrado.dropna(subset=["Nota"])
            df_servicos_filtrado["Nota"] = df_servicos_filtrado["Nota"].astype(int)
            st.success(f"Arquivo de Acompanhamento de Serviços processado com sucesso! {len(df_servicos_filtrado)} notas fiscais encontradas.")
        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {str(e)}")
            raise e

def processar_extrato():
    global rows
    rows = 100

    if uploaded_extrato is None:
         st.error("Nenhum arquivo de Extrato carregado")
         return
    if uploaded_servicos is None:
         st.error("Nenhum arquivo de Acompanhamento de Serviços carregado")
         return
    if df_servicos_filtrado is None or df_servicos_filtrado.empty:
         st.error("Nenhum arquivo de Acompanhamento de Serviços processado")
         return
    if extrato is None or len(extrato) == 0:
         st.error("Nenhum arquivo de Extrato processado")
         return

    if len(extrato) > 0:
        colunas = ["Data", "Cód. Conta Debito", "Cód. Conta Credito", "Valor", "Cód. Histórico", "Complemento Histórico", "Inicia Lote", "Código Matriz/Filial", "Centro de Custo Débito", "Centro de Custo Crédito"]
        extrato_conciliado = pd.DataFrame(columns=colunas)
        for ext in extrato:
            df_result = df_servicos_filtrado[df_servicos_filtrado["Nota"].isin(ext["notas"])]
            if df_result.empty:
                erro = f"Nenhuma nota no arquivo Acompanhamentos de Serviços foi encontrada. Dados do extrado: data={ext['data']}, notas={ext['notas']}, valor={ext['valor']}, nfs={ext['nfs']}, lancamento={ext['lançamento']}, linha={ext['index']}"
                errorReport.append(erro)
            else:
                ext["cliente"] = df_result['Cliente'].values[0]
                extrato_conciliado.loc[len(extrato_conciliado)] = [ext["data"], "10008", "31577", ext["valor"], "132", ext["nfs"] + ext["cliente"], "", 2194, "",""]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            pd.DataFrame(extrato_conciliado).to_excel(writer, index=False, sheet_name="Resultado")
        output.seek(0)

        st.dataframe(extrato_conciliado, use_container_width=True, hide_index=True)

        st.download_button(
            type="primary",
            label="Baixar Arquivo Processado",
            data=output,
            file_name="Planilha Escrituração Dominio.xls",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/check:"
        )

        if len(errorReport) > 0:
            df_error = pd.DataFrame(errorReport, columns=["Erro"])
            st.dataframe(df_error, use_container_width=True, hide_index=True)
            output_error = BytesIO()
            with pd.ExcelWriter(output_error, engine="xlsxwriter") as writer:
                df_error.to_excel(writer, index=False, sheet_name="Erros")
            output_error.seek(0)

            st.download_button(
                type="primary",
                label="Baixar Arquivo de Erros",
                data=output_error,
                file_name="erros.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/error:"
            )

def main():
    try:
        st.title("Processamento de Extrato e Acompanhamento de Serviços")

        st.markdown(
            """
            Este sistema foi criado para processamento de arquivos de planilha Excel contendo informações financeiras.
            Funcionamento:
            - Permite o upload de dois arquivos Excel: Extrato e Acompanhamento de Serviços
            - Extrai números de notas fiscais dos lançamentos do extrato
            - Relaciona estas notas com os registros do arquivo de serviços
            - Gera um extrato conciliado com informações completas
            Dados processados:
            - Data das movimentações
            - Descrição dos lançamentos
            - Valores monetários
            - Documentos fiscais relacionados
            Verificações de segurança:
            - Confirma existência de movimentações
            - Identifica arquivos vazios
            - Registra problemas para investigação posterior
            Tipos de movimentações tratadas:
            - SISPAG
            - TED
            - PIX
            Resultado final:
            - Apresenta relatório na tela
            - Oferece download do arquivo processado em Excel
            - Disponibiliza relatório de erros quando necessário

            Ao final, o script apresenta um relatório consolidado na tela e oferece a opção de baixar o resultado em formato Excel.
            """
        )
        processar_notas_fiscais()
        bt = st.button("Processar", disabled=uploaded_extrato is None or uploaded_servicos is None or df_servicos_filtrado is None or len(extrato) == 0)
        if bt:
            processar_extrato()
    except Exception as e:
        st.error(f"Ocorreu um erro inesperado na aplicação: {str(e)}")

if __name__ == "__main__":
    main()
