import streamlit as st
import time

# Título do app
st.title("Formulário Simples")

# Início do formulário
with st.form(key='formulario_usuarios'):
    nome = st.text_input("Nome")
    email = st.text_input("E-mail")
    mensagem = st.text_area("Mensagem")
    botao_enviar = st.form_submit_button("Enviar")

# Processando os dados do formulário
if botao_enviar:
    st.success(f"Obrigado, {nome}! Sua mensagem foi enviada com sucesso.")
    st.write("Detalhes enviados:")
    st.write(f"- Nome: {nome}")
    st.write(f"- E-mail: {email}")
    st.write(f"- Mensagem: {mensagem}")
    time.sleep(3)
    st.rerun()
