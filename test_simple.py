import streamlit as st

st.set_page_config(
    page_title="Teste Simples",
    page_icon="📊",
    layout="wide"
)

st.title("🚀 Teste de Deploy Streamlit Cloud")
st.write("Este é um teste simples para verificar se o deploy funciona.")

if st.button("Clique aqui"):
    st.success("✅ Botão funcionando!")

st.info("Se você está vendo isso, o deploy está funcionando!")
