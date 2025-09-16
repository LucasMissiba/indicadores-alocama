import streamlit as st

st.set_page_config(
    page_title="ALOCAMA - Teste",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🏥 ALOCAMA - Sistema de Contratos")
st.subtitle("Teste de Funcionamento")

st.success("✅ App funcionando corretamente!")

if st.button("Teste de Botão"):
    st.balloons()
