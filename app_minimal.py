import streamlit as st

st.set_page_config(
    page_title="Dashboard Alocama",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard de Contratos | Alocama")
st.subheader("Sistema de Indicadores")

st.success("✅ App funcionando no Streamlit Cloud!")

st.info("Se você está vendo esta mensagem, o deploy está funcionando perfeitamente.")

# Dados de exemplo simples
import pandas as pd
import plotly.express as px

data = {
    'Mês': ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun'],
    'Faturamento': [100, 120, 110, 130, 125, 140]
}

df = pd.DataFrame(data)

# Gráfico simples
fig = px.bar(df, x='Mês', y='Faturamento', title='Faturamento Mensal')
st.plotly_chart(fig, use_container_width=True)

# Métricas
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total", "R$ 725.000")
with col2:
    st.metric("Crescimento", "15%")
with col3:
    st.metric("Status", "✅ Ativo")

st.markdown("""
<div style="text-align: center; margin-top: 50px; padding: 20px;">
    <p><strong>Dashboard desenvolvido por Lucas Missiba</strong></p>
    <p>Alocama · Setor de Contratos</p>
</div>
""", unsafe_allow_html=True)
