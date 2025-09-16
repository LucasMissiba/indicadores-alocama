import streamlit as st
import pandas as pd
import plotly.express as px
import time

# Configuração da página
st.set_page_config(
    page_title="Dashboard Alocama - Simplificado",
    page_icon="📊",
    layout="wide"
)

# Tela de carregamento simples
if "loading_complete" not in st.session_state:
    st.session_state["loading_complete"] = False

if not st.session_state["loading_complete"]:
    # CSS para tela de carregamento
    st.markdown("""
    <style>
        .stApp {
            background: #000000 !important;
        }
        .main .block-container {
            padding: 0 !important;
            max-width: 100% !important;
        }
        body {
            overflow: hidden !important;
        }
        .loading-screen {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: #000000;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            z-index: 9999;
            color: white;
            font-family: 'Segoe UI', 'Roboto', sans-serif;
        }
        .title {
            font-size: 4.5rem;
            font-weight: 700;
            background: linear-gradient(45deg, #2563eb, #3b82f6, #60a5fa);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            text-shadow: 0 0 30px rgba(37, 99, 235, 0.5);
            margin-bottom: 1rem;
            letter-spacing: 4px;
            animation: glow 2s ease-in-out infinite alternate;
        }
        .subtitle {
            font-size: 1.3rem;
            color: #e5e7eb;
            margin-bottom: 3rem;
            font-weight: 300;
            letter-spacing: 1px;
            opacity: 0.9;
        }
        .progress-container {
            width: 400px;
            height: 6px;
            background: rgba(26, 26, 26, 0.8);
            border-radius: 3px;
            overflow: hidden;
            margin-bottom: 1.5rem;
            box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.3);
        }
        .progress-bar {
            height: 100%;
            background: linear-gradient(90deg, #2563eb 0%, #3b82f6 50%, #60a5fa 100%);
            border-radius: 3px;
            box-shadow: 0 0 15px rgba(37, 99, 235, 0.6);
            transition: width 0.3s ease;
        }
        .percentage {
            font-size: 1.8rem;
            font-weight: 600;
            color: #2563eb;
            text-shadow: 0 0 10px rgba(37, 99, 235, 0.8);
            margin-bottom: 1rem;
            letter-spacing: 2px;
        }
        .status {
            font-size: 1rem;
            color: #e5e7eb;
            opacity: 0.8;
            font-weight: 300;
        }
        @keyframes glow {
            0% { text-shadow: 0 0 30px rgba(37, 99, 235, 0.5); }
            100% { text-shadow: 0 0 40px rgba(37, 99, 235, 0.8), 0 0 60px rgba(59, 130, 246, 0.4); }
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Container para a tela de carregamento
    loading_container = st.empty()
    
    # Simular carregamento
    for i in range(101):
        if i < 20:
            status = "Inicializando sistema..."
        elif i < 40:
            status = "Carregando dados..."
        elif i < 60:
            status = "Processando informações..."
        elif i < 80:
            status = "Preparando visualizações..."
        else:
            status = "Finalizando carregamento..."
        
        with loading_container.container():
            st.markdown(f"""
            <div class="loading-screen">
                <div class="title">ALOCAMA</div>
                <div class="subtitle">Sistema de Contratos</div>
                <div class="progress-container">
                    <div class="progress-bar" style="width: {i}%;"></div>
                </div>
                <div class="percentage">{i}%</div>
                <div class="status">{status}</div>
            </div>
            """, unsafe_allow_html=True)
        
        time.sleep(0.03)
    
    # Fade out final
    with loading_container.container():
        st.markdown("""
        <div class="loading-screen" style="opacity: 1; transition: opacity 2s ease-out;">
            <div class="title">ALOCAMA</div>
            <div class="subtitle">Sistema de Contratos</div>
            <div class="progress-container">
                <div class="progress-bar" style="width: 100%;"></div>
            </div>
            <div class="percentage">100%</div>
            <div class="status">Carregamento concluído!</div>
        </div>
        """, unsafe_allow_html=True)
    
    time.sleep(1.5)
    st.session_state["loading_complete"] = True
    st.rerun()
    return

# Dashboard principal simplificado
st.title("📊 Dashboard de Contratos | Alocama")
st.subheader("Sistema de Indicadores")

# Dados de exemplo
data = {
    'Mês': ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho'],
    'Faturamento': [100000, 120000, 110000, 130000, 125000, 140000],
    'Vidas': [50, 55, 52, 58, 56, 60]
}

df = pd.DataFrame(data)

# Gráfico de faturamento
st.subheader("💰 Faturamento Mensal")
fig_fat = px.line(df, x='Mês', y='Faturamento', title='Evolução do Faturamento')
st.plotly_chart(fig_fat, use_container_width=True)

# Gráfico de vidas
st.subheader("👥 Vidas Atendidas")
fig_vidas = px.bar(df, x='Mês', y='Vidas', title='Vidas Atendidas por Mês')
st.plotly_chart(fig_vidas, use_container_width=True)

# Métricas
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Faturamento Total", "R$ 725.000", "15%")
with col2:
    st.metric("Vidas Totais", "331", "8%")
with col3:
    st.metric("Ticket Médio", "R$ 2.190", "6%")

# Footer
st.markdown("""
<div style="text-align: center; margin-top: 50px; padding: 20px; background: rgba(255,255,255,0.05); border-radius: 10px;">
    <p><strong>Dashboard desenvolvido por Lucas Missiba</strong></p>
    <p>Alocama · Setor de Contratos</p>
</div>
""", unsafe_allow_html=True)
