# =================================================================
# PASSO 1 — PRIMEIRO IMPORTAR TUDO
# =================================================================
import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

# =================================================================
# PASSO 2 — CONFIGURAÇÃO DA PÁGINA
# =================================================================
st.set_page_config(page_title="Emissor de Memorandos", page_icon="📄", layout="wide")

# =================================================================
# PASSO 3 — AGORA SIM O CSS DEPOIS DO st.set_page_config
# =================================================================
st.markdown("""
<style>
/* ===== BARRA DE ROLAGEM MAIS VISÍVEL ===== */
::-webkit-scrollbar {
    width: 14px !important;
    height: 14px !important;
}

::-webkit-scrollbar-track {
    background: #2c3e50 !important;
    border-radius: 7px !important;
}

::-webkit-scrollbar-thumb {
    background: #3498db !important;
    border-radius: 7px !important;
    border: 2px solid #2c3e50 !important;
}

::-webkit-scrollbar-thumb:hover {
    background: #2980b9 !important;
}

/* ===== DESTAQUE NOS EXPANDERS ===== */
div[data-testid="stExpander"] {
    border: 1px solid #4a5568 !important;
    border-radius: 8px !important;
    margin-bottom: 12px !important;
    background: #1e293b !important;
}

div[data-testid="stExpander"]:hover {
    border-color: #3498db !important;
    box-shadow: 0 2px 8px rgba(52, 152, 219, 0.15) !important;
}

/* ===== BOTÕES MAIS DESTAQUES ===== */
button[kind="primary"] {
    background: #3498db !important;
    border: none !important;
    font-weight: 500 !important;
}

button[kind="primary"]:hover {
    background: #2980b9 !important;
}

/* ===== MELHORAR LEITURA GERAL ===== */
.block-container {
    max-width: 100% !important;
    padding-top: 1.5rem !important;
    padding-bottom: 3rem !important;
}

/* ===== LINHA DE SEPARAÇÃO MAIS VISÍVEL ===== */
hr {
    border: none !important;
    border-top: 1px solid #4a5568 !important;
    margin: 1.5rem 0 !important;
}
</style>
""", unsafe_allow_html=True)

# =================================================================
# RESTO DO SEU CÓDIGO AQUI
# =================================================================
# [cole aqui o restante do código — funções, interface, etc.]
