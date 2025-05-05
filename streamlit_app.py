import streamlit as st
import pandas as pd
import time
from collections import Counter
import itertools

# Configura layout e font
st.set_page_config(page_title="Word Analysis", layout="wide")

# Inietta Google Font Poppins + supporto per tema chiaro/scuro
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins&display=swap');

    html, body, [class*="css"]  {
        font-family: 'Poppins', sans-serif;
        transition: all 0.3s ease;
    }

    .st-dark {
        background-color: #1e1e1e !important;
    }

    .metric-box {
        padding: 2rem;
        border-radius: 15px;
        background-color: white;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 2rem;
    }

    [data-theme="dark"] .metric-box {
        background-color: #2d2d2d;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# Selettore per il tema
theme = st.radio("Tema", ["Chiaro", "Scuro"], horizontal=True)

# Forza la classe CSS in base alla scelta (usando hack JS/CSS)
st.markdown(f"""
    <script>
    const root = window.parent.document.documentElement;
    root.setAttribute('data-theme', '{'dark' if theme == "Scuro" else 'light'}');
    </script>
""", unsafe_allow_html=True)

# File uploader
st.title("Analisi Parole")
uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])

if uploaded_file:
    start_time = time.time()

    df = pd.read_excel(uploaded_file)

    # Colonne da analizzare
    columns_to_analyze = ['Title', 'Description', 'B01', 'B02', 'B03', 'B04', 'B05', 'B06', 'B07']

    missing = [col for col in columns_to_analyze if col not in df.columns]
    if missing:
        st.error(f"Colonne mancanti: {', '.join(missing)}")
        st.stop()

    # Estrai e concatena testo
    text_data = [
        " ".join(row[columns_to_analyze].dropna().astype(str))
        for _, row in df.iterrows()
    ]
    all_text = " ".join(text_data)

    # Tokenizza
    words = all_text.split()
    word_count = len(words)
    elapsed_time = round(time.time() - start_time, 2)

    # Mostra risultati
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="metric-box">
            <h2 style='margin: 0;'>🔤 Parole Analizzate</h2>
            <h1 style='margin: 0; font-size: 3rem;'>{word_count}</h1>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
        <div class="metric-box">
            <h2 style='margin: 0;'>⏱️ Tempo Impiegato</h2>
            <h1 style='margin: 0; font-size: 3rem;'>{elapsed_time} sec</h1>
        </div>
        """, unsafe_allow_html=True)
