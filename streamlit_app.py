import streamlit as st
import pandas as pd
from collections import Counter
import itertools
from io import BytesIO
import time

# Configura pagina e font
st.set_page_config(page_title="Analisi Testuale", layout="wide")

# Font Poppins e stile base
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins&display=swap');
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
    }

    .kpi-box {
        padding: 2rem;
        border-radius: 1rem;
        border: 1px solid rgba(128, 128, 128, 0.1);
        background-color: rgba(255, 255, 255, 0.02);
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        text-align: center;
        margin-bottom: 1rem;
    }

    .kpi-label {
        font-size: 1rem;
        color: inherit;
        margin-bottom: 0.5rem;
    }

    .kpi-value {
        font-size: 2.6rem;
        font-weight: 600;
        color: inherit;
    }
    </style>
""", unsafe_allow_html=True)

# Titolo
st.title("🔍 Analisi di Parole, Coppie e Triple")

# Upload
uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])

if uploaded_file:
    start_time = time.time()

    # Leggi Excel
    df = pd.read_excel(uploaded_file)
    columns_to_analyze = ['Title', 'Description', 'B01', 'B02', 'B03', 'B04', 'B05', 'B06', 'B07']

    # Controlla colonne mancanti
    missing = [col for col in columns_to_analyze if col not in df.columns]
    if missing:
        st.error(f"Colonne mancanti: {', '.join(missing)}")
        st.stop()

    # Estrai testo
    text_data = [
        " ".join(row[columns_to_analyze].dropna().astype(str))
        for _, row in df.iterrows()
    ]
    text = " ".join(text_data)

    # Tokenizza
    words = text.split()
    word_count = len(words)

    # Conta parole, coppie, triple
    word_counts = Counter(words)
    word_pairs = Counter(zip(words, itertools.islice(words, 1, None)))
    word_triplets = Counter(zip(words, itertools.islice(words, 1, None), itertools.islice(words, 2, None)))

    # Prepara struttura uniforme
    max_len = max(len(word_counts), len(word_pairs), len(word_triplets))
    word_items = word_counts.most_common(max_len)
    pair_items = word_pairs.most_common(max_len)
    triplet_items = word_triplets.most_common(max_len)

    word_items += [(None, None)] * (max_len - len(word_items))
    pair_items += [(None, None)] * (max_len - len(pair_items))
    triplet_items += [(None, None)] * (max_len - len(triplet_items))

    # Crea DataFrame risultati
    results_df = pd.DataFrame({
        'Parole': [w for w, _ in word_items],
        'Frequenza': [c for _, c in word_items],
        'Coppie': [' '.join(p) if p and all(p) else None for p, _ in pair_items],
        'Frequenza Coppie': [c for _, c in pair_items],
        'Triple': [' '.join(t) if t and all(t) else None for t, _ in triplet_items],
        'Frequenza Triple': [c for _, c in triplet_items],
    })

    elapsed = round(time.time() - start_time, 2)

    # KPI Box
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
            <div class="kpi-box">
                <div class="kpi-label">🔤 Parole Analizzate</div>
                <div class="kpi-value">{word_count}</div>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
            <div class="kpi-box">
                <div class="kpi-label">⏱️ Tempo Impiegato</div>
                <div class="kpi-value">{elapsed} sec</div>
            </div>
        """, unsafe_allow_html=True)

    # Anteprima risultati
    st.subheader("📄 Anteprima Risultati")
    st.dataframe(results_df.head(20))

    # Download Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        results_df.to_excel(writer, index=False, sheet_name='Analisi')
    output.seek(0)

    st.download_button(
        label="📥 Scarica risultati in Excel",
        data=output,
        file_name="analisi_testuale.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
