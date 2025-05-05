import streamlit as st
import pandas as pd
from collections import Counter
import itertools
from io import BytesIO
import time

# Configura layout e font
st.set_page_config(page_title="Analisi Testuale", layout="wide")
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins&display=swap');
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
        background-color: #f8f9fb;
    }
    .metric-box {
        padding: 2rem;
        border-radius: 15px;
        background-color: white;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🔍 Analisi di Parole, Coppie e Triple")

uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])

if uploaded_file:
    start_time = time.time()

    df = pd.read_excel(uploaded_file)
    columns_to_analyze = ['Title', 'Description', 'B01', 'B02', 'B03', 'B04', 'B05', 'B06', 'B07']

    missing = [col for col in columns_to_analyze if col not in df.columns]
    if missing:
        st.error(f"Colonne mancanti: {', '.join(missing)}")
        st.stop()

    # Estrai testo
    text_data = []
    for _, row in df.iterrows():
        text_data.append(" ".join(row[columns_to_analyze].dropna().astype(str)))
    text_data = " ".join(text_data)

    # Tokenizza
    words = text_data.split()
    word_count = len(words)

    # Frequenze
    word_counts = Counter(words)
    word_pairs = Counter(zip(words, itertools.islice(words, 1, None)))
    word_triplets = Counter(zip(words, itertools.islice(words, 1, None), itertools.islice(words, 2, None)))

    # Lunghezza uniforme per l'export
    max_len = max(len(word_counts), len(word_pairs), len(word_triplets))
    word_items = word_counts.most_common(max_len)
    pair_items = word_pairs.most_common(max_len)
    triplet_items = word_triplets.most_common(max_len)

    # Padding
    word_items += [(None, None)] * (max_len - len(word_items))
    pair_items += [(None, None)] * (max_len - len(pair_items))
    triplet_items += [(None, None)] * (max_len - len(triplet_items))

    # DataFrame
    results_df = pd.DataFrame({
        'Parole': [w for w, _ in word_items],
        'Frequenza': [c for _, c in word_items],
        'Coppie': [' '.join(p) if p and all(p) else None for p, _ in pair_items],
        'Frequenza Coppie': [c for _, c in pair_items],
        'Triple': [' '.join(t) if t and all(t) else None for t, _ in triplet_items],
        'Frequenza Triple': [c for _, c in triplet_items],
    })

    # Tempo
    elapsed = round(time.time() - start_time, 2)

    # Box risultati
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="metric-box">
            <h3>🔤 Parole Analizzate</h3>
            <h1>{word_count}</h1>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-box">
            <h3>⏱️ Tempo di Analisi</h3>
            <h1>{elapsed} sec</h1>
        </div>
        """, unsafe_allow_html=True)

    # Anteprima
    st.subheader("📄 Anteprima Risultati")
    st.dataframe(results_df.head(20))

    # Download Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        results_df.to_excel(writer, index=False)
    output.seek(0)

    st.download_button(
        label="📥 Scarica risultati in Excel",
        data=output,
        file_name='analisi_testuale.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
