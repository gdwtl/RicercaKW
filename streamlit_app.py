import streamlit as st
import pandas as pd
from collections import Counter
import itertools
from io import BytesIO

# Titolo dell'app
st.title("Analisi di Parole, Coppie e Triple")

# Upload del file
uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])

if uploaded_file:
    # Leggi il file Excel
    df = pd.read_excel(uploaded_file)

    # Specifica le colonne da analizzare
    columns_to_analyze = ['Title', 'Description', 'B01', 'B02', 'B03', 'B04', 'B05', 'B06', 'B07']

    # Estrai e concatena i testi
    text_data = []
    for _, row in df.iterrows():
        text_data.append(" ".join(row[columns_to_analyze].dropna().astype(str)))
    text_data = " ".join(text_data)

    # Tokenizza
    words = text_data.split()

    # Conta parole, coppie e triple
    word_counts = Counter(words)
    word_pairs = Counter(zip(words, itertools.islice(words, 1, None)))
    word_triplets = Counter(zip(words, itertools.islice(words, 1, None), itertools.islice(words, 2, None)))

    # Preparazione risultati
    max_len = max(len(word_counts), len(word_pairs), len(word_triplets))
    word_items = word_counts.most_common(max_len)
    pair_items = word_pairs.most_common(max_len)
    triplet_items = word_triplets.most_common(max_len)

    word_items += [(None, None)] * (max_len - len(word_items))
    pair_items += [(None, None)] * (max_len - len(pair_items))
    triplet_items += [(None, None)] * (max_len - len(triplet_items))

    results_df = pd.DataFrame({
        'Words': [word for word, _ in word_items],
        'Word Count': [count for _, count in word_items],
        'Word Pairs': [' '.join(pair) if pair and all(pair) else None for pair, _ in pair_items],
        'Pair Count': [count for _, count in pair_items],
        'Triplets': [' '.join(triplet) if triplet and all(triplet) else None for triplet, _ in triplet_items],
        'Triplet Count': [count for _, count in triplet_items]
    })

    # Mostra il DataFrame
    st.subheader("Risultati dell'Analisi")
    st.dataframe(results_df)

    # Crea un buffer per il download
    output = BytesIO()
    results_df.to_excel(output, index=False)
    output.seek(0)

    # Bottone per scaricare il file
    st.download_button(
        label="📥 Scarica risultati in Excel",
        data=output,
        file_name='word_analysis_results.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
