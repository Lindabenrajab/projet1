import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import tempfile
import io

st.set_page_config(page_title="Analyse ventes contrat", layout="wide")
st.title("📊 Analyse des Ventes - Contrats et Assurances")

uploaded_file = st.file_uploader("📂 Téléchargez votre fichier Excel ou CSV", type=["xlsx", "xls", "csv"])

def detect_column(possible_names, df_cols):
    for name in possible_names:
        for col in df_cols:
            if name.lower().strip() == col.lower().strip():
                return col
    return None

def create_pdf_with_images(summary_text, image_buffers):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for line in summary_text.split('\n'):
        pdf.cell(0, 10, line, ln=True)

    for img_buf in image_buffers:
        tmp_img = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp_img.write(img_buf.getbuffer())
        tmp_img.close()
        pdf.image(tmp_img.name, w=180)
        pdf.ln(10)

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    pdf.output(tmp_pdf.name)
    return tmp_pdf.name

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.subheader("Aperçu du fichier")
        st.dataframe(df.head())

        synonyms = {
            "date": ["date", "date début", "date de vente"],
            "revenu": ["prime total ttc", "revenu", "ca"],
            "marge": ["marge", "marge distributeur ttc"],
            "produit": ["produit", "device", "type", "categorie"],
            "assureur": ["part assureur", "part", "taux"],
            "distributeur": ["distributeur", "revendeur", "client", "point de vente"]
        }

        df_cols = list(df.columns)
        col_date = detect_column(synonyms["date"], df_cols)
        col_revenu = detect_column(synonyms["revenu"], df_cols)
        col_marge = detect_column(synonyms["marge"], df_cols)
        col_produit = detect_column(synonyms["produit"], df_cols)
        col_assureur = detect_column(synonyms["assureur"], df_cols)
        col_distrib = detect_column(synonyms["distributeur"], df_cols)

        st.markdown("### 🔧 Confirmez ou ajustez les colonnes")
        col_date = st.selectbox("Colonne de Date", df_cols, index=df_cols.index(col_date) if col_date else 0)
        col_revenu = st.selectbox("Colonne de Revenu", df_cols, index=df_cols.index(col_revenu) if col_revenu else 0)
        col_marge = st.selectbox("Colonne de Marge Distributeur", df_cols, index=df_cols.index(col_marge) if col_marge else 0)
        col_produit = st.selectbox("Colonne de Produit / Device", df_cols, index=df_cols.index(col_produit) if col_produit else 0)
        col_assureur = st.selectbox("Colonne Part Assureur", df_cols, index=df_cols.index(col_assureur) if col_assureur else 0)
        col_distrib = st.selectbox("Colonne de Distributeur", df_cols, index=df_cols.index(col_distrib) if col_distrib else 0)

        df["Date"] = pd.to_datetime(df[col_date], dayfirst=True, errors='coerce')
        df["Revenu"] = pd.to_numeric(df[col_revenu], errors='coerce')
        df["Marge"] = pd.to_numeric(df[col_marge], errors='coerce')
        df["Produit"] = df[col_produit].astype(str)
        df["Assureur"] = pd.to_numeric(df[col_assureur], errors='coerce')
        df["Distributeur"] = df[col_distrib].astype(str)

        st.sidebar.header("🎛️ Filtres")
        min_date, max_date = df["Date"].min().date(), df["Date"].max().date()
        start_date = st.sidebar.date_input("🗓️ Date début", min_value=min_date, max_value=max_date, value=min_date)
        end_date = st.sidebar.date_input("📅 Date fin", min_value=min_date, max_value=max_date, value=max_date)

        produits_dispo = df["Produit"].unique().tolist()
        selected_produits = st.sidebar.multiselect("📦 Produits à afficher", produits_dispo, default=produits_dispo)

        distributeurs_dispo = df["Distributeur"].unique().tolist()
        selected_distributeurs = st.sidebar.multiselect("🏪 Distributeurs à afficher", distributeurs_dispo, default=distributeurs_dispo)

        df_filtered = df[
            (df["Date"].dt.date >= start_date) &
            (df["Date"].dt.date <= end_date) &
            (df["Produit"].isin(selected_produits)) &
            (df["Distributeur"].isin(selected_distributeurs))
        ]

        st.markdown("### 📌 Résumé détaillé de l'activité")
        total_revenu = df_filtered["Revenu"].sum()
        total_marge = df_filtered["Marge"].sum()
        nb_contrats = len(df_filtered)
        revenu_moyen = df_filtered["Revenu"].mean()
        marge_moyenne = df_filtered["Marge"].mean()
        date_min = df_filtered["Date"].min().date()
        date_max = df_filtered["Date"].max().date()
        nb_jours = (date_max - date_min).days + 1
        top_produit = df_filtered.groupby("Produit")["Revenu"].sum().sort_values(ascending=False).head(1).index[0]

        row1 = st.columns(4)
        row1[0].metric("💰 Revenu Total", f"{total_revenu:,.2f} TND")
        row1[1].metric("📈 Marge Totale", f"{total_marge:,.2f} TND")
        row1[2].metric("📊 Revenu Moyen / Contrat", f"{revenu_moyen:,.2f} TND")
        row1[3].metric("💼 Marge Moyenne / Contrat", f"{marge_moyenne:,.2f} TND")

        row2 = st.columns(4)
        row2[0].metric("📑 Nombre de Contrats", nb_contrats)
        row2[1].metric("🗓️ Période", f"{date_min} ➜ {date_max}")
        row2[2].metric("📆 Jours couverts", nb_jours)
        row2[3].metric("🏆 Top Produit", top_produit)

        st.markdown("### 📆 Évolution du revenu par date")
        revenu_par_jour = df_filtered.groupby(df_filtered["Date"].dt.date)["Revenu"].sum()
        st.line_chart(revenu_par_jour)

        st.markdown("### 🥇 Top 10 Produits par Revenu")
        top_produits = df_filtered.groupby("Produit")["Revenu"].sum().sort_values(ascending=False).head(10)
        fig1, ax1 = plt.subplots()
        sns.barplot(x=top_produits.values, y=top_produits.index, ax=ax1, palette="crest")
        ax1.set_xlabel("Revenu (TND)")
        st.pyplot(fig1)

        st.markdown("### 🧁 Répartition des revenus par produit")
        revenus_par_produit = df_filtered.groupby("Produit")["Revenu"].sum()
        fig2, ax2 = plt.subplots()
        ax2.pie(revenus_par_produit, labels=revenus_par_produit.index, autopct="%1.1f%%", startangle=90)
        ax2.axis('equal')
        st.pyplot(fig2)

        st.markdown("### 🏦 Moyenne des Parts Assureurs")
        assureur_avg = df_filtered["Assureur"].mean()
        st.info(f"📊 Moyenne Part Assureur : {assureur_avg:.2f} %")

        # Export Excel
        st.markdown("### 📤 Exporter les données filtrées")

        @st.cache_data
        def to_excel(data):
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                data.to_excel(writer, index=False, sheet_name='Analyse')
            output.seek(0)
            return output

        excel_data = to_excel(df_filtered)

        st.download_button(
            label="📥 Télécharger en Excel",
            data=excel_data,
            file_name="analyse_ventes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Rapport PDF complet (texte + graphiques)
        st.markdown("### 🧾 Générer un rapport PDF complet (texte + graphiques)")

        summary_text = (
            f"Rapport de Ventes\n"
            f"Période : {start_date} à {end_date}\n"
            f"Revenu Total : {total_revenu:,.2f} TND\n"
            f"Marge Totale : {total_marge:,.2f} TND\n"
            f"Nombre de Contrats : {nb_contrats}\n"
            f"Top Produit : {top_produit}\n"
        )

        # Préparer les images en mémoire pour PDF
        buf1 = io.BytesIO()
        fig1.savefig(buf1, format="png")
        plt.close(fig1)
        buf1.seek(0)

        buf2 = io.BytesIO()
        fig2.savefig(buf2, format="png")
        plt.close(fig2)
        buf2.seek(0)

        if st.button("📄 Télécharger le rapport PDF complet"):
            pdf_path = create_pdf_with_images(summary_text, [buf1, buf2])
            with open(pdf_path, "rb") as f:
                st.download_button(
                    label="📥 Télécharger le PDF",
                    data=f.read(),
                    file_name="rapport_complet_ventes.pdf",
                    mime="application/pdf"
                )

    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture ou de l’analyse du fichier : {e}")

else:
    st.info("🕐 Veuillez uploader un fichier Excel ou CSV pour commencer.")
