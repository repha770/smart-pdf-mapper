import io
import pdfplumber
import pandas as pd
import streamlit as st

st.set_page_config(page_title="PDF → Excel avec modèle", layout="wide")

st.title("🧾 PDF → Excel avec modèle & mapping manuel")

st.markdown("""
Ce site fait :
1. Tu saisis **tes noms de colonnes** (ce que tu veux dans l'Excel final).
2. Tu uploades un **PDF modèle** pour récupérer sa structure (colonnes).
3. Tu uploades ton **PDF à extraire** (plusieurs pages possibles).
4. Tu **mappe chaque colonne finale** avec une colonne du modèle.
5. Tu télécharges un **Excel** structuré.
""")

# -----------------------------
# FONCTION D'EXTRACTION PDF
# -----------------------------
def extract_tables_from_pdf(uploaded_pdf):
    """Retourne un DataFrame concaténé avec toutes les tables trouvées dans un PDF."""
    all_tables = []

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for t in tables:
                df = pd.DataFrame(t)
                if df.empty:
                    continue
                # On suppose que la première ligne est l'entête
                df.columns = df.iloc[0]
                df = df[1:]
                df["__page__"] = page_number
                all_tables.append(df)

    if not all_tables:
        return None

    df = pd.concat(all_tables, ignore_index=True)
    df.columns = df.columns.astype(str)
    return df


# -----------------------------
# 1) SAISIE DES NOMS DE COLONNES
# -----------------------------
st.subheader("1️⃣ Tes noms de colonnes (pour l'Excel)")

colonnes_input = st.text_input(
    "Saisis les noms de colonnes, séparés par des virgules :",
    value="nom,prenom,telephone,email"
)

colonnes_finales = []
if colonnes_input.strip():
    colonnes_finales = [c.strip() for c in colonnes_input.split(",") if c.strip()]

if colonnes_finales:
    st.success(f"Colonnes finales : {colonnes_finales}")
else:
    st.warning("Saisis au moins une colonne pour continuer.")


# -----------------------------
# 2) PDF MODELE
# -----------------------------
st.subheader("2️⃣ PDF modèle (pour récupérer la structure)")

pdf_modele = st.file_uploader("Choisis le PDF modèle", type=["pdf"])

df_modele = None
colonnes_modele_pdf = []

if pdf_modele is not None:
    df_modele = extract_tables_from_pdf(pdf_modele)
    if df_modele is None or df_modele.empty:
        st.error("Aucune table détectée dans le PDF modèle.")
    else:
        st.write("Aperçu du PDF modèle (tables détectées) :")
        st.dataframe(df_modele.head(30))

        # Colonnes disponibles dans le PDF modèle (on enlève la colonne __page__)
        colonnes_modele_pdf = [c for c in df_modele.columns if c != "__page__"]
        st.info(f"Colonnes détectées dans le modèle : {colonnes_modele_pdf}")


# -----------------------------
# 3) PDF A EXTRAIRE
# -----------------------------
st.subheader("3️⃣ PDF à extraire (plusieurs pages possibles)")

pdf_extract = st.file_uploader("Choisis le PDF à extraire", type=["pdf"])

df_extract = None
if pdf_extract is not None:
    df_extract = extract_tables_from_pdf(pdf_extract)
    if df_extract is None or df_extract.empty:
        st.error("Aucune table détectée dans le PDF à extraire.")
    else:
        st.write("Aperçu du PDF à extraire (tables détectées) :")
        st.dataframe(df_extract.head(30))


# -----------------------------
# 4) MAPPING & EXPORT
# -----------------------------
if colonnes_finales and df_modele is not None and df_extract is not None and colonnes_modele_pdf:
    st.subheader("4️⃣ Mapping de tes colonnes ↔ éléments du PDF modèle")

    # On s'assure que les colonnes du DF extrait sont des strings
    df_extract.columns = df_extract.columns.astype(str)

    # On suppose que la structure des colonnes du PDF extrait
    # est la même que celle du PDF modèle
    options_source = ["-- Aucune --"] + colonnes_modele_pdf

    mapping = {}
    st.markdown("Associe chaque **colonne finale** à une **colonne du PDF modèle** :")

    for col_finale in colonnes_finales:
        # Auto-suggestion si le même nom existe dans le modèle
        default_index = 0
        if col_finale in colonnes_modele_pdf:
            default_index = options_source.index(col_finale)

        choix = st.selectbox(
            f"Source pour la colonne finale **{col_finale}**",
            options_source,
            index=default_index,
            key=f"map_{col_finale}",
        )
        if choix != "-- Aucune --":
            mapping[col_finale] = choix

    if mapping:
        st.write("Mapping utilisé (colonne finale → colonne du modèle) :")
        st.json(mapping)

        # Construction du DataFrame final, dans l'ordre de TES colonnes
        df_final = pd.DataFrame()
        for col_finale in colonnes_finales:
            if col_finale in mapping:
                src = mapping[col_finale]
                # On prend la colonne correspondante dans le PDF extrait
                if src in df_extract.columns:
                    df_final[col_finale] = df_extract[src].astype(str).fillna("")
                else:
                    # Si jamais la colonne n'existe pas dans l'extrait, on met vide
                    df_final[col_finale] = ""
            else:
                df_final[col_finale] = ""

        st.subheader("Aperçu du résultat final (Excel)")
        st.dataframe(df_final.head(50))

        # Export Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Données")
        buffer.seek(0)

        st.download_button(
            label="📥 Télécharger l'Excel final",
            data=buffer,
            file_name="export_pdf_modele.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Mappe au moins une colonne pour pouvoir générer l'Excel.")
elif pdf_extract is not None and (not colonnes_finales or df_modele is None):
    st.info("Il manque soit tes colonnes finales, soit le PDF modèle, soit les deux.")
