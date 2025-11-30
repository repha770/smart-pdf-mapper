import io
import pdfplumber
import pandas as pd
import streamlit as st

st.set_page_config(page_title="PDF → Excel avec modèle", layout="wide")

st.title("🧾 PDF → Excel basé sur un modèle de colonnes")

st.markdown("""
Ce site fait :
1. Upload **PDF modèle** (optionnel, juste pour toi).
2. Upload **modèle de colonnes** (Excel/CSV ou saisie manuelle).
3. Upload **PDF à extraire**.
4. Mapping colonnes modèle ↔ colonnes extraites.
5. Export en **Excel** propre selon ton modèle.
""")


def extract_tables_from_pdf(uploaded_pdf):
    """Retourne un DataFrame concaténé avec toutes les tables trouvées dans un PDF Streamlit."""
    all_tables = []

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for t in tables:
                df = pd.DataFrame(t)
                if df.empty:
                    continue
                # On suppose que la 1ère ligne = entêtes
                df.columns = df.iloc[0]
                df = df[1:]
                df["__page__"] = page_number
                all_tables.append(df)

    if not all_tables:
        return None

    df = pd.concat(all_tables, ignore_index=True)
    df.columns = df.columns.astype(str)
    return df


# ---------- 1) PDF MODELE (juste pour info) ----------
st.subheader("1️⃣ PDF modèle (optionnel)")
pdf_model = st.file_uploader("PDF modèle (facultatif)", type=["pdf"])
if pdf_model is not None:
    st.info("PDF modèle chargé (il sert uniquement de référence visuelle, pas d'extraction).")


# ---------- 2) MODELE DE COLONNES ----------
st.subheader("2️⃣ Modèle de colonnes")

col_file = st.file_uploader(
    "Fichier modèle de colonnes (Excel ou CSV). Sinon, laisse vide et saisis à la main en dessous.",
    type=["xlsx", "xls", "csv"]
)

colonnes_modele = []

if col_file is not None:
    ext = col_file.name.split(".")[-1].lower()
    if ext in ["xlsx", "xls"]:
        df_cols = pd.read_excel(col_file)
    else:
        df_cols = pd.read_csv(col_file)

    st.write("Aperçu du fichier de colonnes :")
    st.dataframe(df_cols.head())

    mode_cols = st.radio(
        "Comment récupérer les colonnes du modèle ?",
        ["Utiliser les en-têtes du fichier", "Utiliser les valeurs d'une colonne"],
        horizontal=True,
    )

    if mode_cols == "Utiliser les en-têtes du fichier":
        colonnes_modele = list(df_cols.columns)
    else:
        col_select = st.selectbox(
            "Colonne contenant la liste des noms de colonnes",
            df_cols.columns
        )
        colonnes_modele = (
            df_cols[col_select]
            .dropna()
            .astype(str)
            .tolist()
        )
else:
    manuel = st.text_input(
        "Ou saisis manuellement les noms de colonnes (séparés par des virgules) :",
        value="nom,prenom,telephone,email"
    )
    if manuel.strip():
        colonnes_modele = [c.strip() for c in manuel.split(",") if c.strip()]

if colonnes_modele:
    st.success(f"Colonnes du modèle : {colonnes_modele}")
else:
    st.warning("Aucune colonne modèle pour l'instant.")


# ---------- 3) PDF A EXTRAIRE ----------
st.subheader("3️⃣ PDF à extraire")

pdf_extract = st.file_uploader("PDF à extraire", type=["pdf"])

df_extrait = None
if pdf_extract is not None:
    df_extrait = extract_tables_from_pdf(pdf_extract)
    if df_extrait is None or df_extrait.empty:
        st.error("Aucune table détectée dans ce PDF.")
    else:
        st.write("Aperçu des données extraites du PDF :")
        st.dataframe(df_extrait.head(50))


# ---------- 4) MAPPING & EXPORT ----------
if df_extrait is not None and colonnes_modele:
    st.subheader("4️⃣ Mapping colonnes modèle ↔ colonnes extraites & export Excel")

    df_extrait.columns = df_extrait.columns.astype(str)
    colonnes_extraites = list(df_extrait.columns)

    st.write("Colonnes extraites du PDF :")
    st.write(colonnes_extraites)

    st.markdown("### Associe chaque colonne du modèle à une colonne extraite")

    mapping = {}
    options_source = ["-- Aucune --"] + colonnes_extraites

    for col_mod in colonnes_modele:
        # tentative d'auto-match si le nom existe déjà
        default_index = 0
        if col_mod in colonnes_extraites:
            default_index = options_source.index(col_mod)

        choix = st.selectbox(
            f"Source pour la colonne modèle **{col_mod}**",
            options_source,
            index=default_index,
        )
        if choix != "-- Aucune --":
            mapping[col_mod] = choix

    if mapping:
        st.write("Mapping utilisé :")
        st.json(mapping)

        # construction du DF final, dans l'ordre du modèle
        df_final = pd.DataFrame()
        for col_mod in colonnes_modele:
            if col_mod in mapping:
                src = mapping[col_mod]
                df_final[col_mod] = df_extrait[src].astype(str).fillna("")
            else:
                df_final[col_mod] = ""

        st.subheader("Aperçu du résultat final")
        st.dataframe(df_final.head(50))

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Données")
        buffer.seek(0)

        st.download_button(
            label="📥 Télécharger l'Excel final",
            data=buffer,
            file_name="export_modele_pdf.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Configure au moins une colonne dans le mapping pour générer l'Excel.")
elif pdf_extract is not None and not colonnes_modele:
    st.info("Tu as chargé le PDF à extraire, mais pas encore le modèle de colonnes.")
