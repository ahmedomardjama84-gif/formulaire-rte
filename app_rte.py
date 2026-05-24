"""
Formulaire collaboratif - Grille de maturite RTE
Application Streamlit : tes collegues remplissent la grille en ligne,
tu vois la synthese en temps reel et tu exportes le bilan en Word.

Lancement local :
    streamlit run app_rte.py

Partage public (sans inscription) :
    cloudflared tunnel --url http://localhost:8501
"""

import sqlite3
import json
from datetime import datetime, date
from io import BytesIO
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# CONFIGURATION
# ============================================================

DB_PATH = "reponses_rte.db"
APP_TITLE = "Grille de maturite RTE - Formulaire collaboratif"
APP_ICON = "📊"

# Couleurs par axe (coherent avec la grille)
AXE_COLORS = {
    1: "#C28533",  # orange
    2: "#1F4E79",  # bleu
    3: "#256D5A",  # vert
    4: "#5A3B7A",  # violet
    5: "#9C2A2A",  # rouge
}

# ============================================================
# DEFINITION DE LA GRILLE
# ============================================================

AXES = {
    1: "Diagnostic territorial",
    2: "Vision territoriale et recits",
    3: "Cooperation territoriale",
    4: "Gouvernance inclusive",
    5: "Redistribuer et transformer",
}

NIVEAUX_LABELS = {
    1: "1 - Tres faible / inexistant",
    2: "2 - Faible / ponctuel",
    3: "3 - Moyen / structure",
    4: "4 - Fort / strategique",
    0: "Non evalue",
}

NIVEAUX_DESC = {
    1: "Aucune dynamique relationnelle avec le territoire. La dimension territoriale est absente de la strategie.",
    2: "Des interactions existent mais restent ponctuelles, portees par des individus isoles. Territoire percu comme espace de localisation.",
    3: "Processus d'ancrage territorial engage et formalise. Des elements manquent encore pour le niveau strategique.",
    4: "Le territoire est au coeur de la strategie. Entreprendre en collectif et en responsabilite pour le bien commun.",
}

CRITERES = [
    # (id, axe, titre)
    ("1.1", 1, "Specification et diagnostic du territoire"),
    ("1.2", 1, "Modelisation des flux territoriaux"),
    ("1.3", 1, "Redevabilites : double materialite (impact territoire <-> organisation)"),
    ("2.1", 2, "Vision territoriale d'ensemble et vision sectorielle"),
    ("2.2", 2, "Besoins des parties prenantes internes et des usagers"),
    ("2.3", 2, "Besoins des communautes peripheriques"),
    ("2.4", 2, "Integration du vivant et des ecosystemes"),
    ("2.5", 2, "Recits de territoire et identite culturelle"),
    ("3.1", 3, "Culture et competences de cooperation territoriale"),
    ("3.2", 3, "Demarches collectives multi-echelles et partenariats formalises"),
    ("4.1", 4, "Gouvernance interne inclusive et transparente"),
    ("4.2", 4, "Processus de concertation et de dialogue"),
    ("4.3", 4, "Evaluation collective et redevabilite partagee"),
    ("5.1", 5, "Reorientation des echanges vers l'economie locale"),
    ("5.2", 5, "Partage de la valeur produite"),
    ("5.3", 5, "Inclusion des communautes peripheriques"),
    ("5.4", 5, "Bilan des transformations : contribution au bien commun"),
]

# Checklists par critere (items justificatifs)
CHECKLISTS = {
    "1.1": [
        "Un diagnostic territorial formalise est-il disponible (ecosysteme, etendue, diversite) ?",
        "Le diagnostic integre-t-il des sources officielles (INSEE, DRAAF, ARS, ORS, etc.) ?",
        "Le diagnostic est-il mis a jour regulierement (au moins tous les 3 ans) ?",
        "Le diagnostic est-il partage avec les parties prenantes du territoire ?",
    ],
    "1.2": [
        "Une cartographie des flux entrants/sortants (economiques, humains, materiels) existe-t-elle ?",
        "L'organisation sait-elle quels metiers et emplois font vivre le territoire ?",
        "L'organisation a-t-elle etabli une cartographie des acteurs locaux ?",
        "L'organisation suit-elle dans le temps la part de son activite qui est locale ?",
    ],
    "1.3": [
        "L'impact du territoire sur l'organisation est-il documente (ressources, contraintes, opportunites) ?",
        "L'impact de l'organisation sur le territoire est-il documente ?",
        "Un bilan de ces redevabilites bidirectionnelles est-il produit ?",
        "Ces redevabilites sont-elles discutees en gouvernance collective ?",
    ],
    "2.1": [
        "Une vision territoriale d'ensemble est-elle referencee par l'organisation ?",
        "Une vision sectorielle propre a l'organisation est-elle formalisee ?",
        "L'articulation entre vision d'ensemble et vision sectorielle est-elle explicite ?",
        "La vision est-elle co-construite avec les parties prenantes ?",
        "La vision integre-t-elle les transitions ecologiques, sociales, economiques et politiques ?",
    ],
    "2.2": [
        "Les besoins des salaries sont-ils identifies et documentes ?",
        "Les besoins des usagers (beneficiaires directs) sont-ils identifies et documentes ?",
        "Des dispositifs de recueil des besoins existent-ils (enquetes, entretiens, ateliers) ?",
        "Les besoins identifies influencent-ils les orientations strategiques ?",
    ],
    "2.3": [
        "Les populations saisonnieres sont-elles prises en compte dans la strategie ?",
        "Les publics fragiles beneficient-ils de dispositifs dedies ?",
        "Les jeunes sont-ils associes aux demarches de l'organisation ?",
        "Les acteurs eloignes ou peu visibles (ESS, artisans, collectifs) sont-ils identifies et associes ?",
    ],
    "2.4": [
        "Les milieux naturels et la biodiversite sont-ils identifies comme enjeux strategiques ?",
        "Un tableau de bord environnemental existe-t-il ?",
        "Des actions d'attenuation des changements climatiques sont-elles engagees ?",
        "Des actions d'adaptation aux changements climatiques sont-elles engagees ?",
        "Des partenariats existent-ils avec des acteurs de l'environnement (parcs, ONF, LPO, etc.) ?",
    ],
    "2.5": [
        "Le projet mobilise-t-il explicitement des recits de territoire ?",
        "L'identite culturelle locale (patrimoine materiel et immateriel) est-elle integree ?",
        "Les recits sont-ils co-construits avec les acteurs locaux ?",
        "Les recits incluent-ils la diversite des acteurs ?",
    ],
    "3.1": [
        "L'organisation a-t-elle une culture explicite de cooperation ?",
        "Les salaries sont-ils formes aux pratiques cooperatives ?",
        "L'organisation promeut-elle une communication horizontale et transversale ?",
        "Les pratiques d'inter-organisation (ateliers communs, temps partages) sont-elles regulieres ?",
    ],
    "3.2": [
        "L'organisation participe-t-elle a au moins un autre dispositif local ?",
        "L'organisation participe-t-elle a au moins 1 dispositif a l'echelle departementale ?",
        "L'organisation participe-t-elle a au moins 1 dispositif a l'echelle regionale ?",
        "L'organisation est-elle impliquee dans des demarches nationales ou europeennes ?",
        "Le role de l'organisation dans ces dispositifs est-il documente ?",
        "L'organisation a-t-elle formalise par convention ses partenariats ?",
    ],
    "4.1": [
        "La gouvernance inclut-elle au moins 3 categories de parties prenantes differentes ?",
        "Les parties prenantes representees couvrent-elles l'ensemble du territoire d'action ?",
        "Les regles de decision sont-elles documentees et transparentes ?",
        "Les salaries sont-ils representes dans les instances decisionnelles ?",
    ],
    "4.2": [
        "Des espaces de concertation formalises existent-ils ?",
        "La frequence de concertation est-elle reguliere (au moins 2 fois par an) ?",
        "Les concertations sont-elles ouvertes aux acteurs au-dela des seuls adherents ?",
        "Une restitution des concertations est-elle produite et diffusee ?",
    ],
    "4.3": [
        "L'organisation produit-elle un bilan territorial qui rend compte au territoire de son impact ?",
        "Le bilan mobilise-t-il des indicateurs verifiables ?",
        "Le bilan territorial est-il co-construit avec les parties prenantes ?",
        "Le bilan territorial est-il rendu public et debattu avec les acteurs du territoire ?",
    ],
    "5.1": [
        "Les fournisseurs locaux sont-ils prioritaires dans les achats ?",
        "Les acteurs de l'ESS sont-ils prioritaires dans les echanges ?",
        "La question de l'emploi sur la filiere locale est-elle documentee ?",
        "Un suivi chiffre des achats locaux est-il produit annuellement ?",
    ],
    "5.2": [
        "La valeur produite est-elle explicitee ?",
        "Le partage de la valeur entre les parties prenantes a-t-il ete discute ?",
        "L'organisation contribue-t-elle au financement de programmes partages sur le territoire ?",
        "Un dispositif philanthropique (mecenat, dons, fondation) est-il engage ?",
        "Les benefices ou excedents sont-ils redistribues vers le territoire ?",
    ],
    "5.3": [
        "Des actions d'inclusion economique ciblent-elles les communautes peripheriques ?",
        "Des actions d'inclusion sociale ciblent-elles ces communautes ?",
        "Les communautes peripheriques sont-elles associees a la conception des actions qui les concernent ?",
        "Des partenariats operationnels existent-ils avec France Travail / France Services / Missions Locales ?",
    ],
    "5.4": [
        "Transformation economique observable (emplois, valeur retenue localement, nouvelles filieres) ?",
        "Transformation politique observable (nouveaux espaces de dialogue, participation accrue, influence) ?",
        "Transformation sociale observable (lien social, acces aux services, inclusion mesurable) ?",
        "Transformation ecologique observable (empreinte reduite, ecosystemes, adaptation climatique) ?",
    ],
}

NB_CRITERES_PAR_AXE = {1: 3, 2: 5, 3: 2, 4: 3, 5: 4}
TOTAL_CRITERES = sum(NB_CRITERES_PAR_AXE.values())  # 17


# ============================================================
# BASE DE DONNEES SQLITE
# ============================================================

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def init_db():
    conn = get_conn()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS reponses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date_soumission TEXT NOT NULL,
            nom_repondant TEXT NOT NULL,
            fonction TEXT,
            organisation TEXT NOT NULL,
            territoire TEXT,
            email TEXT,
            commentaire_global TEXT,
            donnees_json TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()

def save_response(meta, criteres_data):
    conn = get_conn()
    conn.execute("""
        INSERT INTO reponses
        (date_soumission, nom_repondant, fonction, organisation, territoire,
         email, commentaire_global, donnees_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now().isoformat(timespec="seconds"),
        meta["nom"],
        meta.get("fonction", ""),
        meta["organisation"],
        meta.get("territoire", ""),
        meta.get("email", ""),
        meta.get("commentaire_global", ""),
        json.dumps(criteres_data, ensure_ascii=False),
    ))
    conn.commit()
    conn.close()

def load_responses():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM reponses ORDER BY date_soumission DESC", conn)
    conn.close()
    return df

def delete_response(rid):
    conn = get_conn()
    conn.execute("DELETE FROM reponses WHERE id = ?", (rid,))
    conn.commit()
    conn.close()

def reset_db():
    conn = get_conn()
    conn.execute("DELETE FROM reponses")
    conn.commit()
    conn.close()


# ============================================================
# CALCULS
# ============================================================

def parse_donnees(donnees_json):
    """Retourne {critere_id: {'niveau': int, 'commentaire': str}}"""
    return json.loads(donnees_json)

def score_axe(donnees, axe):
    notes = []
    for c_id, _, _ in CRITERES:
        c_axe = int(c_id.split(".")[0])
        if c_axe != axe:
            continue
        n = donnees.get(c_id, {}).get("niveau", 0)
        if n > 0:
            notes.append(n)
    return round(sum(notes) / len(notes), 2) if notes else 0

def score_global(donnees):
    notes = [donnees.get(c[0], {}).get("niveau", 0) for c in CRITERES]
    notes = [n for n in notes if n > 0]
    return round(sum(notes) / len(notes), 2) if notes else 0

def qualif(score):
    if score == 0: return "Non evalue"
    if score < 1.5: return "1 - Tres faible"
    if score < 2.5: return "2 - Faible / ponctuel"
    if score < 3.5: return "3 - Moyen / structure"
    return "4 - Fort / strategique"

def aggregate_responses(df):
    """Retourne dict {critere_id: {'niveau_moyen': x, 'nb': n, 'commentaires': [...]}}"""
    agg = {c[0]: {"niveaux": [], "commentaires": []} for c in CRITERES}
    for _, row in df.iterrows():
        donnees = parse_donnees(row["donnees_json"])
        for c_id, _, _ in CRITERES:
            n = donnees.get(c_id, {}).get("niveau", 0)
            com = donnees.get(c_id, {}).get("commentaire", "")
            if n > 0:
                agg[c_id]["niveaux"].append(n)
            if com.strip():
                agg[c_id]["commentaires"].append(
                    f"[{row['nom_repondant']}] {com.strip()}"
                )
    result = {}
    for c_id, data in agg.items():
        if data["niveaux"]:
            result[c_id] = {
                "niveau_moyen": round(sum(data["niveaux"]) / len(data["niveaux"]), 2),
                "niveau_min": min(data["niveaux"]),
                "niveau_max": max(data["niveaux"]),
                "nb_reponses": len(data["niveaux"]),
                "commentaires": data["commentaires"],
            }
        else:
            result[c_id] = {
                "niveau_moyen": 0, "niveau_min": 0, "niveau_max": 0,
                "nb_reponses": 0, "commentaires": [],
            }
    return result


# ============================================================
# EXPORT WORD
# ============================================================

def set_cell_bg(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color_hex.lstrip("#"))
    tcPr.append(shd)

def export_synthese_docx(df):
    """Genere le bilan agrege en .docx."""
    doc = Document()
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2)
        s.top_margin = s.bottom_margin = Cm(1.8)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)

    # Titre
    p = doc.add_paragraph()
    r = p.add_run("Synthese collective - Grille de maturite RTE")
    r.font.size = Pt(20); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("1F4E79")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    r = p.add_run(f"Agregation de {len(df)} reponses")
    r.font.size = Pt(12); r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("7D8B96")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    r = p.add_run(f"Document genere le {date.today().strftime('%d/%m/%Y')}")
    r.font.size = Pt(10); r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("7D8B96")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Profil des repondants
    p = doc.add_paragraph(); r = p.add_run("Profil des repondants")
    r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("1F4E79")

    tbl = doc.add_table(rows=len(df) + 1, cols=4)
    headers = ["#", "Nom", "Fonction", "Date"]
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        set_cell_bg(cell, "1F4E79")
        p = cell.paragraphs[0]
        r = p.add_run(h); r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
    for ri, (_, row) in enumerate(df.iterrows(), start=1):
        for ci, val in enumerate([
            str(ri), row["nom_repondant"], row["fonction"] or "-",
            row["date_soumission"][:10]
        ]):
            cell = tbl.cell(ri, ci)
            p = cell.paragraphs[0]
            r = p.add_run(str(val)); r.font.size = Pt(9.5)

    # Scores par axe
    doc.add_paragraph()
    p = doc.add_paragraph(); r = p.add_run("Scores moyens par axe")
    r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("1F4E79")

    # Calcul scores moyens par axe sur tous les repondants
    all_scores = []
    for _, row in df.iterrows():
        d = parse_donnees(row["donnees_json"])
        all_scores.append({a: score_axe(d, a) for a in AXES})
    df_scores = pd.DataFrame(all_scores)
    scores_moyens = df_scores.mean()

    tbl = doc.add_table(rows=len(AXES) + 2, cols=4)
    headers = ["Axe", "Score moyen", "Niveau", "Nb criteres"]
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        set_cell_bg(cell, "1F4E79")
        p = cell.paragraphs[0]
        r = p.add_run(h); r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
    for ri, axe_num in enumerate(AXES, start=1):
        score = scores_moyens.get(axe_num, 0)
        color = AXE_COLORS[axe_num].lstrip("#")
        for ci, val in enumerate([
            f"Axe {axe_num} - {AXES[axe_num]}",
            f"{score:.2f} / 4",
            qualif(score),
            str(NB_CRITERES_PAR_AXE[axe_num]),
        ]):
            cell = tbl.cell(ri, ci)
            if ci == 0:
                set_cell_bg(cell, color)
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(10)
            if ci == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor.from_string("FFFFFF")
    # Score global
    score_g = scores_moyens.mean()
    for ci, val in enumerate([
        "SCORE GLOBAL RTE",
        f"{score_g:.2f} / 4",
        qualif(score_g),
        str(TOTAL_CRITERES),
    ]):
        cell = tbl.cell(len(AXES) + 1, ci)
        set_cell_bg(cell, "C9A961")
        p = cell.paragraphs[0]
        r = p.add_run(val); r.font.bold = True; r.font.size = Pt(11)

    # Detail par critere
    doc.add_page_break()
    p = doc.add_paragraph(); r = p.add_run("Detail par critere")
    r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("1F4E79")

    agg = aggregate_responses(df)
    for c_id, axe, titre in CRITERES:
        a = agg[c_id]
        if a["nb_reponses"] == 0:
            continue
        color = AXE_COLORS[axe].lstrip("#")
        # Bandeau critere
        tbl = doc.add_table(rows=1, cols=1)
        cell = tbl.cell(0, 0)
        set_cell_bg(cell, color)
        p = cell.paragraphs[0]
        r = p.add_run(f"  {c_id}  ")
        r.font.size = Pt(11); r.font.bold = True
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
        r = p.add_run(titre)
        r.font.size = Pt(11); r.font.bold = True
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
        r = p.add_run(
            f"   ·   Moyenne : {a['niveau_moyen']:.2f}/4   "
            f"(min {a['niveau_min']} / max {a['niveau_max']}, "
            f"sur {a['nb_reponses']} reponses)"
        )
        r.font.size = Pt(10); r.font.bold = True
        r.font.color.rgb = RGBColor.from_string("FFE4A0")

        # Commentaires
        if a["commentaires"]:
            for com in a["commentaires"]:
                p = doc.add_paragraph(style="List Bullet")
                r = p.add_run(com); r.font.size = Pt(10)
        doc.add_paragraph()

    # Export to bytes
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ============================================================
# UI : CSS / STYLE
# ============================================================

def apply_custom_css():
    st.markdown("""
    <style>
    .stApp {
        background-color: #FAFAFA;
    }
    .axe-banner {
        padding: 12px 20px;
        border-radius: 8px;
        margin: 20px 0 10px 0;
        color: white;
        font-weight: 700;
        font-size: 18px;
    }
    .critere-box {
        background: white;
        padding: 16px;
        border-radius: 8px;
        border-left: 4px solid #1F4E79;
        margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .critere-id {
        font-weight: 700;
        color: #1F4E79;
        margin-right: 6px;
    }
    .niveau-desc {
        font-size: 12px;
        color: #6B7280;
        font-style: italic;
        margin-top: 4px;
    }
    .kpi {
        background: white;
        padding: 16px;
        border-radius: 8px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }
    .kpi-val {
        font-size: 32px;
        font-weight: 800;
        color: #1F4E79;
    }
    .kpi-lbl {
        font-size: 12px;
        color: #6B7280;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .stProgress > div > div > div > div {
        background-color: #1F4E79;
    }
    </style>
    """, unsafe_allow_html=True)


# ============================================================
# UI : PAGE FORMULAIRE
# ============================================================

def page_formulaire():
    st.title("📝 Remplir la grille de maturite RTE")
    st.markdown(
        "Cette grille evalue le degre de **Responsabilite Territoriale** d'une organisation, "
        "selon 5 axes et 17 criteres. Pour chaque critere, choisissez un niveau de 1 a 4 "
        "et apportez si possible une justification ou un exemple concret."
    )

    st.markdown("---")
    st.subheader("👤 Vos coordonnees")
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom et prenom *", placeholder="ex. Jean Dupont")
        fonction = st.text_input("Fonction", placeholder="ex. Directeur, charge de mission")
    with col2:
        orga = st.text_input("Organisation evaluee *",
                             value="Office de Tourisme Sud Vienne Poitou")
        territoire = st.text_input("Territoire concerne",
                                   value="Communaute de Communes Vienne et Gartempe")
    email = st.text_input("Email (optionnel)", placeholder="optionnel - pour vous renvoyer la synthese")

    st.markdown("---")
    st.subheader("📊 Evaluation par critere")

    # Stockage temporaire des reponses dans le state
    if "form_data" not in st.session_state:
        st.session_state.form_data = {}

    # Affichage des criteres regroupes par axe
    for axe_num, axe_label in AXES.items():
        color = AXE_COLORS[axe_num]
        st.markdown(
            f"<div class='axe-banner' style='background:{color}'>"
            f"AXE {axe_num} : {axe_label.upper()}"
            f"</div>", unsafe_allow_html=True
        )
        criteres_axe = [(c, a, t) for c, a, t in CRITERES if a == axe_num]
        for c_id, _, titre in criteres_axe:
            with st.expander(f"**{c_id}** — {titre}", expanded=False):
                # Checklist d'aide
                st.markdown("**Quelques questions pour vous aider :**")
                for q in CHECKLISTS[c_id]:
                    st.markdown(f"- {q}")
                st.markdown("")

                # Selecteur de niveau
                niveau = st.radio(
                    f"Votre evaluation du critere **{c_id}**",
                    options=[0, 1, 2, 3, 4],
                    format_func=lambda x: NIVEAUX_LABELS[x],
                    horizontal=False,
                    key=f"niveau_{c_id}",
                    index=0,
                )
                if niveau > 0:
                    st.caption(f"💡 {NIVEAUX_DESC[niveau]}")

                # Commentaire libre
                commentaire = st.text_area(
                    "Justification / observations / exemples concrets",
                    key=f"com_{c_id}",
                    placeholder="(optionnel) Quels elements observes vous font choisir ce niveau ?",
                    height=80,
                )

                st.session_state.form_data[c_id] = {
                    "niveau": niveau,
                    "commentaire": commentaire,
                }

    # Commentaire global
    st.markdown("---")
    commentaire_global = st.text_area(
        "Commentaire global / remarques transversales",
        placeholder="(optionnel) Remarques transversales, points qui ne rentrent pas dans la grille...",
        height=100,
    )

    # Apercu score
    score_g = score_global(st.session_state.form_data)
    nb_repondus = sum(1 for c in CRITERES
                      if st.session_state.form_data.get(c[0], {}).get("niveau", 0) > 0)

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    col1.metric("Criteres remplis", f"{nb_repondus} / {TOTAL_CRITERES}")
    col2.metric("Votre score global provisoire", f"{score_g:.2f} / 4")
    col3.metric("Niveau", qualif(score_g))
    st.progress(nb_repondus / TOTAL_CRITERES)

    # Validation
    if st.button("✅ Soumettre ma reponse", type="primary", use_container_width=True):
        if not nom.strip():
            st.error("⚠️ Merci d'indiquer votre nom.")
        elif not orga.strip():
            st.error("⚠️ Merci d'indiquer l'organisation evaluee.")
        elif nb_repondus == 0:
            st.error("⚠️ Merci d'evaluer au moins un critere.")
        else:
            save_response(
                meta={
                    "nom": nom, "fonction": fonction, "organisation": orga,
                    "territoire": territoire, "email": email,
                    "commentaire_global": commentaire_global,
                },
                criteres_data=st.session_state.form_data,
            )
            st.success(f"🎉 Merci {nom} ! Votre reponse a bien ete enregistree.")
            st.balloons()
            st.session_state.form_data = {}


# ============================================================
# UI : PAGE DASHBOARD
# ============================================================

def page_dashboard():
    st.title("📊 Dashboard - Synthese collective en temps reel")

    df = load_responses()
    if len(df) == 0:
        st.info("👋 Aucune reponse pour l'instant. Partagez le lien du formulaire avec vos collegues !")
        return

    # KPI globaux
    st.subheader("Vue d'ensemble")
    all_scores = []
    for _, row in df.iterrows():
        d = parse_donnees(row["donnees_json"])
        all_scores.append({a: score_axe(d, a) for a in AXES})
    df_scores = pd.DataFrame(all_scores)
    scores_moyens = df_scores.mean()
    score_g = scores_moyens.mean()

    col1, col2, col3, col4 = st.columns(4)
    col1.markdown(
        f"<div class='kpi'><div class='kpi-val'>{len(df)}</div>"
        f"<div class='kpi-lbl'>Reponses</div></div>",
        unsafe_allow_html=True)
    col2.markdown(
        f"<div class='kpi'><div class='kpi-val'>{score_g:.2f}</div>"
        f"<div class='kpi-lbl'>Score moyen / 4</div></div>",
        unsafe_allow_html=True)
    col3.markdown(
        f"<div class='kpi'><div class='kpi-val'>{qualif(score_g).split(' - ')[0]}</div>"
        f"<div class='kpi-lbl'>Niveau moyen</div></div>",
        unsafe_allow_html=True)
    col4.markdown(
        f"<div class='kpi'><div class='kpi-val'>{int(df_scores.notna().sum().sum())}</div>"
        f"<div class='kpi-lbl'>Evaluations</div></div>",
        unsafe_allow_html=True)

    st.markdown("---")

    # Radar chart : scores moyens par axe
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("Profil RTE moyen (radar)")
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(
            r=[scores_moyens[a] for a in AXES],
            theta=[f"Axe {a}<br>{AXES[a][:25]}" for a in AXES],
            fill="toself",
            name="Moyenne",
            line=dict(color="#1F4E79", width=3),
            fillcolor="rgba(31, 78, 121, 0.3)",
        ))
        fig.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 4], tickfont=dict(size=10)),
                angularaxis=dict(tickfont=dict(size=11)),
            ),
            showlegend=False,
            height=420,
            margin=dict(l=60, r=60, t=20, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Scores moyens par axe")
        bar_data = pd.DataFrame({
            "Axe": [f"Axe {a} - {AXES[a]}" for a in AXES],
            "Score": [scores_moyens[a] for a in AXES],
            "Couleur": [AXE_COLORS[a] for a in AXES],
        })
        fig = px.bar(
            bar_data, x="Score", y="Axe", orientation="h",
            color="Axe", color_discrete_sequence=[AXE_COLORS[a] for a in AXES],
            text="Score",
        )
        fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig.update_layout(
            xaxis_range=[0, 4.5], showlegend=False, height=420,
            margin=dict(l=20, r=20, t=20, b=20),
            xaxis_title="Score moyen / 4", yaxis_title=None,
        )
        st.plotly_chart(fig, use_container_width=True)

    # Detail par critere
    st.markdown("---")
    st.subheader("Detail par critere")
    agg = aggregate_responses(df)
    rows = []
    for c_id, axe, titre in CRITERES:
        a = agg[c_id]
        rows.append({
            "Critere": c_id,
            "Axe": f"Axe {axe}",
            "Titre": titre,
            "Niveau moyen": f"{a['niveau_moyen']:.2f}" if a["nb_reponses"] else "—",
            "Min - Max": f"{a['niveau_min']} - {a['niveau_max']}" if a["nb_reponses"] else "—",
            "Nb reponses": a["nb_reponses"],
        })
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    # Reponses individuelles
    st.markdown("---")
    st.subheader("Comparaison entre repondants")
    if len(df) > 1:
        comparison = []
        for _, row in df.iterrows():
            d = parse_donnees(row["donnees_json"])
            entry = {"Repondant": row["nom_repondant"]}
            for a in AXES:
                entry[f"Axe {a}"] = score_axe(d, a)
            entry["Global"] = score_global(d)
            comparison.append(entry)
        df_comp = pd.DataFrame(comparison)
        st.dataframe(df_comp, hide_index=True, use_container_width=True)

        # Radar comparatif
        st.subheader("Profil compare (radar superpose)")
        fig = go.Figure()
        for _, row in df.iterrows():
            d = parse_donnees(row["donnees_json"])
            fig.add_trace(go.Scatterpolar(
                r=[score_axe(d, a) for a in AXES],
                theta=[f"Axe {a}" for a in AXES],
                fill="toself",
                name=row["nom_repondant"],
                opacity=0.5,
            ))
        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 4])),
            showlegend=True,
            height=420,
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Au moins 2 reponses sont necessaires pour la comparaison.")

    # Commentaires
    st.markdown("---")
    st.subheader("💬 Commentaires recoltes par critere")
    selected_c = st.selectbox(
        "Choisir un critere pour voir les commentaires",
        options=[c[0] for c in CRITERES],
        format_func=lambda x: f"{x} - {dict((c[0], c[2]) for c in CRITERES)[x]}",
    )
    coms = agg[selected_c]["commentaires"]
    if coms:
        for com in coms:
            st.markdown(f"> {com}")
    else:
        st.info("Aucun commentaire sur ce critere pour l'instant.")

    # Export Word
    st.markdown("---")
    st.subheader("📥 Exporter la synthese")
    buf = export_synthese_docx(df)
    st.download_button(
        label="📄 Telecharger la synthese en Word (.docx)",
        data=buf,
        file_name=f"Synthese_RTE_{date.today().isoformat()}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )


# ============================================================
# UI : PAGE ADMINISTRATION
# ============================================================

def page_admin():
    st.title("⚙️ Administration")
    df = load_responses()

    st.subheader(f"Reponses enregistrees : {len(df)}")
    if len(df) == 0:
        st.info("Aucune reponse enregistree.")
        return

    # SAUVEGARDE - bouton prioritaire en haut
    st.warning(
        "⚠️ **Pense a sauvegarder regulierement** les reponses. "
        "Si l'app est redeployee, la base est reinitialisee. "
        "Telecharge un backup CSV ci-dessous au moins une fois par semaine."
    )
    col1, col2 = st.columns(2)
    with col1:
        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="💾 Telecharger backup CSV (toutes les reponses)",
            data=csv_data,
            file_name=f"backup_RTE_{date.today().isoformat()}.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )
    with col2:
        # Export JSON complet
        json_data = df.to_json(orient="records", force_ascii=False, indent=2).encode("utf-8")
        st.download_button(
            label="📦 Backup JSON complet (avec details)",
            data=json_data,
            file_name=f"backup_RTE_{date.today().isoformat()}.json",
            mime="application/json",
            use_container_width=True,
        )

    st.markdown("---")
    # Tableau des reponses
    display = df[["id", "date_soumission", "nom_repondant", "fonction",
                  "organisation", "email"]].copy()
    st.dataframe(display, hide_index=True, use_container_width=True)

    # Suppression
    st.markdown("---")
    st.subheader("Supprimer une reponse")
    rid = st.number_input("ID a supprimer", min_value=1, step=1)
    if st.button("🗑️ Supprimer cette reponse"):
        delete_response(rid)
        st.success(f"Reponse #{rid} supprimee. Rechargez la page.")

    st.markdown("---")
    st.subheader("⚠️ Reinitialiser toute la base")
    st.warning("Cette action supprime TOUTES les reponses. Irreversible.")
    confirm = st.text_input("Tapez 'SUPPRIMER' pour confirmer")
    if st.button("🔥 Reinitialiser la base"):
        if confirm == "SUPPRIMER":
            reset_db()
            st.success("Base reinitialisee.")
        else:
            st.error("Tapez 'SUPPRIMER' exactement pour confirmer.")


# ============================================================
# MAIN
# ============================================================

def main():
    st.set_page_config(
        page_title=APP_TITLE, page_icon=APP_ICON, layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_custom_css()
    init_db()

    with st.sidebar:
        st.title("🧭 Navigation")
        st.markdown("**Grille de maturite RTE**")
        st.caption("Responsabilite Territoriale des Entreprises")
        st.markdown("---")
        page = st.radio(
            "Choisir la page",
            options=["📝 Formulaire", "📊 Dashboard", "⚙️ Admin"],
            label_visibility="collapsed",
        )
        st.markdown("---")
        df = load_responses()
        st.metric("Reponses recues", len(df))
        st.markdown("---")
        st.caption("Stage M2 IE IAE Poitiers")
        st.caption("Chaire TerrESS - Bordeaux Sciences Agro")

    if page == "📝 Formulaire":
        page_formulaire()
    elif page == "📊 Dashboard":
        page_dashboard()
    elif page == "⚙️ Admin":
        page_admin()


if __name__ == "__main__":
    main()
