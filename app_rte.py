"""
Formulaire collaboratif - Grille de maturite RTE
Application Streamlit avec notation automatique :
  - Chaque question : Observe (1pt) / Partiel (0.5pt) / Non observe (0pt)
  - Score critere = (somme points / nb questions) * 4
  - Score axe = moyenne des criteres
  - Score global = moyenne des criteres

Lancement local :
    streamlit run app_rte.py
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

AXE_COLORS = {
    1: "#C28533",  # orange
    2: "#1F4E79",  # bleu
    3: "#256D5A",  # vert
    4: "#5A3B7A",  # violet
    5: "#9C2A2A",  # rouge
}

POINTS = {"obs": 1.0, "part": 0.5, "non": 0.0}
STATUT_LABELS = {
    "obs": "Observe",
    "part": "Partiel",
    "non": "Non observe",
}
STATUT_LABELS_VISUAL = {
    "obs": "✅ Observe",
    "part": "🟡 Partiel",
    "non": "❌ Non observe",
}
STATUT_COLORS = {
    "obs": "#2E8B73",
    "part": "#C9A961",
    "non": "#C0392B",
}

AXES = {
    1: "Diagnostic territorial",
    2: "Vision territoriale et recits",
    3: "Cooperation territoriale",
    4: "Gouvernance inclusive",
    5: "Redistribuer et transformer",
}

CRITERES = [
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

# Checklists : la base. Chaque item devient une question notee.
CHECKLISTS_RAW = {
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

# Genere les checklists avec IDs (1.1.1, 1.1.2, ...)
CHECKLISTS = {
    c_id: [
        {"id": f"{c_id}.{i+1}", "label": label}
        for i, label in enumerate(items)
    ]
    for c_id, items in CHECKLISTS_RAW.items()
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
            donnees_json TEXT NOT NULL,
            schema_version INTEGER DEFAULT 2
        )
    """)
    # Migration : ajout colonne schema_version si manquante
    cur = conn.execute("PRAGMA table_info(reponses)")
    cols = [r[1] for r in cur.fetchall()]
    if "schema_version" not in cols:
        conn.execute("ALTER TABLE reponses ADD COLUMN schema_version INTEGER DEFAULT 1")
    conn.commit()
    conn.close()

def save_response(meta, criteres_data):
    conn = get_conn()
    conn.execute("""
        INSERT INTO reponses
        (date_soumission, nom_repondant, fonction, organisation, territoire,
         email, commentaire_global, donnees_json, schema_version)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        datetime.now().isoformat(timespec="seconds"),
        meta["nom"],
        meta.get("fonction", ""),
        meta["organisation"],
        meta.get("territoire", ""),
        meta.get("email", ""),
        meta.get("commentaire_global", ""),
        json.dumps(criteres_data, ensure_ascii=False),
        2,
    ))
    conn.commit()
    conn.close()

def load_responses():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM reponses ORDER BY date_soumission DESC", conn)
    conn.close()
    # Ne conserver que les reponses au schema V2 (nouveau format)
    if "schema_version" in df.columns:
        df = df[df["schema_version"] == 2].reset_index(drop=True)
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
# CALCUL DES SCORES (notation automatique)
# ============================================================

def score_critere(questions_data):
    """
    Score automatique d'un critere sur /4.
    questions_data : {q_id: {'statut': 'obs'|'part'|'non', 'preuve': str}}
    """
    pts = []
    for q_id, q_data in (questions_data or {}).items():
        s = q_data.get("statut")
        if s in POINTS:
            pts.append(POINTS[s])
    if not pts:
        return None  # aucune question repondue
    return round((sum(pts) / len(pts)) * 4, 2)

def nb_questions_repondues(questions_data):
    return sum(1 for q in (questions_data or {}).values()
               if q.get("statut") in POINTS)

def score_axe(criteres_data, axe):
    """Moyenne des scores des criteres de cet axe (criteres non repondus ignores)."""
    scores = []
    for c_id, c_axe, _ in CRITERES:
        if c_axe != axe:
            continue
        c_data = criteres_data.get(c_id, {})
        questions = c_data.get("questions", {})
        s = score_critere(questions)
        if s is not None:
            scores.append(s)
    return round(sum(scores) / len(scores), 2) if scores else None

def score_global(criteres_data):
    """Moyenne des scores de tous les criteres (les non repondus sont ignores)."""
    scores = []
    for c_id, _, _ in CRITERES:
        c_data = criteres_data.get(c_id, {})
        questions = c_data.get("questions", {})
        s = score_critere(questions)
        if s is not None:
            scores.append(s)
    return round(sum(scores) / len(scores), 2) if scores else None

def qualif(score):
    if score is None: return "Non evalue"
    if score < 1.5: return "1 - Tres faible / inexistant"
    if score < 2.5: return "2 - Faible / ponctuel"
    if score < 3.5: return "3 - Moyen / structure"
    return "4 - Fort / strategique"


# ============================================================
# AGREGATION DES REPONSES
# ============================================================

def parse_donnees(donnees_json):
    return json.loads(donnees_json)

def aggregate_responses(df):
    """
    Retourne pour chaque critere :
        - score_moyen (moyenne des scores individuels)
        - score_min, score_max
        - nb_reponses (nb de personnes ayant repondu au critere)
        - questions : pour chaque question, distribution des statuts + preuves collectees
    """
    result = {}
    for c_id, _, _ in CRITERES:
        # Scores individuels sur ce critere
        scores_indiv = []
        # Pour chaque question : compteurs + preuves
        questions_agg = {
            q["id"]: {"obs": 0, "part": 0, "non": 0, "preuves": []}
            for q in CHECKLISTS[c_id]
        }
        for _, row in df.iterrows():
            d = parse_donnees(row["donnees_json"])
            c_data = d.get(c_id, {})
            questions = c_data.get("questions", {})
            s = score_critere(questions)
            if s is not None:
                scores_indiv.append(s)
            for q_id, q_data in questions.items():
                if q_id not in questions_agg:
                    continue
                statut = q_data.get("statut")
                if statut in POINTS:
                    questions_agg[q_id][statut] += 1
                preuve = (q_data.get("preuve") or "").strip()
                if preuve:
                    questions_agg[q_id]["preuves"].append(
                        f"[{row['nom_repondant']}] {preuve}"
                    )
        result[c_id] = {
            "score_moyen": round(sum(scores_indiv)/len(scores_indiv), 2) if scores_indiv else None,
            "score_min": min(scores_indiv) if scores_indiv else None,
            "score_max": max(scores_indiv) if scores_indiv else None,
            "nb_reponses": len(scores_indiv),
            "questions": questions_agg,
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

def add_para_styled(container, text, size=10, bold=False, italic=False,
                    color_hex=None, align=None):
    p = container.add_paragraph()
    if align is not None:
        p.alignment = align
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    if color_hex:
        r.font.color.rgb = RGBColor.from_string(color_hex.lstrip("#"))
    return p

def set_cell_border(cell, color_hex="BFBFBF", size="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), size)
        b.set(qn("w:color"), color_hex)
        tcBorders.append(b)
    tcPr.append(tcBorders)


def _commentaire_axe(axe_num, score, scores_moyens, agg, df):
    """Genere un commentaire synthetique auto pour un axe."""
    if score is None:
        return "Non evalue par les repondants."

    # Trouve les criteres faibles et forts de cet axe
    criteres_axe = [(c, t) for c, a, t in CRITERES if a == axe_num]
    crit_scores = []
    for c_id, titre in criteres_axe:
        s = agg[c_id]["score_moyen"]
        if s is not None:
            crit_scores.append((c_id, titre, s))

    # Compare aux autres axes
    valid_axes = {a: s for a, s in scores_moyens.items() if s is not None}
    rang = sorted(valid_axes.values(), reverse=True).index(score) + 1 if score in valid_axes.values() else None

    parts = []

    # Verdict global
    if score >= 3.5:
        parts.append("Point fort majeur du profil RTE")
    elif score >= 2.5:
        parts.append("Axe a un niveau de structuration moyen")
    elif score >= 1.5:
        parts.append("Axe ponctuel a renforcer")
    else:
        parts.append("Axe tres faible, demarche a engager")

    # Si plusieurs axes : indique le rang
    if rang and len(valid_axes) >= 3:
        if rang == 1:
            parts.append("(score le plus eleve)")
        elif rang == len(valid_axes):
            parts.append("(score le plus faible)")

    # Critere a renforcer
    if crit_scores:
        crit_faible = min(crit_scores, key=lambda x: x[2])
        crit_fort = max(crit_scores, key=lambda x: x[2])
        if crit_faible[2] < crit_fort[2] - 0.5:
            parts.append(f". Le critere {crit_faible[0]} ressort comme le moins bien evalue.")
        else:
            parts.append(".")
    else:
        parts.append(".")

    return " ".join(parts)


def _h1_doc(doc, text, color_hex="1F4E79"):
    p = doc.add_paragraph()
    r = p.add_run(text); r.font.size = Pt(20); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(color_hex)
    return p

def _h2_doc(doc, text, color_hex="1F4E79"):
    p = doc.add_paragraph()
    r = p.add_run(text); r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(color_hex)
    return p

def _axe_banner_doc(doc, axe_num, label):
    color = AXE_COLORS[axe_num].lstrip("#")
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0); set_cell_bg(cell, color)
    p = cell.paragraphs[0]
    r = p.add_run(f"  AXE {axe_num} : {label.upper()}")
    r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("FFFFFF")
    doc.add_paragraph()


def export_synthese_docx(df):
    """
    Genere la synthese en .docx au format du Grille_RTE_OT_SudVienne.docx :
      1. Titre + metadonnees
      2. Section 5 : Grille - bloc par critere (score auto + elements + recommandations)
      3. Section 6 : Synthese des resultats (tableau Axe / Score / Niveau / Commentaire)
      4. Annexe : Checklists detaillees par critere
    """
    if len(df) == 0:
        # Doc minimal si pas de donnees
        doc = Document()
        add_para_styled(doc, "Aucune reponse a synthetiser pour l'instant.",
                        size=14, color_hex="9C2A2A", italic=True,
                        align=WD_ALIGN_PARAGRAPH.CENTER)
        buf = BytesIO(); doc.save(buf); buf.seek(0)
        return buf

    doc = Document()
    for s in doc.sections:
        s.left_margin = s.right_margin = Cm(2)
        s.top_margin = s.bottom_margin = Cm(1.8)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)

    # Recuperation org/territoire (depuis la 1ere reponse)
    orga = df.iloc[0]["organisation"] if len(df) > 0 else "—"
    territoire = df.iloc[0]["territoire"] if len(df) > 0 else "—"

    # Calculs globaux
    agg = aggregate_responses(df)
    all_scores = []
    for _, row in df.iterrows():
        d = parse_donnees(row["donnees_json"])
        all_scores.append({a: score_axe(d, a) for a in AXES})
    df_scores = pd.DataFrame(all_scores)
    scores_moyens = {a: (df_scores[a].dropna().mean() if df_scores[a].dropna().shape[0] > 0 else None)
                     for a in AXES}
    valid = [s for s in scores_moyens.values() if s is not None]
    score_g = sum(valid) / len(valid) if valid else None

    # ============================================================
    # 1. TITRE + METADONNEES
    # ============================================================
    p = doc.add_paragraph()
    r = p.add_run("Grille de maturite RTE - Synthese collective")
    r.font.size = Pt(22); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("1F4E79")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    r = p.add_run(orga); r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = RGBColor.from_string("2C3E50")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    r = p.add_run("Modele V4 - Chaire TerrESS / Bordeaux Sciences Agro")
    r.font.size = Pt(11); r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("7D8B96")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Metadonnees
    repondants_list = ", ".join(
        f"{row['nom_repondant']}" + (f" ({row['fonction']})" if row['fonction'] else "")
        for _, row in df.iterrows()
    )
    meta = [
        ("Organisation evaluee", orga),
        ("Territoire concerne", territoire or "—"),
        ("Date de generation", date.today().strftime("%d / %m / %Y")),
        ("Nombre de repondants", str(len(df))),
        ("Repondants", repondants_list),
        ("Methode", "Notation automatique : Observe=1 / Partiel=0,5 / Non observe=0. "
                    "Score critere = (moyenne / nb questions) x 4."),
    ]
    tbl = doc.add_table(rows=len(meta), cols=2)
    tbl.columns[0].width = Cm(5)
    tbl.columns[1].width = Cm(12)
    for ri, (k, v) in enumerate(meta):
        c0 = tbl.cell(ri, 0); c1 = tbl.cell(ri, 1)
        set_cell_bg(c0, "DCE6F1")
        set_cell_border(c0); set_cell_border(c1)
        p = c0.paragraphs[0]
        r = p.add_run(k); r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string("1F4E79")
        p = c1.paragraphs[0]
        r = p.add_run(v); r.font.size = Pt(10)

    # Note
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run(
        "Synthese generee automatiquement a partir des reponses collectees "
        "via le formulaire en ligne. Cette synthese constitue un point de depart "
        "pour un atelier de co-evaluation avec les parties prenantes."
    )
    r.font.size = Pt(10); r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("7D8B96")

    # ============================================================
    # 2. SECTION 5 : GRILLE PAR CRITERE
    # ============================================================
    doc.add_page_break()
    _h1_doc(doc, "5. Grille de maturite : evaluation par critere")
    p = doc.add_paragraph()
    r = p.add_run(
        "Pour chaque critere : score automatique calcule a partir des reponses "
        "collectees, elements d'analyse (preuves citees par les repondants) "
        "et points a ameliorer."
    )
    r.font.size = Pt(10); r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("7D8B96")
    doc.add_paragraph()

    GRILLE_PRINCIPALE = [
        (1, "Diagnostic territorial", ["1.1", "1.2", "1.3"]),
        (2, "Vision territoriale et recits", ["2.1", "2.2", "2.3", "2.4", "2.5"]),
        (3, "Cooperation territoriale", ["3.1", "3.2"]),
        (4, "Gouvernance inclusive", ["4.1", "4.2", "4.3"]),
        (5, "Redistribuer et transformer", ["5.1", "5.2", "5.3", "5.4"]),
    ]
    titre_map = {c[0]: c[2] for c in CRITERES}

    for axe_num, axe_label, criteres_ids in GRILLE_PRINCIPALE:
        _axe_banner_doc(doc, axe_num, axe_label)
        color = AXE_COLORS[axe_num].lstrip("#")
        for c_id in criteres_ids:
            a = agg[c_id]
            score = a["score_moyen"]
            titre = titre_map[c_id]

            # Tableau 3 rangees (header, analyse, reco)
            tbl = doc.add_table(rows=3, cols=1)

            # Header
            c0 = tbl.cell(0, 0); set_cell_bg(c0, color); set_cell_border(c0, color)
            p = c0.paragraphs[0]
            r = p.add_run(f"  {c_id}  ")
            r.font.size = Pt(12); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string("FFFFFF")
            r = p.add_run(titre)
            r.font.size = Pt(12); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string("FFFFFF")
            if score is not None:
                r = p.add_run(
                    f"      Score : {score:.2f} / 4   |   Niveau : {qualif(score).split(' - ')[0]}   |   {a['nb_reponses']} reponses"
                )
            else:
                r = p.add_run("      Score : non evalue")
            r.font.size = Pt(11); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string("FFE4A0")

            # Analyse : preuves agregees
            c1 = tbl.cell(1, 0); set_cell_bg(c1, "F4F6F8"); set_cell_border(c1)
            p = c1.paragraphs[0]
            r = p.add_run("Elements d'analyse (preuves citees par les repondants) : ")
            r.font.size = Pt(10); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string(color)

            # On liste les preuves question par question
            preuves_total = 0
            for q in CHECKLISTS[c_id]:
                qa = a["questions"][q["id"]]
                if qa["preuves"]:
                    preuves_total += len(qa["preuves"])
                    p = c1.add_paragraph()
                    r = p.add_run(f"• {q['id']} : ")
                    r.font.size = Pt(9.5); r.font.bold = True
                    r = p.add_run(" ; ".join(qa["preuves"]))
                    r.font.size = Pt(9.5)
            if preuves_total == 0:
                p = c1.add_paragraph()
                r = p.add_run("Aucune preuve documentaire renseignee par les repondants.")
                r.font.size = Pt(10); r.font.italic = True
                r.font.color.rgb = RGBColor.from_string("7D8B96")

            # Reco : auto - basee sur questions faibles
            c2 = tbl.cell(2, 0); set_cell_bg(c2, "FFFAEB"); set_cell_border(c2)
            p = c2.paragraphs[0]
            r = p.add_run("Points a ameliorer (auto-deduits) : ")
            r.font.size = Pt(10); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string("B7791F")

            questions_faibles = []
            for q in CHECKLISTS[c_id]:
                qa = a["questions"][q["id"]]
                total = qa["obs"] + qa["part"] + qa["non"]
                if total == 0:
                    continue
                q_score = (qa["obs"] * 1.0 + qa["part"] * 0.5) / total
                if q_score < 0.5:  # plus de "non" et "partiel" que "observe"
                    questions_faibles.append((q["id"], q["label"], qa))

            if questions_faibles:
                for q_id, q_label, qa in questions_faibles[:4]:
                    p = c2.add_paragraph()
                    r = p.add_run(f"• {q_id} : ")
                    r.font.size = Pt(9.5); r.font.bold = True
                    r = p.add_run(
                        f"{q_label} "
                        f"(✓{qa['obs']} | ~{qa['part']} | ✗{qa['non']})"
                    )
                    r.font.size = Pt(9.5); r.font.italic = True
            else:
                p = c2.add_paragraph()
                r = p.add_run("Aucun point d'amelioration prioritaire identifie sur ce critere.")
                r.font.size = Pt(10); r.font.italic = True
                r.font.color.rgb = RGBColor.from_string("2E8B73")

            doc.add_paragraph()
        doc.add_paragraph()

    # ============================================================
    # 3. SECTION 6 : SYNTHESE DES RESULTATS (FORMAT IDENTIQUE A L'EXISTANT)
    # ============================================================
    doc.add_page_break()
    _h1_doc(doc, "6. Synthese des resultats")

    tbl = doc.add_table(rows=len(AXES) + 2, cols=4)
    tbl.columns[0].width = Cm(6)
    tbl.columns[1].width = Cm(2.5)
    tbl.columns[2].width = Cm(3.5)
    tbl.columns[3].width = Cm(5)

    # Header
    for ci, h in enumerate(["Axe", "Score moyen", "Niveau", "Commentaire synthetique"]):
        cell = tbl.cell(0, ci); set_cell_bg(cell, "1F4E79"); set_cell_border(cell)
        p = cell.paragraphs[0]
        r = p.add_run(h); r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Lignes axes
    for ri, axe_num in enumerate(AXES, start=1):
        color = AXE_COLORS[axe_num].lstrip("#")
        score = scores_moyens[axe_num]
        nb_crit = NB_CRITERES_PAR_AXE[axe_num]
        commentaire = _commentaire_axe(axe_num, score, scores_moyens, agg, df)

        # Col 1 : Axe (colore)
        cell = tbl.cell(ri, 0); set_cell_bg(cell, color); set_cell_border(cell)
        p = cell.paragraphs[0]
        r = p.add_run(f"Axe {axe_num} - {AXES[axe_num]}")
        r.font.bold = True; r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string("FFFFFF")
        p2 = cell.add_paragraph()
        r = p2.add_run(f"({nb_crit} criteres)")
        r.font.size = Pt(9); r.font.italic = True
        r.font.color.rgb = RGBColor.from_string("FFFFFF")

        # Col 2 : Score
        cell = tbl.cell(ri, 1); set_cell_bg(cell, "F4F6F8"); set_cell_border(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"{score:.2f} / 4" if score is not None else "—")
        r.font.size = Pt(11); r.font.bold = True
        r.font.color.rgb = RGBColor.from_string(color)

        # Col 3 : Niveau
        cell = tbl.cell(ri, 2); set_cell_bg(cell, "F4F6F8"); set_cell_border(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(qualif(score))
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string("2C3E50")

        # Col 4 : Commentaire
        cell = tbl.cell(ri, 3); set_cell_bg(cell, "F4F6F8"); set_cell_border(cell)
        p = cell.paragraphs[0]
        r = p.add_run(commentaire); r.font.size = Pt(10)
        r.font.color.rgb = RGBColor.from_string("2C3E50")

    # Ligne globale doree
    ri = len(AXES) + 1
    for ci, val in enumerate([
        "SCORE GLOBAL RTE",
        f"{score_g:.2f} / 4" if score_g is not None else "—",
        qualif(score_g),
        f"Moyenne ponderee des 5 axes ({TOTAL_CRITERES} criteres au total)",
    ]):
        cell = tbl.cell(ri, ci); set_cell_bg(cell, "C9A961"); set_cell_border(cell)
        p = cell.paragraphs[0]
        r = p.add_run(val); r.font.bold = True; r.font.size = Pt(11)
        r.font.color.rgb = RGBColor.from_string("2C3E50")
        if ci in (1, 2):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Lecture d'ensemble
    doc.add_paragraph()
    _h2_doc(doc, "Lecture d'ensemble")
    if score_g is not None:
        p = doc.add_paragraph()
        r = p.add_run(
            f"Avec un score global de {score_g:.2f}/4, l'organisation evaluee se situe au niveau "
            f"\"{qualif(score_g)}\" sur la base des {len(df)} reponses collectees."
        )
        r.font.size = Pt(11)

        # Points forts
        axes_tries = sorted(
            ((a, s) for a, s in scores_moyens.items() if s is not None),
            key=lambda x: x[1], reverse=True,
        )
        if axes_tries:
            doc.add_paragraph()
            r = doc.add_paragraph().add_run("Points forts :")
            r.font.bold = True; r.font.size = Pt(11)
            r.font.color.rgb = RGBColor.from_string("2E8B73")
            for axe_num, sc in axes_tries[:2]:
                if sc >= 2.5:
                    p = doc.add_paragraph(style="List Bullet")
                    r = p.add_run(
                        f"Axe {axe_num} - {AXES[axe_num]} ({sc:.2f}/4) : "
                        f"{qualif(sc).lower()}."
                    )
                    r.font.size = Pt(10.5)

            # Axes a renforcer
            r = doc.add_paragraph().add_run("Axes a renforcer :")
            r.font.bold = True; r.font.size = Pt(11)
            r.font.color.rgb = RGBColor.from_string("9C2A2A")
            for axe_num, sc in axes_tries[::-1][:2]:
                if sc < 3.0:
                    p = doc.add_paragraph(style="List Bullet")
                    r = p.add_run(
                        f"Axe {axe_num} - {AXES[axe_num]} ({sc:.2f}/4) : "
                        f"{qualif(sc).lower()}."
                    )
                    r.font.size = Pt(10.5)

    # ============================================================
    # 4. ANNEXE : DETAIL DES REPONSES PAR CRITERE
    # ============================================================
    doc.add_page_break()
    _h1_doc(doc, "Annexe - Detail des reponses par critere")
    p = doc.add_paragraph()
    r = p.add_run(
        "Pour chaque critere : distribution des reponses question par question "
        "et preuves collectees. ✓ Observe = 1 pt | ~ Partiel = 0,5 pt | ✗ Non observe = 0 pt."
    )
    r.font.size = Pt(10); r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("7D8B96")

    for axe_num, axe_label, criteres_ids in GRILLE_PRINCIPALE:
        _axe_banner_doc(doc, axe_num, axe_label)
        color = AXE_COLORS[axe_num].lstrip("#")
        for c_id in criteres_ids:
            a = agg[c_id]
            titre = titre_map[c_id]

            # Bandeau critere
            tbl = doc.add_table(rows=1, cols=1)
            cell = tbl.cell(0, 0); set_cell_bg(cell, color)
            p = cell.paragraphs[0]
            r = p.add_run(f"  {c_id}  ")
            r.font.size = Pt(11); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string("FFFFFF")
            r = p.add_run(titre)
            r.font.size = Pt(11); r.font.bold = True
            r.font.color.rgb = RGBColor.from_string("FFFFFF")
            if a["score_moyen"] is not None:
                r = p.add_run(
                    f"   ·   Score : {a['score_moyen']:.2f}/4   ·   {a['nb_reponses']} reponses"
                )
                r.font.size = Pt(10); r.font.bold = True
                r.font.color.rgb = RGBColor.from_string("FFE4A0")

            # Tableau questions detail
            questions = CHECKLISTS[c_id]
            tbl = doc.add_table(rows=len(questions) + 1, cols=4)
            tbl.columns[0].width = Cm(7.5)
            tbl.columns[1].width = Cm(3)
            tbl.columns[2].width = Cm(5)
            tbl.columns[3].width = Cm(1.5)

            for ci, h in enumerate(["Action a documenter", "Distribution",
                                    "Preuves collectees", "Score q."]):
                cell = tbl.cell(0, ci); set_cell_bg(cell, "DCE6F1"); set_cell_border(cell)
                p = cell.paragraphs[0]
                r = p.add_run(h); r.font.bold = True; r.font.size = Pt(10)
                r.font.color.rgb = RGBColor.from_string("1F4E79")

            for ri, q in enumerate(questions, start=1):
                qa = a["questions"][q["id"]]
                total = qa["obs"] + qa["part"] + qa["non"]
                q_score = ((qa["obs"] * 1.0 + qa["part"] * 0.5) / total * 4) if total > 0 else None

                # Question
                c0 = tbl.cell(ri, 0); set_cell_border(c0); c0.text = ""
                p = c0.paragraphs[0]
                r = p.add_run(f"{q['id']} "); r.font.bold = True; r.font.size = Pt(9.5)
                r.font.color.rgb = RGBColor.from_string("1F4E79")
                r = p.add_run(q["label"]); r.font.size = Pt(9.5)

                # Distribution
                c1 = tbl.cell(ri, 1); set_cell_border(c1); c1.text = ""
                p = c1.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(f"✓ {qa['obs']}")
                r.font.size = Pt(9.5); r.font.bold = True
                r.font.color.rgb = RGBColor.from_string("2E8B73")
                r = p.add_run("  |  "); r.font.size = Pt(9.5)
                r = p.add_run(f"~ {qa['part']}")
                r.font.size = Pt(9.5); r.font.bold = True
                r.font.color.rgb = RGBColor.from_string("C9A961")
                r = p.add_run("  |  "); r.font.size = Pt(9.5)
                r = p.add_run(f"✗ {qa['non']}")
                r.font.size = Pt(9.5); r.font.bold = True
                r.font.color.rgb = RGBColor.from_string("C0392B")

                # Preuves
                c2 = tbl.cell(ri, 2); set_cell_border(c2); c2.text = ""
                if qa["preuves"]:
                    p = c2.paragraphs[0]
                    r = p.add_run(qa["preuves"][0])
                    r.font.size = Pt(9); r.font.italic = True
                    for preuve in qa["preuves"][1:]:
                        p = c2.add_paragraph()
                        r = p.add_run(preuve)
                        r.font.size = Pt(9); r.font.italic = True
                else:
                    p = c2.paragraphs[0]
                    r = p.add_run("—")
                    r.font.size = Pt(9); r.font.color.rgb = RGBColor.from_string("BFBFBF")

                # Score q.
                c3 = tbl.cell(ri, 3); set_cell_border(c3); c3.text = ""
                p = c3.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if q_score is not None:
                    r = p.add_run(f"{q_score:.2f}")
                    r.font.size = Pt(10); r.font.bold = True
                else:
                    r = p.add_run("—"); r.font.size = Pt(9.5)
            doc.add_paragraph()

    # Note finale
    doc.add_paragraph()
    p = doc.add_paragraph()
    r = p.add_run(
        f"--- Synthese generee automatiquement le {date.today().strftime('%d/%m/%Y')} ---"
    )
    r.font.size = Pt(9); r.font.italic = True
    r.font.color.rgb = RGBColor.from_string("7D8B96")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
    .q-row {
        background: white;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 8px;
        border-left: 3px solid #1F4E79;
    }
    .q-id {
        font-weight: 700;
        color: #1F4E79;
        font-size: 11px;
    }
    .q-label {
        color: #1A1A1A;
        font-size: 14px;
    }
    .crit-score {
        background: #F4F6F8;
        padding: 10px 16px;
        border-radius: 6px;
        margin-top: 12px;
        font-weight: 600;
        color: #1F4E79;
        text-align: center;
    }
    .stProgress > div > div > div > div {
        background-color: #1F4E79;
    }
    div[data-testid="stRadio"] > label > div > p {
        font-size: 12px;
    }
    </style>
    """, unsafe_allow_html=True)


# ============================================================
# UI : PAGE FORMULAIRE
# ============================================================

def page_formulaire():
    st.title("📝 Remplir la grille de maturite RTE")
    st.markdown(
        "**Mode d'emploi** : pour chaque question, indique si la pratique est "
        "**Observee** (clairement en place), **Partielle** (en cours ou incomplete) "
        "ou **Non observee** (absente). Si tu peux, ajoute une **preuve** "
        "(document, exemple, source). Les scores se calculent automatiquement."
    )
    st.markdown(
        "**Notation auto** : ✅ Observe = 1 pt · 🟡 Partiel = 0,5 pt · ❌ Non observe = 0 pt · "
        "Score critere = (moyenne / nb questions) × 4."
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
    email = st.text_input("Email (optionnel)",
                          placeholder="optionnel - pour recevoir la synthese")

    st.markdown("---")
    st.subheader("📊 Evaluation par critere")

    if "form_data" not in st.session_state:
        st.session_state.form_data = {}

    # Statut options pour les radios
    statut_options = [None, "obs", "part", "non"]
    statut_format = {
        None: "— Pas evalue",
        "obs": "✅ Observe",
        "part": "🟡 Partiel",
        "non": "❌ Non observe",
    }

    for axe_num, axe_label in AXES.items():
        color = AXE_COLORS[axe_num]
        st.markdown(
            f"<div class='axe-banner' style='background:{color}'>"
            f"AXE {axe_num} : {axe_label.upper()}"
            f"</div>", unsafe_allow_html=True
        )
        criteres_axe = [(c, a, t) for c, a, t in CRITERES if a == axe_num]
        for c_id, _, titre in criteres_axe:
            # Calcul du score actuel si reponses partielles
            current_qs = st.session_state.form_data.get(c_id, {}).get("questions", {})
            current_score = score_critere(current_qs)
            score_label = (f"  ·  Score : {current_score:.2f}/4" if current_score is not None
                           else "  ·  Score : —")
            nb_rep = nb_questions_repondues(current_qs)
            total_q = len(CHECKLISTS[c_id])

            with st.expander(
                f"**{c_id}** — {titre}   ({nb_rep}/{total_q} questions{score_label})",
                expanded=False,
            ):
                if c_id not in st.session_state.form_data:
                    st.session_state.form_data[c_id] = {"questions": {}, "commentaire": ""}

                # En-tete de colonnes
                ch, cs, cp = st.columns([4, 3, 4])
                with ch:
                    st.markdown("**Question**")
                with cs:
                    st.markdown("**Statut**")
                with cp:
                    st.markdown("**Preuve / source / observation**")
                st.markdown("---")

                for q in CHECKLISTS[c_id]:
                    cq, cstat, cpr = st.columns([4, 3, 4])
                    with cq:
                        st.markdown(
                            f"<span class='q-id'>{q['id']}</span><br>"
                            f"<span class='q-label'>{q['label']}</span>",
                            unsafe_allow_html=True
                        )
                    with cstat:
                        statut = st.radio(
                            label=f"Statut {q['id']}",
                            options=statut_options,
                            format_func=lambda x: statut_format[x],
                            key=f"st_{q['id']}",
                            label_visibility="collapsed",
                            horizontal=False,
                            index=0,
                        )
                    with cpr:
                        preuve = st.text_area(
                            label=f"Preuve {q['id']}",
                            key=f"pr_{q['id']}",
                            placeholder="Document, exemple, source...",
                            label_visibility="collapsed",
                            height=80,
                        )
                    st.session_state.form_data[c_id]["questions"][q["id"]] = {
                        "statut": statut,
                        "preuve": preuve,
                    }
                    st.markdown("")  # petit espace

                # Commentaire libre du critere
                commentaire = st.text_area(
                    "Commentaire global sur ce critere (facultatif)",
                    key=f"com_{c_id}",
                    placeholder="Remarques transversales, points nuances...",
                    height=70,
                )
                st.session_state.form_data[c_id]["commentaire"] = commentaire

                # Score auto-calcule
                final_qs = st.session_state.form_data[c_id]["questions"]
                final_score = score_critere(final_qs)
                final_nb = nb_questions_repondues(final_qs)
                if final_score is not None:
                    st.markdown(
                        f"<div class='crit-score'>"
                        f"🧮 <b>Score automatique : {final_score:.2f} / 4</b>  ·  "
                        f"{final_nb}/{total_q} questions repondues  ·  "
                        f"Niveau : {qualif(final_score)}"
                        f"</div>", unsafe_allow_html=True
                    )

    # ----- Recap global avant soumission -----
    st.markdown("---")
    commentaire_global = st.text_area(
        "Commentaire global / remarques transversales",
        placeholder="(optionnel) Remarques qui ne rentrent pas dans la grille...",
        height=100,
    )

    # Stats globales
    score_g = score_global(st.session_state.form_data)
    nb_criteres_repondus = sum(
        1 for c_id, _, _ in CRITERES
        if score_critere(st.session_state.form_data.get(c_id, {}).get("questions", {})) is not None
    )

    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    col1.metric("Criteres evalues", f"{nb_criteres_repondus} / {TOTAL_CRITERES}")
    col2.metric("Votre score global provisoire",
                f"{score_g:.2f} / 4" if score_g is not None else "—")
    col3.metric("Niveau", qualif(score_g))
    st.progress(nb_criteres_repondus / TOTAL_CRITERES if TOTAL_CRITERES else 0)

    # Validation
    if st.button("✅ Soumettre ma reponse", type="primary", use_container_width=True):
        if not nom.strip():
            st.error("⚠️ Merci d'indiquer votre nom.")
        elif not orga.strip():
            st.error("⚠️ Merci d'indiquer l'organisation evaluee.")
        elif nb_criteres_repondus == 0:
            st.error("⚠️ Merci d'evaluer au moins une question.")
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

    # Scores par axe
    all_scores = []
    for _, row in df.iterrows():
        d = parse_donnees(row["donnees_json"])
        all_scores.append({a: score_axe(d, a) for a in AXES})
    df_scores = pd.DataFrame(all_scores)

    def mean_safe(serie):
        valid = serie.dropna()
        return valid.mean() if len(valid) > 0 else None

    scores_moyens = {a: mean_safe(df_scores[a]) for a in AXES}
    valid_scores = [s for s in scores_moyens.values() if s is not None]
    score_g = sum(valid_scores) / len(valid_scores) if valid_scores else None

    # KPI
    st.subheader("Vue d'ensemble")
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(
        f"<div class='kpi'><div class='kpi-val'>{len(df)}</div>"
        f"<div class='kpi-lbl'>Reponses</div></div>",
        unsafe_allow_html=True)
    c2.markdown(
        f"<div class='kpi'><div class='kpi-val'>"
        f"{score_g:.2f}" if score_g is not None else "—"
        f"</div><div class='kpi-lbl'>Score moyen / 4</div></div>",
        unsafe_allow_html=True)
    c3.markdown(
        f"<div class='kpi'><div class='kpi-val'>"
        f"{qualif(score_g).split(' - ')[0]}"
        f"</div><div class='kpi-lbl'>Niveau moyen</div></div>",
        unsafe_allow_html=True)
    nb_axes_repondus = sum(1 for s in scores_moyens.values() if s is not None)
    c4.markdown(
        f"<div class='kpi'><div class='kpi-val'>{nb_axes_repondus}/5</div>"
        f"<div class='kpi-lbl'>Axes evalues</div></div>",
        unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns(2)

    # Radar moyen
    with col1:
        st.subheader("Profil RTE moyen (radar)")
        valid_axes = [a for a in AXES if scores_moyens[a] is not None]
        if valid_axes:
            fig = go.Figure()
            fig.add_trace(go.Scatterpolar(
                r=[scores_moyens[a] for a in valid_axes],
                theta=[f"Axe {a}<br>{AXES[a][:25]}" for a in valid_axes],
                fill="toself", name="Moyenne",
                line=dict(color="#1F4E79", width=3),
                fillcolor="rgba(31, 78, 121, 0.3)",
            ))
            fig.update_layout(
                polar=dict(
                    radialaxis=dict(visible=True, range=[0, 4], tickfont=dict(size=10)),
                    angularaxis=dict(tickfont=dict(size=11)),
                ),
                showlegend=False, height=420, margin=dict(l=60, r=60, t=20, b=20),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Pas assez de donnees pour le radar.")

    # Barres horizontales
    with col2:
        st.subheader("Scores moyens par axe")
        valid_axes = [a for a in AXES if scores_moyens[a] is not None]
        if valid_axes:
            bar_data = pd.DataFrame({
                "Axe": [f"Axe {a} - {AXES[a]}" for a in valid_axes],
                "Score": [scores_moyens[a] for a in valid_axes],
            })
            fig = px.bar(
                bar_data, x="Score", y="Axe", orientation="h",
                color="Axe", color_discrete_sequence=[AXE_COLORS[a] for a in valid_axes],
                text="Score",
            )
            fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
            fig.update_layout(
                xaxis_range=[0, 4.5], showlegend=False, height=420,
                margin=dict(l=20, r=20, t=20, b=20),
                xaxis_title="Score moyen / 4", yaxis_title=None,
            )
            st.plotly_chart(fig, use_container_width=True)

    # Tableau scores par critere
    st.markdown("---")
    st.subheader("Scores par critere")
    agg = aggregate_responses(df)
    rows = []
    for c_id, axe, titre in CRITERES:
        a = agg[c_id]
        rows.append({
            "Critere": c_id,
            "Axe": f"Axe {axe}",
            "Titre": titre,
            "Score moyen / 4": f"{a['score_moyen']:.2f}" if a["score_moyen"] is not None else "—",
            "Min": f"{a['score_min']:.2f}" if a["score_min"] is not None else "—",
            "Max": f"{a['score_max']:.2f}" if a["score_max"] is not None else "—",
            "Nb reponses": a["nb_reponses"],
        })
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    # Comparaison entre repondants
    if len(df) > 1:
        st.markdown("---")
        st.subheader("Comparaison entre repondants")
        comparison = []
        for _, row in df.iterrows():
            d = parse_donnees(row["donnees_json"])
            entry = {"Repondant": row["nom_repondant"]}
            for a in AXES:
                s = score_axe(d, a)
                entry[f"Axe {a}"] = f"{s:.2f}" if s is not None else "—"
            sg = score_global(d)
            entry["Global"] = f"{sg:.2f}" if sg is not None else "—"
            comparison.append(entry)
        st.dataframe(pd.DataFrame(comparison), hide_index=True, use_container_width=True)

        # Radar comparatif
        st.subheader("Radars superposes")
        fig = go.Figure()
        for _, row in df.iterrows():
            d = parse_donnees(row["donnees_json"])
            scores = [score_axe(d, a) for a in AXES]
            # remplace les None par 0 pour l'affichage
            scores_plot = [s if s is not None else 0 for s in scores]
            fig.add_trace(go.Scatterpolar(
                r=scores_plot,
                theta=[f"Axe {a}" for a in AXES],
                fill="toself", name=row["nom_repondant"], opacity=0.5,
            ))
        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 4])),
            showlegend=True, height=460,
        )
        st.plotly_chart(fig, use_container_width=True)

    # Detail par critere : distribution des statuts par question
    st.markdown("---")
    st.subheader("🔍 Detail par critere : distribution des statuts par question")
    selected_c = st.selectbox(
        "Choisir un critere",
        options=[c[0] for c in CRITERES],
        format_func=lambda x: f"{x} - {dict((c[0], c[2]) for c in CRITERES)[x]}",
    )
    a = agg[selected_c]
    if a["nb_reponses"] == 0:
        st.info("Aucune reponse sur ce critere pour l'instant.")
    else:
        st.markdown(
            f"**Score moyen** : `{a['score_moyen']:.2f} / 4`  ·  "
            f"min `{a['score_min']:.2f}` / max `{a['score_max']:.2f}`  ·  "
            f"**{a['nb_reponses']} reponses**  ·  "
            f"Niveau : {qualif(a['score_moyen'])}"
        )
        for q in CHECKLISTS[selected_c]:
            qa = a["questions"][q["id"]]
            total = qa["obs"] + qa["part"] + qa["non"]
            q_score = ((qa["obs"] * 1.0 + qa["part"] * 0.5) / total * 4) if total > 0 else None
            with st.container():
                st.markdown(f"**{q['id']}** {q['label']}")
                if total > 0:
                    cols = st.columns(4)
                    cols[0].metric("✅ Observe", qa["obs"])
                    cols[1].metric("🟡 Partiel", qa["part"])
                    cols[2].metric("❌ Non observe", qa["non"])
                    cols[3].metric("Score q./4",
                                   f"{q_score:.2f}" if q_score is not None else "—")
                    if qa["preuves"]:
                        with st.expander(f"📎 {len(qa['preuves'])} preuves collectees"):
                            for preuve in qa["preuves"]:
                                st.markdown(f"- {preuve}")
                else:
                    st.caption("Aucune reponse a cette question.")
                st.markdown("---")

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

    st.warning(
        "⚠️ **Pense a sauvegarder regulierement** les reponses. "
        "Si l'app est redeployee, la base est reinitialisee. "
        "Telecharge un backup CSV ci-dessous au moins une fois par semaine."
    )
    col1, col2 = st.columns(2)
    with col1:
        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "💾 Telecharger backup CSV", data=csv_data,
            file_name=f"backup_RTE_{date.today().isoformat()}.csv",
            mime="text/csv", type="primary", use_container_width=True,
        )
    with col2:
        json_data = df.to_json(orient="records", force_ascii=False, indent=2).encode("utf-8")
        st.download_button(
            "📦 Backup JSON complet", data=json_data,
            file_name=f"backup_RTE_{date.today().isoformat()}.json",
            mime="application/json", use_container_width=True,
        )

    st.markdown("---")
    display = df[["id", "date_soumission", "nom_repondant", "fonction",
                  "organisation", "email"]].copy()
    st.dataframe(display, hide_index=True, use_container_width=True)

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
        st.markdown(
            "**Notation auto** :  \n"
            "✅ Observe = 1 pt  \n"
            "🟡 Partiel = 0,5 pt  \n"
            "❌ Non observe = 0 pt  \n  \n"
            "*Score critere = (moyenne / nb questions) × 4*"
        )
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
