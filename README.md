# Formulaire collaboratif - Grille de maturité RTE

Application web pour évaluer collectivement la maturité RTE (Responsabilité Territoriale des Entreprises) d'une organisation.

## Stack

- **Streamlit** — interface web
- **Plotly** — graphiques interactifs (radar, barres)
- **SQLite** — stockage des réponses
- **python-docx** — export de la synthèse en Word

## Structure

- `app_rte.py` — application principale
- `requirements.txt` — dépendances Python
- `.streamlit/config.toml` — thème et configuration

## Contexte

Stage M2 Intelligence Économique / IAE Poitiers — Chaire TerrESS / Bordeaux Sciences Agro.

Grille de maturité RTE V4 — 5 axes, 17 critères, basée sur les travaux de Filippi (2022, 2024), Girardot (2004), Chaire TerrESS (2024).
