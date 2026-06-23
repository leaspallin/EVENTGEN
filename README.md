# Générateur de propositions — ALL IN

Application web (Streamlit) qui remplit le **vrai template PowerPoint ALL IN** et
génère une proposition **.pptx** identique au masque. Le **PDF** se fait en 1 clic
depuis PowerPoint (Fichier ▸ Exporter ▸ PDF) — rendu identique.

## Ce que l'équipe fait
1. Ouvre l'URL de l'app.
2. Remplit : client (+ logo), dates, participants, programme, espace, devis.
3. Clique **Générer** → télécharge le **PowerPoint** (puis Export PDF dans PowerPoint).

Aucune installation, aucun réglage. Le rendu est le template, pas une imitation.

## Mettre en ligne (une seule fois, ~5 min, fiable)
1. Crée un dépôt **GitHub** (compte de l'entreprise, pas perso) et dépose
   **le contenu de ce dossier à la RACINE du dépôt** (pas dans un sous-dossier) :
   `app.py`, `generator.py`, `requirements.txt`,
   `template/template.pptx`, `assets/fonts/`, `.streamlit/config.toml`.
   ⚠ `app.py` doit être visible directement à la racine du dépôt.
2. Va sur **share.streamlit.io** → *New app* → choisis le dépôt →
   *Main file path* = **`app.py`** → *Deploy*.
3. Déploiement léger (pas de LibreOffice) : ~1-2 min. Récupère l'URL `…streamlit.app`.

## PDF
Par défaut, l'app livre le **.pptx** et le PDF se fait dans PowerPoint (1 clic, fidélité parfaite).
> Option avancée (PDF généré par le serveur) : ajouter un fichier `packages.txt` contenant
> `libreoffice-impress`. ⚠ Cela alourdit fortement le build et peut faire échouer le
> déploiement sur Streamlit Cloud (limite de ressources). Non recommandé.

## Mettre à jour le template
Remplace `template/template.pptx` (même structure de slides) et pousse sur GitHub :
l'app se met à jour seule. Le mapping des formes est dans `generator.py` (dict `S`).

## Points de vigilance
- **Propriété** : dépôt GitHub + Streamlit sur un compte **entreprise**, pas perso,
  pour que l'outil survive à un départ.
- **Police** : Poppins est embarquée (`assets/fonts/`) et installée au démarrage pour
  le rendu PDF côté PowerPoint/serveur. Ne pas supprimer ce dossier.
- **Logo client** : PNG transparent au format paysage ; substitué en conservant le cadre.
