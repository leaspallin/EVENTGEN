"""
ALL IN — Générateur de propositions séminaire / event.
Formulaire web -> remplit le VRAI template PowerPoint -> exporte PPTX + PDF.
"""
import os, shutil, subprocess
import streamlit as st
import pandas as pd
import generator as g

HERE = os.path.dirname(os.path.abspath(__file__))

# mots-clés permettant de détecter qu'une ligne de devis concerne un module
DEVIS_KEYWORDS = {
    "espace":        ["espace", "salle", "séminaire", "seminaire", "bercy", "toronto", "privatisation"],
    "restauration":  ["restaur", "déjeuner", "dejeuner", "petit-déj", "petit-dej", "tata suzanne",
                      "boisson", "cocktail", "traiteur", "pause"],
    "padel":         ["padel", "pétanque", "petanque", "ping-pong", "ping pong", "americano"],
    "intervention":  ["intervention", "ambassadeur", "tsonga", "ascione", "moreau", "conférenc", "conferenc"],
    "vestiaires":    ["vestiaire"],
    "accessibilite": ["navette", "transfert", "parking"],
}

# ---------------------------------------------------------------- polices
@st.cache_resource
def ensure_fonts():
    """Installe Poppins (embarquée) pour que LibreOffice rende le template fidèlement."""
    src = os.path.join(HERE, "assets", "fonts")
    if not os.path.isdir(src):
        return
    dst = os.path.expanduser("~/.fonts")
    os.makedirs(dst, exist_ok=True)
    for f in os.listdir(src):
        if f.lower().endswith((".ttf", ".otf")):
            shutil.copy(os.path.join(src, f), os.path.join(dst, f))
    try:
        subprocess.run(["fc-cache", "-f"], check=False,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

ensure_fonts()

# ---------------------------------------------------------------- page
st.set_page_config(page_title="ALL IN — Générateur de propositions",
                   page_icon="🎾", layout="wide")

st.markdown("""
<style>
:root{ --gold:#f4bf88; }
.stApp{ background:#0b0c0e; }
h1,h2,h3,h4,label,p,span,div{ color:#ececef; }
.block-container{ padding-top:2rem; max-width:1150px; }
.allin-title{ font-weight:800; letter-spacing:.5px; font-size:1.6rem; }
.allin-title b{ color:var(--gold); }
.allin-sub{ color:#9a9a9a; font-size:.9rem; margin-bottom:1.2rem; }
.stButton>button, .stDownloadButton>button{ border-radius:9px; font-weight:600; }
.stDownloadButton>button{ background:linear-gradient(90deg,#f4bf88,#e0a86a); color:#1a1206; border:none; }
hr{ border-color:#23252c; }
.totbox{ background:#15161a; border:1px solid #2a2c33; border-radius:10px; padding:14px 16px; }
.totbox .l{ display:flex; justify-content:space-between; padding:3px 0; color:#cfcfd4; }
.totbox .big{ color:var(--gold); font-weight:700; border-top:1px solid #2a2c33; margin-top:6px; padding-top:8px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="allin-title">Générateur de propositions <b>· ALL IN</b></div>', unsafe_allow_html=True)
st.markdown('<div class="allin-sub">Pôle séminaire / event — remplis le formulaire, télécharge le PPTX et le PDF générés depuis le vrai template.</div>', unsafe_allow_html=True)

col_form, col_out = st.columns([1.25, 1], gap="large")

with col_form:
    st.subheader("Sections à inclure")
    st.caption("Coche les modules à faire apparaître dans la proposition. "
               "La couverture est toujours incluse ; les sections décochées sont retirées du PowerPoint.")
    include = {}
    _mcols = st.columns(3)
    for _i, (_key, _label, _idx) in enumerate(g.MODULES):
        with _mcols[_i % 3]:
            include[_key] = st.checkbox(_label, value=True, key=f"inc_{_key}")
    st.divider()

    st.subheader("Client & événement")
    c1, c2 = st.columns(2)
    client = c1.text_input("Nom du client", "BYMYCAR")
    dates = c2.text_input("Dates (affichage)", "3 ET 4 JUIN 2026")
    c3, c4 = st.columns(2)
    pax = c3.number_input("Nombre de participants", min_value=1, value=35, step=1)
    espace_label = c4.selectbox("Espace séminaire",
                                ["Bercy", "Toronto", "Bercy + Toronto"], index=2)
    logo_file = st.file_uploader("Logo client (couverture) — PNG/JPG, idéalement transparent",
                                 type=["png", "jpg", "jpeg"])

    st.subheader("Programme")
    programme_txt = st.text_area(
        "Une ligne par créneau (ex : 8h30 – Accueil petit déjeuner)",
        "8h30 – Accueil petit déjeuner\n"
        "9h00 – Début des relances en salle de séminaire\n"
        "12h00 – Déjeuner au restaurant Tata Suzanne\n"
        "13h30 – Reprise des relances\n"
        "17h30 – Activités",
        height=140)

    with st.expander("Photo de la salle (optionnel — remplace la photo d'espace du template)"):
        espace_photo_file = st.file_uploader("Photo salle", type=["png", "jpg", "jpeg"],
                                             key="espacephoto")

    st.subheader("Devis")
    st.caption("Saisis les montants **facturés** HT. Type **Forfait** = prix fixe ; "
               "**Par personne** = × participants. La remise sert d'ancrage (tarif initial barré).")
    devis_df = st.data_editor(
        pd.DataFrame([
            {"Libellé": "La privatisation des espaces séminaire Bercy + Toronto", "Montant € HT": 1386.0, "Type": "Forfait"},
            {"Libellé": "Les 2 petits-déjeuners et 2 déjeuners au restaurant Tata Suzanne", "Montant € HT": 3304.35, "Type": "Forfait"},
            {"Libellé": "L'organisation de l'activité padel (terrains + matériel)", "Montant € HT": 2205.0, "Type": "Forfait"},
            {"Libellé": "Terrain de pétanque et table de ping-pong : offerts", "Montant € HT": 0.0, "Type": "Forfait"},
        ]),
        num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Montant € HT": st.column_config.NumberColumn(format="%.2f", min_value=0.0),
            "Type": st.column_config.SelectboxColumn(options=["Forfait", "Par personne"]),
        })
    c5, c6 = st.columns(2)
    remise = c5.number_input("Remise affichée (%)", min_value=0.0, max_value=90.0, value=10.0, step=0.5)
    tva = c6.number_input("TVA (%)", min_value=0.0, value=20.0, step=0.5)

# -------------------------------------------------- assemble data
espace_map = {"Bercy": "bercy", "Toronto": "toronto", "Bercy + Toronto": "combo"}
devis_lines = []
for _, row in devis_df.iterrows():
    lbl = str(row.get("Libellé", "") or "").strip()
    if not lbl:
        continue
    devis_lines.append({
        "label": lbl,
        "montant": float(row.get("Montant € HT", 0) or 0),
        "type": "pax" if row.get("Type") == "Par personne" else "forfait",
    })

data = {
    "client": client,
    "dates": dates,
    "pax": int(pax),
    "logo_bytes": logo_file.getvalue() if logo_file else None,
    "programme": [l for l in programme_txt.splitlines() if l.strip()],
    "espace": espace_map[espace_label],
    "espace_photo_bytes": espace_photo_file.getvalue() if espace_photo_file else None,
    "devis_lines": devis_lines,
    "remise": float(remise),
    "tva": float(tva),
    "include": include,
}

# -------------------------------------------------- live devis + génération
with col_out:
    st.subheader("Récapitulatif devis")
    d = g.compute_devis(devis_lines, int(pax), float(remise), float(tva))
    def e(n): return f"{n:,.2f} €".replace(",", " ").replace(".", ",")
    anchor = (f'<div class="l"><span>Tarif / pers. (initial → remisé)</span>'
              f'<b>{e(d["pp_brut"])} → {e(d["pp_net"])}</b></div>') if remise > 0 else \
             f'<div class="l"><span>Tarif / personne</span><b>{e(d["pp_net"])}</b></div>'
    st.markdown(f"""
    <div class="totbox">
      {anchor}
      <div class="l"><span>Total HT</span><b>{e(d['net'])}</b></div>
      <div class="l"><span>TVA {tva:g}%</span><b>{e(d['tva'])}</b></div>
      <div class="l big"><span>Total TTC</span><b>{e(d['ttc'])}</b></div>
    </div>
    """, unsafe_allow_html=True)

    # cohérence devis <-> sections : alerter si une ligne concerne un module décoché
    _labels = [l["label"].lower() for l in devis_lines]
    _alerts = []
    for _key, _kws in DEVIS_KEYWORDS.items():
        if not include.get(_key, True):
            _hits = [devis_lines[_i]["label"] for _i, _lab in enumerate(_labels)
                     if any(_k in _lab for _k in _kws)]
            if _hits:
                _modlabel = next(lb for k, lb, _ in g.MODULES if k == _key)
                _alerts.append((_modlabel, _hits))
    for _modlabel, _hits in _alerts:
        st.warning(f"⚠ Section **« {_modlabel} »** décochée, mais le devis la facture : "
                   + " ; ".join(f"« {h} »" for h in _hits))

    st.write("")
    safe = "".join(ch if ch.isalnum() else "_" for ch in (client or "PROPOSITION")).strip("_") or "ALLIN"
    if st.button("⚙️ Générer la proposition", type="primary", use_container_width=True):
        if not data["devis_lines"]:
            st.error("Ajoute au moins une ligne de devis.")
        else:
            try:
                with st.spinner("Remplissage du template…"):
                    pptx_bytes, _ = g.build(data)
                st.session_state["pptx"] = pptx_bytes
                st.session_state["safe"] = safe
                st.session_state["pdf"] = None
                # PDF best-effort : seulement si un convertisseur est présent
                if g.pdf_available():
                    try:
                        with st.spinner("Conversion PDF…"):
                            st.session_state["pdf"] = g.to_pdf(pptx_bytes)
                    except Exception:
                        st.session_state["pdf"] = None
                st.success("Proposition générée ✔")
            except Exception as ex:
                st.error(f"Erreur : {ex}")

    if st.session_state.get("pptx"):
        s = st.session_state["safe"]
        st.download_button("⬇️ Télécharger le PowerPoint (.pptx)", st.session_state["pptx"],
                           file_name=f"PROPOSITION_{s}.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                           use_container_width=True)
        if st.session_state.get("pdf"):
            st.download_button("⬇️ Télécharger le PDF", st.session_state["pdf"],
                               file_name=f"PROPOSITION_{s}.pdf", mime="application/pdf",
                               use_container_width=True)
        else:
            st.info("**PDF** : ouvre le .pptx puis **Fichier ▸ Exporter ▸ PDF** "
                    "(1 clic, rendu identique au template).")

st.markdown("<hr>", unsafe_allow_html=True)
st.caption("Rendu identique au template ALL IN (masque, charte, devis). "
           "Logo client et photo de salle substitués en conservant le cadre d'origine.")
