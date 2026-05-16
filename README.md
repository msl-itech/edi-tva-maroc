# EDI TVA Maroc — Générateur de déclaration (Streamlit)

Cabinet **MSL-iTECH** — Mapping Odoo → template Excel `.xlsm` (avec VBA + XML Map DGI).

L'app Streamlit prend en entrée un export Excel des factures fournisseurs depuis Odoo et produit un `.xlsm` pré-rempli ; le comptable ouvre ce fichier dans Excel Windows et déclenche la macro / XML Map pour générer le XML SIMPL-TVA à uploader sur le portail DGI.

> Spec complète et décisions verrouillées : [`SPEC.md`](SPEC.md).

---

## Installation locale

```bash
python -m venv .venv
.venv\Scripts\activate            # Windows PowerShell
pip install -r requirements.txt
streamlit run app.py
```

L'app s'ouvre sur http://localhost:8501.

## Workflow utilisateur

1. **Paramètres déclaration** (formulaire) : Raison sociale, IF, Année, Période (mois), Régime (Débit ou Encaissement).
2. **Upload** du fichier Odoo `.xlsx` (12 colonnes attendues, voir SPEC §"Structure de l'export Odoo").
3. **Validation** : l'app affiche le nombre de lignes valides + un tableau d'anomalies bloquantes ligne par ligne. **Tant qu'une seule anomalie persiste, la génération est bloquée.**
4. **Génération** : clic sur "🚀 Générer" → téléchargement du `.xlsm`.
5. **Côté Excel Windows** : ouvrir le `.xlsm`, autoriser les macros, déclencher la macro / la commande d'export XML du XML Map → fichier XML prêt pour le portail SIMPL-TVA.

## Anomalies détectées

Liste exhaustive des règles bloquantes (voir `validate_and_transform` dans `app.py`) :

- Champs Odoo manquants : Référence / Libellé / Partenaire / HT / Taxe / Total / IF / ICE / Méthode de paiement / Date de paiement / Date de facturation
- Méthode de paiement hors `{ESPECES, CHEQUE, PRELEVEMENT, VIREMENT, LCN}` (case-insensitive, espaces tolérés)
- IF non convertible en entier
- ICE non numérique ou > 15 chiffres (sinon paddé à gauche par `zfill(15)`)
- Taux non détectable dans la colonne *Lignes de facture/Taxes* alors que `Taxe ≠ 0`
- Incohérence `HT + Taxe ≠ Total` avec tolérance 0,05 DH

Les lignes exonérées (Taxe = 0 et colonne taxe vide) sont **incluses** dans le relevé avec un taux 0 %.

## Préservation VBA + XML Map (critique)

openpyxl conserve `xl/vbaProject.bin` quand on charge avec `keep_vba=True`, mais **drop** :
- `xl/xmlMaps.xml`
- `xl/tables/tableSingleCells1.xml`

L'app patche le ZIP en post-save (`_preserve_xml_map` dans `app.py`) : ré-injection des deux parties + ajout des `<Override>` dans `[Content_Types].xml` et des relations dans `xl/_rels/workbook.xml.rels` + `xl/worksheets/_rels/sheet1.xml.rels`. Vérification automatique : `python scripts/test_sample.py`.

> **Avertissement Excel attendu** : les colonnes D (`M_HT`) et E (`TVA`) ont une `calculatedColumnFormula` dans le template. Comme on écrit la valeur de la cellule directement, Excel peut afficher un coin vert "valeur incohérente avec la formule calculée". La valeur exportée dans le XML est bien celle écrite par l'app — l'avertissement est cosmétique.

## Test E2E

```bash
python scripts/test_sample.py
```

Sur le sample fourni (`samples/Odoo_template.xlsx`, 80 lignes) : ~43 valides / ~37 anomalies. Vérifie aussi la préservation de `xl/vbaProject.bin`, `xl/xmlMaps.xml`, `xl/tables/*`.

## Déploiement Streamlit Community Cloud

1. **Pousser** sur GitHub (repo public requis pour le tier gratuit).
2. Aller sur https://share.streamlit.io → **New app**.
3. Sélectionner le repo, branch `main`, main file path `app.py`.
4. Click **Deploy**. URL générée du type `https://<account>-edi-tva-maroc.streamlit.app`.

Streamlit Cloud installe automatiquement les deps de `requirements.txt`. Aucune secret/config nécessaire en v1 (pas d'auth, multi-utilisateurs accédant à la même URL).

## Sécurité (v1)

Aucune authentification — URL publique = tout visiteur peut générer un `.xlsm`. Aucune donnée n'est persistée côté serveur (traitement 100 % en mémoire). Pour durcir : ajouter `streamlit-authenticator` + `.streamlit/secrets.toml`, ou déployer en privé sur un VPS interne MSL-iTECH.

## Structure du repo

```
edi-tva-maroc/
├── app.py                                    # Streamlit app (mono-fichier)
├── requirements.txt
├── README.md                                 # ce fichier
├── SPEC.md                                   # spec gelée
├── templates/EDI_MAROCAINE_XML_GENERATOR.xlsm
├── samples/Odoo_template.xlsx
├── scripts/
│   ├── inspect_files.py                      # inspection du template + sample
│   └── test_sample.py                        # test E2E (validation + ZIP check)
└── docs/superpowers/plans/2026-05-16-edi-tva-maroc.md
```
