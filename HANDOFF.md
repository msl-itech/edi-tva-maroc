# HANDOFF.md — État du projet EDI TVA Maroc

> Document de passation pour reprise du projet par un autre agent (Codex CLI / autre).
> Dernière mise à jour : **16 mai 2026**.
> Mainteneur : El Houssine BOUHMAIDA · Cabinet MSL-iTECH (Marrakech).

---

## 🎯 But du projet

App Streamlit qui transforme un export Excel Odoo de factures fournisseurs en un fichier .xlsm rempli, basé sur un template DGI (Direction Générale des Impôts Maroc). Le .xlsm contient une **macro VBA** et un **XML Map** natif Excel qui, une fois cliqués dans Excel Windows, génèrent un XML de déclaration TVA prêt à uploader sur le portail **SIMPL-TVA** de la DGI.

**Chaîne complète** :
```
Odoo .xlsx  →  Streamlit (mapping/validation)  →  .xlsm rempli  
            →  Excel Windows + macro VBA       →  XML DGI  →  SIMPL-TVA
```

---

## ✅ Statut actuel : V1 DÉPLOYÉE EN PRODUCTION

**URL live** : https://edi-tva-maroc.streamlit.app
**Repo GitHub** : https://github.com/msl-itech/edi-tva-maroc (public, MIT)
**Stack** : Python 3.13 (cloud) / 3.14 (local dev) · Streamlit 1.57 · openpyxl 3.1 · pandas 2.3

### Ce qui fonctionne (validé manuellement)

- ✅ Upload export Odoo .xlsx (12 colonnes attendues)
- ✅ Validation stricte ligne par ligne (rapport d'anomalies)
- ✅ Mapping complet Odoo → EDI (13 colonnes)
- ✅ Génération .xlsm avec préservation contrôlée par `scripts/test_sample.py` :
  - VBA (`vbaProject.bin`)
  - XML Map DGI (`xmlMaps.xml`)
  - Bouton "Générer XML" (`xl/drawings/drawing1.xml`)
  - Header bindings (`tableSingleCells1.xml`)
  - Images / logos
- ✅ Style table alternance bleue sur toutes les lignes (pas seulement les 8 premières)
- ✅ Ligne Total avec formules SUBTOTAL
- ✅ Hauteurs de lignes uniformes (17.25)
- ✅ Largeurs de colonnes auto-fit dynamiques
- ✅ XML DGI généré dans Excel Windows = conforme schéma SIMPL-TVA

### Ce qui n'a PAS encore été testé en prod

- ⚠️ End-to-end **sur la version déployée Streamlit Cloud** : upload réel → génération → download (testé seulement en script local)
- ⚠️ Comportement multi-utilisateurs simultanés
- ⚠️ Performance sur fichiers Odoo > 500 lignes (testé sur 80 lignes max)

---

## 📂 Structure du projet

```
edi-tva-maroc/
├── app.py                      # ← Application Streamlit principale (point d'entrée)
├── requirements.txt            # streamlit, openpyxl, pandas
├── README.md                   # Doc utilisateur (généré + édité)
├── SPEC.md                     # ⭐ Spec métier exhaustive — LIRE EN PREMIER
├── HANDOFF.md                  # Ce fichier (état du projet)
├── AGENTS.md                   # Règles de travail pour agents IA
├── LICENSE                     # MIT
├── .gitignore                  # Ignore .claude/, scripts/out_test.xlsm
├── templates/
│   └── EDI_MAROCAINE_XML_GENERATOR.xlsm   # Template avec VBA + XML Map (NE PAS MODIFIER)
├── samples/
│   └── Odoo_template.xlsx                 # Sample de test (80 lignes)
├── scripts/
│   ├── inspect_files.py                   # Utilitaire inspection des .xlsm
│   ├── test_sample.py                     # Test mapping + contrat préservation .xlsm
│   └── out_test.xlsm                      # Sortie de test (ignoré par git)
└── docs/
    └── superpowers/plans/                 # Plans de travail Claude Code (historique)
```

---

## 🔑 Fichiers critiques à connaître

| Fichier | Rôle | Doit-on le modifier ? |
|---------|------|----------------------|
| `app.py` | App Streamlit complète (UI + mapping + validation + injection) | OUI pour features/bugfix |
| `templates/EDI_MAROCAINE_XML_GENERATOR.xlsm` | Template DGI avec VBA et XML Map | **JAMAIS** |
| `samples/Odoo_template.xlsx` | Sample de test (80 lignes) | À régénérer si format Odoo change |
| `SPEC.md` | Spec métier complète (17 décisions verrouillées) | Modifier UNIQUEMENT avec aval du mainteneur |
| `scripts/test_sample.py` | Test isolé mapping + contrat de préservation `.xlsm` | OUI pour ajouter des cas de test |

---

## 🏛️ Décisions architecturales clés (résumé — détails dans SPEC.md)

1. **Validation stricte** : si un champ requis manque → bloquer génération avec rapport ligne par ligne (pas de fix automatique)
2. **Pas de filtre période** : l'utilisateur filtre dans Odoo en amont
3. **Lignes Taxe=0 incluses** avec taux 0% (assurances exonérées)
4. **ICE en string 15 chars** (zéros à gauche préservés, format cellule `@`)
5. **TAUX en float décimal** (0.20) avec format cellule `0%` → XML envoie 0.2, affichage Excel 20%
6. **HT/TVA/TTC écrits directs depuis Odoo** (écrasent les formules calculées du template)
7. **Méthodes paiement** : ESPECES=1, CHEQUE=2, PRELEVEMENT=3, VIREMENT=4, LCN=5
8. **Header saisi manuellement** à chaque génération (pas de config multi-client persistante)

---

## 🛠️ Approche technique actuelle : openpyxl + réparation ZIP post-save

**Problème historique** : `openpyxl` load_workbook + save **perd** les drawings (boutons, shapes) et certains XML internes du .xlsm.

**Solution implémentée aujourd'hui** : ce n'est pas un PATCH ZIP strict. Le code :
1. Charge le template en mémoire avec `openpyxl.load_workbook(..., keep_vba=True)`
2. Modifie les cellules, styles, largeurs, hauteurs, range du tableau et formules via openpyxl
3. Sauvegarde en mémoire avec `wb.save(...)`
4. Répare ensuite l'archive ZIP générée pour réinjecter/écraser les parties critiques que openpyxl supprime ou réécrit mal

Réparations actuelles dans `app.py` :
- Réinjecter si absent :
  - `xl/xmlMaps.xml` — XML Map DGI
  - `xl/tables/tableSingleCells1.xml` — bindings header C3..C6
- Écraser depuis le template :
  - `xl/drawings/drawing1.xml` — bouton "Générer XML" avec macro `[0]!Export216`
- Patcher les relations/content types :
  - `[Content_Types].xml`
  - `xl/_rels/workbook.xml.rels`
  - `xl/worksheets/_rels/sheet1.xml.rels`

**Contrat de sécurité** : `python scripts/test_sample.py` vérifie maintenant explicitement la préservation `.xlsm` : ZIP valide, VBA/XML Map/tableSingleCells/drawing byte-identiques, relations drawing + médias, content type macro-enabled, range/autofilter du tableau, chargement openpyxl `keep_vba=True`, formules de totaux et formats ICE/TAUX/dates.

Une implémentation PATCH ZIP stricte (modifier directement `xl/worksheets/sheet1.xml` et `xl/tables/table1.xml` sans `wb.save`) reste une **tâche future possible de durcissement**, pas l'état courant du projet.

Fichiers du template à NE JAMAIS toucher :
- `xl/vbaProject.bin` — la macro
- `xl/xmlMaps.xml` — le mapping XML DGI
- `xl/drawings/drawing1.xml` — le bouton
- `xl/ctrlProps/*.xml` — propriétés du bouton, si un futur template en contient
- `xl/media/*` — logos
- `xl/tables/tableSingleCells1.xml` — bindings du header

Fichiers que la logique de génération modifie actuellement via openpyxl :
- `xl/worksheets/sheet1.xml` — le contenu du tableau
- `xl/tables/table1.xml` — la range du tableau et la totals row

### Note named ranges / print areas

Investigation du 16 mai 2026 :
- Le template contient deux defined names invalides côté feuille CA :
  - `_xlnm._FilterDatabase` → `CA!#REF!`
  - `_xlnm.Print_Area` → `CA!#REF!`
- Le fichier généré ne conserve plus ces defined names CA invalides. Il garde uniquement la zone d'impression EDI.
- La zone d'impression EDI reste `EDI!$A$1:$M$17` même quand le tableau généré s'étend au-delà de la ligne 17.
- Ce n'est pas considéré bloquant pour la génération XML : l'export XML dépend de `xl/tables/table1.xml`, du XML Map et de `xl/tables/tableSingleCells1.xml`, pas de la zone d'impression.
- Exemple validé avec 43 lignes valides : `Tableau5.ref = A8:M52` et `autoFilter.ref = A8:M51`. C'est intentionnel : la ligne 52 est la ligne de totaux, incluse dans la table mais exclue de l'autofilter.
- Amélioration optionnelle future : mettre à jour la zone d'impression EDI en `A1:M{last_row}` si l'impression ou l'export PDF devient un besoin utilisateur.

## Manual Validation Checklist - Excel Windows + Streamlit Cloud

### Test Matrix

Run the full checklist with these input cases:

- [ ] 0 valid rows: all rows invalid, generation must be blocked.
- [ ] 1 valid row: generation must succeed, table must contain 1 data row + totals row.
- [ ] Around 43 valid rows: use current sample-like volume.
- [ ] 500+ valid rows: validate performance, table expansion, download, and Excel behavior.

### Local / Streamlit App Flow

For each test file:

- [ ] Open the app locally or on Streamlit Cloud.
- [ ] Fill declaration header:
  - [ ] Raison sociale
  - [ ] IF
  - [ ] Année
  - [ ] Régime (`1 = TVA mensuel`, `2 = TVA trimestriel`)
  - [ ] Période (mois 1-12 ou trimestre 1-4 selon le régime)
- [ ] Upload the real Odoo `.xlsx` export.
- [ ] Confirm the app reads the file without crash.
- [ ] Confirm the app shows expected input row count.
- [ ] Confirm invalid rows appear in the anomaly report with line number, reference, partner, and errors.
- [ ] Confirm generation is blocked when anomalies exist.
- [ ] If generation is expected, correct/filter the input so there are no blocking anomalies.
- [ ] Generate the `.xlsm`.
- [ ] Download the `.xlsm`.
- [ ] Record generated filename and timestamp.

### Excel Windows Validation

For each generated `.xlsm`:

- [ ] Open the file in Microsoft Excel on Windows.
- [ ] Confirm Excel does not report corruption or repair the workbook.
- [ ] Enable macros / content when prompted.
- [ ] Confirm the "Générer XML" button is visible.
- [ ] Confirm logos/images are visible.
- [ ] Confirm the EDI table is populated with the expected number of rows.
- [ ] Confirm the totals row is present.
- [ ] Confirm TAUX displays as percent, e.g. `20%`, while preserving decimal XML behavior.
- [ ] Confirm ICE values keep leading zeros and 15-character format.
- [ ] Confirm dates display correctly.
- [ ] Click the "Générer XML" button.
- [ ] Confirm Excel generates an XML file.
- [ ] Open the XML file in a text editor.
- [ ] Confirm XML contains expected header fields:
  - [ ] identifiantFiscal
  - [ ] annee
  - [ ] periode
  - [ ] regime
- [ ] Confirm XML contains expected `rd` rows.
- [ ] Confirm row count in XML matches valid EDI rows, not including totals.
- [ ] Confirm sample fields are present:
  - [ ] ord
  - [ ] num
  - [ ] des
  - [ ] mht
  - [ ] tva
  - [ ] ttc
  - [ ] refF/if
  - [ ] refF/nom
  - [ ] refF/ice
  - [ ] tx
  - [ ] mp/id
  - [ ] dpai
  - [ ] dfac

### Streamlit Cloud Parity

For the same validated input file:

- [ ] Generate `.xlsm` locally.
- [ ] Generate `.xlsm` from Streamlit Cloud.
- [ ] Open both files in Excel Windows.
- [ ] Confirm both files show the button.
- [ ] Confirm both files generate XML successfully.
- [ ] Confirm both XML outputs have the same row count and key values.
- [ ] Confirm no Streamlit Cloud-specific download or file corruption issue.

### Failure Evidence To Capture

For any failure, save:

- [ ] Input Odoo `.xlsx` used.
- [ ] Generated `.xlsm`.
- [ ] Generated XML, if any.
- [ ] Screenshot of Streamlit error/anomaly screen.
- [ ] Screenshot of Excel warning, corruption repair dialog, or missing button.
- [ ] Screenshot of macro error, if any.
- [ ] Browser download filename and timestamp.
- [ ] Streamlit Cloud logs or console output, if available.
- [ ] Local terminal output from:

```powershell
python scripts/test_sample.py
```

- [ ] Notes with exact steps to reproduce.

---

## 🐛 Historique des bugs résolus

| # | Bug | Solution appliquée |
|---|-----|--------------------|
| 1 | Style table alternance perdu après la 8e ligne | Copie programmatique du `_style` d'une row du milieu pour toutes les nouvelles rows |
| 2 | Ligne de totaux effacée et vide | Préservation du style row 17 + formules SUBTOTAL réinjectées sur la nouvelle dernière ligne |
| 3 | Bouton "Générer XML" disparu du fichier généré | Réparation ZIP post-save : `drawing1.xml` est écrasé depuis le template après `wb.save()` |
| 4 | Hauteurs de lignes hétérogènes | `row_dimensions[r].height = 17.25` sur toutes les data rows |
| 5 | Colonnes trop étroites pour noms fournisseurs longs | Auto-fit dynamique : `width = max_len * 1.1 + 2`, plancher = template, plafond = 50 |

---

## 🎯 Tâches potentielles (non engagées, par priorité de valeur)

### Priorité haute
1. **Tester end-to-end le déploiement Streamlit Cloud** (upload Odoo réel → download .xlsm → vérif Excel)
2. **Documenter dans README.md** le workflow utilisateur étape par étape (captures d'écran)
3. **Ajouter une option de test avec sample embarqué** dans l'app (bouton "Tester avec un fichier exemple")

### Priorité moyenne
4. **Authentification utilisateurs** (streamlit-authenticator + secrets.toml) — l'URL prod est actuellement publique
5. **Multi-clients persistant** : config JSON avec liste des sociétés du cabinet (raison sociale, IF) pour éviter saisie manuelle
6. **Génération XML directe en Python** (option B discutée dans la conversation initiale) : éviter le passage par Excel local
7. **Support feuille CA** (ventes/encaissements) en plus de l'EDI (achats/déductions)

### Priorité basse
8. **Tests automatisés** (pytest) sur le module de validation/mapping
9. **CI GitHub Actions** : test à chaque PR
10. **Logging** + suivi des générations (qui a généré quoi, quand)

---

## 🧰 Commandes utiles

### Installation locale
```powershell
pip install -r requirements.txt
```

### Lancer l'app en local
```powershell
python -m streamlit run app.py
```
*(sur Windows si `streamlit` n'est pas dans le PATH)*

### Tester le mapping sans UI
```powershell
python scripts/test_sample.py
```

### Inspecter un .xlsm (debug)
```powershell
python scripts/inspect_files.py
```

### Workflow Git standard
```powershell
git checkout -b feature/nom-de-la-feature   # nouvelle branche
# ... modifs ...
git status
git add .
git commit -m "feat: description claire"
git push origin feature/nom-de-la-feature
# Puis créer une PR sur GitHub
```

### Déploiement Streamlit Cloud
Connecté automatiquement au repo `msl-itech/edi-tva-maroc`, branche `main`.
**Tout push sur `main` redéploie automatiquement** la version prod.
→ Ne pas pusher de WIP sur `main`, utiliser des branches feature.

---

## ⚠️ Contraintes absolues à respecter

1. **Ne jamais modifier** `templates/EDI_MAROCAINE_XML_GENERATOR.xlsm` (c'est le template DGI sacré)
2. **Ne jamais committer** de données client réelles (anonymiser les samples)
3. **Ne jamais committer** de secrets (`.streamlit/secrets.toml` est gitignored)
4. **Ne jamais casser** la stratégie actuelle openpyxl + réparation ZIP post-save — si on doit modifier app.py, garder cette logique sauf aval explicite
5. **Toujours tester** sur `samples/Odoo_template.xlsx` AVANT de pousser
6. **Toujours vérifier** par script que les fichiers critiques `.xlsm` sont préservés (`python scripts/test_sample.py`, puis `scripts/inspect_files.py` si besoin de debug)

---

## 🤝 Contexte de cette passation

Le projet a été spécifié et démarré dans **Claude.ai web** (chat de spec/architecture), puis l'implémentation a été déléguée à **Claude Code** en local. Pour économiser les tokens, la suite des travaux passe maintenant sur **Codex CLI** dans VS Code.

**Toutes les décisions métier sont locked dans SPEC.md** — pas besoin de re-questionner ces points.

**Pour toute reprise** :
1. Lire SPEC.md en entier
2. Lire ce HANDOFF.md
3. Lire app.py (~ structure principale)
4. Lancer `python scripts/test_sample.py` pour vérifier que tout marche localement
5. Tester l'app live sur https://edi-tva-maroc.streamlit.app

---

**Fin du HANDOFF.md.** Pour modifications majeures, demander validation au mainteneur.
