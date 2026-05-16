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
- ✅ Génération .xlsm avec préservation totale :
  - VBA (`vbaProject.bin`)
  - XML Map DGI (`xmlMaps.xml`)
  - Bouton "Générer XML" (drawings + ctrlProps)
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
│   ├── test_sample.py                     # Test du mapping sur le sample
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
| `scripts/test_sample.py` | Test isolé du mapping (sans Streamlit) | OUI pour ajouter des cas de test |

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

## 🛠️ Approche technique : PATCH ZIP (critique à comprendre)

**Problème historique** : `openpyxl` load_workbook + save **perd** les drawings (boutons, shapes) et certains XML internes du .xlsm.

**Solution implémentée** : ne PAS faire `load + save` complet. À la place :
1. Copier le template binaire intégralement vers une nouvelle archive ZIP en mémoire
2. Ne modifier QUE 3 fichiers internes :
   - `xl/worksheets/sheet1.xml` (data + heights + widths)
   - `xl/tables/table1.xml` (range + totals)
   - (Et le header header bindings dans `xl/tables/tableSingleCells1.xml` reste identique)
3. Préserver TOUS les autres fichiers tels quels (vbaProject.bin, xmlMaps.xml, drawings, ctrlProps, media, etc.)

**Cette approche est la SEULE qui fonctionne** pour préserver le bouton "Générer XML" du template. Toute reprise du code doit conserver cette approche.

Fichiers du template à NE JAMAIS toucher :
- `xl/vbaProject.bin` — la macro
- `xl/xmlMaps.xml` — le mapping XML DGI
- `xl/drawings/*.xml` — le bouton
- `xl/ctrlProps/*.xml` — propriétés du bouton
- `xl/media/*` — logos
- `xl/tables/tableSingleCells1.xml` — bindings du header

Fichiers acceptables à modifier :
- `xl/worksheets/sheet1.xml` — le contenu du tableau
- `xl/tables/table1.xml` — la range du tableau et la totals row

---

## 🐛 Historique des bugs résolus

| # | Bug | Solution appliquée |
|---|-----|--------------------|
| 1 | Style table alternance perdu après la 8e ligne | Copie programmatique du `_style` d'une row du milieu pour toutes les nouvelles rows |
| 2 | Ligne de totaux effacée et vide | Préservation du style row 17 + formules SUBTOTAL réinjectées sur la nouvelle dernière ligne |
| 3 | Bouton "Générer XML" disparu du fichier généré | Bascule complète vers l'approche **PATCH ZIP** au lieu de openpyxl load+save |
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
4. **Ne jamais casser** l'approche PATCH ZIP — si on doit modifier app.py, garder cette logique
5. **Toujours tester** sur `samples/Odoo_template.xlsx` AVANT de pousser
6. **Toujours vérifier** par script que les fichiers binaires critiques sont préservés (`scripts/inspect_files.py`)

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
