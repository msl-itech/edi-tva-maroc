# AGENTS.md — Règles de travail pour agents IA

> Ce fichier est lu en premier par les agents IA (Codex CLI, Claude Code, etc.) avant de commencer à travailler sur ce projet.

---

## Vue d'ensemble

App Streamlit qui transforme un export Odoo Excel en fichier .xlsm DGI Maroc pré-rempli, avec préservation totale du VBA et XML Map natif Excel. Détails complets dans **SPEC.md** et **HANDOFF.md** à la racine.

---

## Ordre de lecture obligatoire avant toute modification

1. `AGENTS.md` (ce fichier) — règles de travail
2. `HANDOFF.md` — état actuel du projet
3. `SPEC.md` — spec métier et décisions verrouillées
4. `app.py` — code principal
5. Fichier(s) spécifique(s) concerné(s) par la tâche

**Ne pas commencer à coder sans avoir lu ces 4 fichiers.**

---

## Commandes principales

```powershell
# Installation
pip install -r requirements.txt

# Lancer l'app en local (sur Windows)
python -m streamlit run app.py

# Test du mapping sans UI (rapide)
python scripts/test_sample.py

# Inspecter un .xlsm généré pour debug
python scripts/inspect_files.py

# Lint (pas configuré par défaut, utiliser ruff ou flake8 si demandé)
# Pas de typecheck obligatoire pour l'instant
```

L'app de production est sur **https://edi-tva-maroc.streamlit.app**. Tout push sur `main` redéploie automatiquement.

---

## Conventions de code

- **Langue des commentaires et messages** : français (le mainteneur est francophone)
- **Strings de l'UI** : français (interface utilisateur)
- **Noms de variables / fonctions** : anglais standard Python (snake_case)
- **Docstrings** : en français acceptable, surtout pour la logique métier
- **Indentation** : 4 espaces (PEP 8)
- **Longueur de ligne** : 100 chars max
- **Imports** : `pathlib`, stdlib, third-party (`pandas`, `streamlit`, `openpyxl`), local — dans cet ordre

---

## Avant de faire des changements

1. **Inspecte d'abord** les fichiers concernés
2. **Explique brièvement** la modification que tu veux faire
3. **Préfère des changements minimaux** plutôt qu'une réécriture
4. **Ne crée pas de nouveaux fichiers** sans nécessité
5. **Ne refactorise pas pour le plaisir** — seulement si demandé
6. **N'ajoute pas de nouvelles dépendances** sans aval explicite (le `requirements.txt` est volontairement minimaliste)

---

## Après avoir fait des changements

1. **Lance `python scripts/test_sample.py`** — doit toujours passer
2. **Vérifie** que l'app se lance sans erreur : `python -m streamlit run app.py`
3. **Si tu as modifié la logique d'injection .xlsm**, lance aussi `scripts/inspect_files.py` pour vérifier que les fichiers binaires critiques (VBA, XML Map, drawings) sont toujours préservés
4. **Résume les fichiers modifiés** et explique les changements en français
5. **Mentionne tout test que tu n'as pas pu exécuter** et pourquoi
6. **Ne commit JAMAIS sans demande explicite** du mainteneur

---

## Contraintes absolues (ne JAMAIS faire)

| ❌ Ne pas faire | ✅ Pourquoi |
|----------------|-----------|
| Modifier `templates/EDI_MAROCAINE_XML_GENERATOR.xlsm` | C'est le template DGI sacré, immuable |
| Faire `load_workbook(...) + save(...)` sur le template | Casse les drawings (bouton "Générer XML") — utiliser l'approche **PATCH ZIP** existante |
| Ajouter une dépendance Python sans accord | Le `requirements.txt` doit rester minimaliste pour Streamlit Cloud |
| Committer des données client réelles | Toujours anonymiser dans samples/ |
| Committer des secrets (.env, secrets.toml) | Sécurité, gitignored |
| Pusher du WIP sur `main` | `main` = prod, auto-déployée. Utiliser des branches feature |
| Renommer les fichiers SPEC.md / HANDOFF.md / AGENTS.md | Ils sont référencés dans la doc utilisateur |
| Re-questionner les 17 décisions de SPEC.md | Elles sont locked. Pour les changer → aval explicite mainteneur |

---

## L'approche PATCH ZIP (technique critique)

Le code dans `app.py` utilise une approche par **patching de l'archive ZIP** du .xlsm plutôt que `openpyxl.load_workbook() + save()`. Cette approche est OBLIGATOIRE car openpyxl perd les drawings (boutons, shapes) lors d'un cycle load+save complet.

Si tu dois modifier la logique d'injection :
- **Garde l'approche PATCH ZIP** intacte
- Modifie uniquement ces fichiers internes du template :
  - `xl/worksheets/sheet1.xml` (data du tableau)
  - `xl/tables/table1.xml` (range et totals)
- **Préserve binaire à 100%** :
  - `xl/vbaProject.bin` (macros)
  - `xl/xmlMaps.xml` (XML Map DGI)
  - `xl/drawings/*` (bouton "Générer XML")
  - `xl/ctrlProps/*` (propriétés bouton)
  - `xl/media/*` (logos)
  - `xl/tables/tableSingleCells1.xml` (bindings header)
- **Vérifie systématiquement** avec `scripts/inspect_files.py` après chaque modification

---

## Workflow Git recommandé

```
1. Nouvelle tâche → créer une branche feature
   git checkout -b feature/nom-de-la-tache

2. Faire les changements minimum nécessaires

3. Tester localement
   python scripts/test_sample.py
   python -m streamlit run app.py

4. Commit avec un message descriptif clair
   git commit -m "feat: description courte" 
   ou "fix: ..." "docs: ..." "refactor: ..."

5. Pousser la branche
   git push origin feature/nom-de-la-tache

6. Créer une PR sur GitHub pour review

7. Merger sur main UNIQUEMENT après validation utilisateur
```

---

## Workflow recommandé pour les questions ouvertes

Plutôt que de prendre une initiative non documentée :

1. **Décris** ce que tu as détecté comme ambigüité
2. **Propose 2-3 options** clairement explicitées
3. **Recommande** celle qui te semble la plus sûre
4. **Attends la décision** du mainteneur avant d'implémenter

Si tu hésites entre 2 approches techniques → arrête-toi et demande. Le mainteneur préfère une question de plus à un refactor non souhaité.

---

## Périmètre v1 vs v2

**v1 (actuelle, déployée)** :
- ✅ Feuille EDI seulement (achats/déductions)
- ✅ Validation stricte + rapport anomalies
- ✅ Génération .xlsm pour Excel local
- ✅ Pas d'auth, pas de multi-client persistant

**v2 (potentielle, NON commencée)** :
- Feuille CA (ventes/encaissements)
- Auth utilisateurs
- Multi-client persistant (config JSON)
- Génération XML directement en Python (option discutée)
- Tests automatisés

Si on te demande une feature v2, **vérifie d'abord** dans HANDOFF.md / SPEC.md qu'elle a été validée par le mainteneur avant d'investir du temps.

---

**Fin de AGENTS.md.** Pour toute évolution de ces règles, demander au mainteneur.
