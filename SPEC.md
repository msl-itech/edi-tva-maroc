# SPEC.md — EDI TVA Maroc · Générateur Streamlit

> **Document de spécification complet** pour reprise dans Claude Code.
> Cabinet **MSL-iTECH** (Marrakech) · Projet client **Ait Oukhali Travaux** et autres clients TVA.
> Version 1.0 — Spec gelée le 16 mai 2026.

---

## 🚀 PROMPT D'INITIALISATION CLAUDE CODE

> Copie-colle ce bloc au démarrage de ta session Claude Code, avec les 2 fichiers Excel attachés (template EDI + sample export Odoo) :

```
Salam Claude. Je veux construire une app Streamlit qui prend en entrée un export 
Excel Odoo de factures fournisseurs et qui remplit un template EDI .xlsm 
(préservant le VBA et un XML Map vers la DGI Maroc). L'utilisateur télécharge 
ensuite le .xlsm, l'ouvre dans Excel Windows, et clique sur la macro pour 
générer le XML de déclaration TVA.

La spec complète est dans SPEC.md (lis-la EN ENTIER avant de coder). Le template 
EDI est dans templates/EDI_MAROCAINE_XML_GENERATOR.xlsm. Un sample d'export Odoo 
est dans samples/Odoo_template.xlsx.

Ta mission :
1. Lis SPEC.md intégralement
2. Crée la structure de projet selon la section "Architecture"
3. Code l'app Streamlit en respectant TOUTES les 14+ décisions verrouillées
4. Teste sur le sample Odoo réel — montre-moi les anomalies détectées
5. Génère un .xlsm de sortie et VÉRIFIE que le VBA + XML Map sont préservés 
   (test critique : ouvrir le fichier généré, vérifier que xl/vbaProject.bin 
   et xl/xmlMaps.xml existent encore dans le ZIP)
6. Initialise git, prépare le repo pour déploiement Streamlit Cloud
7. Donne-moi les instructions de déploiement

IMPORTANT : ne dévie d'AUCUNE décision verrouillée sans me demander. Si tu 
détectes une ambiguïté ou un edge case non couvert, pose-moi la question 
avant de coder.

GO.
```

---

## 📋 CONTEXTE MÉTIER

### Le problème
La déclaration TVA mensuelle (régime débit) ou trimestrielle (régime encaissement) au Maroc se fait via le portail **SIMPL-TVA** de la DGI (Direction Générale des Impôts). Les entreprises doivent soumettre :
- Un **Relevé de déductions** (achats avec TVA déductible) — *périmètre de cette v1*
- Un **État des encaissements** (ventes avec TVA collectée) — *out of scope v1*

Ces relevés peuvent être saisis manuellement OU importés via un **fichier XML** au format DGI.

### Le workflow actuel (manuel, douloureux)
1. Le comptable exporte les factures fournisseurs depuis **Odoo** (en Excel)
2. Il **recopie manuellement** ligne par ligne dans un template Excel .xlsm
3. Ce template contient un **XML Map** Excel branché sur le schéma DGI
4. Il clique sur "Exporter XML" dans Excel → fichier XML prêt à uploader sur SIMPL-TVA

### Le workflow cible (avec cette app)
1. Exporter depuis Odoo (.xlsx)
2. Ouvrir l'app Streamlit dans le navigateur
3. Saisir 5 paramètres de déclaration (raison sociale, IF, année, période, régime)
4. Uploader le .xlsx Odoo
5. L'app valide les lignes, affiche les anomalies bloquantes s'il y en a
6. Si tout est OK → bouton "Générer" → téléchargement du .xlsm pré-rempli
7. Ouvrir le .xlsm dans Excel Windows local → clic sur le bouton macro → XML DGI prêt
8. Upload sur SIMPL-TVA

### Pourquoi Streamlit ne génère pas le XML directement
- Streamlit tourne sur Linux (Streamlit Cloud)
- Les macros VBA ne s'exécutent que dans Microsoft Excel (Windows/Mac)
- `xlwings` nécessiterait Excel installé → impossible sur cloud Linux
- Donc Streamlit fait **uniquement le mapping** ; la macro VBA + XML Map du template gèrent la suite

---

## 🔒 DÉCISIONS VERROUILLÉES (NE PAS MODIFIER SANS AVAL)

### Périmètre fonctionnel
1. **v1 = feuille EDI uniquement** (achats/déductions). Feuille CA (ventes/encaissements) explicitement **hors scope** v1.
2. **Pas de filtre de période côté app**. L'utilisateur filtre dans Odoo avant l'export.
3. **Validation stricte** : si un champ requis manque sur une ligne → générer un rapport ligne par ligne et **bloquer la génération** du .xlsm tant que ce n'est pas corrigé.

### Header de la déclaration (saisie manuelle dans l'UI)
4. Les 5 champs du header EDI sont saisis dans Streamlit à chaque génération (pas de config par client) :
   - `Raison sociale` (texte libre)
   - `Identifiant fiscal IF` (texte/nombre)
   - `Année` (dropdown 2016–2030)
   - `Période` (dropdown 1–12, représente le mois)
   - `Régime` (radio : 1 = Encaissement, 2 = Débit)

### Mapping des colonnes (Odoo → EDI)

| EDI col | EDI label (XML field) | ← Source Odoo (12 cols) | Transformation |
|--------:|:----------------------|:------------------------|:---------------|
| A | `OR` (ord) | (auto) | Séquence 1..N, type int |
| B | `FACT_NUM` (num) | A `Référence` | string, strip |
| C | `DESIGNATION` (des) | B `Libellé` | string, strip |
| D | `M_HT` (mht) | C `Montant hors taxes` | **float, écrase la formule Excel** |
| E | `TVA` (tva) | D `Taxe` | **float, écrase la formule Excel** |
| F | `M_TTC` (ttc) | E `Total` | float (direct, pas de calcul) |
| G | `IF` (if) | G `IF` | **converti en int** (xmlDataType=int) |
| H | `LIB_FRSS` (nom) | H `Partenaire` | string, strip |
| I | `ICE_FRS` (ice) | I `ICE` | **string 15 caractères, zéros à gauche préservés** |
| J | `TAUX` (tx) | F `Lignes de facture/Taxes` | **Parse "20% 146" → 0.20** (float), cellule formatée `0%` |
| K | `ID_PAIE` (mp/id) | J `Méthode de paiement` | Lookup → 1-5 (voir table ci-dessous) |
| L | `DATE_PAIE` (dpai) | K `Date de paiement` | datetime, format `yyyy-mm-dd` |
| M | `DATE_FAC` (dfac) | L `Date de facturation` | datetime, format `yyyy-mm-dd` |

### Table de mapping des méthodes de paiement
```
ESPECES      → 1   (Espèce)
CHEQUE       → 2   (Chèque)
PRELEVEMENT  → 3   (Prélèvement)
VIREMENT     → 4   (Virement)
LCN          → 5   (Effet / Lettre de Change Normalisée)
```
- Toute autre valeur → **anomalie bloquante**
- Valeur vide → **anomalie bloquante**

### Règle TAUX (extraction depuis "Lignes de facture/Taxes")
5. Regex de parsing : `r'(\d+(?:[.,]\d+)?)\s*%'` sur la chaîne Odoo.
   - `"20% 146"` → 0.20
   - `"10% 150"` → 0.10
   - `"20% S 140"` → 0.20 (les anciennes "S 140" ont été éliminées dans le dernier export, mais robustesse maintenue)
6. Si colonne vide ET `Taxe = 0` → **taux 0%** accepté (ligne exonérée/hors champ incluse au relevé)
7. Si colonne vide ET `Taxe ≠ 0` → **anomalie bloquante**

### Convention de stockage TAUX
8. La cellule J de l'EDI doit contenir un **float décimal** (ex: `0.20`), avec format de cellule `0%`.
   - Stockage interne Excel = 0.20
   - Affichage utilisateur = "20 %"
   - Export XML DGI = `<tx>0.2</tx>` (xmlDataType="float")
9. **NE PAS** stocker `20` ou `"20%"` → casse le XML map.

### Lignes Taxe = 0
10. **INCLURE** dans le relevé avec taux 0%. (Cas typiques : assurances, opérations hors champ.)

### Format ICE
11. **String 15 caractères**, zéros à gauche préservés (ex: `"003724512000016"`).
12. La cellule I doit avoir le format `@` (texte) pour empêcher Excel de retirer les zéros.

### Cohérence comptable
13. Vérifier `HT + Taxe ≈ Total` avec tolérance **0,05 DH**. Si écart > 0,05 → anomalie bloquante (la ligne sera dans le rapport).

### Tech stack
14. **Streamlit Community Cloud** (déploiement gratuit, URL publique).
15. **openpyxl** avec `keep_vba=True` pour préserver le VBA et le XML Map natif Excel.
16. **pandas** pour la lecture du fichier Odoo et le data wrangling.
17. **Multi-utilisateurs** : pas d'auth pour v1 (à ajouter si demande prod). Plusieurs comptables peuvent accéder à la même URL.

---

## 📂 STRUCTURE DE FICHIERS ATTENDUE

```
edi-tva-maroc/
├── app.py                                 # Application Streamlit principale (point d'entrée)
├── requirements.txt                       # streamlit, openpyxl, pandas
├── README.md                              # Guide de déploiement + utilisation
├── SPEC.md                                # CE fichier (à inclure dans le repo)
├── .gitignore                             # __pycache__, .venv, .DS_Store, *.pyc
├── templates/
│   └── EDI_MAROCAINE_XML_GENERATOR.xlsm   # Template embarqué (NE PAS MODIFIER)
├── samples/
│   └── Odoo_template.xlsx                 # Échantillon de test (80 lignes)
└── lib/                                   # (optionnel si on découpe le code)
    ├── mapper.py                          # Logique de mapping Odoo → EDI
    ├── validator.py                       # Règles de validation + rapport anomalies
    └── injector.py                        # Écriture dans .xlsm (préservation VBA)
```

**Note Claude Code** : commence par une approche monofichier (`app.py` seul) pour simplicité. Découpe en modules `lib/` seulement si `app.py` dépasse ~500 lignes.

---

## ⚙️ requirements.txt

```
streamlit>=1.30,<2.0
openpyxl>=3.1,<4.0
pandas>=2.0,<3.0
```

**Note** : pas besoin de `xlrd`, `python-docx`, ou autres. Le minimum suffit.

---

## 🧠 LOGIQUE MÉTIER DÉTAILLÉE

### Algorithme de validation + transformation

Pour chaque ligne du DataFrame Odoo (boucle `for idx, row in df.iterrows()`) :

**Étape 1 — Vérifier les champs obligatoires** :
```
Référence, Libellé, HT, Taxe, Total, Partenaire, ICE, IF,
Date de paiement, Date de facturation, Méthode de paiement
```
Si l'un est vide → ajouter à `anomalies[]` et `continue`.

**Étape 2 — Vérifier la méthode de paiement** :
- Doit être dans `{ESPECES, CHEQUE, PRELEVEMENT, VIREMENT, LCN}` (case-insensitive)
- Sinon → anomalie

**Étape 3 — Extraire le TAUX** :
- Si colonne "Lignes de facture/Taxes" non vide → `re.search(r'(\d+(?:[.,]\d+)?)\s*%', s)` → division par 100
- Si vide ET Taxe = 0 → taux = 0.0
- Sinon → anomalie

**Étape 4 — Vérifier cohérence comptable** :
- `abs((HT + Taxe) - Total) > 0.05` → anomalie

**Étape 5 — Conversions de type** :
- IF en `int` (ex: `"76146183"` → `76146183`) — anomalie si non convertible
- ICE en `str` (préserve zéros)
- Dates en `datetime.datetime` natif Python (pas pandas.Timestamp pour la cellule openpyxl)

**Étape 6 — Construire la ligne EDI** :
- OR = position dans le tableau valide (1..N)
- Tous les autres champs mappés selon la table ci-dessus

### Pseudocode validation
```python
def validate_and_transform(df_odoo) -> tuple[pd.DataFrame, list[dict]]:
    anomalies = []
    valid_rows = []
    for idx, row in df_odoo.iterrows():
        line_num = idx + 2  # +2 car ligne Excel = index pandas + header
        errors = []
        # ... vérifications ...
        if errors:
            anomalies.append({
                "Ligne Excel": line_num,
                "Référence": row.get("Référence", "(vide)"),
                "Partenaire": row.get("Partenaire", "(vide)"),
                "Erreurs": " ; ".join(errors),
            })
            continue
        # ... construction ligne EDI ...
        valid_rows.append({...})
    return pd.DataFrame(valid_rows), anomalies
```

### Pseudocode injection dans .xlsm

```python
def build_edi_xlsm(df_edi, header: dict, template_bytes: bytes) -> bytes:
    wb = load_workbook(BytesIO(template_bytes), keep_vba=True)
    ws = wb["EDI"]

    # 1. Remplir le header (cellules C2 à C6)
    ws["C2"] = header["raison_sociale"]
    ws["C3"] = header["if"]
    ws["C4"] = int(header["annee"])
    ws["C5"] = int(header["periode"])
    ws["C6"] = int(header["regime"])

    # 2. Effacer les anciennes données du tableau (rows 9 to current max)
    table = ws.tables["Tableau5"]
    max_existing_row = int(table.ref.split(":")[1][1:])
    for r in range(9, max_existing_row + 1):
        for c in range(1, 14):  # cols A à M
            ws.cell(row=r, column=c).value = None

    # 3. Écrire les nouvelles lignes EDI
    for i, edi_row in enumerate(df_edi.itertuples(index=False)):
        r = 9 + i
        ws.cell(r, 1, int(edi_row.OR))
        ws.cell(r, 2, str(edi_row.FACT_NUM))
        # ... etc pour les 13 colonnes ...
        ws.cell(r, 10, float(edi_row.TAUX))
        ws.cell(r, 10).number_format = "0%"        # affichage 20 %
        ws.cell(r, 9).number_format = "@"          # ICE en texte
        ws.cell(r, 12).number_format = "yyyy-mm-dd"  # DATE_PAIE
        ws.cell(r, 13).number_format = "yyyy-mm-dd"  # DATE_FAC

    # 4. Ajuster la plage du tableau
    n = len(df_edi)
    last_row = 8 + n + 1  # header (8) + n data rows + 1 totals row
    table.ref = f"A8:M{last_row}"
    if table.autoFilter:
        table.autoFilter.ref = f"A8:M{last_row - 1}"

    # 5. Retourner les bytes
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
```

---

## ⚠️ PIÈGES TECHNIQUES À ÉVITER

### Préservation VBA + XML Map (CRITIQUE)
- **Toujours** charger avec `load_workbook(path, keep_vba=True)`
- **Toujours** sauvegarder via `wb.save()` (openpyxl conservera VBA et XML Map automatiquement)
- **Test obligatoire** après génération : décompresser le `.xlsm` (c'est un ZIP) et vérifier que ces fichiers existent encore :
  - `xl/vbaProject.bin` (VBA)
  - `xl/xmlMaps.xml` (XML Map DGI)
  - `xl/tables/table1.xml` (définition Tableau5)

```bash
# Test de validation après génération
unzip -l output.xlsm | grep -E "vbaProject|xmlMaps|tables"
# Doit afficher les 3 fichiers
```

### Formules calculées dans la Table
- Le Tableau5 a `calculatedColumnFormula` sur les colonnes D (M_HT) et E (TVA) :
  - D: `=Tableau5[[#This Row],[M_TTC]]/(1+Tableau5[[#This Row],[TAUX]])`
  - E: `=Tableau5[[#This Row],[M_HT]]*Tableau5[[#This Row],[TAUX]]`
- **Quand on écrit une valeur dans la cellule, openpyxl ne supprime pas la formule calculée de la table**.
- Comportement Excel : à l'ouverture, Excel respecte la valeur cellule explicite, mais peut afficher un coin vert "valeur inconsistante avec la formule calculée".
- **Acceptable** : valeur correcte, juste un avertissement visuel. À documenter dans le README.

### Range de la table
- Le tableau d'origine fait `A8:M17` (8 lignes data + 1 totals)
- Si on écrit 80 lignes data, il faut : `A8:M89` (1 header + 80 data + 1 totals)
- L'autoFilter doit faire `A8:M88` (exclut la ligne totaux)
- **Si on ne met pas à jour la range** → certaines lignes data ne seront pas dans le XML !

### Type pandas → openpyxl
- `pd.Timestamp` → convertir en `datetime.datetime` natif via `.to_pydatetime()` avant `ws.cell().value = ...`
- `pd.NA`, `np.nan`, `pd.NaT` → convertir en `None` (Excel les acceptera comme cellule vide)

### Encodage Méthode de paiement
- Faire `.strip().upper()` avant comparaison dans le dict de lookup
- Gérer les espaces ou caractères invisibles ("VIREMENT " avec trailing space)

### Charges du template
- Charger les bytes du template UNE FOIS par génération via `Path.read_bytes()` puis passer à `BytesIO`
- Ne **JAMAIS** écrire sur le fichier template original — toujours travailler en mémoire (BytesIO) et retourner des bytes

---

## 🎨 UI STREAMLIT — STRUCTURE ATTENDUE

### Layout
```
📑 EDI TVA Maroc — Générateur de fichier de déclaration
Cabinet MSL-iTECH · Mapping Odoo → Template EDI .xlsm

1️⃣  Paramètres de la déclaration
   [Raison sociale]    [IF]
   [Année ▼]    [Période ▼]    [Régime ⦿ Débit ○ Encaissement]
   [✅ Valider paramètres]

2️⃣  Upload export Odoo (.xlsx)
   [Glisse-dépose ici]
   ✅ 80 lignes chargées
   ▶ Aperçu des 10 premières lignes (expander)

3️⃣  Validation et contrôle
   📊 80 lignes en entrée · 65 valides · 15 anomalies
   ⚠️ Génération bloquée — 15 ligne(s) en anomalie :
   [Tableau des anomalies avec colonnes : Ligne Excel | Référence | Partenaire | Erreurs]

4️⃣  Génération du fichier EDI .xlsm
   [🚀 Générer] (désactivé si anomalies)
   ⬇️ EDI_TVA_AitOukhaliTravaux_2026_M05_20260516_1432.xlsm
```

### Composants Streamlit recommandés
- `st.form` pour le bloc paramètres (évite les reruns intempestifs)
- `st.file_uploader` pour l'export Odoo
- `st.metric` pour le résumé (in / valides / anomalies)
- `st.dataframe` pour le tableau des anomalies
- `st.download_button` pour le téléchargement final
- `st.session_state` pour conserver le header entre interactions

### Convention de nommage du fichier généré
```
EDI_TVA_{raison_sociale_safe}_{annee}_M{periode:02d}_{timestamp}.xlsm

ex: EDI_TVA_AitOukhaliTravaux_2026_M05_20260516_1432.xlsm
```
Où `raison_sociale_safe` = `re.sub(r'[^A-Za-z0-9]', '_', raison_sociale)[:30]`

---

## 🧪 TESTS DE VALIDATION

### Sur le sample fourni (`samples/Odoo_template.xlsx`, 80 lignes)
Données attendues après analyse :
- **Anomalies attendues** : ~30-40 lignes (10 sans méthode paiement + 15 sans date paiement + 9 sans ICE + 20 sans IF, avec recouvrements)
- **Lignes valides** : ~40-50
- **Taux extraits** : 20% (68 lignes), 10% (11 lignes)

### Tests à automatiser
1. Le fichier généré contient bien `xl/vbaProject.bin` (VBA préservé)
2. Le fichier généré contient bien `xl/xmlMaps.xml` (XML Map préservée)
3. La range de Tableau5 a bien été mise à jour selon le nombre de lignes
4. Les dates sont au format `datetime`, pas string
5. Le TAUX est stocké en float (0.20, 0.10) avec format `0%`
6. L'ICE est en string 15 chars avec zéros préservés
7. Le fichier généré s'ouvre dans Excel sans erreur (à valider manuellement sur Windows)

### Test manuel critique (côté Windows)
1. Ouvrir le `.xlsm` généré dans Excel Windows
2. Activer les macros si demandé
3. Cliquer sur le bouton de la macro/XML map → vérifier que le XML s'export correctement
4. Ouvrir le XML dans un éditeur → vérifier la structure DGI :
   ```xml
   <DeclarationReleveDeduction>
     <releveDeductions>
       <rd>
         <ord>1</ord>
         <num>2605069</num>
         <des>GASOIL10PPM</des>
         <mht>435000</mht>
         <tva>43500</tva>
         <ttc>478500</ttc>
         <refF>
           <if>76146183</if>
           <nom>STE KAMED Transport</nom>
           <ice>000076447000091</ice>
         </refF>
         <tx>0.10</tx>
         <mp><id>4</id></mp>
         <dpai>2026-05-15</dpai>
         <dfac>2026-05-15</dfac>
       </rd>
       <!-- ... autres lignes ... -->
     </releveDeductions>
   </DeclarationReleveDeduction>
   ```

---

## 🚢 DÉPLOIEMENT STREAMLIT COMMUNITY CLOUD

### Pré-requis
- Compte GitHub avec accès au repo
- Compte Streamlit Cloud (gratuit) : https://share.streamlit.io
- Repo doit être **public** (ou compte payant pour repo privé)

### Étapes
1. **Initialiser git** :
   ```bash
   cd edi-tva-maroc/
   git init
   git add .
   git commit -m "v1: EDI TVA Maroc generator initial commit"
   ```

2. **Créer le repo GitHub** (via `gh` CLI ou interface web) :
   ```bash
   gh repo create msl-itech/edi-tva-maroc --public --source=. --remote=origin --push
   ```

3. **Déployer sur Streamlit Cloud** :
   - Aller sur https://share.streamlit.io
   - "New app" → choisir le repo `msl-itech/edi-tva-maroc`
   - Branch : `main`
   - Main file path : `app.py`
   - Click "Deploy"
   - URL générée : `https://msl-itech-edi-tva-maroc.streamlit.app`

4. **Tester en ligne** :
   - Uploader un export Odoo
   - Vérifier que le `.xlsm` se télécharge et fonctionne dans Excel

### Considérations sécurité (à discuter post-v1)
- Pas d'auth en v1 → URL publique = tout le monde peut utiliser l'app
- Si données sensibles → ajouter `streamlit-authenticator` + `.streamlit/secrets.toml`
- Alternative : déployer en privé sur un VPS interne MSL-iTECH

---

## 📦 EXTRAIT DE STRUCTURE DU TEMPLATE EDI (référence)

### Feuilles du `.xlsm`
- **EDI** : feuille principale (header + Tableau5 + XML Map)
- **CA** : feuille ventes (hors scope v1, ne pas toucher)
- **Feuil1** : lookup paiement (ne pas toucher)

### Cellules header EDI (à écrire)
- `C2` : Raison Sociale
- `C3` : Identifiant Fiscal
- `C4` : Année (int, validation list 2016-2030)
- `C5` : Période (int, validation list 1-12)
- `C6` : Régime (int, 1=Encais, 2=Débit)

### Tableau5 (à écrire)
- Range initial : `A8:M17` (1 header + 8 data + 1 totals)
- Header (row 8) : `OR | FACT_NUM | DESIGNATION | M_HT | TVA | M_TTC | IF | LIB_FRSS | ICE_FRS | TAUX | ID_PAIE | DATE_PAIE | DATE_FAC`
- Première ligne data : row 9
- À mettre à jour selon le nombre de lignes injectées

### XML Map (NE PAS TOUCHER, mais bon à savoir)
Le XML Map est dans `xl/xmlMaps.xml` et lie chaque colonne de Tableau5 à un xpath du schéma DGI :
```
A (OR)         → /DeclarationReleveDeduction/releveDeductions/rd/ord
B (FACT_NUM)   → /DeclarationReleveDeduction/releveDeductions/rd/num
C (DESIGNATION)→ /DeclarationReleveDeduction/releveDeductions/rd/des
D (M_HT)       → /DeclarationReleveDeduction/releveDeductions/rd/mht
E (TVA)        → /DeclarationReleveDeduction/releveDeductions/rd/tva
F (M_TTC)      → /DeclarationReleveDeduction/releveDeductions/rd/ttc
G (IF)         → /DeclarationReleveDeduction/releveDeductions/rd/refF/if
H (LIB_FRSS)   → /DeclarationReleveDeduction/releveDeductions/rd/refF/nom
I (ICE_FRS)    → /DeclarationReleveDeduction/releveDeductions/rd/refF/ice
J (TAUX)       → /DeclarationReleveDeduction/releveDeductions/rd/tx
K (ID_PAIE)    → /DeclarationReleveDeduction/releveDeductions/rd/mp/id
L (DATE_PAIE)  → /DeclarationReleveDeduction/releveDeductions/rd/dpai
M (DATE_FAC)   → /DeclarationReleveDeduction/releveDeductions/rd/dfac
```

---

## 🔧 STRUCTURE DE L'EXPORT ODOO (référence)

12 colonnes attendues (ordre strict) :

| # | Header Odoo (FR) | Type | Required | Notes |
|---|------------------|------|----------|-------|
| A | Référence | str | ✅ | Numéro facture fournisseur |
| B | Libellé | str | ✅ | Description |
| C | Montant hors taxes | float | ✅ | HT en DH |
| D | Taxe | float | ✅ | Montant TVA en DH (peut être 0) |
| E | Total | float | ✅ | TTC = HT + Taxe |
| F | Lignes de facture/Taxes | str | ⚠️ | Ex: "20% 146" — vide accepté si Taxe=0 |
| G | IF | str→int | ✅ | Identifiant fiscal fournisseur |
| H | Partenaire | str | ✅ | Nom du fournisseur |
| I | ICE | str | ✅ | 15 chars avec zéros à gauche |
| J | Méthode de paiement | str | ✅ | ESPECES/CHEQUE/PRELEVEMENT/VIREMENT/LCN |
| K | Date de paiement | date | ✅ | datetime |
| L | Date de facturation | date | ✅ | datetime |

---

## 📝 CHANGELOG / DÉCISIONS PRISES

| Date | Décision | Raison |
|------|----------|--------|
| 2026-05-16 | v1 = EDI seul, pas de CA | Périmètre raisonnable pour MVP |
| 2026-05-16 | Validation stricte avec rapport | Préfère bloquer que générer un XML rejeté DGI |
| 2026-05-16 | TAUX parsé depuis colonne G Odoo | Nouvelle colonne fournie par Odoo, plus fiable que ratio Taxe/HT |
| 2026-05-16 | HT, TVA, TTC écrits en valeurs directes | Plus robuste que formules (anciennes anomalies "20% S 140") |
| 2026-05-16 | TAUX stocké en 0.20 + format 0% | Convention DGI + template d'origine |
| 2026-05-16 | ICE en string 15 chars | Préserver zéros à gauche, conforme spec DGI |
| 2026-05-16 | Streamlit Cloud, multi-users sans auth en v1 | Simplicité, à durcir si besoin |

---

## 🆘 EN CAS DE PROBLÈME

### Bugs courants à anticiper
1. **VBA disparu après génération** → vérifier `keep_vba=True` à `load_workbook` ET à `save`
2. **XML Map cassée** → ne pas modifier les `xmlColumnPr` des colonnes Tableau5, ne pas supprimer la table
3. **Excel affiche "fichier corrompu"** → souvent un format de cellule invalide ou un range tableau désaligné
4. **Date stockée comme nombre** → forcer le `number_format = "yyyy-mm-dd"` sur la cellule
5. **ICE perd ses zéros** → mettre `number_format = "@"` (texte) AVANT d'écrire la valeur
6. **TAUX affiche 0,2 au lieu de 20%** → vérifier `number_format = "0%"` sur la cellule

### Logs et debugging recommandés
- Logger chaque étape clé avec `st.write()` en mode dev
- Stocker le fichier généré temporairement pour inspection : `out.getvalue()` peut être inspecté via `BytesIO` + `unzip`
- Pour tester en local : `streamlit run app.py` (lance sur localhost:8501)

---

## ✅ DEFINITION OF DONE (v1)

- [ ] L'app se lance sans erreur avec `streamlit run app.py`
- [ ] L'utilisateur peut saisir les 5 champs header
- [ ] L'utilisateur peut uploader un .xlsx Odoo
- [ ] Le rapport d'anomalies affiche ligne par ligne les erreurs (si présentes)
- [ ] Si pas d'anomalies, le bouton "Générer" produit un .xlsm téléchargeable
- [ ] Le .xlsm généré conserve le VBA (vérifié via unzip)
- [ ] Le .xlsm généré conserve le XML Map (vérifié via unzip)
- [ ] Ouvert dans Excel Windows, le .xlsm permet de générer un XML DGI valide
- [ ] Le repo est sur GitHub, public
- [ ] L'app est déployée sur Streamlit Cloud avec une URL fonctionnelle
- [ ] Le README.md explique : installation locale, déploiement, utilisation

---

**Fin de SPEC.md** · Toute modification de ce document doit être validée par El Houssine BOUHMAIDA (MSL-iTECH).
