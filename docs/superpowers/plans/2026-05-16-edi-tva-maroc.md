# EDI TVA Maroc — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Streamlit app that maps an Odoo vendor-invoice export to the pre-existing `EDI_MAROCAINE_XML_GENERATOR.xlsm` template, preserving its VBA + XML Map so the user can generate the SIMPL-TVA DGI XML from Excel.

**Architecture:** Single-file `app.py` (per SPEC note: split only if > 500 lines). Loads template bytes once, runs `validate_and_transform()` on the Odoo DataFrame, blocks on anomalies, then `build_edi_xlsm()` writes via `openpyxl(keep_vba=True)` into a `BytesIO`. The template is never modified on disk.

**Tech Stack:** Python 3.10+, Streamlit 1.30+, openpyxl 3.1+ (`keep_vba=True`), pandas 2.0+.

**Locked decisions** — see `SPEC.md` §"DÉCISIONS VERROUILLÉES". Do not deviate. ICE strategy confirmed in chat: load with `dtype=str` to preserve original, then `zfill(15)` safety net.

---

## File Structure

- Create: `app.py` — Streamlit entry point + all logic (header form, uploader, validation, anomaly report, .xlsm generation, download)
- Create: `requirements.txt` — pinned deps per SPEC
- Modify: `README.md` — replace stub with deployment + usage guide
- Already exists: `templates/EDI_MAROCAINE_XML_GENERATOR.xlsm` (do not modify)
- Already exists: `samples/Odoo_template.xlsx` (do not modify)
- Created (kept): `scripts/inspect_files.py` — one-shot inspector for spec validation; safe to ship

---

## Task 1: requirements.txt

**Files:**
- Create: `requirements.txt`

- [ ] **Step 1: Write requirements**

```txt
streamlit>=1.30,<2.0
openpyxl>=3.1,<4.0
pandas>=2.0,<3.0
```

---

## Task 2: Skeleton app.py — constants & imports

**Files:**
- Create: `app.py`

- [ ] **Step 1: Write imports + constants**

```python
"""EDI TVA Maroc — Streamlit generator (MSL-iTECH)."""
from __future__ import annotations

import datetime as _dt
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "templates" / "EDI_MAROCAINE_XML_GENERATOR.xlsm"

PAYMENT_METHOD_MAP = {
    "ESPECES": 1,
    "CHEQUE": 2,
    "PRELEVEMENT": 3,
    "VIREMENT": 4,
    "LCN": 5,
}

ODOO_REQUIRED_COLS = [
    "Référence", "Libellé", "Montant hors taxes", "Taxe", "Total",
    "Lignes de facture/Taxes", "IF", "Partenaire", "ICE",
    "Méthode de paiement", "Date de paiement", "Date de facturation",
]

TAUX_REGEX = re.compile(r"(\d+(?:[.,]\d+)?)\s*%")
TOLERANCE_DH = 0.05
TABLE_NAME = "Tableau5"
DATA_START_ROW = 9   # template's first data row
HEADER_ROW = 8
```

---

## Task 3: Odoo loader (preserves ICE/IF/Référence as strings)

**Files:**
- Modify: `app.py` (append)

- [ ] **Step 1: Define `load_odoo` reader**

```python
def load_odoo(uploaded_bytes: bytes) -> pd.DataFrame:
    """Read Odoo xlsx forcing str dtype on ICE/IF/Référence to keep leading zeros."""
    df = pd.read_excel(
        BytesIO(uploaded_bytes),
        dtype={"Référence": str, "IF": str, "ICE": str},
    )
    return df
```

---

## Task 4: Helpers — payment lookup, taux parse, normalisations

**Files:**
- Modify: `app.py` (append)

- [ ] **Step 1: Add helpers**

```python
def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return False


def _clean_numeric_str(v) -> str:
    """For columns we forced to str: drop trailing '.0' that pandas adds when
    the underlying Excel cell was numeric."""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def normalize_payment(v) -> str | None:
    if _is_blank(v):
        return None
    return str(v).strip().upper()


def parse_taux(raw, taxe: float) -> tuple[float | None, str | None]:
    """Returns (taux, error). taux is a float in [0,1]."""
    if _is_blank(raw):
        if abs(taxe) < 1e-9:
            return 0.0, None
        return None, "Colonne 'Lignes de facture/Taxes' vide alors que Taxe ≠ 0"
    m = TAUX_REGEX.search(str(raw))
    if not m:
        return None, f"Taux non détectable dans '{raw}'"
    val = float(m.group(1).replace(",", ".")) / 100.0
    return val, None


def normalize_ice(v) -> tuple[str | None, str | None]:
    if _is_blank(v):
        return None, "ICE manquant"
    s = _clean_numeric_str(v)
    if not s.isdigit():
        return None, f"ICE non numérique: '{s}'"
    if len(s) > 15:
        return None, f"ICE > 15 chiffres ({len(s)}): '{s}'"
    return s.zfill(15), None


def normalize_if(v) -> tuple[int | None, str | None]:
    if _is_blank(v):
        return None, "IF manquant"
    s = _clean_numeric_str(v)
    try:
        return int(s), None
    except ValueError:
        return None, f"IF non convertible en entier: '{s}'"


def to_pydate(v) -> _dt.datetime | None:
    if _is_blank(v):
        return None
    if isinstance(v, _dt.datetime):
        return v
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    try:
        ts = pd.to_datetime(v, errors="raise")
        return ts.to_pydatetime() if pd.notna(ts) else None
    except (ValueError, TypeError):
        return None
```

---

## Task 5: Core — `validate_and_transform`

**Files:**
- Modify: `app.py` (append)

- [ ] **Step 1: Implement validation/transformation loop**

```python
def validate_and_transform(df: pd.DataFrame) -> tuple[pd.DataFrame, list[dict]]:
    missing_cols = [c for c in ODOO_REQUIRED_COLS if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Colonnes Odoo manquantes: {missing_cols}")

    anomalies: list[dict] = []
    valid_rows: list[dict] = []

    for idx, row in df.iterrows():
        line_num = int(idx) + 2  # +2: header is row 1 in Excel
        errors: list[str] = []

        ref = row["Référence"]
        if _is_blank(ref):
            errors.append("Référence manquante")
        libelle = row["Libellé"]
        if _is_blank(libelle):
            errors.append("Libellé manquant")
        partner = row["Partenaire"]
        if _is_blank(partner):
            errors.append("Partenaire manquant")

        try:
            ht = float(row["Montant hors taxes"])
            if pd.isna(ht):
                raise ValueError
        except (ValueError, TypeError):
            ht = None
            errors.append("HT manquant ou non numérique")
        try:
            taxe = float(row["Taxe"])
            if pd.isna(taxe):
                raise ValueError
        except (ValueError, TypeError):
            taxe = None
            errors.append("Taxe manquante ou non numérique")
        try:
            total = float(row["Total"])
            if pd.isna(total):
                raise ValueError
        except (ValueError, TypeError):
            total = None
            errors.append("Total manquant ou non numérique")

        if_val, if_err = normalize_if(row["IF"])
        if if_err:
            errors.append(if_err)
        ice_val, ice_err = normalize_ice(row["ICE"])
        if ice_err:
            errors.append(ice_err)

        payment_raw = row["Méthode de paiement"]
        payment_norm = normalize_payment(payment_raw)
        if payment_norm is None:
            errors.append("Méthode de paiement manquante")
            payment_id = None
        elif payment_norm not in PAYMENT_METHOD_MAP:
            errors.append(f"Méthode de paiement inconnue: '{payment_raw}'")
            payment_id = None
        else:
            payment_id = PAYMENT_METHOD_MAP[payment_norm]

        dpai = to_pydate(row["Date de paiement"])
        if dpai is None:
            errors.append("Date de paiement manquante")
        dfac = to_pydate(row["Date de facturation"])
        if dfac is None:
            errors.append("Date de facturation manquante")

        taux_err = None
        taux_val = None
        if taxe is not None:
            taux_val, taux_err = parse_taux(row["Lignes de facture/Taxes"], taxe)
            if taux_err:
                errors.append(taux_err)

        if (
            ht is not None
            and taxe is not None
            and total is not None
            and abs((ht + taxe) - total) > TOLERANCE_DH
        ):
            errors.append(
                f"Incohérence HT+TVA vs Total: {ht:.2f}+{taxe:.2f}={ht+taxe:.2f} ≠ {total:.2f}"
            )

        if errors:
            anomalies.append({
                "Ligne Excel": line_num,
                "Référence": "(vide)" if _is_blank(ref) else str(ref).strip(),
                "Partenaire": "(vide)" if _is_blank(partner) else str(partner).strip(),
                "Erreurs": " ; ".join(errors),
            })
            continue

        valid_rows.append({
            "OR": len(valid_rows) + 1,
            "FACT_NUM": str(ref).strip(),
            "DESIGNATION": str(libelle).strip(),
            "M_HT": ht,
            "TVA": taxe,
            "M_TTC": total,
            "IF": if_val,
            "LIB_FRSS": str(partner).strip(),
            "ICE_FRS": ice_val,
            "TAUX": taux_val,
            "ID_PAIE": payment_id,
            "DATE_PAIE": dpai,
            "DATE_FAC": dfac,
        })

    df_edi = pd.DataFrame(valid_rows)
    return df_edi, anomalies
```

---

## Task 6: Core — `build_edi_xlsm`

**Files:**
- Modify: `app.py` (append)

- [ ] **Step 1: Implement injector**

```python
def build_edi_xlsm(df_edi: pd.DataFrame, header: dict, template_bytes: bytes) -> bytes:
    wb = load_workbook(BytesIO(template_bytes), keep_vba=True)
    ws = wb["EDI"]

    # 1. Header cells
    ws["C2"] = header["raison_sociale"]
    ws["C3"] = header["if"]
    ws["C4"] = int(header["annee"])
    ws["C5"] = int(header["periode"])
    ws["C6"] = int(header["regime"])

    # 2. Clear previous data rows (rows 9..end of current table ref)
    table = ws.tables[TABLE_NAME]
    last_existing = int(table.ref.split(":")[1][1:])
    for r in range(DATA_START_ROW, last_existing + 1):
        for c in range(1, 14):
            ws.cell(row=r, column=c).value = None

    # 3. Write new data rows
    n = len(df_edi)
    for i, edi_row in enumerate(df_edi.itertuples(index=False)):
        r = DATA_START_ROW + i
        ws.cell(r, 1, int(edi_row.OR))
        ws.cell(r, 2, str(edi_row.FACT_NUM))
        ws.cell(r, 3, str(edi_row.DESIGNATION))
        ws.cell(r, 4, float(edi_row.M_HT))
        ws.cell(r, 5, float(edi_row.TVA))
        ws.cell(r, 6, float(edi_row.M_TTC))
        ws.cell(r, 7, int(edi_row.IF))
        ws.cell(r, 8, str(edi_row.LIB_FRSS))
        ice_cell = ws.cell(r, 9, str(edi_row.ICE_FRS))
        ice_cell.number_format = "@"
        taux_cell = ws.cell(r, 10, float(edi_row.TAUX))
        taux_cell.number_format = "0%"
        ws.cell(r, 11, int(edi_row.ID_PAIE))
        dpai_cell = ws.cell(r, 12, edi_row.DATE_PAIE)
        dpai_cell.number_format = "yyyy-mm-dd"
        dfac_cell = ws.cell(r, 13, edi_row.DATE_FAC)
        dfac_cell.number_format = "yyyy-mm-dd"

    # 4. Update table range (header + n data + 1 totals row)
    last_row = HEADER_ROW + n + 1
    table.ref = f"A{HEADER_ROW}:M{last_row}"
    if table.autoFilter is not None:
        table.autoFilter.ref = f"A{HEADER_ROW}:M{last_row - 1}"

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
```

---

## Task 7: Streamlit UI

**Files:**
- Modify: `app.py` (append)

- [ ] **Step 1: Implement UI**

```python
def _safe_name(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "_", s)[:30]


def main() -> None:
    st.set_page_config(page_title="EDI TVA Maroc", page_icon="📑", layout="wide")
    st.title("📑 EDI TVA Maroc — Générateur de fichier de déclaration")
    st.caption("Cabinet MSL-iTECH · Mapping Odoo → Template EDI .xlsm")

    if not TEMPLATE_PATH.exists():
        st.error(f"Template introuvable : {TEMPLATE_PATH}")
        st.stop()

    # --- 1. Paramètres déclaration ---
    st.subheader("1️⃣  Paramètres de la déclaration")
    with st.form("decl_form"):
        c1, c2 = st.columns(2)
        with c1:
            raison_sociale = st.text_input("Raison sociale", key="raison_sociale")
        with c2:
            if_decl = st.text_input("Identifiant fiscal (IF)", key="if_decl")
        c3, c4, c5 = st.columns(3)
        with c3:
            annee = st.selectbox("Année", list(range(2016, 2031)), index=10)
        with c4:
            periode = st.selectbox("Période (mois)", list(range(1, 13)), index=0)
        with c5:
            regime = st.radio("Régime", [2, 1], format_func=lambda v: "Débit" if v == 2 else "Encaissement", horizontal=True)
        submitted = st.form_submit_button("✅ Valider paramètres")
    if submitted:
        if not raison_sociale.strip() or not if_decl.strip():
            st.error("Raison sociale et IF sont obligatoires.")
        else:
            st.session_state["header"] = {
                "raison_sociale": raison_sociale.strip(),
                "if": if_decl.strip(),
                "annee": int(annee),
                "periode": int(periode),
                "regime": int(regime),
            }
            st.success("Paramètres validés.")

    header = st.session_state.get("header")
    if header is None:
        st.info("Saisissez et validez les paramètres pour continuer.")
        return

    # --- 2. Upload Odoo ---
    st.subheader("2️⃣  Upload export Odoo (.xlsx)")
    up = st.file_uploader("Glisse-dépose le fichier Odoo", type=["xlsx"])
    if up is None:
        return

    try:
        df_odoo = load_odoo(up.getvalue())
    except Exception as e:
        st.error(f"Lecture impossible : {e}")
        return
    st.success(f"{len(df_odoo)} ligne(s) chargée(s).")
    with st.expander("Aperçu des 10 premières lignes"):
        st.dataframe(df_odoo.head(10), use_container_width=True)

    # --- 3. Validation ---
    st.subheader("3️⃣  Validation et contrôle")
    try:
        df_edi, anomalies = validate_and_transform(df_odoo)
    except ValueError as e:
        st.error(str(e))
        return

    m1, m2, m3 = st.columns(3)
    m1.metric("Lignes en entrée", len(df_odoo))
    m2.metric("Lignes valides", len(df_edi))
    m3.metric("Anomalies", len(anomalies))

    if anomalies:
        st.warning(f"⚠️ Génération bloquée — {len(anomalies)} ligne(s) en anomalie :")
        st.dataframe(pd.DataFrame(anomalies), use_container_width=True)
        return

    # --- 4. Génération ---
    st.subheader("4️⃣  Génération du fichier EDI .xlsm")
    template_bytes = TEMPLATE_PATH.read_bytes()
    if st.button("🚀 Générer", type="primary"):
        with st.spinner("Génération du .xlsm…"):
            xlsm_bytes = build_edi_xlsm(df_edi, header, template_bytes)
        ts = _dt.datetime.now().strftime("%Y%m%d_%H%M")
        fname = f"EDI_TVA_{_safe_name(header['raison_sociale'])}_{header['annee']}_M{header['periode']:02d}_{ts}.xlsm"
        st.download_button(
            "⬇️ Télécharger le .xlsm",
            data=xlsm_bytes,
            file_name=fname,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )
        st.success(f"Fichier prêt : {fname}")


if __name__ == "__main__":
    main()
```

---

## Task 8: Test on Odoo sample (E2E in-script)

**Files:**
- Create: `scripts/test_sample.py` — one-shot integration test (kept in repo)

- [ ] **Step 1: Write the script**

```python
"""End-to-end test: validate sample + generate xlsm + verify VBA/XML Map."""
import zipfile
from pathlib import Path

import sys
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import load_odoo, validate_and_transform, build_edi_xlsm  # noqa: E402

SAMPLE = ROOT / "samples" / "Odoo_template.xlsx"
TEMPLATE = ROOT / "templates" / "EDI_MAROCAINE_XML_GENERATOR.xlsm"

print("Loading sample…")
df = load_odoo(SAMPLE.read_bytes())
print(f"  shape={df.shape}")

print("Validating…")
df_edi, anomalies = validate_and_transform(df)
print(f"  valid={len(df_edi)}  anomalies={len(anomalies)}")
for a in anomalies[:10]:
    print(f"  L{a['Ligne Excel']}: {a['Erreurs']}")
if len(anomalies) > 10:
    print(f"  ... and {len(anomalies) - 10} more")

print("Generating xlsm (using all-valid stub: skip anomalies)…")
header = {
    "raison_sociale": "AitOukhaliTravaux",
    "if": "12345678",
    "annee": 2026,
    "periode": 5,
    "regime": 2,
}
xlsm = build_edi_xlsm(df_edi, header, TEMPLATE.read_bytes())
out_path = ROOT / "scripts" / "out_test.xlsm"
out_path.write_bytes(xlsm)
print(f"  wrote {out_path} ({len(xlsm)} bytes)")

print("Checking ZIP contents (VBA + XML Map + tables)…")
with zipfile.ZipFile(out_path) as zf:
    names = zf.namelist()
required = ["xl/vbaProject.bin", "xl/xmlMaps.xml"]
tables = [n for n in names if n.startswith("xl/tables/")]
for r in required:
    print(f"  {r}: {'OK' if r in names else 'MISSING'}")
print(f"  tables: {tables}")
assert all(r in names for r in required), "VBA or XML Map missing!"
assert tables, "No table definitions!"
print("ALL CHECKS PASS")
```

- [ ] **Step 2: Run it**

```bash
python scripts/test_sample.py
```

Expected: prints anomaly count (~30-40 per spec), "ALL CHECKS PASS".

---

## Task 9: README

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Replace stub with full guide** — install local, run streamlit, déployer Streamlit Cloud, utiliser dans Excel Windows.

---

## Task 10: Git commit (only when user explicitly asks)

Per user CLAUDE.md: "Never run `git commit` unless I explicitly ask."

- [ ] **Step 1:** Hold off on `git commit` and `git push`. Show deployment instructions and wait for user signal.

---

## Self-Review

- ✅ Header `C2:C6` cells written
- ✅ All 13 EDI columns mapped (OR, FACT_NUM, DESIGNATION, M_HT, TVA, M_TTC, IF, LIB_FRSS, ICE_FRS, TAUX, ID_PAIE, DATE_PAIE, DATE_FAC)
- ✅ ICE: `dtype=str` + `_clean_numeric_str` + `zfill(15)` + `number_format="@"`
- ✅ IF: int conversion, anomaly if non-convertible
- ✅ TAUX: regex `(\d+(?:[.,]\d+)?)\s*%` /100, format `0%`, taux=0 accepted when Taxe=0
- ✅ Payment method lookup ESPECES/CHEQUE/PRELEVEMENT/VIREMENT/LCN, case+space tolerant
- ✅ Dates → `datetime` via `to_pydatetime()`, format `yyyy-mm-dd`
- ✅ Cohérence HT+TVA≈Total tolérance 0.05
- ✅ Strict blocking validation: anomalies → no generation
- ✅ Table range update + autoFilter
- ✅ VBA preservation via `keep_vba=True`
- ✅ Test script verifies xl/vbaProject.bin + xl/xmlMaps.xml + xl/tables/
- ✅ Filename convention `EDI_TVA_{safe}_{Y}_M{MM}_{ts}.xlsm`
- ✅ st.form for header, st.file_uploader, st.metric, st.dataframe, st.download_button
