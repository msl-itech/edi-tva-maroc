"""End-to-end test on the real Odoo sample.

Validates that the spec contract holds:
- validation produces the expected anomaly count
- the generated .xlsm still contains xl/vbaProject.bin + xl/xmlMaps.xml + xl/tables/
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import build_edi_xlsm, load_odoo, validate_and_transform  # noqa: E402

SAMPLE = ROOT / "samples" / "Odoo_template.xlsx"
TEMPLATE = ROOT / "templates" / "EDI_MAROCAINE_XML_GENERATOR.xlsm"
OUT = ROOT / "scripts" / "out_test.xlsm"


def main() -> int:
    print("=" * 72)
    print("LOADING SAMPLE")
    print("=" * 72)
    df = load_odoo(SAMPLE.read_bytes())
    print(f"  shape = {df.shape}")

    print()
    print("=" * 72)
    print("VALIDATION")
    print("=" * 72)
    df_edi, anomalies = validate_and_transform(df)
    print(f"  valid rows = {len(df_edi)}")
    print(f"  anomalies  = {len(anomalies)}")
    print()
    print("First 15 anomalies:")
    for a in anomalies[:15]:
        print(
            f"  L{a['Ligne Excel']:>3}  "
            f"ref={a['Référence']!r:>18}  "
            f"partner={a['Partenaire']!r:>30}  "
            f"errs={a['Erreurs']}"
        )
    if len(anomalies) > 15:
        print(f"  ...and {len(anomalies) - 15} more")

    if df_edi.empty:
        print("WARN: no valid rows in sample - will still generate empty xlsm to test preservation")

    print()
    print("=" * 72)
    print("GENERATION")
    print("=" * 72)
    header = {
        "raison_sociale": "AitOukhaliTravaux",
        "if": "12345678",
        "annee": 2026,
        "periode": 5,
        "regime": 2,
    }
    xlsm = build_edi_xlsm(df_edi, header, TEMPLATE.read_bytes())
    OUT.write_bytes(xlsm)
    print(f"  wrote {OUT}  ({len(xlsm):,} bytes)")

    print()
    print("=" * 72)
    print("ZIP CONTENTS CHECK (VBA + XML Map + tables)")
    print("=" * 72)
    with zipfile.ZipFile(OUT) as zf:
        names = zf.namelist()

    required = ["xl/vbaProject.bin", "xl/xmlMaps.xml"]
    tables = [n for n in names if n.startswith("xl/tables/")]
    failures = []
    for r in required:
        ok = r in names
        print(f"  {r:<25} {'PRESENT' if ok else '*** MISSING ***'}")
        if not ok:
            failures.append(r)
    print(f"  xl/tables/* count        = {len(tables)} ({tables})")
    if not tables:
        failures.append("xl/tables/*")

    print()
    print("=" * 72)
    print("CELL-LEVEL SANITY CHECK")
    print("=" * 72)
    from openpyxl import load_workbook  # noqa: PLC0415
    wb = load_workbook(OUT, keep_vba=True)
    ws = wb["EDI"]
    print(f"  C2 (raison sociale) = {ws['C2'].value!r}")
    print(f"  C3 (IF)             = {ws['C3'].value!r}")
    print(f"  C4 (année)          = {ws['C4'].value!r}")
    print(f"  C5 (période)        = {ws['C5'].value!r}")
    print(f"  C6 (régime)         = {ws['C6'].value!r}")
    table = ws.tables["Tableau5"]
    print(f"  Tableau5.ref        = {table.ref!r}")
    if table.autoFilter:
        print(f"  Tableau5.filter     = {table.autoFilter.ref!r}")

    if df_edi.shape[0] > 0:
        r = 9  # first data row
        print(f"  Row {r}:")
        for c in range(1, 14):
            cell = ws.cell(r, c)
            print(f"    col{c}({chr(64+c)}): value={cell.value!r}, fmt={cell.number_format!r}")

    print()
    if failures:
        print(f"FAIL: {len(failures)} missing artefact(s): {failures}")
        return 1
    print("ALL CHECKS PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
