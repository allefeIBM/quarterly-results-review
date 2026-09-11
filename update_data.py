#!/usr/bin/env python3
"""
update_data.py — Regenera o data.json a partir dos arquivos .xlsx da pasta.

Uso:
    python3 update_data.py

O script lê todos os .xlsx, extrai os KPIs e salva o data.json atualizado.
Em seguida, faz commit e push automático para o GitHub (opcional).
"""

import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).parent


# ── Helpers de parsing ────────────────────────────────────────────────────────

def find_header_row(rows):
    for i, row in enumerate(rows):
        r = [str(c).strip().lower() if c else "" for c in row]
        if "event name" in r and "brief" in r:
            return i
    return None


def col_index(headers, *candidates):
    for name in candidates:
        nl = name.lower()
        for i, h in enumerate(headers):
            if h and h.lower() == nl:
                return i
        for i, h in enumerate(headers):
            if h and h.lower().startswith(nl):
                return i
        for i, h in enumerate(headers):
            if h and nl in h.lower():
                return i
    return None


def col_index_exact(headers, name):
    nl = name.upper()
    for i, h in enumerate(headers):
        hc = h.upper().split("(")[0].strip() if h else ""
        if hc == nl:
            return i
    return None


def col_index_total(headers):
    for exact in ("totals", "total", "total sql"):
        for i, h in enumerate(headers):
            if h and h.lower() == exact:
                return i
    for i, h in enumerate(headers):
        hl = h.lower() if h else ""
        if hl.startswith("total") and hl != "total sqlr" and "sqlr" not in hl:
            return i
    return None


# ── Extração principal ────────────────────────────────────────────────────────

def extract_all_events():
    events = {}
    xlsx_files = sorted(
        p.name for p in BASE_DIR.glob("*.xlsx") if not p.name.startswith("~$")
    )

    if not xlsx_files:
        print("⚠  Nenhum arquivo .xlsx encontrado na pasta.")
        sys.exit(1)

    print(f"📂 {len(xlsx_files)} arquivo(s) encontrado(s):\n")

    for filename in xlsx_files:
        path = BASE_DIR / filename
        print(f"  → {filename}")
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:
            print(f"     ⚠  Erro ao abrir: {e}")
            continue

        sheet_count = 0
        row_count = 0

        for sheet_name in wb.sheetnames:
            if "option" in sheet_name.lower() or sheet_name.lower() == "sheet1":
                continue

            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            hi = find_header_row(rows)
            if hi is None:
                continue

            headers = [str(c).strip() if c else "" for c in rows[hi]]
            icol = {
                "name":  col_index(headers, "event name"),
                "brief": col_index(headers, "brief"),
                "ut15":  col_index(headers, "ut 15"),
                "ut17":  col_index(headers, "ut 17"),
                "cq":    col_index_exact(headers, "cq"),
                "nq":    col_index_exact(headers, "nq"),
                "nq1":   col_index_exact(headers, "nq+1"),
                "nq2":   col_index_exact(headers, "nq+2"),
                "total": col_index_total(headers),
            }

            if icol["name"] is None or icol["brief"] is None:
                continue

            quarter = next(
                (q for q in ["Q2", "Q3", "Q4", "Q1"] if q in sheet_name.upper()), "Q2"
            )
            cur_name = cur_brief = None

            for row in rows[hi + 1:]:
                nv = row[icol["name"]]  if icol["name"]  is not None else None
                bv = row[icol["brief"]] if icol["brief"] is not None else None
                if nv and str(nv).strip():
                    cur_name = str(nv).strip()
                if bv and str(bv).strip():
                    cur_brief = str(bv).strip()
                if not cur_brief:
                    continue

                def gn(k):
                    if icol.get(k) is None:
                        return None
                    v = row[icol[k]]
                    try:
                        return float(v) if v is not None else None
                    except (TypeError, ValueError):
                        return None

                cq, nq, nq1, nq2, tot = gn("cq"), gn("nq"), gn("nq1"), gn("nq2"), gn("total")
                if all(x is None for x in [cq, nq, nq1, nq2, tot]):
                    continue

                ut15 = str(row[icol["ut15"]]).strip() if icol.get("ut15") is not None and row[icol["ut15"]] else ""
                ut17 = str(row[icol["ut17"]]).strip() if icol.get("ut17") is not None and row[icol["ut17"]] else ""

                kpi_row = {
                    "event_name": cur_name or "",
                    "brief": cur_brief,
                    "ut15": ut15,
                    "ut17": ut17,
                    "cq":   cq,
                    "nq":   nq,
                    "nq1":  nq1,
                    "nq2":  nq2,
                    "total": tot,
                    "quarter": quarter,
                    "source_file": filename,
                }

                key = cur_brief.upper()
                if key not in events:
                    events[key] = {
                        "event_name": cur_name or "",
                        "brief": cur_brief,
                        "kpi_rows": [],
                    }
                events[key]["kpi_rows"].append(kpi_row)
                if cur_name:
                    events[key]["event_name"] = cur_name

                row_count += 1
            sheet_count += 1

        print(f"     ✓ {sheet_count} aba(s) · {row_count} KPI row(s)")

    return events


# ── Escrita do data.json ──────────────────────────────────────────────────────

def write_data_json(events):
    output = {
        "generated": datetime.now().isoformat()[:16],
        "events": events,
    }
    out_path = BASE_DIR / "data.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))
    size_kb = out_path.stat().st_size / 1024
    print(f"\n✅ data.json gerado — {len(events)} eventos · {size_kb:.1f} KB")
    return out_path


# ── Git commit + push ─────────────────────────────────────────────────────────

def git_push(out_path):
    print()
    answer = input("Deseja fazer commit e push para o GitHub agora? [s/N] ").strip().lower()
    if answer not in ("s", "sim", "y", "yes"):
        print("ℹ  Push ignorado. Rode 'git add data.json && git commit -m \"data: update\" && git push' manualmente.")
        return

    try:
        subprocess.run(["git", "add", str(out_path)], check=True, cwd=BASE_DIR)
        msg = f"data: update KPI projections — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        subprocess.run(["git", "commit", "-m", msg], check=True, cwd=BASE_DIR)
        subprocess.run(["git", "push"], check=True, cwd=BASE_DIR)
        print("🚀 Push concluído! A página será atualizada em ~1 minuto.")
    except subprocess.CalledProcessError as e:
        print(f"⚠  Erro no git: {e}")
        print("   Rode manualmente: git add data.json && git commit -m 'data: update' && git push")


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 60)
    print("  Quarterly Results Review — Atualizar data.json")
    print("=" * 60)
    print()

    events = extract_all_events()
    out_path = write_data_json(events)
    git_push(out_path)
