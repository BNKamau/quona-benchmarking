"""
push_comp_updates_to_turso.py
Pushes comp updates to Turso (quona_exit_comps):
  - Fix Fawry gross margin (7.4% -> 44.0%, estimated)
  - Delete Zettle from exit_comps and portfolio_comp_mapping
  - Insert iKhokha into exit_comps
  - Map iKhokha and DPO Group to Yoco in portfolio_comp_mapping
Schema matches local SQLite: comp_id TEXT PK, no id integer column.
"""
import requests

URL   = "https://quona-exit-comps-bnkamau.aws-eu-west-1.turso.io"
TOKEN = (
    "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9"
    ".eyJhIjoicnciLCJpYXQiOjE3ODA1NzU1OTEsImlkIjoiMDE5ZTkyOGQtZGEwMS03NTJlLWE4"
    "NjktZGNjNzNjNjBmYjVlIiwicmlkIjoiNDE5YjMzNTMtNmQ0Ny00ZTlkLTg3NzctODA3MzIx"
    "MDk2NTJiIn0.8pvd--6kEb1E_GocArR0VtUh3i6RLfqd4aE8i7R2wI_nZJmq6rg-yc3vxtol"
    "xl_Y2Di4oOhcoSUlzhFLa8LcAA"
)
HEADERS = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}


def _arg(val):
    if val is None:
        return {"type": "null", "value": None}
    if isinstance(val, int):
        return {"type": "integer", "value": str(val)}
    if isinstance(val, float):
        return {"type": "float", "value": val}
    return {"type": "text", "value": str(val)}


def run(sql, args=None):
    stmt = {"sql": sql}
    if args:
        stmt["args"] = [_arg(v) for v in args]
    body = {"requests": [{"type": "execute", "stmt": stmt}, {"type": "close"}]}
    r = requests.post(f"{URL}/v2/pipeline", headers=HEADERS, json=body, timeout=15)
    r.raise_for_status()
    result = r.json()["results"][0]
    if result.get("type") == "error":
        raise Exception(result["error"]["message"])
    return result["response"]["result"]


def scalar(sql, args=None):
    rows = run(sql, args)["rows"]
    return rows[0][0]["value"] if rows else None


# ── 1. Fix Fawry gross margin ─────────────────────────────────────────────────
run(
    "UPDATE exit_comps SET gross_margin_pct = 44.0, "
    "notes_caveats = ?, updated_at = datetime('now') "
    "WHERE company_name = 'Fawry'",
    [
        "Gross margin 44.0% is a conservative estimate based on FY2021 trajectory "
        "(FT Partners IPO filing); EBITDA margin 25.0% confirmed ($9.2M EBITDA / $36.7M revenue). "
        "30x oversubscribed IPO; now $155M+ revenue with 43%+ EBITDA margins; grew without VC funding."
    ],
)
print("Fawry gross margin fixed")

# ── 2. Delete Zettle ─────────────────────────────────────────────────────────
run("DELETE FROM portfolio_comp_mapping WHERE comp_id = 'PB-Zettle'")
run("DELETE FROM exit_comps WHERE comp_id = 'PB-Zettle'")
print("Zettle deleted")

# ── 3. Insert iKhokha ────────────────────────────────────────────────────────
run("DELETE FROM exit_comps WHERE comp_id = 'QC-iKhokha'")
run(
    """INSERT INTO exit_comps (
        comp_id, company_name, sub_sector, geography, country,
        founded_year, exit_status, exit_year, exit_type, acquirer_exchange,
        exit_ev_usd_m, revenue_at_exit_usd_m,
        gross_margin_pct, ebitda_margin_pct,
        ev_revenue_multiple, ev_ebitda_multiple,
        key_narrative_drivers, pre_exit_moves, notes_caveats,
        data_confidence, data_source,
        is_clean_exit, use_for_margins, use_for_multiples
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
    [
        "QC-iKhokha", "iKhokha", "Payments / mPOS", "SSA", "South Africa",
        2012, "Exited", 2023, "Strategic M&A", "MTN",
        94.0, 21.0,
        None, None,
        4.5, None,
        "South Africa leading SME mPOS provider; card readers, business account, and working capital products; "
        "strong SME merchant network; acquired by MTN to anchor its fintech and merchant acquiring strategy",
        "Built out business account and capital products alongside core mPOS; expanded merchant base nationwide",
        "MTN acquired iKhokha ~$94M (2023); revenue estimated ~$21M at 4-5x ARR (midpoint used); "
        "gross and EBITDA margins not publicly disclosed",
        "Medium", "Online Research",
        1, 0, 1,
    ],
)
print("iKhokha inserted")

# ── 4. Map iKhokha and DPO Group to Yoco ─────────────────────────────────────
run("DELETE FROM portfolio_comp_mapping WHERE portfolio_company = 'Yoco' AND comp_id IN ('QC-iKhokha','PB-DPO_Group')")
run(
    "INSERT INTO portfolio_comp_mapping (portfolio_company, comp_id, relevance_score, mapping_rationale) "
    "VALUES (?,?,?,?)",
    [
        "Yoco", "QC-iKhokha", 5,
        "SA mPOS direct comp - SME merchant acquiring, card reader, business account; "
        "most geographically comparable; acquired by MTN $94M at 4-5x ARR",
    ],
)
run(
    "INSERT INTO portfolio_comp_mapping (portfolio_company, comp_id, relevance_score, mapping_rationale) "
    "VALUES (?,?,?,?)",
    [
        "Yoco", "PB-DPO_Group", 3,
        "Pan-African PSP exit - 21-country merchant acquiring, 60K+ merchants, "
        "acquired by Network International $291M; infrastructure and multi-country scale comparable",
    ],
)
print("Yoco mappings added")

# ── 5. Verify ─────────────────────────────────────────────────────────────────
fawry_gm    = scalar("SELECT gross_margin_pct FROM exit_comps WHERE company_name = 'Fawry'")
zettle_cnt  = scalar("SELECT COUNT(*) FROM exit_comps WHERE comp_id = 'PB-Zettle'")
ikhokha_ev  = scalar("SELECT ev_revenue_multiple FROM exit_comps WHERE comp_id = 'QC-iKhokha'")
dpo_ev      = scalar("SELECT ev_revenue_multiple FROM exit_comps WHERE comp_id = 'PB-DPO_Group'")
yoco_maps   = run(
    "SELECT comp_id, relevance_score FROM portfolio_comp_mapping WHERE portfolio_company = 'Yoco' "
    "ORDER BY relevance_score DESC"
)["rows"]

print(f"\nFawry gross_margin_pct : {fawry_gm}  (expected 44.0)")
print(f"Zettle count           : {zettle_cnt}  (expected 0)")
print(f"iKhokha ev_rev_mult    : {ikhokha_ev}  (expected 4.5)")
print(f"DPO Group ev_rev_mult  : {dpo_ev}  (expected 11.7)")
print("Yoco comp mappings:")
for row in yoco_maps:
    print(f"  {row[0]['value']:40s}  rel={row[1]['value']}")

print("\nAll done.")
