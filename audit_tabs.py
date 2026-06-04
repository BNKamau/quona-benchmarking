# audit_tabs.py
import re

with open("app.py", "r", encoding="utf-8") as f:
    src = f.read()

COMPANIES = [
    "cowrywise", "vertofx", "lulalend", "yoco",
    "twinco", "maxsoko", "khazna", "enza", "sava",
]

tab_pattern = re.compile(
    r"def _render_(\w+)_exit_tab\(\).*?(?=\ndef _render_|\ndef render_exit_tab|\Z)",
    re.DOTALL,
)
tabs = {m.group(1).lower(): m.group(0) for m in tab_pattern.finditer(src)}

print(f"Found exit tab functions: {sorted(tabs.keys())}\n")

CHECKS = {
    "pathway_card":      r"def _pathway_card",
    "expander_pathways": r'st\.expander\("Exit Pathways',
    "implied_valuation": r"Implied Valuation",
    "acquirer_universe": r"Acquirer Universe",
    "affinity_cache":    r"affinity_cache\s*=\s*st\.session_state",
    "sync_affinity_btn": r'st\.button\("Sync Affinity"',
    "header_row":        r"def _header_row",
    "buyer_row":         r"def _buyer_row",
    "affinity_override": r"affinity_override",
    "tab_local":         r"tab_local",
    "tab_global":        r"tab_global",
    "next_steps":        r"Next Steps Generator",
    "buyer_actions":     r"_BUYER_ACTIONS",
    "generate_button":   r'st\.button\("Generate',
    "stale_check":       r'note\.get\("stale"\)',
    "5_col_layout":      r'st\.columns\(\[2, 2, 3, 1, 2\]\)',
}

flags = []
results = {}
for company, body in sorted(tabs.items()):
    results[company] = {}
    for check, pattern in CHECKS.items():
        found = bool(re.search(pattern, body))
        results[company][check] = found
        if not found:
            flags.append((company, check))

header = f"{'Check':<22}" + "".join(f"{c[:8]:<10}" for c in sorted(tabs.keys()))
print(header)
print("-" * len(header))
for check in CHECKS:
    row = f"{check:<22}"
    for company in sorted(tabs.keys()):
        val = results[company][check]
        row += f"{'OK':<10}" if val else f"{'MISSING':<10}"
    print(row)

print("\n=== Routing in render_exit_tab() ===")
routing_block = re.search(
    r"def render_exit_tab.*?(?=\ndef |\Z)", src, re.DOTALL
).group(0)
for company in COMPANIES:
    found = company.lower() in routing_block.lower()
    status = "OK" if found else "MISSING ROUTE"
    print(f"  {company:<15} {status}")

print("\n=== SUPPORTED_COMPANIES + PARSERS in excel_parsers.py ===")
try:
    with open("parsers/excel_parsers.py", "r", encoding="utf-8") as f:
        parser_src = f.read()
    for company in COMPANIES:
        in_supported = company.lower() in parser_src.lower()
        in_parsers   = (f'"{company.title()}"' in parser_src or
                        f'"{company.upper()}"' in parser_src)
        print(f"  {company:<15} supported={'OK' if in_supported else 'MISSING'}  "
              f"parser_key={'OK' if in_parsers else 'MISSING'}")
except FileNotFoundError:
    print("  parsers/excel_parsers.py not found")

print(f"\n{'='*50}")
print(f"Total flags: {len(flags)}")
for company, check in sorted(flags):
    print(f"  MISSING  {company:<15} {check}")
