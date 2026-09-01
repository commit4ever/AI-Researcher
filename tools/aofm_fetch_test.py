from __future__ import annotations

import hashlib
import io
import json
import re
import sys
import zipfile
from pathlib import Path
from urllib.parse import urljoin

import requests
from openpyxl import load_workbook

DATA_HUB = "https://www.aofm.gov.au/data-hub"
PINNED_BONDS_XLSX = "https://www.aofm.gov.au/sites/default/files/2025-06-20/treasury%20bonds%20-%20issuance.xlsx"
UA = "sovereign-bond-auction-monitor/0.1 (+GitHub Actions; official-source verification)"


def get(url: str) -> requests.Response:
    r = requests.get(url, timeout=45, allow_redirects=True, headers={"User-Agent": UA, "Accept": "*/*"})
    print(json.dumps({
        "request_url": url,
        "status": r.status_code,
        "final_url": r.url,
        "content_type": r.headers.get("content-type"),
        "content_length": len(r.content),
    }))
    r.raise_for_status()
    return r


def discover_bond_xlsx(html: str) -> list[str]:
    hrefs = re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.I)
    urls = [urljoin(DATA_HUB, h) for h in hrefs]
    matches = [u for u in urls if "treasury" in u.lower() and "bond" in u.lower() and u.lower().split("?")[0].endswith(".xlsx")]
    return list(dict.fromkeys(matches))


def validate_xlsx(data: bytes) -> dict:
    if len(data) < 4 or data[:2] != b"PK":
        raise RuntimeError(f"response is not an XLSX/ZIP payload; first bytes={data[:20]!r}")
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise RuntimeError("PK signature present but payload is not a valid ZIP/XLSX")
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    sample = {}
    for ws in wb.worksheets[:3]:
        rows = []
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 12), values_only=True):
            rows.append([v.isoformat() if hasattr(v, "isoformat") else v for v in row[:12]])
        sample[ws.title] = rows
    return {"sheet_names": sheet_names, "sample": sample}


def main() -> int:
    out = Path("aofm_test_output")
    out.mkdir(exist_ok=True)

    hub = get(DATA_HUB)
    discovered = discover_bond_xlsx(hub.text)
    print(json.dumps({"discovered_bond_xlsx": discovered}, indent=2))

    candidates = discovered or [PINNED_BONDS_XLSX]
    if PINNED_BONDS_XLSX not in candidates:
        candidates.append(PINNED_BONDS_XLSX)

    failures = []
    for url in candidates:
        try:
            r = get(url)
            meta = validate_xlsx(r.content)
            sha = hashlib.sha256(r.content).hexdigest()
            xlsx_path = out / "treasury_bonds_issuance.xlsx"
            xlsx_path.write_bytes(r.content)
            report = {
                "source_url": url,
                "final_url": r.url,
                "http_status": r.status_code,
                "content_type": r.headers.get("content-type"),
                "byte_length": len(r.content),
                "sha256": sha,
                **meta,
            }
            (out / "result.json").write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
            print(json.dumps(report, indent=2, default=str))
            return 0
        except Exception as e:
            failures.append({"url": url, "error": repr(e)})
            print(json.dumps(failures[-1]))

    (out / "failures.json").write_text(json.dumps(failures, indent=2), encoding="utf-8")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
