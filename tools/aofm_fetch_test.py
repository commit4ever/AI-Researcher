from __future__ import annotations

import hashlib
import io
import json
import re
import zipfile
from pathlib import Path
from urllib.parse import urljoin

import requests
from openpyxl import load_workbook

DATA_HUB = "https://www.aofm.gov.au/data-hub"
PINNED_BONDS_XLSX = "https://www.aofm.gov.au/sites/default/files/2025-06-20/treasury%20bonds%20-%20issuance.xlsx"
UA = "sovereign-bond-auction-monitor/0.1 (+GitHub Actions; official-source verification)"
TIMEOUT = 12


def get(url: str) -> requests.Response:
    r = requests.get(
        url,
        timeout=TIMEOUT,
        allow_redirects=True,
        headers={"User-Agent": UA, "Accept": "*/*"},
    )
    print(json.dumps({
        "request_url": url,
        "status": r.status_code,
        "final_url": r.url,
        "content_type": r.headers.get("content-type"),
        "content_length": len(r.content),
    }), flush=True)
    r.raise_for_status()
    return r


def discover_bond_xlsx(html: str) -> list[str]:
    hrefs = re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.I)
    urls = [urljoin(DATA_HUB, h) for h in hrefs]
    matches = [
        u for u in urls
        if "treasury" in u.lower()
        and "bond" in u.lower()
        and u.lower().split("?")[0].endswith(".xlsx")
    ]
    return list(dict.fromkeys(matches))


def validate_xlsx(data: bytes) -> dict:
    if len(data) < 4 or data[:2] != b"PK":
        raise RuntimeError(
            f"response is not an XLSX/ZIP payload; first bytes={data[:20]!r}"
        )
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise RuntimeError("PK signature present but payload is not a valid ZIP/XLSX")

    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sample = {}
    for ws in wb.worksheets[:3]:
        rows = []
        for row in ws.iter_rows(
            min_row=1,
            max_row=min(ws.max_row or 1, 15),
            values_only=True,
        ):
            rows.append([
                v.isoformat() if hasattr(v, "isoformat") else v
                for v in row[:15]
            ])
        sample[ws.title] = rows
    return {"sheet_names": wb.sheetnames, "sample": sample}


def main() -> int:
    out = Path("aofm_test_output")
    out.mkdir(exist_ok=True)

    report = {
        "data_hub": {"url": DATA_HUB},
        "xlsx_attempts": [],
    }

    discovered = []
    try:
        hub = get(DATA_HUB)
        discovered = discover_bond_xlsx(hub.text)
        report["data_hub"].update({
            "ok": True,
            "http_status": hub.status_code,
            "final_url": hub.url,
            "discovered_bond_xlsx": discovered,
        })
    except Exception as exc:
        report["data_hub"].update({
            "ok": False,
            "error": repr(exc),
        })
        print(json.dumps(report["data_hub"]), flush=True)

    candidates = list(dict.fromkeys(discovered + [PINNED_BONDS_XLSX]))

    successful_xlsx = None
    for url in candidates:
        attempt = {"url": url}
        try:
            r = get(url)
            meta = validate_xlsx(r.content)
            sha = hashlib.sha256(r.content).hexdigest()
            xlsx_path = out / "treasury_bonds_issuance.xlsx"
            xlsx_path.write_bytes(r.content)
            attempt.update({
                "ok": True,
                "final_url": r.url,
                "http_status": r.status_code,
                "content_type": r.headers.get("content-type"),
                "byte_length": len(r.content),
                "sha256": sha,
                **meta,
            })
            successful_xlsx = attempt
            report["xlsx_attempts"].append(attempt)
            break
        except Exception as exc:
            attempt.update({"ok": False, "error": repr(exc)})
            report["xlsx_attempts"].append(attempt)
            print(json.dumps(attempt), flush=True)

    report["success"] = successful_xlsx is not None
    (out / "result.json").write_text(
        json.dumps(report, indent=2, default=str), encoding="utf-8"
    )
    print(json.dumps(report, indent=2, default=str), flush=True)
    return 0 if report["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
