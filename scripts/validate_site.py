from __future__ import annotations

import csv
import hashlib
import html
import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]


def fail(message):
    print(f"ERROR: {message}")
    raise SystemExit(1)


def main():
    manifest = json.loads((ROOT / "manifest.json").read_text(encoding="utf-8"))
    workbook = ROOT / "data" / manifest["independent_input"]
    if not workbook.exists():
        fail(f"missing workbook: {workbook}")
    actual_hash = hashlib.sha256(workbook.read_bytes()).hexdigest()
    if actual_hash != manifest["input_workbook_sha256"]:
        fail("workbook hash does not match manifest")

    html_files = [path for path in ROOT.rglob("*.html") if "sources" not in path.relative_to(ROOT).parts]
    if len(html_files) != 46:
        fail(f"expected 46 HTML pages, found {len(html_files)}")
    href_pattern = re.compile(r"""(?:href|src)=["']([^"']+)["']""", re.IGNORECASE)
    links_checked = 0
    missing = []
    for page in html_files:
        text = page.read_text(encoding="utf-8")
        for raw in href_pattern.findall(text):
            target = html.unescape(raw)
            parsed = urlsplit(target)
            if parsed.scheme or target.startswith("#") or target.startswith("data:"):
                continue
            relative = unquote(parsed.path)
            if not relative:
                continue
            destination = (page.parent / relative).resolve()
            links_checked += 1
            if not destination.exists():
                missing.append(f"{page.relative_to(ROOT)} -> {target}")
    if missing:
        fail("missing local links: " + "; ".join(missing[:8]))

    register_path = ROOT / "data" / "source_register.csv"
    with register_path.open(encoding="utf-8-sig", newline="") as handle:
        register = list(csv.DictReader(handle))
    retrieved = [row for row in register if row["retrieval_status"] in {"retrieved", "retrieved_via_alternate"}]
    external = [row for row in register if row["retrieval_status"] == "external_mirror"]
    failed = [row for row in register if row["retrieval_status"] == "failed"]
    for row in retrieved + external + failed:
        if not row["snapshot_path"]:
            fail(f"source {row['source_id']} has no snapshot path")
        if not (ROOT / row["snapshot_path"]).exists():
            fail(f"source {row['source_id']} snapshot path is missing")
    for row in external:
        if not row.get("snapshot_source_url", "").startswith(("http://", "https://")):
            fail(f"external mirror {row['source_id']} has no mirror URL")
    if len(register) != manifest["site"]["source_urls"]:
        fail("source register count does not match manifest")
    if len(retrieved) != manifest["site"]["retrieved_snapshots"]:
        fail("retrieved snapshot count does not match manifest")
    if len(external) != manifest["site"].get("external_mirrors", 0):
        fail("external mirror count does not match manifest")
    if len(failed) != manifest["site"]["failed_retrievals"]:
        fail("failed snapshot count does not match manifest")

    panel_json = json.loads((ROOT / "data" / "panel.json").read_text(encoding="utf-8"))
    if len(panel_json) != manifest["site"]["panel_rows"]:
        fail("panel row count does not match manifest")
    evidence_path = ROOT / "data" / "panel_evidence.csv"
    with evidence_path.open(encoding="utf-8-sig", newline="") as handle:
        evidence = list(csv.DictReader(handle))
    expected_evidence = sum(
        1
        for row in panel_json
        for key, _ in [
            ("year", "Year"),
            ("population", "Population"),
            ("foreign_born", "Foreign-born"),
            ("foreign_born_pct_pop", "Foreign-born %"),
            ("foreign_nationals", "Foreign nationals"),
            ("foreign_nationals_pct_pop", "Foreign nationals %"),
            ("irregular_stock", "Irregular stock"),
            ("irregular_proxy_overstayers", "Overstayer proxy"),
            ("irregular_proxy_detections", "Detection proxy"),
        ]
        if row.get(key) not in (None, "")
    )
    if len(evidence) != expected_evidence or len(evidence) != manifest["site"].get("panel_evidence_rows"):
        fail(f"Panel evidence count mismatch: expected {expected_evidence}, found {len(evidence)}")
    evidence_by_iso = {}
    for row in evidence:
        evidence_by_iso[row["iso3"]] = evidence_by_iso.get(row["iso3"], 0) + 1
        if row["retrieval_status"] in {"failed", "unmapped", ""}:
            fail(f"Panel value {row['cell_id']} has no evidence target")
        if row["downloadable_local"].lower() == "true":
            if not row["local_path"] or not (ROOT / row["local_path"]).exists():
                fail(f"Panel value {row['cell_id']} has a missing local evidence file")
        elif not row["evidence_href"].startswith(("http://", "https://")):
            fail(f"Panel value {row['cell_id']} has no external evidence URL")

    country_pages = list((ROOT / "countries").glob("*.html"))
    if len(country_pages) != manifest["site"]["country_pages"]:
        fail("country page count does not match manifest")
    for page in country_pages:
        iso3 = page.stem
        expected = evidence_by_iso.get(iso3, 0)
        text = page.read_text(encoding="utf-8")
        if "Panel data" not in text:
            fail(f"{page.name} is missing the Panel data section")
        links = re.findall(r'<a class="evidence-link" href="([^"]+)"', text)
        if len(links) != expected:
            fail(f"{page.name} has {len(links)} evidence links; expected {expected}")

    author_phrase = "Co-work of"
    for page in html_files:
        if author_phrase not in page.read_text(encoding="utf-8"):
            fail(f"{page.relative_to(ROOT)} is missing the co-work note")

    print(
        f"Validated {len(html_files)} HTML pages, {links_checked} local links, "
        f"{len(register)} source records ({len(retrieved)} local snapshots, {len(external)} external mirrors, {len(failed)} retrieval records), "
        f"{len(evidence)} Panel evidence links, and workbook hash {actual_hash}."
    )


if __name__ == "__main__":
    main()
