# Independent migration-healthcare evidence site

This is a separately generated static website for the migration/population panel. It was built on 2026-08-17 from the user-provided workbook:

outputs/20260817_migration_panel_final/migration_population_panel_40countries_2010-2022_final.xlsx

The website lives in its own folder:

openai-work/migration-healthcare-evidence-site

It does not read or copy the prior Claude-generated website directory. The build script independently reads the workbook and independently requests the distinct source URLs recorded in the workbook's Source Audit and data sheets. Retrieved response bytes are stored under sources/; blocked mirrors retain a local pointer and a direct mirror URL. The Panel evidence CSV/JSON gives every displayed nonblank Panel value a clickable evidence target.

The local Python environment uses a proxy whose CA certificate is unavailable to Python's CA bundle. For this snapshot pass only, HTTPS retrieval uses an unverified TLS context; this limitation is recorded in data/source_register.csv and build_summary.json. The original URL, response bytes, and SHA-256 are preserved.

## Browse

After GitHub Pages is enabled, the site is available at the repository's Pages URL. The repository is intended to be public so editors and reviewers can download the workbook, CSV/JSON exports, and snapshots without local access.

## Build

    python scripts/build_site.py --workbook <path-to-workbook> --output . --workers 8

The output manifest is manifest.json; SHA-256 checksums are in SHA256SUMS.txt.

## Co-work note

This website is the co-work of Prof. Raymond Kuo at National Taiwan University (https://raymond.cph.ntu.edu.tw/) and OpenAI GTP-5.6-luna.
