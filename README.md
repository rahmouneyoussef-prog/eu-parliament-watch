# EU Parliament Watch — robust v2

Recent-only monitoring of European Parliament publications with Telegram alerts.

## Important fixes in v2

- Extracts the **real document title** from the surrounding H3/card, not only the generic “PDF (137 KB)” anchor.
- Prioritises keyword-bearing titles and direct `RegData`/DOCEO/PDF/XML documents.
- Raises the candidate limit from 900 to 6000 so late-listed committees such as PETI are not cut off.
- Sends a title-level alert even when PDF download/text extraction temporarily fails.
- Does **not** mark failed downloads as seen, so they are retried later.
- Keeps compact one-document-per-message Telegram alerts and Excel attachment.

## Validation case

The parser includes a regression test for the PETI page pattern that contains:

`PETI_CM(2026)790891 — impact of the EU-Morocco Association Agreement...`

## Install

Replace these files in the GitHub repository:

- `watch.py`
- `requirements.txt`
- `README.md`

Keep your existing `keywords.txt`, secrets, `data/seen.json`, and workflow.
