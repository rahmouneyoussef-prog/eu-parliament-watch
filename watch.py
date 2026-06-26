import os
import re
import json
import hashlib
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader
import pandas as pd

BASE = "https://www.europarl.europa.eu"
DATA_DIR = Path("data")
OUTPUT_DIR = Path("output")
SEEN_FILE = DATA_DIR / "seen.json"
KEYWORDS_FILE = Path("keywords.txt")

# Points de départ. On pourra enrichir ensuite avec Open Data / SPARQL.
START_URLS = [
    "https://www.europarl.europa.eu/doceo/recent-documents",
    "https://www.europarl.europa.eu/plenary/en/parliamentary-questions.html?tabType=wq",
    "https://www.europarl.europa.eu/plenary/en/agendas.html",
    "https://www.europarl.europa.eu/plenary/en/texts-adopted.html",
    "https://www.europarl.europa.eu/news/en/press-room",
]

HEADERS = {
    "User-Agent": "EU-Parliament-Watch/1.0 (+https://github.com/)"
}


def ensure_dirs():
    DATA_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)


def load_keywords():
    if not KEYWORDS_FILE.exists():
        return []
    return [line.strip() for line in KEYWORDS_FILE.read_text(encoding="utf-8").splitlines() if line.strip()]


def load_seen():
    if SEEN_FILE.exists():
        return json.loads(SEEN_FILE.read_text(encoding="utf-8"))
    return {"items": {}}


def save_seen(seen):
    SEEN_FILE.write_text(json.dumps(seen, ensure_ascii=False, indent=2), encoding="utf-8")


def fetch(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r


def normalize_url(href, base_url):
    if not href:
        return None
    return urljoin(base_url, href)


def extract_links(page_url, html):
    soup = BeautifulSoup(html, "html.parser")
    links = set()
    for a in soup.find_all("a", href=True):
        u = normalize_url(a["href"], page_url)
        if not u:
            continue
        if "europarl.europa.eu" not in u:
            continue
        if any(x in u for x in ["/doceo/document/", "/plenary/", "/news/"]):
            links.add(u.split("#")[0])
    return sorted(links)


def text_from_html(html):
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    return soup.get_text(" ", strip=True)


def text_from_pdf_bytes(content):
    tmp = DATA_DIR / "tmp.pdf"
    tmp.write_bytes(content)
    try:
        reader = PdfReader(str(tmp))
        parts = []
        for i, page in enumerate(reader.pages, start=1):
            try:
                txt = page.extract_text() or ""
            except Exception:
                txt = ""
            if txt.strip():
                parts.append(f"\n[page {i}]\n{txt}")
        return "\n".join(parts)
    finally:
        try:
            tmp.unlink()
        except FileNotFoundError:
            pass


def get_doc_text(url):
    r = fetch(url)
    ctype = r.headers.get("content-type", "").lower()
    if url.lower().endswith(".pdf") or "application/pdf" in ctype:
        return text_from_pdf_bytes(r.content), "pdf"
    return text_from_html(r.text), "html"


def find_hits(text, keywords):
    lower = text.lower()
    hits = []
    for kw in keywords:
        if kw.lower() in lower:
            hits.append(kw)
    return hits


def snippet_for(text, keyword, radius=220):
    m = re.search(re.escape(keyword), text, re.IGNORECASE)
    if not m:
        return ""
    start = max(0, m.start() - radius)
    end = min(len(text), m.end() + radius)
    snippet = text[start:end]
    return re.sub(r"\s+", " ", snippet).strip()


def item_id(url):
    return hashlib.sha256(url.encode("utf-8")).hexdigest()[:16]


def telegram_send(message):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        print("Telegram secrets missing; printing only.")
        print(message)
        return
    api = f"https://api.telegram.org/bot{token}/sendMessage"
    r = requests.post(api, data={"chat_id": chat_id, "text": message, "disable_web_page_preview": True}, timeout=30)
    r.raise_for_status()


def discover_candidate_urls():
    candidates = set(START_URLS)
    for url in START_URLS:
        try:
            r = fetch(url)
            candidates.update(extract_links(url, r.text))
        except Exception as e:
            print(f"Discovery error on {url}: {e}")
    # garder une taille raisonnable pour GitHub Actions
    return sorted(candidates)[:250]


def classify_url(url):
    u = url.lower()
    if "-asw_" in u or "-asw" in u:
        return "Answer / ASW"
    if "/document/e-" in u:
        return "Written question"
    if "agenda" in u:
        return "Agenda"
    if "texts-adopted" in u:
        return "Texts adopted"
    if u.endswith(".pdf"):
        return "PDF document"
    return "EP document"


def write_excel(rows):
    if not rows:
        return None
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    path = OUTPUT_DIR / f"ep_watch_results_{today}.xlsx"
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="alerts")
    return path


def main():
    ensure_dirs()
    keywords = load_keywords()
    if not keywords:
        raise RuntimeError("No keywords in keywords.txt")

    seen = load_seen()
    seen_items = seen.setdefault("items", {})
    rows = []
    alert_count = 0

    urls = discover_candidate_urls()
    print(f"Candidates: {len(urls)}")

    for url in urls:
        iid = item_id(url)
        if iid in seen_items:
            continue
        try:
            text, fmt = get_doc_text(url)
            hits = find_hits(text, keywords)
        except Exception as e:
            print(f"Error reading {url}: {e}")
            continue

        # Marquer comme vu même sans hit pour éviter de relire sans cesse les mêmes pages.
        seen_items[iid] = {"url": url, "checked_at": datetime.now(timezone.utc).isoformat(), "matched": bool(hits)}

        if not hits:
            continue

        first_hit = hits[0]
        snip = snippet_for(text, first_hit)
        doc_type = classify_url(url)
        row = {
            "detected_at_utc": datetime.now(timezone.utc).isoformat(),
            "type": doc_type,
            "format": fmt,
            "keywords": ", ".join(hits),
            "url": url,
            "snippet": snip,
        }
        rows.append(row)
        alert_count += 1

        msg = (
            "🚨 Alerte Parlement européen\n\n"
            f"Type: {doc_type}\n"
            f"Mots-clés: {', '.join(hits)}\n"
            f"Lien: {url}\n\n"
            f"Extrait: {snip[:900]}"
        )
        telegram_send(msg)

    excel_path = write_excel(rows)
    if excel_path:
        telegram_send(f"📊 Export Excel généré: {excel_path}")

    save_seen(seen)
    print(f"Done. Alerts: {alert_count}")


if __name__ == "__main__":
    main()
