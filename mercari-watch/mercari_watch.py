#!/usr/bin/env python3
"""Vigila los enlaces de búsqueda de Mercari de tu Google Sheets.

Por cada celda con un enlace de jp.mercari.com/search carga la búsqueda (ordenada por
más recientes y solo en venta), compara con los anuncios ya vistos y, si hay nuevos,
colorea la celda y le pone una nota con el detalle.

Variables de entorno: SPREADSHEET_ID, GOOGLE_CREDENTIALS (JSON de la cuenta de servicio),
SHEETS (opcional, pestañas separadas por "|").
Uso local de prueba:  python mercari_watch.py --dry-run --limit 3
"""
import argparse, json, os, random, re, sys, time
import urllib.parse as up
from datetime import datetime
from zoneinfo import ZoneInfo

STATE_TITLE = "_mercari_estado"
ALERT_COLOR = {"red": 1.0, "green": 0.84, "blue": 0.35}
MAX_SEEN = 300
MAX_NOTE = 1800
NO_RESULT = "出品された商品がありません"
TZ = ZoneInfo("Europe/Madrid")

ITEM_SELECTOR = 'a[href*="/item/m"], a[href*="/shops/product/"]'
EXTRACT_JS = r"""() => {
  const out = [], seen = new Set();
  for (const a of document.querySelectorAll('a[href*="/item/m"], a[href*="/shops/product/"]')) {
    const m = (a.getAttribute('href') || '').match(/\/(?:item|shops\/product)\/([A-Za-z0-9]+)/);
    if (!m || seen.has(m[1])) continue;
    seen.add(m[1]);
    const el = a.querySelector('[aria-label]');
    const img = a.querySelector('img');
    let label = (el && el.getAttribute('aria-label')) || a.getAttribute('aria-label')
              || (img && img.alt) || a.innerText || '';
    out.push({id: m[1], href: a.href, label: label.replace(/\s+/g, ' ').trim().slice(0, 140)});
  }
  return out;
}"""


def normalize_url(url):
    """Fuerza 'solo en venta' + 'más recientes' para que la comparación sea fiable."""
    p = up.urlparse(url)
    q = up.parse_qs(p.query, keep_blank_values=True)
    q["status"] = ["on_sale"]
    q["sort"] = ["created_time"]
    q["order"] = ["desc"]
    return up.urlunparse(p._replace(query=up.urlencode(q, doseq=True)))


# ---------------------------------------------------------------- Google Sheets
def open_sheet():
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_CREDENTIALS"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds).open_by_key(os.environ["SPREADSHEET_ID"])


def collect_links(sh, only):
    fields = ("sheets(properties(title,sheetId),data(startRow,startColumn,rowData(values("
              "hyperlink,note,userEnteredValue/formulaValue,textFormatRuns/format/link))))")
    meta = sh.fetch_sheet_metadata(params={"includeGridData": "true", "fields": fields})
    links = []
    for s in meta["sheets"]:
        title = s["properties"]["title"]
        if title == STATE_TITLE or (only and title not in only):
            continue
        for d in s.get("data", []):
            r0, c0 = d.get("startRow", 0), d.get("startColumn", 0)
            for r, row in enumerate(d.get("rowData", [])):
                for c, cell in enumerate(row.get("values", [])):
                    url = cell.get("hyperlink")
                    if not url:
                        for run in cell.get("textFormatRuns", []):
                            url = run.get("format", {}).get("link", {}).get("uri") or url
                    if not url:
                        f = cell.get("userEnteredValue", {}).get("formulaValue", "")
                        m = re.search(r'HYPERLINK\(\s*"([^"]+)"', f, re.I)
                        url = m.group(1) if m else None
                    if url and "mercari.com" in url and "/search" in url:
                        links.append(dict(sheetId=s["properties"]["sheetId"], title=title,
                                          row=r0 + r, col=c0 + c, url=url,
                                          note=cell.get("note", "")))
    return links


def load_state(sh):
    import gspread
    try:
        ws = sh.worksheet(STATE_TITLE)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(STATE_TITLE, rows=500, cols=4)
        ws.update(range_name="A1", values=[["clave", "ids_vistos", "ultima_revision", "estado"]])
        ws.hide()
    state = {}
    for row in ws.get_all_values()[1:]:
        if row and row[0]:
            try:
                state[row[0]] = json.loads(row[1]) if len(row) > 1 and row[1] else []
            except json.JSONDecodeError:
                state[row[0]] = []
    return ws, state


def save_state(ws, state, status):
    now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M")
    rows = [["clave", "ids_vistos", "ultima_revision", "estado"]]
    for k, ids in state.items():
        rows.append([k, json.dumps(ids[:MAX_SEEN]), now, status.get(k, "")])
    ws.clear()
    ws.update(range_name="A1", values=rows)


# ---------------------------------------------------------------- Mercari
def scrape(page, url):
    """Devuelve lista de anuncios (puede ser []), o None si falló/bloquearon."""
    from playwright.sync_api import TimeoutError as PWTimeout
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_selector(ITEM_SELECTOR, timeout=25000)
    except PWTimeout:
        try:
            if NO_RESULT in page.inner_text("body"):
                return []
        except Exception:
            pass
        return None
    except Exception as e:
        print("   error:", e)
        return None
    page.wait_for_timeout(1200)
    return page.evaluate(EXTRACT_JS)


def build_note(old_note, new_items):
    stamp = datetime.now(TZ).strftime("%d/%m %H:%M")
    lines = [f"🆕 {len(new_items)} anuncio(s) nuevo(s) · {stamp}"]
    for it in new_items[:5]:
        lines.append(f"• {it['label'] or it['id']}\n  {it['href']}")
    if len(new_items) > 5:
        lines.append(f"… y {len(new_items) - 5} más")
    text = "\n".join(lines)
    if old_note.startswith("🆕"):
        text += "\n\n— anteriores —\n" + old_note
    return text[:MAX_NOTE]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="no escribe en la hoja")
    ap.add_argument("--limit", type=int, default=0, help="solo los N primeros enlaces")
    ap.add_argument("--url", help="probar una sola URL de Mercari sin tocar Google Sheets")
    args = ap.parse_args()

    from playwright.sync_api import sync_playwright

    if args.url:
        links, sh, state_ws, state = [dict(url=args.url, title="-", row=0, col=0,
                                           sheetId=0, note="")], None, None, {}
    else:
        sh = open_sheet()
        only = [s for s in os.environ.get("SHEETS", "").split("|") if s]
        links = collect_links(sh, only)
        state_ws, state = load_state(sh)
        print(f"{len(links)} enlaces de Mercari encontrados")
    if args.limit:
        links = links[:args.limit]

    requests_, status, failures = [], {}, 0
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(
            locale="ja-JP", timezone_id="Asia/Tokyo",
            viewport={"width": 1366, "height": 900},
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"))
        ctx.route("**/*", lambda r: r.abort()
                  if r.request.resource_type in ("image", "media", "font") else r.continue_())
        page = ctx.new_page()

        for i, ln in enumerate(links, 1):
            url = normalize_url(ln["url"])
            items = scrape(page, url)
            tag = f"[{i}/{len(links)}] {ln['title']} r{ln['row']+1}c{ln['col']+1}"
            if items is None:
                failures += 1
                status[url] = "ERROR: no se pudo leer (¿bloqueo?)"
                print(f"{tag}: FALLO")
            else:
                ids = [it["id"] for it in items]
                print(f"{tag}: {len(ids)} anuncios" + (f" · ej: {items[0]['label'][:50]}" if items else ""))
                if url not in state:
                    state[url] = ids
                    status[url] = "línea base"
                else:
                    seen = set(state[url])
                    new = [it for it in items if it["id"] not in seen]
                    state[url] = ids + [x for x in state[url] if x not in set(ids)]
                    status[url] = f"ok, {len(new)} nuevos"
                    if new:
                        print(f"   🆕 {len(new)} nuevos")
                        requests_.append({"updateCells": {
                            "range": {"sheetId": ln["sheetId"], "startRowIndex": ln["row"],
                                      "endRowIndex": ln["row"] + 1, "startColumnIndex": ln["col"],
                                      "endColumnIndex": ln["col"] + 1},
                            "rows": [{"values": [{
                                "note": build_note(ln["note"], new),
                                "userEnteredFormat": {"backgroundColor": ALERT_COLOR}}]}],
                            "fields": "note,userEnteredFormat.backgroundColor"}})
            time.sleep(random.uniform(2.5, 5.5))
        browser.close()

    if sh and not args.dry_run:
        if requests_:
            sh.batch_update({"requests": requests_})
        save_state(state_ws, state, status)
    print(f"Hecho: {len(requests_)} celdas avisadas, {failures} fallos de {len(links)}")
    if links and failures > len(links) / 2:
        sys.exit(1)  # GitHub te mandará un email de "workflow failed"


if __name__ == "__main__":
    main()
