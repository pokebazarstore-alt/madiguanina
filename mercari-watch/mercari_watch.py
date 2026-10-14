#!/usr/bin/env python3
"""Vigila los enlaces de búsqueda de Mercari de tu Google Sheets.

Por cada celda con un enlace de jp.mercari.com/search carga la búsqueda (ordenada por
más recientes y solo en venta), compara con los anuncios ya vistos y, si hay nuevos,
colorea la celda y le pone una nota con el detalle. En cada ejecución quita de la nota
los anuncios que ya no están a la venta; si no queda ninguno, limpia la nota y el color.

Variables de entorno: SPREADSHEET_ID, GOOGLE_CREDENTIALS (JSON de la cuenta de servicio),
SHEETS (opcional, pestañas separadas por "|").
Uso local de prueba:  python mercari_watch.py --dry-run --limit 3
"""
import argparse, json, os, random, re, sys, time
import urllib.parse as up
from datetime import datetime
from zoneinfo import ZoneInfo

STATE_TITLE = "_mercari_estado"
COLORS_TITLE = "_mercari_colores"
PENDING_TITLE = "_mercari_pendientes"
# Color según nº de anuncios sin revisar: 1-5 verde, 6-15 amarillo, 16 o más rojo
COLOR_GREEN = {"red": 0.58, "green": 0.84, "blue": 0.52}
COLOR_YELLOW = {"red": 1.0, "green": 0.84, "blue": 0.35}
COLOR_RED = {"red": 0.93, "green": 0.40, "blue": 0.38}


def alert_color(n):
    return COLOR_GREEN if n <= 5 else COLOR_YELLOW if n <= 15 else COLOR_RED

MAX_SEEN = 300
TARGET_ITEMS = 90
MAX_ALERT_BATCH = 25  # si salen más 'nuevos' de golpe, se considera lectura anómala
STATE_VERSION = "2"
MAX_NOTE = 1800
MAX_LISTED = 8
MISS_TO_DROP = 2  # ejecuciones seguidas sin aparecer (vendido/retirado) antes de quitarlo de la nota
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
    label = label.replace(/\s+/g, ' ').replace(/のサムネイル$/, '').trim().slice(0, 140);
    const pm = (a.innerText || '').replace(/\s+/g, '').match(/[\d,]{3,}/);
    out.push({id: m[1], href: a.href, label: label, price: pm ? pm[0] : ''});
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


def hex_color(bg):
    """Color de la API -> '#rrggbb'. Blanco/sin color -> '' (= 'sin color')."""
    if not bg:
        return ""
    r, g, b = (round(bg.get(k, 0) * 255) for k in ("red", "green", "blue"))
    return "" if (r, g, b) == (255, 255, 255) else f"#{r:02x}{g:02x}{b:02x}"


def alert_hexes():
    return {hex_color(c) for c in (COLOR_GREEN, COLOR_YELLOW, COLOR_RED)}


def hex_to_rgb(h):
    return {"red": int(h[1:3], 16) / 255, "green": int(h[3:5], 16) / 255,
            "blue": int(h[5:7], 16) / 255}


def hidden_sheet(sh, title, header):
    import gspread
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title, rows=500, cols=len(header))
        ws.update(range_name="A1", values=[header])
        ws.hide()
        return ws


def load_colors(sh):
    ws = hidden_sheet(sh, COLORS_TITLE, ["celda (hoja|fila|col)", "color_original"])
    return ws, {r[0]: (r[1] if len(r) > 1 else "") for r in ws.get_all_values()[1:] if r and r[0]}


def save_colors(ws, colors):
    ws.clear()
    ws.update(range_name="A1", values=[["celda (hoja|fila|col)", "color_original"]]
              + [[k, v] for k, v in colors.items()])


def load_pending(sh):
    ws = hidden_sheet(sh, PENDING_TITLE, ["hoja|url", "anuncios_sin_revisar"])
    out = {}
    for r in ws.get_all_values()[1:]:
        if r and r[0] and len(r) > 1 and r[1]:
            try:
                out[r[0]] = json.loads(r[1])
            except json.JSONDecodeError:
                pass
    return ws, out


def save_pending(ws, pending):
    ws.clear()
    ws.update(range_name="A1", values=[["hoja|url", "anuncios_sin_revisar"]]
              + [[k, json.dumps(v, ensure_ascii=False)] for k, v in pending.items() if v])


def collect_links(sh, only):
    fields = ("sheets(properties(title,sheetId),data(startRow,startColumn,rowData(values("
              "hyperlink,note,userEnteredFormat/backgroundColor,userEnteredValue/formulaValue,"
              "textFormatRuns/format/link))))")
    meta = sh.fetch_sheet_metadata(params={"includeGridData": "true", "fields": fields})
    links = []
    for s in meta["sheets"]:
        title = s["properties"]["title"]
        if title in (STATE_TITLE, COLORS_TITLE, PENDING_TITLE) or (only and title not in only):
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
                                          note=cell.get("note", ""),
                                          bg=hex_color(cell.get("userEnteredFormat", {})
                                                       .get("backgroundColor"))))
    return links


def load_state(sh):
    ws = hidden_sheet(sh, STATE_TITLE, ["clave", "ids_vistos", "ultima_revision", "estado"])
    state = {}
    rows = ws.get_all_values()[1:]
    version = next((r[1] for r in rows if r and r[0] == "__version__"), None)
    if version != STATE_VERSION:  # formato antiguo: empezar de cero (sin avisos falsos)
        print("Estado antiguo detectado: se reinicia el punto de partida")
        return ws, state
    for row in rows:
        if row and row[0] and row[0] != "__version__":
            try:
                state[row[0]] = json.loads(row[1]) if len(row) > 1 and row[1] else []
            except json.JSONDecodeError:
                state[row[0]] = []
    return ws, state


def save_state(ws, state, status):
    now = datetime.now(TZ).strftime("%Y-%m-%d %H:%M")
    rows = [["clave", "ids_vistos", "ultima_revision", "estado"],
            ["__version__", STATE_VERSION, now, ""]]
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
    page.wait_for_timeout(1000)
    last = 0
    for k in range(8):  # baja para que Mercari cargue más anuncios (~90)
        n = page.locator(ITEM_SELECTOR).count()
        if n >= TARGET_ITEMS or (k > 0 and n == last):
            break
        last = n
        page.mouse.wheel(0, 5000)
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        page.wait_for_timeout(1100)
    return page.evaluate(EXTRACT_JS)


# ---------------------------------------------------------------- Lógica de avisos
def build_note(pend):
    lines = [f"🆕 {len(pend)} anuncio(s) a la venta sin revisar"]
    for p in pend[:MAX_LISTED]:
        price = f" · ¥{p['price']}" if p.get("price") else ""
        lines.append(f"• {p['label'] or p['id']}{price}\n  {p['href']}\n  (detectado {p['ts']})")
    if len(pend) > MAX_LISTED:
        lines.append(f"… y {len(pend) - MAX_LISTED} más")
    return "\n".join(lines)[:MAX_NOTE]


def cell_range(ln):
    return {"sheetId": ln["sheetId"], "startRowIndex": ln["row"], "endRowIndex": ln["row"] + 1,
            "startColumnIndex": ln["col"], "endColumnIndex": ln["col"] + 1}


def process_link(ln, url, items, state, pending, colors, status, requests_, now=None):
    """Actualiza estado/pendientes de un enlace con la lectura `items` (lista) y añade a
    `requests_` los cambios de nota/color necesarios. Devuelve nº de anuncios nuevos."""
    now = now or datetime.now(TZ).strftime("%d/%m %H:%M")
    ids = [it["id"] for it in items]
    cur = set(ids)
    pkey = f"{ln['sheetId']}|{url}"
    ckey = f"{ln['sheetId']}|{ln['row']}|{ln['col']}"
    flagged = ln["note"].startswith("🆕")

    new = []
    if url not in state:
        state[url] = ids
        status[url] = "línea base"
    else:
        seen = set(state[url])
        # Los anuncios nuevos van los primeros (orden "más recientes"). Solo se cuentan los que
        # están antes del primer anuncio ya visto; así, si se vende uno y sube otro viejo desde
        # el final de la lista, no se confunde con un anuncio nuevo.
        new = []
        for it in items:
            if it["id"] in seen:
                break
            new.append(it)
        state[url] = ids + [x for x in state[url] if x not in cur]
        status[url] = f"ok, {len(new)} nuevos"
        if len(new) > MAX_ALERT_BATCH:
            print(f"   ⚠ {len(new)} 'nuevos' de golpe: lectura anómala, se ignora")
            status[url] = "lectura anómala (ignorada)"
            new = []

    # pendientes: si la nota ya no tiene aviso (la borraste), se empieza limpio
    pend = pending.get(pkey, []) if flagged else []
    kept = []
    for p in pend:
        if p["id"] in cur:
            p["miss"] = 0
            kept.append(p)
        else:
            p["miss"] = p.get("miss", 0) + 1  # vendido/retirado: se quita al repetirse
            if p["miss"] < MISS_TO_DROP:
                kept.append(p)
            else:
                print(f"   ✖ vendido/retirado: {p['label'][:40]}")
    pend = kept
    have = {p["id"] for p in pend}
    pend = [dict(id=it["id"], label=it["label"], price=it.get("price", ""), href=it["href"],
                 ts=now, miss=0) for it in new if it["id"] not in have] + pend
    if new:
        print(f"   🆕 {len(new)} nuevos")
    if pend:
        pending[pkey] = pend
    else:
        pending.pop(pkey, None)

    if pend:
        note = build_note(pend)
        color = alert_color(len(pend))
        if note != ln["note"] or ln["bg"] != hex_color(color):
            if ckey not in colors:  # recordar el color original una sola vez
                colors[ckey] = "" if ln["bg"] in alert_hexes() else ln["bg"]
            requests_.append({"updateCells": {
                "range": cell_range(ln),
                "rows": [{"values": [{"note": note,
                                      "userEnteredFormat": {"backgroundColor": color}}]}],
                "fields": "note,userEnteredFormat.backgroundColor"}})
    elif flagged:  # ya no queda nada a la venta: limpiar nota y devolver el color
        orig = colors.pop(ckey, "")
        value = {"userEnteredFormat": {"backgroundColor": hex_to_rgb(orig)}} if orig else {}
        requests_.append({"updateCells": {
            "range": cell_range(ln), "rows": [{"values": [value]}],
            "fields": "note,userEnteredFormat.backgroundColor"}})
        print("   ✔ todo vendido: aviso retirado")
    return len(new)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="no escribe en la hoja")
    ap.add_argument("--limit", type=int, default=0, help="solo los N primeros enlaces")
    ap.add_argument("--url", help="probar una sola URL de Mercari sin tocar Google Sheets")
    args = ap.parse_args()

    from playwright.sync_api import sync_playwright

    if args.url:
        links = [dict(url=args.url, title="-", row=0, col=0, sheetId=0, note="", bg="")]
        sh = state_ws = colors_ws = pending_ws = None
        state, colors, pending = {}, {}, {}
    else:
        sh = open_sheet()
        only = [s for s in os.environ.get("SHEETS", "").split("|") if s]
        links = collect_links(sh, only)
        state_ws, state = load_state(sh)
        colors_ws, colors = load_colors(sh)
        pending_ws, pending = load_pending(sh)
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
                print(f"{tag}: {len(items)} anuncios"
                      + (f" · ej: {items[0]['label'][:50]}" if items else ""))
                process_link(ln, url, items, state, pending, colors, status, requests_)
            time.sleep(random.uniform(2.5, 5.5))
        browser.close()

    if sh and not args.dry_run:
        if requests_:
            sh.batch_update({"requests": requests_})
        save_state(state_ws, state, status)
        save_colors(colors_ws, colors)
        save_pending(pending_ws, pending)
    print(f"Hecho: {len(requests_)} celdas actualizadas, {failures} fallos de {len(links)}")
    if links and failures > len(links) / 2:
        sys.exit(1)  # GitHub te mandará un email de "workflow failed"


if __name__ == "__main__":
    main()
