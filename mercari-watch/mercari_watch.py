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
import argparse, base64, json, os, random, re, sys, time
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
MAX_LISTED = 15
SHOW_STATE = os.environ.get("SHOW_STATE", "1") == "1"  # 1 = línea de estado (IA/descripción) entre precio y %; 0 = quitarla
YEN_TO_EUR = float(os.environ.get("YEN_TO_EUR", "0.006"))   # 1000 ¥ = 6 €
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-flash-latest")
AI_MAX_PER_RUN = int(os.environ.get("AI_MAX_PER_RUN", "40"))  # tope de análisis por ejecución
AI_DELAY = 4.5  # segundos entre llamadas (respeta el límite por minuto de la capa gratuita)
AI_MAX_IMAGES = 3
DESC_MAX_PER_RUN = int(os.environ.get("DESC_MAX_PER_RUN", "60"))  # anuncios a leer por descripción (sin IA) por ejecución
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
              "hyperlink,note,formattedValue,effectiveValue/numberValue,userEnteredFormat/backgroundColor,userEnteredValue/formulaValue,"
              "textFormatRuns/format/link))))")
    meta = sh.fetch_sheet_metadata(params={"includeGridData": "true", "fields": fields})
    links = []
    for s in meta["sheets"]:
        title = s["properties"]["title"]
        if title in (STATE_TITLE, COLORS_TITLE, PENDING_TITLE) or (only and title not in only):
            continue
        for d in s.get("data", []):
            r0, c0 = d.get("startRow", 0), d.get("startColumn", 0)
            price_col = None  # columna cuyo encabezado (fila 1) contiene "PRECIO"
            if r0 == 0 and d.get("rowData"):
                for c, cell in enumerate(d["rowData"][0].get("values", [])):
                    if "PRECIO" in (cell.get("formattedValue") or "").upper():
                        price_col = c
                        break
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
                        vals = row.get("values", [])
                        ref = None
                        if price_col is not None and price_col < len(vals):
                            ref = vals[price_col].get("effectiveValue", {}).get("numberValue")
                        links.append(dict(sheetId=s["properties"]["sheetId"], title=title, ref=ref,
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
def yen_to_int(price):
    try:
        return int(re.sub(r"[^\d]", "", price or "") or 0) or None
    except ValueError:
        return None


def fmt_eur(yen):
    if not yen:
        return "precio ?"
    e = yen * YEN_TO_EUR
    return f"{e:.0f} €" if e >= 10 else f"{e:.1f} €".replace(".", ",")


def build_note(pend, ref=None):
    lines = [f"🆕 {len(pend)} anuncio(s) a la venta sin revisar"]
    for p in sorted(pend, key=lambda p: p.get("pos") or 999)[:MAX_LISTED]:
        yen = yen_to_int(p.get("price"))
        seen_at = f" · visto {p['ts']}" if p.get("ts") else ""
        lines.append((f"• Anuncio {p['pos']}" if p.get("pos") else f"• {(p.get('label') or p['id'])[:40]}") + seen_at)
        lines.append(f"  {fmt_eur(yen)}")
        if SHOW_STATE:
            ai = p.get("ai")
            if isinstance(ai, dict) and (ai.get("nota") is not None or ai.get("resumen")):
                lines.append(f"  IA: {ai['nota']}/10 · {ai.get('resumen', '')}" if ai.get("nota") is not None
                             else f"  IA: {ai['resumen']}")
            elif isinstance(p.get("desc"), dict):
                d = p["desc"]
                lines.append(f"  Descripción: ~{d['nota']}/10 · {d['resumen']}" if d.get("nota") is not None
                             else f"  Descripción: {d.get('resumen', '')}")
        lines.append(f"  {yen * YEN_TO_EUR / ref * 100:.0f}% de mi ref NM" if yen and ref else "  sin ref NM")
    if len(pend) > MAX_LISTED:
        lines.append(f"… y {len(pend) - MAX_LISTED} más")
    return "\n".join(lines)[:MAX_NOTE]


# ---------------------------------------------------------------- Estado por descripción (sin IA)
FIELD_BASE = {  # campo "商品の状態" de Mercari -> (nota, texto)
    "新品、未使用": (10, "nueva/sin usar"),
    "未使用に近い": (9, "casi sin usar"),
    "目立った傷や汚れなし": (8, "sin daños visibles"),
    "やや傷や汚れあり": (6, "algo dañada"),
    "傷や汚れあり": (4, "dañada"),
    "全体的に状態が悪い": (2, "mal estado general"),
}
FIELD_RE = re.compile("|".join(sorted(map(re.escape, FIELD_BASE), key=len, reverse=True)))
NEG_RULES = [  # (regex, penalización, etiqueta en español)
    (r"白欠け|白かけ|白カケ|フチ白|縁白|エッジ傷", 1, "bordes blanqueados"),
    (r"角(潰れ|つぶれ|ツブレ|傷|スレ|折れ)", 1, "esquinas dañadas"),
    (r"折れ|折り目|折目|曲がり|反り|シワ|しわ|皺|ヨレ|歪み", 2, "dobleces/arrugas"),
    (r"擦れ|スレ|こすれ", 1, "roces"),
    (r"キズ|傷|きず|スクラッチ|ひっかき|引っかき", 1, "arañazos"),
    (r"汚れ|シミ|しみ|黄ばみ|ヤケ|日焼け|変色", 1, "manchas/amarilleo"),
    (r"凹み|ヘコミ|へこみ|打痕", 1, "abolladuras"),
    (r"水濡れ|ペン書き|落書き|破れ|破損|ダメージ|剥がれ|ベタつき", 2, "daños graves"),
    (r"ジャンク|状態悪|プレイ用|プレー用|ボロボロ", 3, "muy jugada"),
]
POS_RE = re.compile(r"美品|極美品|綺麗|きれい|キレイ")
_NEG_WORDS = r"(?:傷|キズ|きず|汚れ|スレ|擦れ|白欠け|白かけ|折れ|シワ|しわ|ヨレ|角潰れ|ダメージ|破れ|凹み|シミ|ヤケ)"
_NO = r"(?:は|等は|などは|も)?\s*(?:ありません|ございません|なし|無し|ないです|ない|見られません|見当たりません|気になりません|ほぼありません)"
NEGATED_RES = [
    re.compile(r"(?:目立った)?(?:傷|キズ)(?:や|・|、|と)(?:汚れ|スレ|擦れ)" + _NO),
    re.compile(_NEG_WORDS + _NO),
]
GRADE_RE = re.compile(r"(PSA|BGS|ARS|CGC)\s*(\d+(?:\.\d)?)", re.I)


def description_region(text):
    i = text.find("商品の説明")
    if i < 0:
        return text
    j = text.find("商品の情報", i)
    return text[i:j] if j > i else text[i:i + 1500]


def text_condition(text, label=""):
    """Estado aproximado a partir del campo de estado de Mercari y las palabras de la descripción.
    Devuelve {"nota": int|None, "resumen": str}."""
    if not text:
        return None
    field = FIELD_RE.search(text)
    base, field_es = FIELD_BASE[field.group(0)] if field else (None, None)
    body = FIELD_RE.sub(" ", description_region(text))
    g = GRADE_RE.search(label + " " + body)
    if g:
        n = min(10, max(1, round(float(g.group(2)))))
        return {"nota": n, "resumen": f"gradada {g.group(1).upper()} {g.group(2)}"}
    if re.search(r"まとめ売り|\d+枚セット|大量|まとめて", label + " " + body):
        return {"nota": None, "resumen": "parece un lote; revisa las fotos"}
    clean = body
    for rx in NEGATED_RES:
        clean = rx.sub(" ", clean)
    flags, penalty = [], 0
    for rx, w, name in NEG_RULES:
        if re.search(rx, clean):
            flags.append(name)
            penalty += w
    bonus = 1 if POS_RE.search(clean) else 0
    if base is None and not flags and not bonus:
        return {"nota": None, "resumen": "sin datos de estado en la descripción"}
    # con estado del vendedor, la descripción solo matiza (media penalización): ya cuenta en su base
    pen = min(penalty, 5) if base is None else (min(penalty, 5) + 1) // 2
    score = (base if base is not None else 7) + bonus - pen
    score = max(1, min(10, score))
    parts = []
    if field_es:
        parts.append(f"declarada {field_es}")
    if flags:
        parts.append("menciona " + ", ".join(flags[:3]))
    elif bonus:
        parts.append("descripción positiva")
    return {"nota": score, "resumen": "; ".join(parts) or "estimación por descripción"}


# ---------------------------------------------------------------- IA (Gemini)
AI_PROMPT = (
    "Eres un experto en cartas Pokémon TCG japonesas vintage. Evalúa el ESTADO FÍSICO de la carta "
    "a partir de las fotos de un anuncio de Mercari y del texto (en japonés). Escala: 10 = perfecta "
    "(Near Mint), 7-8 = ligeros defectos, 5-6 = desgaste visible, 3-4 = muy jugada, 1-2 = destrozada. "
    "Fíjate en bordes blanqueados, esquinas, arrugas o dobleces, arañazos en el holo, manchas, "
    "humedad, deformación o escritura. Si el anuncio es un lote o las fotos no permiten juzgar, "
    "usa nota null y explícalo. Responde SOLO con JSON: "
    '{"nota": <entero 1-10 o null>, "resumen": "<máximo 14 palabras en español>"}.'
)


class AIBudget:
    def __init__(self, n, desc=None):
        self.left = n
        self.desc_left = DESC_MAX_PER_RUN if desc is None else desc
        self.blocked = False  # se activa si Gemini devuelve límite alcanzado (429)


def item_context(page, href):
    """Abre el anuncio y devuelve (texto visible, [urls de fotos])."""
    try:
        page.goto(href, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_selector('img[src*="mercdn"]', timeout=15000)
    except Exception:
        return "", []
    page.wait_for_timeout(800)
    text = page.inner_text("body")[:4000]
    srcs = page.evaluate("""() => {
        const id = location.pathname.split('/').pop();
        return [...document.querySelectorAll('img')]
          .map(i => i.currentSrc || i.src || '')
          .filter(s => s.includes('mercdn') && s.includes(id));
    }""")
    by_idx = {}
    for u in srcs:
        m = re.search(r"_(\d+)\.(?:jpe?g|png|webp)", u)
        if not m:
            continue
        k = int(m.group(1))
        if k not in by_idx or ("/orig/" in u and "/orig/" not in by_idx[k]):
            by_idx[k] = u
    return text, [by_idx[k] for k in sorted(by_idx)][:AI_MAX_IMAGES]


def gemini_assess(text, image_urls, label):
    import requests
    parts = [{"text": f"{AI_PROMPT}\n\nTítulo: {label}\nTexto del anuncio:\n{text[:1800]}"}]
    for u in image_urls:
        try:
            r = requests.get(u, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
            if r.ok and len(r.content) < 5_000_000:
                parts.append({"inline_data": {"mime_type": r.headers.get("content-type", "image/jpeg").split(";")[0],
                                              "data": base64.b64encode(r.content).decode()}})
        except Exception:
            pass
    if len(parts) == 1:
        return {"nota": None, "resumen": "sin fotos legibles"}
    resp = requests.post(
        f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent",
        headers={"x-goog-api-key": os.environ["GEMINI_API_KEY"], "Content-Type": "application/json"},
        json={"contents": [{"parts": parts}],
              "generationConfig": {"responseMimeType": "application/json", "temperature": 0.2}},
        timeout=90)
    if resp.status_code == 429:
        raise RuntimeError("limit")
    resp.raise_for_status()
    out = json.loads(resp.json()["candidates"][0]["content"]["parts"][0]["text"])
    nota = out.get("nota")
    return {"nota": int(nota) if isinstance(nota, (int, float)) else None,
            "resumen": str(out.get("resumen", ""))[:120]}


def enrich_items(page, items, budget):
    """Para cada anuncio pendiente sin análisis: intenta la IA (si hay clave y cupo); si no, estima el
    estado leyendo la descripción (gratis)."""
    has_key = bool(os.environ.get("GEMINI_API_KEY"))
    for p in items:
        if p.get("ai") is not None:
            continue
        can_ai = has_key and budget.left > 0 and not budget.blocked
        need_desc = p.get("desc") is None and budget.desc_left > 0
        if not can_ai and not need_desc:
            continue
        text, urls = item_context(page, p["href"])
        if can_ai:
            try:
                p["ai"] = gemini_assess(text, urls, p["label"])
                print(f"   IA: {p['label'][:30]} → {p['ai']}")
            except RuntimeError:
                budget.blocked = True
                print("   ⚠ Gemini: límite alcanzado; el resto usa la descripción y se analiza luego")
            except Exception as e:
                print(f"   ⚠ IA falló ({type(e).__name__}); se reintentará")
            budget.left -= 1
            time.sleep(AI_DELAY)
        if p.get("ai") is None and p.get("desc") is None and budget.desc_left > 0:
            p["desc"] = text_condition(text, p["label"])
            budget.desc_left -= 1
            if p["desc"]:
                print(f"   Desc: {p['label'][:30]} → {p['desc']}")
            if not can_ai:
                time.sleep(1.5)


def cell_range(ln):
    return {"sheetId": ln["sheetId"], "startRowIndex": ln["row"], "endRowIndex": ln["row"] + 1,
            "startColumnIndex": ln["col"], "endColumnIndex": ln["col"] + 1}


def process_link(ln, url, items, state, pending, colors, status, requests_, now=None, enrich=None):
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
    posmap = {it["id"]: i + 1 for i, it in enumerate(items)}  # puesto en la lista de la búsqueda
    for p in pend:
        p["pos"] = posmap.get(p["id"], p.get("pos"))
    pend.sort(key=lambda p: p.get("pos") or 999)
    if pend:
        pending[pkey] = pend
    else:
        pending.pop(pkey, None)

    if pend and enrich:
        enrich(pend)
    if pend:
        note = build_note(pend, ln.get("ref"))
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
        ai_page = ctx.new_page()
        budget = AIBudget(AI_MAX_PER_RUN)

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
                process_link(ln, url, items, state, pending, colors, status, requests_,
                             enrich=(lambda pend: enrich_items(ai_page, pend, budget)) if SHOW_STATE else None)
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
