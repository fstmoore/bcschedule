#!/usr/bin/env python3
"""Sheet -> ICS. Stdlib only. Usage: python3 schedule_to_ics.py --group 1Д-22 --start 2026-09-14 --weeks 4 -o out.ics"""
import argparse, glob, html, io, os, re, urllib.request, urllib.parse, uuid, datetime, zipfile
import xml.etree.ElementTree as ET

SHEET_ID = "1SXdz3k3Ect865_IIL3vm-Ia1LvNhK3ls"
MOODLE_URL = "http://78.137.2.119:2929/mod/forum/discuss.php?d=45"
BASE = "https://schdl.eu.cc"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# bells: https://drive.google.com/file/d/1piTUM_2_NWuXF9ZYWc0IbIno0exC-ArY/view (sheet pair N = bell N)
TIMES = {0: ("08:00", "08:50"), 1: ("09:00", "10:20"), 2: ("10:30", "11:50"), 3: ("12:00", "13:20"),
         4: ("13:40", "15:00"), 5: ("15:10", "16:30"), 6: ("16:40", "18:00"), 7: ("18:10", "19:00"),
         # ponytail: bells PDF ends at VII; 8/9 extrapolated +80min, fix when college publishes them
         8: ("19:10", "20:30"), 9: ("20:40", "22:00")}

def _colletters(ref):
    col = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n - 1

def _numstr(v):
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else v
    except ValueError:
        return v

def _sheet_grid(z, path, shared):
    root = ET.fromstring(z.read(path))
    cells = {}
    for c in root.iter(NS + "c"):
        ref = c.get("r")
        r = int(re.search(r"\d+", ref).group(0)) - 1
        col = _colletters(ref)
        t = c.get("t")
        if t == "inlineStr":
            v = "".join(x.text or "" for x in c.iter(NS + "t"))
        else:
            v = c.findtext(NS + "v") or ""
            v = shared[int(v)] if t == "s" and v.isdigit() else _numstr(v)
        if v:
            cells[(r, col)] = v
    if not cells:
        return []
    maxr = max(r for r, _ in cells)
    maxc = max(c for _, c in cells)
    return [[cells.get((r, c), "") for c in range(maxc + 1)] for r in range(maxr + 1)]

def fetch_sheets(sheet_id=SHEET_ID):
    """(name, grid) per visible sheet, via xlsx export (csv export only gives the first tab)."""
    raw = urllib.request.urlopen(
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx", timeout=60).read()
    z = zipfile.ZipFile(io.BytesIO(raw))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    shared = ["".join(x.text or "" for x in si.iter(NS + "t"))
              for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(NS + "si")]
    out = []
    for s in wb.iter(NS + "sheet"):
        if s.get("state") != "visible":
            continue
        path = "xl/" + rels[s.get(REL + "id")].split("xl/")[-1]
        out.append((s.get("name"), _sheet_grid(z, path, shared)))
    return out

def _ngrp(g):
    return re.sub(r"[\s\-–—]+", "", g).upper()

def _cell_lines(cell):
    lines = []
    for p in re.split(r"<br[^>]*>|<p[^>]*>|</p>", cell):
        t = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", p))).strip()
        if t:
            lines.append(t)
    return lines

def fetch_replacements(moodle=MOODLE_URL):
    """Moodle post -> google doc IDs -> {(date, GROUP, pair): (subject, teacher, room, cancelled)}. Never raises."""
    try:
        page = urllib.request.urlopen(moodle, timeout=30).read().decode("utf-8", "replace")
        ids = list(dict.fromkeys(re.findall(r"docs\.google\.com/document/d/([\w\-]+)", page)))
        out = {}
        for doc in ids:
            raw = urllib.request.urlopen(
                f"https://docs.google.com/document/d/{doc}/export?format=html", timeout=60).read().decode("utf-8", "replace")
            tables = re.findall(r"<table.*?</table>", raw, re.S)
            dates = [f"{y}-{m}-{d}" for d, m, y in re.findall(r"(\d{2})\.(\d{2})\.(\d{4})", raw)]
            for i, t in enumerate(tables):
                if i >= len(dates):
                    break
                for r in re.findall(r"<tr.*?</tr>", t, re.S)[1:]:  # skip header
                    cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", r, re.S)
                    if len(cells) < 5:
                        continue
                    cls = [_cell_lines(c) for c in cells[:5]]
                    groups = [p for ln in cls[0] for p in re.split(r"[,;\s]+", ln) if p]
                    pairs = [int(x) for x in re.findall(r"\d+", " ".join(cls[1]))]
                    subj = " / ".join(cls[2])
                    if not groups or not pairs or not subj or "за розкладом" in subj.lower():
                        continue
                    teach, room = ", ".join(cls[3]), ", ".join(cls[4])
                    # ponytail: multi-group rows join teachers/rooms; per-group mapping when college publishes it
                    for g in groups:
                        for p in pairs:
                            out[(dates[i], _ngrp(g), p)] = (subj, teach, room, "відмін" in subj.lower())
        return out
    except Exception:
        return {}  # ponytail: regen without replacements beats no regen; alert when this happens often

def apply_replacements(ics, file_group, repl):
    """Replace/cancel/add bell slots per replacements. Bell N = TIMES[N], absolute dates."""
    fg = _ngrp(file_group)
    mine = [((d, p), v) for (d, gn, p), v in repl.items()
            if p in TIMES and (fg == gn or fg.startswith(gn) or gn.startswith(fg))]
    if not mine:
        return ics
    body, tail = ics.rsplit("END:VCALENDAR", 1)
    head, *events = body.split("BEGIN:VEVENT")
    drop = {f"DTSTART:{d.replace('-', '')}T{TIMES[p][0].replace(':', '')}00" for (d, p), _ in mine}
    kept = "".join("BEGIN:VEVENT" + e for e in events if not any(x in e for x in drop))
    new = []
    for (d, p), (subj, teach, room, cancelled) in mine:
        if cancelled:
            continue
        s, e = TIMES[p]
        ymd = d.replace("-", "")
        uid = uuid.uuid5(uuid.NAMESPACE_URL, f"zamini|{file_group}|{d}|{p}|{subj}")
        desc = ", ".join(x for x in (teach, f"ауд. {room}" if room else "") if x)
        new.append("\r\n".join(["BEGIN:VEVENT", f"UID:{uid}@sheet",
            f"DTSTAMP:{datetime.datetime.now(datetime.timezone.utc):%Y%m%dT%H%M%SZ}",
            f"DTSTART:{ymd}T{s.replace(':', '')}00", f"DTEND:{ymd}T{e.replace(':', '')}00",
            f"SUMMARY:{subj} (заміна)", f"DESCRIPTION:{desc}", f"LOCATION:{room}", "END:VEVENT"]))
    return head + kept + ("\r\n".join(new) + "\r\n" if new else "") + "END:VCALENDAR" + tail

def sheet_groups(grid):
    """Header cells may hold several groups sharing one block ('1Т-21 3Т-22')."""
    groups = []
    for c in grid[0] if grid else []:
        for part in c.split():
            if part not in groups:
                groups.append(part)
    return groups

def parse_group(rows, group):
    g = next((i for i, c in enumerate(rows[0]) if c.strip() == group), None)
    if g is None:  # fallback: substring match
        g = next(i for i, c in enumerate(rows[0]) if group in c)
    g //= 6; g *= 6
    subj_c, num_c = g + 2, g + 1
    # ponytail: day boundary = pair numbers wrapping down in THIS group's column;
    # the old B-col '1' breaks when groups run different pairs per day
    day_of, day, prev = {}, -1, None
    for r in range(len(rows)):
        v = rows[r][num_c].strip() if num_c < len(rows[r]) else ""
        if v.isdigit():
            n = int(v)
            if day < 0:
                day = 0
            elif prev is not None and n < prev:
                day += 1
            prev = n
        day_of[r] = day
    lessons = []
    for r in range(1, len(rows)):
        row = rows[r]
        def cell(c):
            return row[c].strip() if c < len(row) else ""
        day = day_of[r]
        if cell(num_c).isdigit():
            n = int(cell(num_c))
            ROOM = r"\d{2,4}[а-яА-Яa-zA-Z]?|СЗ"
            def subj_at(srow, c):
                return rows[srow][c].strip() if 0 <= srow < len(rows) and c < len(rows[srow]) else ""
            def teach_at(c, rws):  # '.' = teacher; rooms/subjects never contain it... except 'Англ. мова' — hence explicit rows
                for rr in rws:
                    v = rows[rr][c].strip() if 0 <= rr < len(rows) and c < len(rows[rr]) else ""
                    if "." in v:
                        return v
                return ""
            def find_sides(srow):
                main = subj_at(srow, subj_c)
                if len(main) > 2 and main not in ("А", "Б") and not re.fullmatch(ROOM, main):
                    return [(subj_c, "", "")]
                sides = []  # A/B split: subjects sit in sub-cols beside the А/Б markers; '-' = no lesson
                for c, tag, sub in ((subj_c + 1, " (А)", "А"), (subj_c + 3, " (Б)", "Б"),
                                    (subj_c - 1, " (А)", "А"), (subj_c + 2, " (Б)", "Б")):
                    v = subj_at(srow, c)
                    if c != num_c and len(v) > 2 and not re.fullmatch(ROOM, v):
                        sides.append((c, tag, sub))
                        if len(sides) == 2:
                            break
                return sides
            # new sheet: alt weeks are stacked halves (r-1 = чис., r+2 = знам.) —
            # unless a number row sits between (then it's the next pair, not знам.);
            # old sheet: both variants in one cell (multiline) — still handled below
            halves = []  # (alt, sub, col, tag, teacher, subjects)
            for srow, trows, alt in ((r - 1, (r, r + 1), 0), (r + 2, (r + 3, r + 4), 1)):
                if alt == 1 and any((rows[rr][num_c].strip() if 0 <= rr < len(rows) and num_c < len(rows[rr]) else "").isdigit()
                                    for rr in (r + 1, r + 2, r + 3, r + 4)):
                    continue
                for c, tag, sub in find_sides(srow):
                    t = teach_at(c, trows)
                    S = [x.strip() for x in subj_at(srow, c).split("\n") if x.strip()]
                    halves.append((alt, sub, c, tag, t, S))
            def room_in(c, rws):
                for rr in rws:
                    m = [ln.strip() for ln in subj_at(rr, c).split("\n")
                         if re.fullmatch(ROOM, ln.strip())]
                    if m:
                        return "\n".join(m)
                return ""
            chis = [h for h in halves if h[0] == 0]
            znam = [h for h in halves if h[0] == 1]
            rooms = {}
            if not znam:  # single pair: room may sit anywhere below (old layout too)
                for _, _, c, _, _, _ in chis:
                    rooms[(0, c)] = room_in(c, (r + 1, r + 2, r + 3, r + 4))
            else:  # split pair: rooms live inside their own half; share across when missing
                for _, _, c, _, _, _ in chis:
                    rooms[(0, c)] = room_in(c, (r + 1, r, r + 4))
                for _, _, c, _, _, _ in znam:
                    rooms[(1, c)] = room_in(c, (r + 3, r + 4))
                for _, _, c, _, _, _ in chis:
                    rooms[(0, c)] = rooms[(0, c)] or rooms.get((1, c), "")
                for _, _, c, _, _, _ in znam:
                    rooms[(1, c)] = rooms[(1, c)] or rooms.get((0, c), "")
            tchis = {h[2]: h[4] for h in chis}
            zkeys = {(h[1], h[2], h[4] or tchis.get(h[2], "")) for h in znam}
            ckeys = {(h[1], h[2], h[4]) for h in chis}
            if not znam:
                for _, sub, c, tag, t, S in chis:
                    room = rooms[(0, c)]
                    if len(S) < 2:
                        lessons.append({"day": day, "pair": n, "sub": sub, "subject": S[0] + tag,
                                        "teacher": t, "room": room})
                        continue
                    T = [x.strip() for x in t.split("\n") if x.strip()] or [""]
                    R = [x.strip() for x in room.split("\n") if x.strip()] or [""]
                    for i, suffix in ((0, " (чис.)"), (1, " (знам.)")):
                        lessons.append({"day": day, "pair": n, "alt": i, "sub": sub,
                                        "subject": S[i] + tag + suffix,
                                        "teacher": T[i] if i < len(T) else T[0],
                                        "room": R[i] if i < len(R) else R[0]})
                continue
            for _, sub, c, tag, t, S in chis:
                room = rooms[(0, c)]
                if (sub, S[0] + tag, t) in zkeys:
                    lessons.append({"day": day, "pair": n, "sub": sub, "subject": S[0] + tag,
                                    "teacher": t, "room": room})
                else:
                    lessons.append({"day": day, "pair": n, "alt": 0, "sub": sub,
                                    "subject": S[0] + tag + " (чис.)", "teacher": t, "room": room})
            for _, sub, c, tag, t, S in znam:
                t = t or tchis.get(c, "")  # знам. teacher missing -> чис. teacher, like the old split
                room = rooms[(1, c)]
                if (sub, S[0] + tag, t) not in ckeys:
                    lessons.append({"day": day, "pair": n, "alt": 1, "sub": sub,
                                    "subject": S[0] + tag + " (знам.)", "teacher": t, "room": room})
    return lessons

def to_ics(lessons, start_monday, weeks=1, group="", flip_weeks=False, sub=""):
    out = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//ics//sheet//EN"]
    for w in range(weeks):
        for L in lessons:
            if L["day"] < 0 or L["pair"] not in TIMES:
                continue
            if sub and L.get("sub", "") not in ("", sub):
                continue
            d = start_monday + datetime.timedelta(days=L["day"] + 7 * w)
            # ponytail: alt parity from absolute ISO week (not w) so weekly regens don't flip weeks
            if L.get("alt") is not None and (d.isocalendar()[1] % 2 == 1) != bool(L["alt"] ^ flip_weeks):
                continue
            s, e = TIMES[L["pair"]]
            dt = lambda dd, t: dd.strftime("%Y%m%d") + "T" + t.replace(":", "") + "00"
            # ponytail: uuid5 (not uuid4) so regen updates events instead of duplicating them
            uid = uuid.uuid5(uuid.NAMESPACE_URL, f"{group}|{sub}|{d}|{L['pair']}|{L['subject']}")
            subject = L["subject"].removesuffix(f" ({sub})") if sub else L["subject"]
            desc = ", ".join(x for x in (L["teacher"], f"ауд. {L['room']}" if L["room"] else "") if x)
            out += ["BEGIN:VEVENT", f"UID:{uid}@sheet", f"DTSTAMP:{datetime.datetime.now(datetime.timezone.utc):%Y%m%dT%H%M%SZ}",
                    f"DTSTART:{dt(d, s)}", f"DTEND:{dt(d, e)}", f"SUMMARY:{subject}",
                    f"DESCRIPTION:{desc}", f"LOCATION:{L['room']}", "END:VEVENT"]
    return "\r\n".join(out + ["END:VCALENDAR"]) + "\r\n"

_HERE = os.path.dirname(os.path.abspath(__file__))
def _asset(name):
    with open(os.path.join(_HERE, name), encoding="utf-8") as f:
        return f.read()
_CSS = _asset("style.css")
_JS = _asset("app.js")  # preview logic, including the ics parser



def render_index(pages):
    """One static page: pick your group, get its .ics link. No deps, no build step."""
    import html
    sheets = []
    for sheet, label, fname in pages:
        if not sheets or sheets[-1][0] != sheet:
            sheets.append((sheet, []))
        sheets[-1][1].append((label, fname))
    chips = ['<button class="chip on" data-s="">Усі</button>'] + [
        f'<button class="chip" data-s="{html.escape(s.strip())}">{html.escape(s.strip())}</button>'
        for s, _ in sheets]
    def card(sheet, label, fname, i):
        enc = urllib.parse.quote(fname)  # Cyrillic filenames percent-encoded so hrefs are valid URLs
        absu = f"{BASE}/calendars/{enc}"
        return (
            f'<div class="card" style="animation-delay:{min(i * 20, 400)}ms" data-g="{html.escape(label.lower())}" data-s="{html.escape(sheet.strip())}">'
            f'<span class="t">{html.escape(label)}</span><span class="row">'
            f'<a class="btn fill" href="https://calendar.google.com/calendar/r?cid={urllib.parse.quote(absu, safe="")}" data-gcal="calendars/{html.escape(fname)}">В Google-календар</a>'
            f'<a class="btn tonal" href="{absu.replace("https://", "webcal://")}" data-sub="calendars/{html.escape(fname)}">Інший календар</a>'
            f'<button class="btn out" data-link="calendars/{html.escape(fname)}">Лінк</button>'
            f'</span></div>')
    secs = "".join(
        f"<h2>{html.escape(s.strip())}</h2><div class='grid'>"
        + "".join(card(s, label, fname, i) for i, (label, fname) in enumerate(items)) + "</div>"
        for s, items in sheets)
    ts = datetime.datetime.now(datetime.timezone.utc).strftime("%d.%m.%Y %H:%M")
    return f"""<!doctype html><html lang="uk"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Розклад, який живе в календарі</title>
<link rel="icon" type="image/svg+xml" href="favicon.svg">
<meta name="description" content="Розклад занять ЧДБК — Черкаського державного фахового бізнес-коледжу (ЧДФБК) — у вигляді календаря: обери групу, додай у Google-календар або будь-який інший. Оновлюється кожні 6 годин. Неофіційний проєкт, не афілійований з коледжем.">
<meta name="keywords" content="розклад, ЧДБК, ЧДФБК, Черкаський державний фаховий бізнес-коледж, розклад занять, календар, групи, пари">
<meta name="theme-color" content="#edf2e6">
<meta property="og:type" content="website">
<meta property="og:title" content="Нормальний розклад, який живе у календарі">
<meta property="og:description" content="Обирай групу — пари самі прийдуть у твій календар. Без гугл-таблички.">
<script data-goatcounter="https://bcschedule.goatcounter.com/count" async src="//gc.zgo.at/count.js"></script>
<meta name="google-site-verification" content="uGk621K82j_O0Twhiou4c1LqOzossy-005O0nijx-Ig" />
<style>{_CSS}</style></head><body><div class="wrap">
<div class="hero"><h1>Нормальний розклад, <br>який живе у календарі</h1>
<svg class="vine" viewBox="0 0 420 34" aria-hidden="true">
<path d="M4 26 C 90 6, 180 30, 260 14 S 380 10, 416 20"/>
<ellipse class="l1" cx="120" cy="14" rx="11" ry="5" transform="rotate(-24 120 14)"/>
<ellipse class="l2" cx="250" cy="20" rx="11" ry="5" transform="rotate(18 250 20)"/>
<ellipse class="l3" cx="350" cy="13" rx="11" ry="5" transform="rotate(-18 350 13)"/>
</svg>
<p class="sub">Оновлюється сам, враховує заміни. Гудбай гугл табличка</p>
<p class="fine"><b>Конфіденційність: нам начхати на твої дані.</b> У прямому сенсі.
Ця сторінка майже нічого не збирає: один безкуковий лічильник переглядів (GoatCounter) —
без імен, пошт та ідентифікаторів. Немає акаунтів, трекерів або сервера, з якого можна злити базу. Єдине, що відбувається, —
твій телефон качає статичний файл з парами. Параноїш — відкрий .ics блокнотом, там лише пари.</p>
<p class="fine"><b>Умови: ми не всевидющі, а ти — дорослий.</b> Розклад береться
з гугл-таблички коледжу, а отже, містить їхні помилки: перенесення, скасування, раптові
«дивіться оновлення». Ми не сидимо на їхніх нарадах (та й на пара нечасто) і фізично не знаємо про все, що відбувається.
Перед важливими парами звіряйся з офіційним розкладом. Прогуляв пару, завалив сесію,
відрахували — твої проблеми, не наші. Ми лише переклали табличку в календар.</p></div>
<div class="search"><input id="q" placeholder="Знайди свою групу…" autocomplete="off"></div>
<div class="chips">{"".join(chips)}</div>
<div id="secs">{secs}</div>
<dialog id="how"><h3>Як додати розклад</h3>
<p>Найпростіше — кнопки вище: «В Google-календар» або «Інший календар».
А якщо треба вручну — ось лінк:</p>
<div class="row"><input id="lk" readonly><button class="btn fill" id="cp">Копіювати</button>
<button class="btn out" id="nope">Закрити</button></div>
<p>iPhone: Параметри → Календар → Облікові записи → Додати передплачений календар.
Google: Календар → Інші календарі → Додати за URL. Android: тільки через Google.</p></dialog>
<dialog id="pvw"><h3 id="pname"></h3>
<div class="chips"><button class="chip" id="pch">Чисельник</button><button class="chip" id="pzn">Знаменник</button></div>
<div id="pgrid"></div>
<div class="row"><button class="btn out" id="pclose">Закрити</button></div></dialog>
<footer>Поливаємо кожні 6 годин — розклад сам росте з гугл-таблички.
<br>Vibecoded in 2 hrs without a wage.
<br>Останнє оновлення: {ts} (UTC).
<br>Розклад ЧДБК — Черкаський державний фаховий бізнес-коледж (ЧДФБК).
<br>Неофіційний проєкт, не афілійований з коледжем.</footer>
<script>{_JS}</script></div></body></html>
"""

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--group", required=True)
    ap.add_argument("--start", default="auto", help="Monday date YYYY-MM-DD or 'auto' (this week's Monday)")
    ap.add_argument("--weeks", type=int, default=1)
    ap.add_argument("--flip-weeks", action="store_true", help="swap чисельник/знаменник if line 1 turns out to be знаменник")
    ap.add_argument("-o", default="schedule.ics")
    ap.add_argument("--sheet", default=SHEET_ID)
    a = ap.parse_args()
    monday = datetime.date.today() - datetime.timedelta(days=datetime.date.today().weekday()) \
        if a.start == "auto" else datetime.date.fromisoformat(a.start)
    rows = fetch_sheets(a.sheet)
    jobs = []
    if a.group == "all":
        for name, grid in rows:
            jobs += [(name, g, grid) for g in sheet_groups(grid)]
    else:
        for _, grid in rows:
            try:
                parse_group(grid, a.group)
                jobs = [(None, a.group, grid)]
                break
            except (StopIteration, IndexError):
                continue
        if not jobs:
            raise SystemExit(f"group {a.group} not found on any sheet")
    pages = []
    repl = fetch_replacements()
    for sheet, g, grid in jobs:
        lessons = parse_group(grid, g)
        subs = sorted({L["sub"] for L in lessons if L.get("sub")}) or [""] if a.group == "all" else [""]
        for sub in subs:  # split A/B subgroups into separate files; shared lessons go in both
            out = a.o if a.group != "all" else os.path.join("calendars", f"{g}-{sub}.ics" if sub else f"{g}.ics")
            os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
            open(out, "w", encoding="utf-8").write(to_ics(lessons, monday, a.weeks, g, a.flip_weeks, sub))
            label = f"{g}-{sub}" if sub else g
            if repl:
                txt = apply_replacements(open(out, encoding="utf-8").read(), label, repl)
                open(out, "w", encoding="utf-8").write(txt)
            print(f"{g}{'-' + sub if sub else ''}: {len(lessons)} pairs/week -> {out}")
            if a.group == "all":
                pages.append((sheet, label, os.path.basename(out)))
    if a.group == "all":
        # ponytail: drop stale files (e.g. graduated groups) so dead links never linger
        keep = {os.path.basename(p[2]) for p in pages}
        for f in glob.glob(os.path.join("calendars", "*.ics")):
            if os.path.basename(f) not in keep:
                os.remove(f)
        open("index.html", "w", encoding="utf-8").write(render_index(pages))
        print(f"index.html ({len(pages)} groups)")
        today = datetime.date.today().isoformat()
        # ponytail: sitemap lists only the page; .ics files are text/calendar, not indexable, and poison sitemap quality
        sm = ['<?xml version="1.0" encoding="UTF-8"?>',
              '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
              f"<url><loc>{BASE}/</loc><lastmod>{today}</lastmod></url>",
              "</urlset>"]
        open("sitemap.xml", "w", encoding="utf-8").write("\n".join(sm) + "\n")
        bots = ["GPTBot", "ChatGPT-User", "ClaudeBot", "anthropic-ai", "PerplexityBot", "Bytespider"]
        open("robots.txt", "w", encoding="utf-8").write(
            "User-agent: *\nAllow: /\n\n"
            "# AI agents explicitly welcome: cite freely, subscribe users to their group calendar\n" +
            "".join(f"User-agent: {b}\nAllow: /\n\n" for b in bots) +
            f"Sitemap: {BASE}/sitemap.xml\n")
        open("llms.txt", "w", encoding="ascii").write(
            "# ChDBK Class Schedule\n\n"
            "> Class schedule of Cherkasy State Vocational Business College "
            "(ChDBK/ChDFBK) as subscribable per-group calendars.\n\n"
            "## Use\n\n"
            "- Pick a group on the index page, add its calendar via Google Calendar "
            "or any webcal client. Subscribed calendars auto-update; downloaded .ics files do not.\n"
            "- Calendars live under /calendars/<GROUP>.ics (3-week rolling window, "
            "upper/lower week variants marked).\n"
            "- Source: the college's public Google Sheet, re-exported every 6 hours. "
            "The sheet may contain errors; this mirror adds none and fixes none.\n"
            "- No accounts, no cookies, no tracking besides cookieless GoatCounter page counts.\n")
