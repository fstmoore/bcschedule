#!/usr/bin/env python3
"""Sheet -> ICS. Stdlib only. Usage: python3 schedule_to_ics.py --group 1Д-22 --start 2026-09-14 --weeks 4 -o out.ics"""
import argparse, glob, html, io, os, re, urllib.request, urllib.parse, uuid, datetime, zipfile
import xml.etree.ElementTree as ET

SHEET_ID = "15lyVaUfBsadJLdFmXAYr_eldqLbcBLjc"  # official sheet, via http://78.137.2.119:2929/mod/forum/discuss.php?d=2#p2
MOODLE_URL = "http://78.137.2.119:2929/mod/forum/discuss.php?d=45"
BASE = "https://schdl.eu.cc"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# bells: https://drive.google.com/file/d/1piTUM_2_NWuXF9ZYWc0IbIno0exC-ArY/view (sheet pair N = bell N)
TIMES = {0: ("08:00", "08:50"), 1: ("09:00", "10:00"), 2: ("10:10", "11:10"), 3: ("11:20", "12:20"),
         4: ("12:40", "13:40"), 5: ("13:50", "14:50"), 6: ("15:00", "16:00"), 7: ("16:10", "17:10"),
         # ponytail: bells PDF ends at VII; 8/9 extrapolated +80min, fix when college publishes them
         8: ("17:20", "18:20"), 9: ("18:30", "19:30")}

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
        if s.get("state", "visible") != "visible":  # absent state = visible; the new sheet omits it
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
_JS = _asset("app.js")  # preview logic, including the ics parser



def render_index(pages):
    """One static page: pick your group, get its .ics link. No deps, no build step."""
    import html
    flat = lambda s: re.sub(r"\s+", " ", s.strip())  # display text; data-s keeps raw spacing for filtering
    sheets = []
    for sheet, label, fname in pages:
        if not sheets or sheets[-1][0] != sheet:
            sheets.append((sheet, []))
        sheets[-1][1].append((label, fname))
    chips = ['<button class="chip on" data-s="">Усі</button>'] + [
        f'<button class="chip" data-s="{html.escape(s.strip())}">{html.escape(flat(s))}</button>'
        for s, _ in sheets]
    def card(sheet, label, fname, i):
        enc = urllib.parse.quote(fname)  # Cyrillic filenames percent-encoded so hrefs are valid URLs
        absu = f"{BASE}/calendars/{enc}"
        return (
            f'<div class="card" style="animation-delay:{min(i * 20, 400)}ms" data-g="{html.escape(label.lower())}" data-s="{html.escape(sheet.strip())}">'
            f'<span class="t">{html.escape(label)}</span><span class="row">'
            f'<a class="btn fill" href="https://calendar.google.com/calendar/r?cid={urllib.parse.quote(absu, safe="")}" data-gcal="calendars/{html.escape(fname)}"><svg class="ic"><use href="#i-gcal"/></svg>В Google-календар</a>'
            f'<a class="btn tonal" href="{absu.replace("https://", "webcal://")}" data-sub="calendars/{html.escape(fname)}"><svg class="ic"><use href="#i-cal"/></svg>Інший календар</a>'
            f'<button class="btn out" data-link="calendars/{html.escape(fname)}"><svg class="ic"><use href="#i-link"/></svg>Лінк</button>'
            f'</span></div>')
    secs = "".join(
        f"<h2>{html.escape(flat(s))}</h2><div class='grid'>"
        + "".join(card(s, label, fname, i) for i, (label, fname) in enumerate(items)) + "</div>"
        for s, items in sheets)
    ts = datetime.datetime.now(datetime.timezone.utc).strftime("%d.%m.%Y %H:%M")
    return f"""<!doctype html><html lang="uk"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>BCSchedule</title>
<link rel="icon" type="image/svg+xml" href="favicon.svg">
<meta name="description" content="Неофіційний розклад занять ЧДБК (ЧДФБК) з можливістю додавання у календар">
<meta name="keywords" content="розклад, ЧДБК, ЧДФБК, Черкаський державний фаховий бізнес-коледж, розклад занять, календар, групи, пари">
<meta name="theme-color" content="#edf2e6">
<meta property="og:type" content="website">
<meta property="og:title" content="Неофіційний розклад пар ЧДБК">
<meta property="og:description" content="Обирай групу — пари самі прийдуть у твій календар">
<script data-goatcounter="https://bcschedule.goatcounter.com/count" async src="//gc.zgo.at/count.js"></script>
<meta name="google-site-verification" content="uGk621K82j_O0Twhiou4c1LqOzossy-005O0nijx-Ig" />
<link rel="stylesheet" href="style.css"></head><body><svg aria-hidden="true" style="position:absolute;width:0;height:0;overflow:hidden"><defs><symbol id="i-search" viewBox="0 -960 960 960"><path d="M796-121 533-384q-30 26-70 40.5T378-329q-108 0-183-75t-75-181q0-106 75-181t182-75q106 0 180.5 75T632-585q0 43-14 83t-42 75l264 262-44 44ZM377-389q81 0 138-57.5T572-585q0-81-57-138.5T377-781q-82 0-139.5 57.5T180-585q0 81 57.5 138.5T377-389Z"/></symbol><symbol id="i-help" viewBox="0 -960 960 960"><path d="M511-258q11-11 11-27t-11-27q-11-11-27-11t-27 11q-11 11-11 27t11 27q11 11 27 11t27-11Zm-62-135h59q0-26 6.5-47.5T555-490q31-26 44-51t13-55q0-53-34.5-85T486-713q-49 0-86.5 24.5T345-621l53 20q11-28 33-43.5t52-15.5q34 0 55 18.5t21 47.5q0 22-13 41.5T508-512q-30 26-44.5 51.5T449-393Zm31 313q-82 0-155-31.5t-127.5-86Q143-252 111.5-325T80-480q0-83 31.5-156t86-127Q252-817 325-848.5T480-880q83 0 156 31.5T763-763q54 54 85.5 127T880-480q0 82-31.5 155T763-197.5q-54 54.5-127 86T480-80Zm0-60q142 0 241-99.5T820-480q0-142-99-241t-241-99q-141 0-240.5 99T140-480q0 141 99.5 240.5T480-140Zm0-340Z"/></symbol><symbol id="i-info" viewBox="0 -960 960 960"><path d="M453-280h60v-240h-60v240Zm50.5-323.2q9.5-9.2 9.5-22.8 0-14.45-9.48-24.22-9.48-9.78-23.5-9.78t-23.52 9.78Q447-640.45 447-626q0 13.6 9.48 22.8 9.48 9.2 23.5 9.2t23.52-9.2ZM480.27-80q-82.74 0-155.5-31.5Q252-143 197.5-197.5t-86-127.34Q80-397.68 80-480.5t31.5-155.66Q143-709 197.5-763t127.34-85.5Q397.68-880 480.5-880t155.66 31.5Q709-817 763-763t85.5 127Q880-563 880-480.27q0 82.74-31.5 155.5Q817-252 763-197.68q-54 54.31-127 86Q563-80 480.27-80Zm.23-60Q622-140 721-239.5t99-241Q820-622 721.19-721T480-820q-141 0-240.5 98.81T140-480q0 141 99.5 240.5t241 99.5Zm-.5-340Z"/></symbol><symbol id="i-gcal" viewBox="0 -960 960 960"><path d="M700-80v-120H580v-60h120v-120h60v120h120v60H760v120h-60Zm-520-80q-24 0-42-18t-18-42v-540q0-24 18-42t42-18h65v-60h65v60h260v-60h65v60h65q24 0 42 18t18 42v302q-15-2-30-2t-30 2v-112H180v350h320q0 15 3 30t8 30H180Zm0-470h520v-130H180v130Zm0 0v-130 130Z"/></symbol><symbol id="i-cal" viewBox="0 -960 960 960"><path d="M180-80q-24 0-42-18t-18-42v-620q0-24 18-42t42-18h65v-60h65v60h340v-60h65v60h65q24 0 42 18t18 42v620q0 24-18 42t-42 18H180Zm0-60h600v-430H180v430Zm0-490h600v-130H180v130Zm0 0v-130 130Zm300 230q-17 0-28.5-11.5T440-440q0-17 11.5-28.5T480-480q17 0 28.5 11.5T520-440q0 17-11.5 28.5T480-400Zm-188.5-11.5Q280-423 280-440t11.5-28.5Q303-480 320-480t28.5 11.5Q360-457 360-440t-11.5 28.5Q337-400 320-400t-28.5-11.5ZM640-400q-17 0-28.5-11.5T600-440q0-17 11.5-28.5T640-480q17 0 28.5 11.5T680-440q0 17-11.5 28.5T640-400ZM480-240q-17 0-28.5-11.5T440-280q0-17 11.5-28.5T480-320q17 0 28.5 11.5T520-280q0 17-11.5 28.5T480-240Zm-188.5-11.5Q280-263 280-280t11.5-28.5Q303-320 320-320t28.5 11.5Q360-297 360-280t-11.5 28.5Q337-240 320-240t-28.5-11.5ZM640-240q-17 0-28.5-11.5T600-280q0-17 11.5-28.5T640-320q17 0 28.5 11.5T680-280q0 17-11.5 28.5T640-240Z"/></symbol><symbol id="i-link" viewBox="0 -960 960 960"><path d="M450-280H280q-83 0-141.5-58.5T80-480q0-83 58.5-141.5T280-680h170v60H280q-58.33 0-99.17 40.76-40.83 40.77-40.83 99Q140-422 180.83-381q40.84 41 99.17 41h170v60ZM325-450v-60h310v60H325Zm185 170v-60h170q58.33 0 99.17-40.76 40.83-40.77 40.83-99Q820-538 779.17-579q-40.84-41-99.17-41H510v-60h170q83 0 141.5 58.5T880-480q0 83-58.5 141.5T680-280H510Z"/></symbol></defs></svg><nav class="nav"><div class="nav-in"><a class="brand" href="index.html">BCSchedule</a><div class="search"><svg class="ic"><use href="#i-search"/></svg><input id="q" placeholder="Знайди свою групу…" autocomplete="off"></div><span class="nav-links"><a href="guide.html">Як додати?</a><a href="info.html">Інфо</a></span></div></nav><div class="wrap">
<div class="hero"><h1>Розклад пар, зроблений студентами для студентів</h1>
<svg class="vine" viewBox="0 0 420 34" aria-hidden="true">
<path d="M4 20 C 90 8, 170 30, 250 15 S 380 10, 416 19"/>
<ellipse class="l1" cx="105" cy="12" rx="11" ry="5" transform="rotate(-24 105 12)"/>
<ellipse class="l2" cx="210" cy="17" rx="6" ry="6"/>
<ellipse class="l3" cx="315" cy="12" rx="11" ry="5" transform="rotate(18 315 12)"/>
</svg>
<p class="sub">З автоматичним оновленням та враховуванням замін</p>
<p class="sub">Доволі багато навчальних закладів надають змогу експортувати розклад зайнять у календар. Це зручно - пари відображаються разом із іншими подіями, а при їх заміні - автоматично переміщуються або зникають. За допомогою магії <s>та костилів</s>, ми перетворили Google Sheet у такий календар. Тепер не потрібно відкривати розклад, документи з замінами та дзвінками - все знаходиться у одному місці (у всіх можливих сенсах)</p>
<span class="row"><a class="btn fill" href="guide.html"><svg class="ic"><use href="#i-help"/></svg>Як додати?</a><a class="btn tonal" href="info.html"><svg class="ic"><use href="#i-info"/></svg>Додаткова інформація</a></span></div>
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
<footer>Оновлюється кожні 6 годин
<br>Останнє оновлення: {ts} (UTC).
<br>Розклад ЧДБК - Черкаський державний фаховий бізнес-коледж (ЧДФБК)
<br><b>Неофіційний проєкт, не афілійований з коледжем</b></footer>
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
        # ponytail: sitemap lists only indexable pages; .ics files are text/calendar, not indexable, and poison sitemap quality
        sm = ['<?xml version="1.0" encoding="UTF-8"?>',
              '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'] + [
              f"<url><loc>{BASE}/{u}</loc><lastmod>{today}</lastmod></url>" for u in ("", "guide.html", "info.html")
        ] + ["</urlset>"]
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
