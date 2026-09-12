#!/usr/bin/env python3
"""Sheet -> ICS. Stdlib only. Usage: python3 schedule_to_ics.py --group 1Д-22 --start 2026-09-14 --weeks 4 -o out.ics"""
import argparse, glob, io, os, re, urllib.request, uuid, datetime, zipfile
import xml.etree.ElementTree as ET

SHEET_ID = "1SXdz3k3Ect865_IIL3vm-Ia1LvNhK3ls"
BASE = "https://fstmoore.github.io/bcschedule"
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

_CSS = """
:root{--paper:#edf2e6;--ink:#1b241a;--moss:#2e5339;--fern:#5f8f60;--pollen:#d9a13b;
--mut:#68765f;--card:#f5f8ef;--line:#c9d4bd}
footer{margin:44px 0 0;color:var(--mut);font-size:14px}
.fine{font-size:14px;color:var(--mut);max-width:70ch}.fine b{color:var(--moss)}
@media(prefers-color-scheme:dark){:root{--paper:#111711;--ink:#e6ecdf;--moss:#9cc184;
--fern:#7ba37e;--pollen:#e0aa45;--mut:#a3ae9c;--card:#1a211a;--line:#33402f}}
*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);
font:17px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif}
body::after{content:"";position:fixed;inset:0;pointer-events:none;opacity:.05;
background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='120' height='120'%3E%3Cfilter id='n'%3E%3CfeTurbulence baseFrequency='.9'/%3E%3C/filter%3E%3Crect width='120' height='120' filter='url(%23n)'/%3E%3C/svg%3E")}
.wrap{max-width:960px;margin:0 auto;padding:24px 20px 80px}
.hero{padding:40px 0 12px}.hero h1{font-family:Fraunces,Georgia,serif;
font-size:clamp(40px,7vw,72px);line-height:1.02;font-weight:600;letter-spacing:-.01em;margin:0}
.hero .sub{font-size:19px;max-width:58ch;color:var(--mut)}
.vine{display:block;width:min(420px,80%);height:34px;margin-top:6px}
.vine path{fill:none;stroke:var(--fern);stroke-width:3;stroke-linecap:round;
stroke-dasharray:600;stroke-dashoffset:600;animation:grow 1.6s ease-out .3s forwards}
.vine ellipse{fill:var(--fern);opacity:0;animation:leaf .5s ease-out forwards}
.vine .l1{animation-delay:1s}.vine .l2{animation-delay:1.3s}.vine .l3{animation-delay:1.5s}
@keyframes grow{to{stroke-dashoffset:0}}
@keyframes leaf{to{opacity:1}}
@media(prefers-reduced-motion:reduce){.vine path{animation:none;stroke-dashoffset:0}.vine ellipse{opacity:1;animation:none}}
.search{display:flex;background:var(--card);border:1.5px solid var(--line);
border-radius:999px;padding:12px 22px;margin:26px 0 14px}
.search:focus-within{border-color:var(--moss)}
input{flex:1;border:0;background:none;color:var(--ink);font:inherit;outline:none;min-width:0}
.chips{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:6px}
.chip{border:1.5px solid var(--line);background:var(--card);color:var(--ink);
border-radius:999px;padding:7px 18px;font:inherit;cursor:pointer}
.chip.on{background:var(--moss);border-color:var(--moss);color:var(--paper)}
h2{font-family:Fraunces,Georgia,serif;font-size:30px;font-weight:600;margin:40px 0 14px}
.grid{display:grid;gap:14px;grid-template-columns:repeat(auto-fill,minmax(250px,1fr))}
.card{background:var(--card);border:1.5px solid var(--line);border-radius:20px;
padding:18px;display:flex;flex-direction:column;gap:14px}
.card .t{font-family:Fraunces,Georgia,serif;font-size:24px;font-weight:600}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.btn{border-radius:999px;padding:9px 18px;font:inherit;font-size:15px;cursor:pointer;text-decoration:none}
.fill{background:var(--moss);color:var(--paper);border:0}
.tonal{background:none;color:var(--ink);border:1.5px solid var(--moss)}
.out{background:none;border:0;color:var(--mut);text-decoration:underline;
text-underline-offset:3px;padding:9px 6px}
:focus-visible{outline:2px solid var(--pollen);outline-offset:2px}
dialog{border:1.5px solid var(--moss);border-radius:18px;background:var(--paper);color:var(--ink);
padding:28px;max-width:min(440px,90vw)}
dialog::backdrop{background:rgba(20,26,19,.45);backdrop-filter:blur(6px)}
dialog h3{font-family:Fraunces,Georgia,serif;font-size:30px;margin:0 0 8px}
dialog p{margin:0 0 16px}
dialog input{flex:1;min-width:0;font:inherit;background:var(--card);color:var(--ink);
border:1.5px solid var(--line);border-radius:12px;padding:8px}
footer{margin-top:56px;color:var(--mut);font-size:15px}
"""

_JS = """
const cards=[...document.querySelectorAll('.card')];
let sheet='';
const apply=()=>{const q=document.getElementById('q').value.trim().toLowerCase();
cards.forEach(c=>c.style.display=(!sheet||c.dataset.s===sheet)&&c.dataset.g.includes(q)?'':'none');
document.querySelectorAll('#secs h2').forEach(h=>{let n=h.nextElementSibling,v=false;
[...n.children].forEach(k=>{if(k.style.display!=='none')v=true});h.style.display=v?'':'none';n.style.display=v?'grid':'none'})};
document.getElementById('q').oninput=apply;
document.querySelectorAll('.chip').forEach(b=>b.onclick=()=>{document.querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));b.classList.add('on');sheet=b.dataset.s;apply()});
const abs=p=>new URL(p,location.href).href;
document.querySelectorAll('[data-gcal]').forEach(a=>a.href='https://calendar.google.com/calendar/r?cid='+encodeURIComponent(abs(a.dataset.gcal)));
document.querySelectorAll('[data-sub]').forEach(a=>a.href=abs(a.dataset.sub).replace(/^https?/,'webcal'));
const dlg=document.getElementById('how'),lk=document.getElementById('lk'),cp=document.getElementById('cp');
document.querySelectorAll('[data-link]').forEach(b=>b.onclick=()=>{lk.value=abs(b.dataset.link);cp.textContent='Копіювати';dlg.showModal()});
cp.onclick=()=>{(navigator.clipboard?navigator.clipboard.writeText(lk.value):Promise.reject()).then(()=>cp.textContent='Готово!',()=>lk.select())};
document.getElementById('nope').onclick=()=>dlg.close();
dlg.onclick=e=>{if(e.target===dlg)dlg.close()};
"""

_VIEW_CSS = """
.week{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-top:16px}
.wday{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:10px}
.wday h3{margin:0 0 8px;font-size:14px;color:var(--moss)}
.ev{border-left:3px solid var(--fern);padding:4px 6px;margin:0 0 8px;font-size:13px}
.ev b{display:block}.ev span{color:var(--mut)}
select{font:inherit;background:var(--card);color:var(--ink);border:1.5px solid var(--moss);border-radius:12px;padding:8px;max-width:100%}
#wl{min-width:110px;text-align:center;color:var(--mut)}
"""

_VIEW_JS = """
let W=0,EV=[];
const f=document.getElementById('f'),grid=document.getElementById('grid'),wl=document.getElementById('wl');
const load=()=>fetch('calendars/'+encodeURIComponent(f.value)).then(r=>r.text()).then(t=>{EV=parse(t);W=0;draw()});
const parse=t=>{t=t.replace(/\\r\\n[ \\t]/g,'');const ev=[];let c={};
for(const l of t.split(/\\r\\n|\\n/)){const i=l.indexOf(':');
if(l==='BEGIN:VEVENT')c={};else if(l==='END:VEVENT'){if(c.dt)ev.push(c);c={}}
else if(i>0){const k=l.slice(0,i),v=l.slice(i+1);
if(k==='DTSTART')c.dt=v;else if(k==='DTEND')c.en=v;else if(k==='SUMMARY')c.s=v;else if(k==='LOCATION')c.l=v;else if(k==='DESCRIPTION')c.d=v}}return ev};
const wk=d=>{const x=new Date(d);x.setHours(0,0,0,0);x.setDate(x.getDate()-((x.getDay()+6)%7));return x.getTime()};
const p=s=>new Date(+s.slice(0,4),+s.slice(4,6)-1,+s.slice(6,8));
const iso=d=>d.getFullYear()+String(d.getMonth()+1).padStart(2,'0')+String(d.getDate()).padStart(2,'0');
const fmt=d=>d.getDate()+'.'+(d.getMonth()+1);
const dn=d=>['Нд','Пн','Вт','Ср','Чт','Пт','Сб'][d.getDay()];
const draw=()=>{const weeks=[...new Set(EV.map(e=>wk(p(e.dt))))].sort((a,b)=>a-b);
if(!weeks.length){grid.innerHTML='Порожньо.';wl.textContent='';return}
W=Math.max(0,Math.min(weeks.length-1,W));const w0=weeks[W];
const days=[...Array(7)].map((_,i)=>new Date(w0+i*864e5));
wl.textContent=fmt(days[0])+' — '+fmt(days[6]);
grid.innerHTML=days.map(d=>{const k=iso(d);
const es=EV.filter(e=>e.dt.slice(0,8)===k).sort((a,b)=>a.dt<b.dt?-1:1);
return '<div class=wday><h3>'+dn(d)+' '+k.slice(6)+'.'+k.slice(4,6)+'</h3>'+(es.map(e=>'<div class=ev><b>'+e.dt.slice(9,11)+':'+e.dt.slice(11,13)+'–'+e.en.slice(9,11)+':'+e.en.slice(11,13)+'</b>'+e.s+(e.l?'<span> · ауд. '+e.l+'</span>':'')+(e.d?'<br><span>'+e.d+'</span>':'')+'</div>').join('')||'<span>—</span>')+'</div>'}).join('')};
f.onchange=load;document.getElementById('pv').onclick=()=>{W--;draw()};document.getElementById('nx').onclick=()=>{W++;draw()};
load();
"""

def render_view(files):
    """Week-grid ICS viewer. File list baked in (static hosting has no directory listing)."""
    import html
    return ("<!doctype html><html lang='uk'><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            "<title>Тиждень сіткою</title>"
            "<script data-goatcounter='https://bcschedule.goatcounter.com/count' async src='//gc.zgo.at/count.js'></script>"
            "<style>" + _CSS + _VIEW_CSS + "</style></head><body><div class='wrap'>"
            "<h1>Що там на тижні</h1>"
            "<p class='sub'><a href='index.html'>← до груп</a></p>"
            "<div class='row'><select id='f'>" + "".join(f"<option>{html.escape(x)}</option>" for x in files) + "</select>"
            "<button class='btn tonal' id='pv'>←</button><span id='wl'></span><button class='btn tonal' id='nx'>→</button></div>"
            "<div id='grid' class='week'></div>"
            "<script>" + _VIEW_JS + "</script></div></body></html>")

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
    def card(sheet, label, fname):
        return (
            f'<div class="card" data-g="{html.escape(label.lower())}" data-s="{html.escape(sheet.strip())}">'
            f'<span class="t">{html.escape(label)}</span><span class="row">'
            f'<a class="btn fill" href="#" data-gcal="calendars/{html.escape(fname)}">В Google-календар</a>'
            f'<a class="btn tonal" href="#" data-sub="calendars/{html.escape(fname)}">Інший календар</a>'
            f'<button class="btn out" data-link="calendars/{html.escape(fname)}">Лінк</button>'
            f'</span></div>')
    secs = "".join(
        f"<h2>{html.escape(s.strip())}</h2><div class='grid'>"
        + "".join(card(s, label, fname) for label, fname in items) + "</div>"
        for s, items in sheets)
    ts = datetime.datetime.now(datetime.timezone.utc).strftime("%d.%m.%Y %H:%M")
    return f"""<!doctype html><html lang="uk"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Розклад, який живе в календарі</title>
<meta name="description" content="Розклад занять ЧДБК — Черкаського державного фахового бізнес-коледжу (ЧДФБК) — у вигляді календаря: обери групу, додай у Google-календар або будь-який інший. Оновлюється кожні 6 годин. Неофіційний проєкт, не афілійований з коледжем.">
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
<footer>Поливаємо кожні 6 годин — розклад сам росте з гугл-таблички.
<br>Vibecoded in 2 hrs without a wage.
<br>Останнє оновлення: {ts} (UTC).
<br><a href='view.html'>Тиждень сіткою</a>
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
    for sheet, g, grid in jobs:
        lessons = parse_group(grid, g)
        subs = sorted({L["sub"] for L in lessons if L.get("sub")}) or [""] if a.group == "all" else [""]
        for sub in subs:  # split A/B subgroups into separate files; shared lessons go in both
            out = a.o if a.group != "all" else os.path.join("calendars", f"{g}-{sub}.ics" if sub else f"{g}.ics")
            os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
            open(out, "w", encoding="utf-8").write(to_ics(lessons, monday, a.weeks, g, a.flip_weeks, sub))
            print(f"{g}{'-' + sub if sub else ''}: {len(lessons)} pairs/week -> {out}")
            if a.group == "all":
                pages.append((sheet, f"{g}-{sub}" if sub else g, os.path.basename(out)))
    if a.group == "all":
        # ponytail: drop stale files (e.g. graduated groups) so dead links never linger
        keep = {os.path.basename(p[2]) for p in pages}
        for f in glob.glob(os.path.join("calendars", "*.ics")):
            if os.path.basename(f) not in keep:
                os.remove(f)
        open("index.html", "w", encoding="utf-8").write(render_index(pages))
        print(f"index.html ({len(pages)} groups)")
        files = sorted({p[2] for p in pages})
        open("view.html", "w", encoding="utf-8").write(render_view(files))
        print(f"view.html ({len(files)} files)")
        today = datetime.date.today().isoformat()
        urls = ["", "view.html"] + [f"calendars/{f}" for f in files]
        sm = ['<?xml version="1.0" encoding="UTF-8"?>',
              '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'] + [
              f"<url><loc>{BASE}/{u}</loc><lastmod>{today}</lastmod></url>" for u in urls
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
