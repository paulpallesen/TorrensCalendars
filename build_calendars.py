#!/usr/bin/env python3
import os, hashlib, sys, json
import pandas as pd
from datetime import datetime, date, time, timedelta, timezone
from zoneinfo import ZoneInfo

DEFAULT_TZ = "Australia/Sydney"  # source timezone for your sheet times
CAL_REQUIRED = ["Title", "Start Date"]

def slugify(s: str) -> str:
    return "".join(ch.lower() if ch.isalnum() else "-" for ch in str(s)).strip("-") or "general"

def ical_escape(s: str) -> str:
    return (str(s)
            .replace("\\", "\\\\")
            .replace(",", "\\,")
            .replace(";", "\\;")
            .replace("\n", "\\n"))

def fmt_date(d: date) -> str:
    return d.strftime("%Y%m%d")

def fmt_utc(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")

def _to_date(val):
    if val is None or val == "": return None
    if isinstance(val, datetime): return val.date()
    if isinstance(val, date): return val
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    try: return datetime.fromisoformat(s).date()
    except Exception: return None

def _to_time(val):
    if val in (None, ""): return None
    if isinstance(val, datetime): return val.time().replace(second=0, microsecond=0)
    if isinstance(val, time): return val.replace(second=0, microsecond=0)
    s = str(val).strip()
    for fmt in ("%H:%M", "%H:%M:%S"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: pass
    try: return datetime.fromisoformat(s).time().replace(second=0, microsecond=0)
    except Exception: return None

def parse_local_datetime(d_val, t_val, tzid: str):
    d = _to_date(d_val)
    if not d: return None
    t = _to_time(t_val) or time(0, 0, 0)
    tz = ZoneInfo(tzid)
    return datetime(d.year, d.month, d.day, t.hour, t.minute, 0, tzinfo=tz)

def truthy(val) -> bool:
    if val is None: return False
    return str(val).strip().lower() in {"true", "yes", "y", "1", "transparent", "free"}

def make_uid(fields):
    h = hashlib.sha1("|".join(str(x) for x in fields).encode("utf-8")).hexdigest()[:16]
    return f"{h}@github-pages"

# ---------- RFC5545 folding (75 octets per line, CRLF) ----------
def fold_ical_line(line: str, limit: int = 75) -> list[str]:
    b = line.encode("utf-8")
    out = []
    while len(b) > limit:
        cut = limit
        while cut > 0 and (b[cut] & 0xC0) == 0x80:  # avoid splitting UTF-8 mid-char
            cut -= 1
        out.append(b[:cut].decode("utf-8"))
        b = b[cut:]
        out.append(" ")  # continuation marker (will be merged)
    if b:
        out.append(b.decode("utf-8"))
    merged = []
    i = 0
    while i < len(out):
        if out[i] == " " and i + 1 < len(out):
            merged.append(" " + out[i+1])
            i += 2
        else:
            merged.append(out[i])
            i += 1
    return merged

def add_prop(lines: list[str], prop: str):
    for part in fold_ical_line(prop):
        lines.append(part)

def build_ics_for_group(df: pd.DataFrame, tzid: str, cal_name: str) -> str:
    now_utc = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines: list[str] = []
    add_prop(lines, "BEGIN:VCALENDAR")
    add_prop(lines, "PRODID:-//Dynamic Calendars//GitHub Pages//EN")
    add_prop(lines, "VERSION:2.0")
    add_prop(lines, "CALSCALE:GREGORIAN")
    add_prop(lines, "METHOD:PUBLISH")
    add_prop(lines, f"X-WR-CALNAME:{ical_escape(cal_name)}")
    add_prop(lines, "X-PUBLISHED-TTL:PT12H")

    for _, r in df.iterrows():
        title = str(r.get("Title", "") or "").strip()
        if not title:
            continue
        sdate = r.get("Start Date")
        if not _to_date(sdate):
            continue

        uid = str(r.get("UID", "") or "")
        cat = str(r.get("Calendar", "") or "").strip()
        stime = r.get("Start Time") or ""
        edate = r.get("End Date") or ""
        etime = r.get("End Time") or ""
        location = str(r.get("Location", "") or "").strip()
        desc = str(r.get("Description", "") or "").strip()
        url = str(r.get("URL", "") or "").strip()
        is_transparent = truthy(r.get("Transparent"))

        is_all_day = (not _to_time(stime) and not _to_time(etime))
        if not uid:
            uid = make_uid([title, sdate, edate, stime, etime, location, cal_name])

        add_prop(lines, "BEGIN:VEVENT")
        add_prop(lines, f"UID:{uid}")
        add_prop(lines, f"DTSTAMP:{now_utc}")
        add_prop(lines, f"SUMMARY:{ical_escape(title)}")
        add_prop(lines, f"TRANSP:{'TRANSPARENT' if is_transparent else 'OPAQUE'}")
        if location: add_prop(lines, f"LOCATION:{ical_escape(location)}")
        if desc:     add_prop(lines, f"DESCRIPTION:{ical_escape(desc)}")
        if url:      add_prop(lines, f"URL:{ical_escape(url)}")

        cats = []
        if cat: cats.append(cat)
        if location: cats.append(location)
        if cats: add_prop(lines, f"CATEGORIES:{ical_escape(','.join(cats))}")

        if is_all_day:
            start_d = _to_date(sdate)
            end_d = _to_date(edate) or start_d
            end_excl = end_d + timedelta(days=1)
            add_prop(lines, f"DTSTART;VALUE=DATE:{fmt_date(start_d)}")
            add_prop(lines, f"DTEND;VALUE=DATE:{fmt_date(end_excl)}")
        else:
            dt_start_local = parse_local_datetime(sdate, stime, tzid)
            dt_end_local   = parse_local_datetime(edate or sdate, etime or stime or "00:00", tzid)
            if not dt_start_local or not dt_end_local:
                add_prop(lines, "END:VEVENT")
                continue
            add_prop(lines, f"DTSTART:{fmt_utc(dt_start_local)}")
            add_prop(lines, f"DTEND:{fmt_utc(dt_end_local)}")

        add_prop(lines, "END:VEVENT")

    add_prop(lines, "END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"

# ------------------------- MAIN -------------------------
def main():
    csv_url = os.environ.get("CSV_URL", "").strip()
    if not csv_url:
        print("ERROR: CSV_URL env var is not set. Set it to your published Google Sheet CSV link.", file=sys.stderr)
        sys.exit(1)

    df = pd.read_csv(csv_url)

    for c in CAL_REQUIRED:
        if c not in df.columns:
            raise SystemExit(f"Missing required column: {c}")

    if "Calendar" not in df.columns:
        df["Calendar"] = "General"
    df["Calendar"] = df["Calendar"].fillna("General").apply(lambda x: x if str(x).strip() else "General")

    outdir = "public"
    os.makedirs(outdir, exist_ok=True)

    categories = sorted(df["Calendar"].dropna().unique().tolist())
    built = []

    for cat in categories:
        sub = df[df["Calendar"] == cat].copy()
        ics = build_ics_for_group(sub, DEFAULT_TZ, cat)
        slug = slugify(cat)
        fname = f"calendar-{slug}.ics"
        with open(os.path.join(outdir, fname), "w", encoding="utf-8", newline="") as outf:
            outf.write(ics)
        built.append((cat, fname, len(sub)))

    ics_all = build_ics_for_group(df, DEFAULT_TZ, "All")
    with open(os.path.join(outdir, "calendar-all.ics"), "w", encoding="utf-8", newline="") as outf:
        outf.write(ics_all)

    feeds = [{"label": "All (combined)", "file": "calendar-all.ics", "count": len(df)}]
    feeds += [{"label": cat, "file": fn, "count": count} for (cat, fn, count) in built]
    feeds_json = json.dumps(feeds, ensure_ascii=False)

    # HTML template with a safe placeholder token (no Python f-string here)
    HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Subscribe to Calendars</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root {
    --orange: #F05A28; --maroon: #4B0A14; --cream: #F6F2EC;
    --ink: #111; --muted: #666; --card: #fff; --stroke: #e3dfda;
  }
  * { box-sizing: border-box; }
  body { margin:0; font-family: system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif; color:var(--ink); background: linear-gradient(180deg, var(--cream), #fff 40%); }
  header { background:var(--maroon); color:#fff; padding:18px 22px; }
  header .brand { display:flex; gap:12px; align-items:center; font-weight:700; letter-spacing:.3px; }
  header .dot { width:10px; height:10px; border-radius:50%; background: var(--orange); }
  main { padding:28px 22px; }
  .card { max-width:980px; margin:0 auto; background:var(--card); border:1px solid var(--stroke); border-radius:16px; padding:22px; box-shadow:0 6px 20px rgba(0,0,0,.05); }
  h1 { margin:0 0 10px; font-size:28px; }
  .lead { color:var(--muted); margin:0 0 18px; }
  .row { margin:16px 0; }
  label { font-weight:600; margin-right:8px; }
  select { font-size:16px; padding:8px 10px; border-radius:10px; border:1px solid var(--stroke); background:#fff; }
  .grid { display:grid; gap:18px; grid-template-columns:1fr; }
  @media(min-width:820px){ .grid { grid-template-columns:1.2fr .8fr; } }
  .btns { display:flex; flex-wrap:wrap; gap:10px; margin-top:8px; }
  a.button { display:inline-flex; align-items:center; gap:10px; padding:10px 14px; border-radius:12px; text-decoration:none; border:1px solid var(--stroke); background:#fff; color:var(--ink); transition: transform .05s ease, box-shadow .2s ease; }
  a.button:hover { transform: translateY(-1px); box-shadow: 0 6px 14px rgba(0,0,0,.08); }
  .button.google { border-color:#DADCE0; } .button.apple { border-color:#D1D1D1; } .button.outlook { border-color:#C7DCF7; }
  small { color:var(--muted); }
  code { background:#faf7f3; padding:4px 6px; border-radius:8px; border:1px solid var(--stroke); }
  .aside { border-left:1px dashed var(--stroke); padding-left:18px; }
  .kicker { color:var(--orange); font-weight:700; font-size:12px; letter-spacing:.6px; }
</style>
</head>
<body>
  <header><div class="brand"><div class="dot"></div><div>Torrens Dynamic Calendars</div></div></header>
  <main>
    <div class="card grid">
      <section>
        <div class="kicker">LIVE FEEDS</div>
        <h1>Subscribe to a calendar</h1>
        <p class="lead">Pick a category, then choose your calendar app. These are <strong>live subscriptions</strong>, not downloads.</p>

        <div class="row"><label for="cal">Calendar:</label><select id="cal"></select></div>

        <div class="row">
          <div class="btns">
            <a id="btn-apple" class="button apple" href="#" title="Subscribe in Apple Calendar / iOS / Outlook desktop">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <rect x="3" y="4" width="18" height="17" rx="4" ry="4" stroke="#555" fill="#fff"/>
                <rect x="3" y="8" width="18" height="13" rx="4" ry="4" fill="#fff" stroke="#555"/>
                <rect x="7" y="2" width="2" height="4" rx="1" fill="#555"/><rect x="15" y="2" width="2" height="4" rx="1" fill="#555"/>
                <rect x="6" y="11" width="12" height="7" fill="var(--orange)" opacity=".9"/>
              </svg><span>Apple / iOS / Outlook (desktop)</span></a>

            <a id="btn-google" class="button google" href="#" target="_blank" rel="noopener" title="Add by URL in Google Calendar">
              <svg width="18" height="18" viewBox="0 0 256 262" aria-hidden="true">
                <path fill="#4285F4" d="M255.9 133.5c0-10.6-.9-18.3-2.8-26.3H130v47.7h71.9c-1.4 11.9-9 29.8-25.9 41.8l-.2 1.6 37.6 29.1 2.6.3c23.8-22 39.9-54.4 39.9-94.2"/>
                <path fill="#34A853" d="M130 261.1c36.3 0 66.8-12 89.1-32.8l-42.4-32.8c-11.3 7.9-26.6 13.4-46.7 13.4-35.6 0-65.7-23.5-76.4-56.2l-1.6.1-41.5 32.1-.5 1.5C31.8 231.5 77.9 261.1 130 261.1"/>
                <path fill="#FBBC05" d="M53.6 152.7c-2.8-8-4.4-16.6-4.4-25.4 0-8.9 1.6-17.5 4.2-25.4l-.1-1.7-42-32.4-1.4.7C3.3 88.5 0 108.7 0 127.3s3.3 38.8 9.7 55.2l43.9-29.8"/>
                <path fill="#EA4335" d="M130 50.5c25.2 0 42.2 10.9 51.9 20l37.8-36.9C196.5 13.7 166.3 0 130 0 77.9 0 31.8 29.6 9.7 72.1l43.9 29.8C64.3 74.3 94.4 50.5 130 50.5"/>
              </svg><span>Google Calendar (web)</span></a>

            <a id="btn-outlookcom" class="button outlook" href="#" target="_blank" rel="noopener" title="Subscribe in Outlook.com (personal)">
              <svg width="18" height="18" viewBox="0 0 24 24" aria-hidden="true">
                <rect x="2" y="6" width="10" height="12" fill="#0A64D8"/><rect x="10" y="6" width="12" height="12" fill="#0F7BFF" opacity=".85"/>
                <circle cx="9" cy="12" r="3" fill="#fff"/>
              </svg><span>Outlook.com (personal)</span></a>

            <a id="btn-o365" class="button outlook" href="#" target="_blank" rel="noopener" title="Subscribe in Outlook 365 (work/school)">
              <svg width="18" height="18" viewBox="0 0 24 24" aria-hidden="true">
                <rect x="2" y="6" width="10" height="12" fill="#0A64D8"/><rect x="10" y="6" width="12" height="12" fill="#0F7BFF" opacity=".85"/>
                <circle cx="9" cy="12" r="3" fill="#fff"/>
              </svg><span>Outlook 365 (work/school)</span></a>
          </div>
          <p class="lead"><small>Google adds feeds under <em>Other calendars</em>. Apple may show a trust prompt the first time.</small></p>
        </div>
      </section>

      <aside class="aside">
        <div class="kicker">LINKS</div>
        <div class="row"><p><strong>Subscribe URL (webcal):</strong><br><code id="sub-url"></code></p><p><small>Use this to subscribe in apps that accept <code>webcal://</code>.</small></p></div>
        <div class="row"><p><strong>Download URL (https):</strong><br><code id="dl-url"></code></p><p><small>One-off import (static). Prefer the buttons or webcal URL for live updates.</small></p></div>
      </aside>
    </div>
  </main>

<script>
  // FEEDS will be injected below by Python safely as JSON:
  const FEEDS = /*__FEEDS__*/;

  const sel = document.getElementById('cal');
  const btnApple = document.getElementById('btn-apple');
  const btnGoogle = document.getElementById('btn-google');
  const btnOutlookCom = document.getElementById('btn-outlookcom');
  const btnO365 = document.getElementById('btn-o365');
  const subCode = document.getElementById('sub-url');
  const dlCode  = document.getElementById('dl-url');

  FEEDS.forEach(function(f) {
    const opt = document.createElement('option');
    opt.value = f.file;
    opt.textContent = f.label + (typeof f.count === 'number' ? ' (' + f.count + ')' : '');
    sel.appendChild(opt);
  });

  function canonicalBase() {
    let href = window.location.href.split('#')[0].split('?')[0];
    if (href.endsWith('index.html')) href = href.slice(0, -'index.html'.length);
    if (!href.endsWith('/')) href += '/';
    return href;
  }

  function computeGoogleRawUrl(icsFile) {
    // Expect site at: https://<owner>.github.io/<repo>/public/index.html
    const u = new URL(canonicalBase());
    const host = u.host;                // <owner>.github.io
    const owner = host.split('.')[0];   // <owner>
    const parts = u.pathname.split('/').filter(Boolean);
    const repo  = parts.length > 0 ? parts[0] : '';
    if (!repo) {
      // If user/organization site without repo path, fallback to Pages URL
      return new URL(icsFile, canonicalBase()).toString();
    }
    return `https://raw.githubusercontent.com/${owner}/${repo}/gh-pages/public/${icsFile}`;
  }

  function updateLinks() {
    const file   = sel.value;
    const base   = canonicalBase();
    const https  = new URL(file, base).toString();                 // Pages URL
    const webcal = 'webcal://' + https.replace(/^https?:\/\//, '');// Apple/desktop
    const rawUrl = computeGoogleRawUrl(file);                      // Google

    const googleAdd = 'https://calendar.google.com/calendar/render?cid='
                    + encodeURIComponent(rawUrl + '?no-cache=' + Date.now());

    const outlookCom = 'https://outlook.live.com/calendar/0/addfromweb'
                     + '?url='  + encodeURIComponent(https)
                     + '&name=' + encodeURIComponent(sel.options[sel.selectedIndex].text);

    const o365 = 'https://outlook.office.com/calendar/0/addfromweb'
               + '?url='  + encodeURIComponent(https)
               + '&name=' + encodeURIComponent(sel.options[sel.selectedIndex].text);

    btnApple.href      = webcal;
    btnGoogle.href     = googleAdd;
    btnOutlookCom.href = outlookCom;
    btnO365.href       = o365;

    subCode.textContent = webcal;
    dlCode.textContent  = https;
  }

  sel.addEventListener('change', updateLinks);
  updateLinks();
</script>
</body>
</html>
"""

    index_html = HTML_TEMPLATE.replace("/*__FEEDS__*/", feeds_json)
    with open(os.path.join(outdir, "index.html"), "w", encoding="utf-8", newline="") as outf:
        outf.write(index_html)

if __name__ == "__main__":
    main()
