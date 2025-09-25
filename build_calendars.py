#!/usr/bin/env python3
import os, sys, hashlib, json
import pandas as pd
from datetime import datetime, date, time, timedelta

DEFAULT_TZ = "Australia/Sydney"

AUS_TZ_VTIMEZONE = """BEGIN:VTIMEZONE
TZID:Australia/Sydney
BEGIN:STANDARD
DTSTART:19700405T030000
TZOFFSETFROM:+1100
TZOFFSETTO:+1000
TZNAME:AEST
RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:19701004T020000
TZOFFSETFROM:+1000
TZOFFSETTO:+1100
TZNAME:AEDT
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=1SU
END:DAYLIGHT
END:VTIMEZONE
"""

CAL_REQUIRED = ["Title", "Start Date"]

def slugify(s: str) -> str:
    return "".join(ch.lower() if ch.isalnum() else "-" for ch in str(s)).strip("-") or "general"

def ical_escape(s: str) -> str:
    return (str(s)
            .replace("\\", "\\\\")
            .replace(",", "\\,")
            .replace(";", "\\;")
            .replace("\n", "\\n"))

def fmt_local(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%S")

def fmt_date(d: date) -> str:
    return d.strftime("%Y%m%d")

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

def parse_datetime(d_val, t_val):
    d = _to_date(d_val)
    if not d: return None
    t = _to_time(t_val) or time(0, 0, 0)
    return datetime(d.year, d.month, d.day, t.hour, t.minute, 0)

def truthy(val) -> bool:
    if val is None: return False
    return str(val).strip().lower() in {"true", "yes", "y", "1", "transparent", "free"}

def make_uid(fields):
    h = hashlib.sha1("|".join(str(x) for x in fields).encode("utf-8")).hexdigest()[:16]
    return f"{h}@netlify-calendars"

def build_ics_for_group(df: pd.DataFrame, tzid: str, cal_name: str) -> str:
    now_utc = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//Dynamic Calendars//Netlify//EN",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{ical_escape(cal_name)}",
        f"X-WR-TIMEZONE:{tzid}",
        "X-PUBLISHED-TTL:PT12H",
        AUS_TZ_VTIMEZONE.strip()
    ]

    for _, r in df.iterrows():
        title = str(r.get("Title", "") or "").strip()
        if not title: continue
        sdate = r.get("Start Date")
        if not _to_date(sdate): continue

        uid = str(r.get("UID", "") or "")
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

        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{now_utc}")
        lines.append(f"SUMMARY:{ical_escape(title)}")
        lines.append(f"TRANSP:{'TRANSPARENT' if is_transparent else 'OPAQUE'}")
        if location: lines.append(f"LOCATION:{ical_escape(location)}")
        if desc:     lines.append(f"DESCRIPTION:{ical_escape(desc)}")
        if url:      lines.append(f"URL:{ical_escape(url)}")

        if is_all_day:
            start_d = _to_date(sdate)
            end_d = _to_date(edate) or start_d
            end_excl = end_d + timedelta(days=1)
            lines.append(f"DTSTART;VALUE=DATE:{fmt_date(start_d)}")
            lines.append(f"DTEND;VALUE=DATE:{fmt_date(end_excl)}")
        else:
            dt_start = parse_datetime(sdate, stime)
            dt_end   = parse_datetime(edate or sdate, etime or stime or "00:00")
            if not dt_start or not dt_end:
                lines.append("END:VEVENT")
                continue
            lines.append(f"DTSTART;TZID={tzid}:{fmt_local(dt_start)}")
            lines.append(f"DTEND;TZID={tzid}:{fmt_local(dt_end)}")

        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"

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
    feeds = []
    for cat in categories:
        sub = df[df["Calendar"] == cat].copy()
        ics = build_ics_for_group(sub, DEFAULT_TZ, cat)
        slug = slugify(cat)
        fname = f"calendar-{slug}.ics"
        with open(os.path.join(outdir, fname), "w", encoding="utf-8", newline="") as fh:
            fh.write(ics)
        feeds.append({"label": cat, "file": fname, "count": int(len(sub))})

    ics_all = build_ics_for_group(df, DEFAULT_TZ, "All")
    with open(os.path.join(outdir, "calendar-all.ics"), "w", encoding="utf-8", newline="") as fh:
        fh.write(ics_all)
    feeds.insert(0, {"label": "All (combined)", "file": "calendar-all.ics", "count": int(len(df))})

    # FUNCTIONALITY-FIRST TEMPLATE (neutral skin; working links)
    html_template = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Subscribe to Calendars</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root {
    --page: #f6f3e7;   /* light background */
    --card: #420318;  /* maroon card */
    --text: #ffffff;
    --apple:  #979797;
    --google: #ea4236;
    --outlook:#0077da;
  }
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 2rem; background: var(--page); color: var(--text); }
  .card { max-width: 900px; padding: 1.5rem; border-radius: 14px; background: var(--card); }
  h1 { margin: 0 0 1rem 0; font-size: 2rem; }
  .row { margin: 1rem 0; }
  .controls { display:flex; align-items:center; gap: .75rem; }
  .controls label { font-weight: 700; min-width: 110px; }
  select {
    padding: .55rem .7rem; border-radius: 10px; font-size: 1rem;
    width: 320px; background: #fff; color: #222; border: 1px solid rgba(255,255,255,.25);
  }
  .btn { display:inline-block; padding:.6rem .95rem; margin-right:.5rem; text-decoration:none; color:#fff; border-radius:10px; font-weight:600; }
  .btn-apple  { background: var(--apple); }
  .btn-google { background: var(--google); }
  .btn-outlook{ background: var(--outlook); }
  code { display:inline-block; background:#fff; color:#222; padding:.3rem .45rem; border-radius:8px; border:1px solid rgba(0,0,0,.1); cursor:pointer; white-space:nowrap; }
  #toast { position: fixed; left:50%; bottom:24px; transform:translateX(-50%); background: rgba(0,0,0,.85); color:#fff; padding:.5rem .75rem; border-radius:8px; font-size:.9rem; opacity:0; transition:opacity .2s ease; }
  #toast.show { opacity:1; }
</style>
</head>
<body>
  <div class="card">
    <h1>Subscribe to a Calendar</h1>
    <div class="row controls">
      <label for="cal">Calendar:</label>
      <select id="cal"></select>
    </div>
    <div class="row">
      <strong>Subscribe with:</strong><br>
      <a id="btn-apple"   class="btn btn-apple"   href="#">Apple Calendar</a>
      <a id="btn-google"  class="btn btn-google"  href="#">Google Calendar</a>
      <a id="btn-outlook" class="btn btn-outlook" href="#">Outlook (Work/Study)</a>
    </div>
    <div class="row">
      <strong>Direct feed URL:</strong> <code id="direct-url" title="Click to copy"></code>
    </div>
  </div>
  <div id="toast">Copied</div>

<script>
  const FEEDS = __FEEDS_JSON__;

  const sel = document.getElementById('cal');
  const btnApple   = document.getElementById('btn-apple');
  const btnGoogle  = document.getElementById('btn-google');
  const btnOutlook = document.getElementById('btn-outlook');
  const directCode = document.getElementById('direct-url');
  const toast      = document.getElementById('toast');

  function showToast(msg) {
    toast.textContent = msg;
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 1500);
  }

  // Since index.html and .ics live in the same folder (/public/), just resolve the filename against this page:
  function icsUrl(file) {
    const base = new URL(window.location.href);
    if (base.pathname.endsWith('/index.html')) {
      base.pathname = base.pathname.slice(0, -'index.html'.length);
    }
    return new URL(file, base).toString();
  }

  FEEDS.forEach(f => {
    const opt = document.createElement('option');
    opt.value = f.file;
    opt.textContent = f.label + (typeof f.count === 'number' ? ` (${f.count})` : '');
    sel.appendChild(opt);
  });

  let CURRENT = { https:'', label:'' };

  function updateLinks() {
    const file  = sel.value;
    const label = sel.options[sel.selectedIndex].text;
    const https = icsUrl(file);

    CURRENT = { https, label };
    // Apple/macOS/iOS & Outlook desktop (webcal)
    btnApple.href = 'webcal://' + https.replace(/^https?:\\/\\//,'');
    btnApple.target = '_self'; btnApple.rel = '';

    // Outlook (Work/Study) composer (pre-fills)
    btnOutlook.href =
      'https://outlook.office.com/owa/?path=/calendar/action/compose&rru=addsubscription'
      + '&url='  + encodeURIComponent(https)
      + '&name=' + encodeURIComponent(label);
    btnOutlook.target = '_blank';
    btnOutlook.rel = 'noopener';

    // Direct URL visible & copyable
    directCode.textContent = https;
  }

  // Google: copy to clipboard + open "Add by URL" page
  btnGoogle.addEventListener('click', async (e) => {
    e.preventDefault();
    try { await navigator.clipboard.writeText(CURRENT.https); } catch (_) {}
    showToast('Feed URL copied. Paste it in Google → Add by URL.');
    window.open('https://calendar.google.com/calendar/u/0/r/settings/addbyurl', '_blank', 'noopener');
  });

  directCode.addEventListener('click', async () => {
    try { await navigator.clipboard.writeText(CURRENT.https); } catch (_) {}
    showToast('Feed URL copied');
  });

  sel.addEventListener('change', updateLinks);
  updateLinks();
</script>
</body>
</html>
"""
    html_out = html_template.replace("__FEEDS_JSON__", json.dumps(feeds, ensure_ascii=False))
    with open(os.path.join(outdir, "index.html"), "w", encoding="utf-8") as fh:
        fh.write(html_out)

if __name__ == "__main__":
    main()
