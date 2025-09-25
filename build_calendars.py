#!/usr/bin/env python3
import os, hashlib, sys
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
    return f"{h}@github-pages"

def build_ics_for_group(df: pd.DataFrame, tzid: str, cal_name: str) -> str:
    now_utc = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//Dynamic Calendars//GitHub Pages//EN",
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

        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{now_utc}")
        lines.append(f"SUMMARY:{ical_escape(title)}")
        lines.append(f"TRANSP:{'TRANSPARENT' if is_transparent else 'OPAQUE'}")
        if location: lines.append(f"LOCATION:{ical_escape(location)}")
        if desc:     lines.append(f"DESCRIPTION:{ical_escape(desc)}")
        if url:      lines.append(f"URL:{ical_escape(url)}")
        if cat:      lines.append(f"CATEGORIES:{ical_escape(cat)}")

        if is_all_day:
            start_d = _to_date(sdate)
            end_d = _to_date(edate) or start_d
            end_excl = end_d + timedelta(days=1)
            lines.append(f"DTSTART;VALUE=DATE:{fmt_date(start_d)}")
            lines.append(f"DTEND;VALUE=DATE:{fmt_date(end_excl)}")
        else:
            dt_start = parse_datetime(sdate, stime)
            dt_end   = parse_datetime(edate or sdate, etime or stime or "00:00")
            if dt_start and dt_end:
                lines.append(f"DTSTART;TZID={tzid}:{fmt_local(dt_start)}")
                lines.append(f"DTEND;TZID={tzid}:{fmt_local(dt_end)}")
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"

def main():
    csv_url = os.environ.get("CSV_URL", "").strip()
    if not csv_url:
        print("ERROR: CSV_URL not set", file=sys.stderr)
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
        with open(os.path.join(outdir, fname), "w", encoding="utf-8") as f_out:
            f_out.write(ics)
        built.append((cat, fname, len(sub)))

    ics_all = build_ics_for_group(df, DEFAULT_TZ, "All")
    with open(os.path.join(outdir, "calendar-all.ics"), "w", encoding="utf-8") as f_all:
        f_all.write(ics_all)

    feeds = [{"label": "All (combined)", "file": "calendar-all.ics", "count": len(df)}]
    feeds += [{"label": cat, "file": fn, "count": count} for (cat, fn, count) in built]

    index_html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Subscribe to Calendars</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  body {{ font-family: 'Segoe UI', Roboto, Arial, sans-serif; margin: 0; padding: 0; background: #f8f9fa; color: #222; }}
  header {{ background: #002147; color: #fff; padding: 1rem 2rem; }}
  header h1 {{ margin: 0; font-size: 1.75rem; }}
  .container {{ max-width: 900px; margin: 2rem auto; padding: 2rem; background: #fff; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,.1); }}
  .row {{ margin: 1rem 0; }}
  label {{ font-weight: 600; display: block; margin-bottom: .5rem; }}
  select {{ font-size: 1rem; padding: .5rem; border: 1px solid #ccc; border-radius: 8px; }}
  .buttons a {{ display:inline-block; padding:.75rem 1rem; margin:.25rem; border-radius:8px; text-decoration:none; font-weight:500; }}
  .apple {{ background:#555; color:#fff; }}
  .google {{ background:#EA4335; color:#fff; }}
  .outlook {{ background:#0078D4; color:#fff; }}
  code {{ background:#f6f6f6; padding:.25rem .4rem; border-radius:6px; }}
  small {{ color:#555; display:block; margin-top:.5rem; }}
</style>
</head>
<body>
  <header>
    <h1>Dynamic Calendars</h1>
  </header>
  <div class="container">
    <div class="row">
      <label for="cal">Select a Calendar:</label>
      <select id="cal"></select>
    </div>

    <div class="row buttons">
      <a id="btn-apple" class="apple" href="#"> Apple / iOS / Outlook (desktop)</a>
      <a id="btn-google" class="google" href="#">Google Calendar (web)</a>
      <a id="btn-outlook" class="outlook" href="#">Outlook.com (web)</a>
      <small>Google requires "Add by URL" — we copy the link for you.</small>
    </div>

    <div class="row">
      <label>Direct feed URL:</label>
      <code id="direct-url"></code>
    </div>
  </div>

<script>
  const FEEDS = {feeds};

  function baseUrl() {{
    const u = new URL(window.location.href);
    if (!u.pathname.endsWith('/')) {{
      u.pathname = u.pathname.substring(0, u.pathname.lastIndexOf('/') + 1);
    }}
    return u;
  }}

  const sel = document.getElementById('cal');
  const btnApple   = document.getElementById('btn-apple');
  const btnGoogle  = document.getElementById('btn-google');
  const btnOutlook = document.getElementById('btn-outlook');
  const directCode = document.getElementById('direct-url');

  FEEDS.forEach(f => {{
    const opt = document.createElement('option');
    opt.value = f.file;
    // Escape braces inside Python f-string so `${{f.count}}` reaches JS as `${f.count}`
    opt.textContent = f.label + (typeof f.count === 'number' ? ` (${{f.count}})` : '');
    sel.appendChild(opt);
  }});

  function updateLinks() {{
    const file = sel.value;
    const base = baseUrl();
    const https = new URL(file, base).toString();

    const webcal = 'webcal://' + https.replace(/^https?:\\/\\//, '');
    const outlook = 'https://outlook.live.com/owa?path=/calendar/action/compose&rru=addsubscription'
                  + '&url=' + encodeURIComponent(https)
                  + '&name=' + encodeURIComponent(sel.options[sel.selectedIndex].text);

    btnApple.href   = webcal;
    btnOutlook.href = outlook;
    directCode.textContent = https;

    btnGoogle.onclick = async function(e) {{
      e.preventDefault();
      try {{
        await navigator.clipboard.writeText(https);
        alert("Calendar URL copied! In Google Calendar → Other calendars → From URL → Paste → Add.");
      }} catch (err) {{
        prompt("Copy this calendar URL manually:", https);
      }}
      const settingsUrl = 'https://calendar.google.com/calendar/u/0/r/settings/addbyurl?cid=' + encodeURIComponent(https);
      window.open(settingsUrl, '_blank');
    }};
  }}

  sel.addEventListener('change', updateLinks);
  updateLinks();
</script>
</body>
</html>
"""
    with open(os.path.join(outdir, "index.html"), "w", encoding="utf-8") as f_index:
        f_index.write(index_html)

if __name__ == "__main__":
    main()
