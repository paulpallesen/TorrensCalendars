#!/usr/bin/env python3
import os, csv, hashlib, io, sys
import pandas as pd
from datetime import datetime, date, time, timedelta
from collections import defaultdict

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
CAL_OPTIONAL = [
    "Calendar", "Start Time", "End Date", "End Time", "Location",
    "Description", "URL", "Transparent", "UID"
]

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
        AUS_TZ_VTIMEZONE.strip()
    ]

    event_count = 0
    for _, r in df.iterrows():
        title = str(r.get("Title", "") or "").strip()
        if not title: continue
        sdate = r.get("Start Date")
        if not _to_date(sdate): continue

        uid = str(r.get("UID", "") or "")
        course = ""  # if you have a Course column, map it in the sheet and include here
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
            uid = make_uid([course, title, sdate, edate, stime, etime, location, cal_name])

        summary = f"{title}"
        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{now_utc}")
        lines.append(f"SUMMARY:{ical_escape(summary)}")
        lines.append(f"TRANSP:{'TRANSPARENT' if is_transparent else 'OPAQUE'}")
        if location: lines.append(f"LOCATION:{ical_escape(location)}")
        if desc:     lines.append(f"DESCRIPTION:{ical_escape(desc)}")
        if url:      lines.append(f"URL:{ical_escape(url)}")

        cats = []
        if cat: cats.append(cat)
        if location: cats.append(location)
        if cats: lines.append(f"CATEGORIES:{ical_escape(','.join(cats))}")

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
        event_count += 1

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"

def main():
    csv_url = os.environ.get("CSV_URL", "").strip()
    if not csv_url:
        print("ERROR: CSV_URL env var is not set. Set it to your published Google Sheet CSV link.", file=sys.stderr)
        sys.exit(1)

    df = pd.read_csv(csv_url)

    # Ensure required columns exist
    for c in CAL_REQUIRED:
        if c not in df.columns:
            raise SystemExit(f"Missing required column: {c}")

    if "Calendar" not in df.columns:
        df["Calendar"] = "General"

    # Normalize empty calendar names
    df["Calendar"] = df["Calendar"].fillna("General").apply(lambda x: x if str(x).strip() else "General")

    # Output dir
    outdir = "public"
    os.makedirs(outdir, exist_ok=True)

    # Build one ICS per Calendar + a combined "All"
    categories = sorted(df["Calendar"].dropna().unique().tolist())
    built = []

    for cat in categories:
        sub = df[df["Calendar"] == cat].copy()
        ics = build_ics_for_group(sub, DEFAULT_TZ, cat)
        slug = slugify(cat)
        fname = f"calendar-{slug}.ics"
        with open(os.path.join(outdir, fname), "w", encoding="utf-8", newline="") as f:
            f.write(ics)
        built.append((cat, fname, len(sub)))

    # Build combined
    ics_all = build_ics_for_group(df, DEFAULT_TZ, "All")
    with open(os.path.join(outdir, "calendar-all.ics"), "w", encoding="utf-8", newline="") as f:
        f.write(ics_all)

    # Generate a simple index.html with dropdown
    options_html = '\n'.join(
        [f'<option value="calendar-all.ics">All (combined)</option>'] +
        [f'<option value="{fn}">{cat} ({count} events)</option>' for cat, fn, count in built]
    )

    index_html = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Dynamic Calendars</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 2rem; }}
.card {{ max-width: 720px; padding: 1.25rem; border: 1px solid #ddd; border-radius: 12px; }}
label, select {{ font-size: 1rem; }}
#link, #webcal {{ display: block; margin-top: 0.75rem; }}
small {{ color: #555; }}
</style>
</head>
<body>
  <div class="card">
    <h1>Subscribe to a Calendar</h1>
    <p>Select a calendar category. Use the <b>Subscribe (webcal)</b> link in Google Calendar/Outlook for a live subscription, or the <b>Direct URL</b> link for manual use.</p>

    <label for="cal">Calendar:</label>
    <select id="cal">{options_html}</select>

    <a id="webcal" href="#" target="_blank">Subscribe (webcal)</a>
    <a id="link" href="#" target="_blank">Direct URL (https)</a>

    <p><small>Note: Calendar apps refresh on their own schedule (Google ~every few hours). Updates come from your Google Sheet via scheduled builds or manual runs.</small></p>
  </div>

<script>
(function() {{
  var base = window.location.href.replace(/\\/[^\\/]*$/, "/"); // ensure trailing slash
  function updateLinks() {{
    var sel = document.getElementById('cal');
    var file = sel.value;
    var https = base + file;
    var webcal = "webcal://" + https.replace(/^https?:\\/\\//, "");
    document.getElementById('link').href = https;
    document.getElementById('webcal').href = webcal;
  }}
  document.getElementById('cal').addEventListener('change', updateLinks);
  updateLinks();
}})();
</script>
</body>
</html>
"""
    with open(os.path.join(outdir, "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)

    print("Built calendars:", [fn for _, fn, _ in built] + ["calendar-all.ics"])
    print("Wrote public/index.html")
    # Create a small robots.txt to be nice
    with open(os.path.join(outdir, "robots.txt"), "w", encoding="utf-8") as f:
        f.write("User-agent: *\nAllow: /\n")

if __name__ == "__main__":
    main()
