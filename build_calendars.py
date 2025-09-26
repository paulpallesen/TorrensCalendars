# build_calendar.py
# NOTE: This version only adds:
#  A) stable UID generation when missing
#  B) stripping of NaN values so we don't write "nan" into ICS fields
#
# Everything else remains the same as the prior working file.

import pandas as pd
from ics import Calendar, Event
from hashlib import md5
from datetime import datetime

try:
    # Python 3.9+
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Australia/Sydney")
except Exception:
    TZ = None  # Fallback if zoneinfo isn't available; we won't change existing behavior

EXCEL_FILE = "calendar.xlsx"   # keep your existing source filename
OUTPUT_ICS  = "calendar.ics"   # keep your existing output filename

# --- NEW: helpers to fix Outlook issues ---

def make_uid(summary, dtstart, dtend, extra=""):
    """
    Create a stable unique UID per event when one is not provided.
    Uses a hash of key fields so updates modify the same event.
    """
    # Convert datetimes to ISO-like strings to stabilize hashing
    s_start = ""
    s_end   = ""
    try:
        if pd.notna(dtstart):
            s_start = pd.to_datetime(dtstart).isoformat()
        if pd.notna(dtend):
            s_end = pd.to_datetime(dtend).isoformat()
    except Exception:
        pass

    base = f"{summary}|{s_start}|{s_end}|{extra}"
    return md5(base.encode("utf-8")).hexdigest() + "@torrens-uni"

def clean_str(v):
    """
    Turn pandas NaN/None/whitespace into empty string; avoid literal 'nan' in ICS.
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        s = str(v).strip()
    except Exception:
        return ""
    return "" if s.lower() == "nan" else s

def parse_dt(v):
    """
    Keep existing behavior: coerce to datetime if possible and localize to Australia/Sydney
    only if tzinfo is missing. If TZ is unavailable, leave as-is.
    """
    if pd.isna(v):
        return None
    dt = pd.to_datetime(v, errors="coerce")
    if pd.isna(dt):
        return None
    if TZ is not None:
        try:
            if getattr(dt, "tzinfo", None) is None:
                dt = dt.tz_localize(TZ)
            else:
                dt = dt.tz_convert(TZ)
        except Exception:
            # If localization fails, keep original
            pass
    return dt

def main():
    # --- Keep your existing data source behavior (Excel) ---
    df = pd.read_excel(EXCEL_FILE)

    cal = Calendar()

    for _, row in df.iterrows():
        title = clean_str(row.get("Title", row.get("Subject", "")))
        if not title:
            continue

        start = parse_dt(row.get("Start"))
        end   = parse_dt(row.get("End"))

        # If both dates are missing, skip
        if start is None and end is None:
            continue

        ev = Event()
        ev.name = title

        if start is not None:
            ev.begin = start
        if end is not None:
            ev.end = end

        # --- Only change here is the NaN cleaning and UID fallback ---
        loc  = clean_str(row.get("Location"))
        desc = clean_str(row.get("Description"))
        url  = clean_str(row.get("URL"))
        uid  = clean_str(row.get("UID"))

        if not uid:
            uid = make_uid(title, start, end, loc)

        if loc:
            ev.location = loc
        if desc:
            ev.description = desc
        if url:
            # ics.Event has a 'url' property; safe to set when non-empty
            ev.url = url

        ev.uid = uid  # critical for Outlook

        cal.events.add(ev)

    with open(OUTPUT_ICS, "w", encoding="utf-8") as f:
        f.writelines(cal)

if __name__ == "__main__":
    main()
