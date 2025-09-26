# build_calendars.py
# Reads a Google Sheet (CSV export URL) and generates calendar.ics
# Includes Outlook-safe fixes: stable UID + removal of "nan" strings

import pandas as pd
from ics import Calendar, Event
from hashlib import md5

# ---- CONFIG --------------------------------------------------------------

# Replace <SHEET_ID> and <GID> with your real IDs (Publish to the web → CSV)
CSV_URL = "https://docs.google.com/spreadsheets/d/<SHEET_ID>/export?format=csv&gid=<GID>"

# If your sheet uses different headers, map them here -> canonical names
# Canonical names expected after mapping: Title, Start, End (optional),
# Location (optional), Description (optional), URL (optional), UID (optional)
COLUMN_MAP = {
    # "Subject": "Title",
    # "Start Time": "Start",
    # "End Time": "End",
    # "Room": "Location",
}

# Time zone (only used to localize if datetimes are naive)
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
    TZ = ZoneInfo("Australia/Sydney")
except Exception:
    TZ = None  # leaves times naïve if zoneinfo isn't available

# ---- Helpers -------------------------------------------------------------

def clean_str(v):
    """Turn NaN/None/whitespace into empty string; avoid literal 'nan'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s

def parse_dt(v):
    """Coerce to pandas datetime; localize to TZ if naïve."""
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
            # If localization fails, keep as-is
            pass
    return dt

def make_uid(summary, dtstart, dtend, extra=""):
    """Stable UID for Outlook: hash of key fields."""
    s_start = ""
    s_end = ""
    try:
        if dtstart is not None:
            s = pd.to_datetime(dtstart, errors="coerce")
            s_start = "" if pd.isna(s) else s.isoformat()
        if dtend is not None:
            e = pd.to_datetime(dtend, errors="coerce")
            s_end = "" if pd.isna(e) else e.isoformat()
    except Exception:
        pass
    base = f"{summary}|{s_start}|{s_end}|{extra}"
    return md5(base.encode("utf-8")).hexdigest() + "@torrens-uni"

# ---- Core build ----------------------------------------------------------

def read_sheet():
    df = pd.read_csv(CSV_URL)
    # Normalize headers and apply optional mapping
    df.columns = [c.strip() for c in df.columns]
    if COLUMN_MAP:
        df = df.rename(columns=COLUMN_MAP)

    # Replace blanks with NA, then drop rows missing essentials
    df = df.replace(r"^\s*$", pd.NA, regex=True)
    if "Title" not in df.columns or "Start" not in df.columns:
        missing = [c for c in ["Title", "Start"] if c not in df.columns]
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    df = df.dropna(subset=["Title", "Start"])

    # Coerce datetimes (Start required; End optional)
    for col in ["Start", "End"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Drop rows where Start failed to parse
    df = df.dropna(subset=["Start"])

    return df

def build_calendar(df: pd.DataFrame) -> Calendar:
    cal = Calendar()
    for _, r in df.iterrows():
        title = clean_str(r.get("Title"))
        if not title:
            continue

        start = parse_dt(r.get("Start"))
        end   = parse_dt(r.get("End")) if "End" in r else None
        if start is None and end is None:
            continue

        ev = Event()
        ev.name = title
        if start is not None:
            ev.begin = start
        if end is not None:
            ev.end = end

        loc  = clean_str(r.get("Location"))
        desc = clean_str(r.get("Description"))
        url  = clean_str(r.get("URL"))
        uid  = clean_str(r.get("UID")) or make_uid(title, start, end, loc)

        if loc:  ev.location = loc
        if desc: ev.description = desc
        if url:  ev.url = url
        ev.uid = uid  # critical for Outlook consistency

        cal.events.add(ev)
    return cal

if __name__ == "__main__":
    df = read_sheet()
    print(f"Rows after cleaning: {len(df)}")
    if not df.empty:
        try:
            print(df.head(3).to_string(index=False))
        except Exception:
            pass
    cal = build_calendar(df)
    with open("calendar.ics", "w", encoding="utf-8") as f:
        f.writelines(cal)
