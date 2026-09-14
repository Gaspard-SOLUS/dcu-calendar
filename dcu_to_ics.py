#!/usr/bin/env python3
"""
dcu_to_ics.py — turn a DCU MyTimetable Excel export into an Apple Calendar feed.

Input   : the .xlsx produced by MyTimetable -> "Multiple weeks" -> EXCEL
          (columns: Module Name, Location, Date, Day, Time range, Weeks, Source)
Enrich  : overrides.csv — session type, lecturer, notes. Keyed on
          module_code|day|start, so it survives a re-export untouched.
Output  : an RFC 5545 iCalendar file, Europe/Dublin, stable UIDs. Each event
          carries GEO coordinates and a tap-to-open Apple Maps link for its
          room, and keeps its DTSTAMP/SEQUENCE unchanged run to run unless its
          content actually changed, so regenerating produces a clean git diff.

The Excel export lists every occurrence with its real date, so no week-number
arithmetic is involved. WEEK1_MONDAY is used only to label events and to print
the coverage report.

Typical weekly workflow
-----------------------
    # 1. re-export from MyTimetable, drop the file in this folder
    # 2. see what moved before touching anything
    python3 dcu_to_ics.py --xlsx Timetables.xlsx --diff DCU_Semester1_2026.ics
    # 3. rebuild and publish
    python3 dcu_to_ics.py --xlsx Timetables.xlsx
    git add -A && git commit -m "Timetable refresh" && git push
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

try:
    import pandas as pd
except ImportError:  # pragma: no cover
    sys.exit("error: pandas is required — pip install pandas openpyxl")

# --------------------------------------------------------------------------- #
# Configuration                                                                #
# --------------------------------------------------------------------------- #

WEEK1_MONDAY = date(2026, 9, 7)  # DCU: "Semester 1 teaching starts"

CALENDAR_NAME = "DCU — Semester 1 2026/27"
CALENDAR_DESC = "DCU Glasnevin — teaching timetable."
TZID = "Europe/Dublin"

CAMPUS_SUFFIX = "DCU Glasnevin Campus, Collins Avenue Ext, Whitehall, Dublin 9"
CAMPUS_LAT, CAMPUS_LON = 53.3858, -6.2565

# University closure days. MyTimetable still lists classes on these dates, but
# DCU states that scheduled classes do not take place on public holidays.
CLOSURES: dict[date, str] = {
    date(2026, 10, 26): "October Bank Holiday — University closed",
}

KEY_DATES: list[tuple[date, date, str]] = [
    (date(2026, 9, 7),   date(2026, 9, 7),   "Semester 1 teaching starts"),
    (date(2026, 10, 26), date(2026, 10, 26), "University closed — public holiday"),
    (date(2026, 11, 27), date(2026, 11, 27), "End of Semester 1 teaching"),
    (date(2026, 11, 30), date(2026, 12, 6),  "Exam study period"),
    (date(2026, 12, 7),  date(2026, 12, 19), "Exam period — Semester 1"),
]

PRODID = "-//DCU Timetable Generator//EN"

REQUIRED_COLUMNS = ["Module Name", "Location", "Date", "Day", "Time range"]

VTIMEZONE = """BEGIN:VTIMEZONE
TZID:Europe/Dublin
X-LIC-LOCATION:Europe/Dublin
BEGIN:DAYLIGHT
TZNAME:IST
TZOFFSETFROM:+0000
TZOFFSETTO:+0100
DTSTART:19700329T010000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZNAME:GMT
TZOFFSETFROM:+0100
TZOFFSETTO:+0000
DTSTART:19701025T020000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE"""



# --------------------------------------------------------------------------- #
# Room codes                                                                   #
# --------------------------------------------------------------------------- #

BUILDINGS: dict[str, dict] = {}
CAMPUSES: dict[str, str] = {}

FLOOR_NAMES = {
    "G": "ground floor", "0": "ground floor", "1": "1st floor", "2": "2nd floor",
    "3": "3rd floor", "4": "4th floor", "5": "5th floor", "6": "6th floor",
}

# GLA.SA301 -> campus 'GLA', tail 'SA301'
ROOM_RE = re.compile(r"^([A-Za-z]{2,3})\.([A-Za-z0-9]+)$")


def load_buildings(path: Path) -> list[str]:
    """Load the campus building table. Absent file = room codes stay raw."""
    global BUILDINGS, CAMPUSES
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        sys.exit(f"error: {path} is not valid JSON — {exc}")
    BUILDINGS = {k.upper(): v for k, v in data.get("buildings", {}).items()}
    CAMPUSES = {k.upper(): v for k, v in data.get("campuses", {}).items()}
    located = sum(1 for b in BUILDINGS.values() if "lat" in b)
    return [f"buildings: {len(BUILDINGS)} known, {located} with coordinates"]


def parse_room(code: str) -> dict | None:
    """'GLA.SA301' -> Stokes Extension, 3rd floor, room 01.

    Building codes are one or two characters, followed by a floor (G or a
    digit) and the room number. The two-character reading is tried first but
    only accepted when the code is known AND a valid floor follows, so
    'GLA.SG16' resolves to Stokes / ground / 16 rather than an unknown 'SG'.
    """
    match = ROOM_RE.match(code.strip())
    if not match:
        return None
    campus, tail = match.group(1).upper(), match.group(2).upper()

    def split_at(width: int) -> tuple[str, str, str] | None:
        head, rest = tail[:width], tail[width:]
        if head in BUILDINGS and rest and rest[0] in FLOOR_NAMES:
            return head, rest[0], rest[1:]
        return None

    parts = split_at(2) or split_at(1)
    if parts is None:
        return None
    building, floor, room = parts
    entry = BUILDINGS[building]

    return {
        "campus": CAMPUSES.get(campus, campus),
        "building_code": building,
        "building": entry.get("name", building),
        "floor": FLOOR_NAMES[floor],
        "room": room,
        "lat": entry.get("lat"),
        "lon": entry.get("lon"),
    }


def describe_rooms(raw: str) -> tuple[str, float | None, float | None, str | None]:
    """'GLA.SG16, GLA.SG15' -> ('GLA.SG16, GLA.SG15 — Stokes Building, ground floor', lat, lon, building)

    Returns the display string, the coordinates of the first resolved
    building, and that building's name — or (raw, None, None, None) when
    nothing resolves.
    """
    codes = [c.strip() for c in raw.split(",") if c.strip()]
    parsed = [(c, parse_room(c)) for c in codes]
    resolved = [p for _, p in parsed if p]
    if not resolved:
        return raw, None, None, None

    # One label per distinct building+floor, in first-seen order.
    labels: list[str] = []
    for info in resolved:
        label = f"{info['building']}, {info['floor']}"
        if label not in labels:
            labels.append(label)
    display = f"{raw} — {' / '.join(labels)}"
    return display, resolved[0]["lat"], resolved[0]["lon"], resolved[0]["building"]


def load_config(path: Path) -> list[str]:
    """Override the module-level settings above from a JSON file.

    Every key is optional: anything absent keeps the default. Returns a list of
    human-readable notes for the run report.
    """
    global WEEK1_MONDAY, CALENDAR_NAME, CALENDAR_DESC, TZID
    global CAMPUS_SUFFIX, CAMPUS_LAT, CAMPUS_LON, CLOSURES, KEY_DATES

    if not path.exists():
        return []
    try:
        cfg = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        sys.exit(f"error: {path} is not valid JSON — {exc}")

    notes = [f"config: {path}"]

    def iso(value: str, field: str) -> date:
        try:
            return date.fromisoformat(value)
        except ValueError:
            sys.exit(f"error: {path}: {field} must be YYYY-MM-DD, got {value!r}")

    if "week1_monday" in cfg:
        WEEK1_MONDAY = iso(cfg["week1_monday"], "week1_monday")
        if WEEK1_MONDAY.weekday() != 0:
            sys.exit(f"error: {path}: week1_monday must be a Monday")
    CALENDAR_NAME = cfg.get("calendar_name", CALENDAR_NAME)
    CALENDAR_DESC = cfg.get("calendar_desc", CALENDAR_DESC)
    TZID = cfg.get("tzid", TZID)
    if TZID != "Europe/Dublin":
        notes.append(f"warning: tzid={TZID!r} but the embedded VTIMEZONE block "
                      f"is hardcoded to Europe/Dublin — event times would be "
                      f"mislabeled. Edit the VTIMEZONE constant to match.")
    CAMPUS_SUFFIX = cfg.get("campus_suffix", CAMPUS_SUFFIX)
    CAMPUS_LAT = cfg.get("campus_lat", CAMPUS_LAT)
    CAMPUS_LON = cfg.get("campus_lon", CAMPUS_LON)

    if "closures" in cfg:
        CLOSURES = {iso(d, "closures"): reason for d, reason in cfg["closures"].items()}
        notes.append(f"{len(CLOSURES)} closure day(s)")
    if "key_dates" in cfg:
        KEY_DATES = [
            (iso(item["start"], "key_dates.start"),
             iso(item.get("end", item["start"]), "key_dates.end"),
             item["title"])
            for item in cfg["key_dates"]
        ]
        notes.append(f"{len(KEY_DATES)} key date(s)")
    return notes


# --------------------------------------------------------------------------- #
# iCalendar primitives                                                         #
# --------------------------------------------------------------------------- #

def escape(value: str) -> str:
    """Escape a TEXT value per RFC 5545 section 3.3.11."""
    return (
        value.replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\n", "\\n")
    )


def param(value: str) -> str:
    """Quote an RFC 5545 parameter value (e.g. X-TITLE="...").

    Parameter values have no escaping mechanism: a quoted string may contain
    anything except a double quote and control characters. So we strip the
    characters that cannot appear rather than backslash-escaping them, which
    some parsers read literally and others reject.
    """
    cleaned = value.replace('"', "'").replace("\n", " ").replace("\r", " ")
    return f'"{cleaned}"'


def fold(line: str) -> str:
    """Fold a content line to 75 octets; continuations get a leading space."""
    raw = line.encode("utf-8")
    if len(raw) <= 75:
        return line
    chunks, start, limit = [], 0, 75
    while start < len(raw):
        end = min(start + limit, len(raw))
        # If the byte right after the cut is a UTF-8 continuation byte, the
        # boundary falls inside a multi-byte character (e.g. the — used in
        # every room description) — back up to before its lead byte instead.
        while end > start and end < len(raw) and (raw[end] & 0xC0) == 0x80:
            end -= 1
        chunks.append(raw[start:end].decode("utf-8"))
        start, limit = end, 74
    return "\r\n ".join(chunks)


def unfold(text: str) -> list[str]:
    """Inverse of fold — used by --diff to read a previously generated file."""
    return re.sub(r"\r?\n[ \t]", "", text).splitlines()


def slugify(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", value)


def stamp_utc() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


# --------------------------------------------------------------------------- #
# Parsing                                                                      #
# --------------------------------------------------------------------------- #

def parse_module(raw: str) -> tuple[str, str]:
    """'EEG1011[1] Engineering Mathematics IV' -> ('EEG1011', 'Engineering ...')"""
    match = re.match(r"\s*([A-Z]{2,4}\d{3,5})(?:\[\d+\])?\s*(.*)", str(raw).strip())
    if not match:
        return str(raw).strip(), ""
    return match.group(1), match.group(2).strip()


def parse_time_range(raw: str) -> tuple[str, str]:
    start, _, end = str(raw).strip().partition("-")
    return start.strip(), end.strip()


def week_number(day: date) -> int:
    return (day - WEEK1_MONDAY).days // 7 + 1


def load_overrides(path: Path | None) -> dict[tuple[str, str, str], dict]:
    """module_code|day|start -> {kind, lecturer, notes}"""
    if not path or not path.exists():
        return {}
    table: dict[tuple[str, str, str], dict] = {}
    with path.open(newline="", encoding="utf-8-sig") as fh:
        # '#' comments are stripped so the file can document itself.
        stripped = (line for line in fh if not line.lstrip().startswith("#"))
        reader = csv.DictReader(stripped)
        for lineno, row in enumerate(reader, start=2):
            row = {k: (v or "").strip() for k, v in row.items() if k}
            if not row.get("module_code"):
                continue
            for field in ("day", "start"):
                if not row.get(field):
                    sys.exit(f"error: {path} line {lineno}: '{field}' is required")
            key = (row["module_code"], row["day"].lower(), row["start"])
            if key in table:
                sys.exit(f"error: {path} line {lineno}: duplicate entry for "
                         f"{row['module_code']} {row['day']} {row['start']}")
            table[key] = row
    return table


def load_sessions(xlsx: Path, overrides: dict, used_overrides: set) -> list[dict]:
    frame = pd.read_excel(xlsx)
    missing = [c for c in REQUIRED_COLUMNS if c not in frame.columns]
    if missing:
        sys.exit(f"error: {xlsx} is missing column(s): {', '.join(missing)}")

    sessions: list[dict] = []
    for _, row in frame.iterrows():
        code, name = parse_module(row["Module Name"])
        start, end = parse_time_range(row["Time range"])
        day = datetime.strptime(str(row["Date"]).strip(), "%d/%m/%Y").date()
        rooms = str(row["Location"]).strip()
        key = (code, str(row["Day"]).strip().lower(), start)
        if key in overrides:
            used_overrides.add(key)
        extra = overrides.get(key, {})
        sessions.append(
            {
                "code": code,
                "name": name,
                "date": day,
                "weekday": str(row["Day"]).strip(),
                "start": start,
                "end": end,
                "rooms": rooms,
                "weeks": str(row.get("Weeks", "")).strip(),
                "kind": extra.get("kind", ""),
                "lecturer": extra.get("lecturer", ""),
                "notes": extra.get("notes", ""),
            }
        )

    # The export can repeat a row when an activity spans two joined rooms.
    seen, unique = set(), []
    for session in sessions:
        key = (session["code"], session["date"], session["start"],
               session["end"], session["rooms"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(session)

    unique.sort(key=lambda s: (s["date"], s["start"], s["code"]))
    return unique


def merge_adjacent(sessions: list[dict]) -> list[dict]:
    """Fuse back-to-back sessions of the same module, room, type and day."""
    merged: list[dict] = []
    for session in sessions:
        previous = next(
            (
                m for m in reversed(merged)
                if m["date"] == session["date"]
                and m["code"] == session["code"]
                and m["rooms"] == session["rooms"]
                and m["kind"] == session["kind"]
                and m["end"] == session["start"]
            ),
            None,
        )
        if previous:
            previous["end"] = session["end"]
        else:
            merged.append(dict(session))
    return merged


# --------------------------------------------------------------------------- #
# Event building                                                               #
# --------------------------------------------------------------------------- #

def make_uid(session: dict) -> str:
    # Identity only — never presentation. Room, session type, lecturer and notes
    # are deliberately excluded so that editing them updates the event in place
    # instead of deleting it and creating a new one.
    fingerprint = "|".join([session["code"], session["weekday"], session["start"]])
    digest = hashlib.sha1(fingerprint.encode("utf-8")).hexdigest()[:10]
    return f"{session['date']:%Y%m%d}-{slugify(session['code'])}-{digest}@dcu.timetable"


def hhmm(value: str) -> str:
    hh, mm = value.split(":")
    return f"{int(hh):02d}{int(mm):02d}00"


def event_fields(session: dict) -> dict:
    """Compute every RFC 5545 field value for a session.

    Used both to render the VEVENT (build_event) and to detect real content
    changes across runs (main), so that DTSTAMP/SEQUENCE only move when
    something a subscriber would notice actually changed.
    """
    title = f"{session['code']} {session['name']}".strip()
    if session["kind"]:
        title = f"{title} ({session['kind']})"

    described, lat, lon, building = describe_rooms(session["rooms"])
    location = f"{described}, {CAMPUS_SUFFIX}" if CAMPUS_SUFFIX else described
    lat = CAMPUS_LAT if lat is None else lat
    lon = CAMPUS_LON if lon is None else lon

    # A short, human pin name ("DCU Stokes Extension") plus the campus postal
    # address, used for the tap-to-open Apple Maps link and the structured
    # location Apple's Calendar app uses for "Time to Leave" alerts.
    campus_first_line = next((p.strip() for p in CAMPUS_SUFFIX.split(",") if p.strip()), "")
    pin_title = f"DCU {building}" if building else (campus_first_line or "DCU Glasnevin Campus")
    pin_address = "\\n".join(p.strip() for p in CAMPUS_SUFFIX.split(",") if p.strip())
    maps_url = f"https://maps.apple.com/?ll={lat},{lon}&q={quote(pin_title)}&z=17"

    parts = [f"Module: {session['code']} — {session['name']}"]
    if session["kind"]:
        parts.append(f"Type: {session['kind']}")
    parts.append(f"Room: {described}")
    if session["lecturer"]:
        parts.append(f"Lecturer: {session['lecturer']}")
    parts.append(f"Teaching week {week_number(session['date'])}")
    if session["notes"]:
        parts.append(session["notes"])
    parts.append(f"Map: {maps_url}")

    return {
        "uid": make_uid(session),
        "summary": title,
        "location": location,
        "description": "\n".join(parts),
        "dtstart": f"{session['date']:%Y%m%d}T{hhmm(session['start'])}",
        "dtend": f"{session['date']:%Y%m%d}T{hhmm(session['end'])}",
        "lat": lat,
        "lon": lon,
        "pin_title": pin_title,
        "pin_address": pin_address,
        "room_raw": session["rooms"],
        "maps_url": maps_url,
    }


def build_event(session: dict, dtstamp: str, alarm: int | None, sequence: int = 0) -> list[str]:
    f = event_fields(session)

    lines = [
        "BEGIN:VEVENT",
        f"UID:{f['uid']}",
        f"DTSTAMP:{dtstamp}",
        f"DTSTART;TZID={TZID}:{f['dtstart']}",
        f"DTEND;TZID={TZID}:{f['dtend']}",
        f"SUMMARY:{escape(f['summary'])}",
        f"LOCATION:{escape(f['location'])}",
        f"DESCRIPTION:{escape(f['description'])}",
        f"CATEGORIES:{escape(session['code'])}",
        "STATUS:CONFIRMED",
        "TRANSP:OPAQUE",
        f"SEQUENCE:{sequence}",
        (
            "X-APPLE-STRUCTURED-LOCATION;VALUE=URI"
            f";X-ADDRESS={param(f['pin_address'])}"
            ";X-APPLE-RADIUS=80;X-APPLE-REFERENCEFRAME=1"
            f";X-TITLE={param(f['pin_title'])}"
            f":geo:{f['lat']},{f['lon']}"
        ),
        f"GEO:{f['lat']};{f['lon']}",
        f"URL;VALUE=URI:{f['maps_url']}",
    ]
    if alarm:
        lines += [
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            f"DESCRIPTION:{escape(f['summary'])}",
            f"TRIGGER:-PT{alarm}M",
            "END:VALARM",
        ]
    lines.append("END:VEVENT")
    return lines


def key_date_uid(start: date, title: str) -> str:
    digest = hashlib.sha1(title.encode("utf-8")).hexdigest()[:10]
    return f"{start:%Y%m%d}-key-{digest}@dcu.timetable"


def build_key_date(start: date, end_inclusive: date, title: str, dtstamp: str,
                   sequence: int = 0) -> list[str]:
    return [
        "BEGIN:VEVENT",
        f"UID:{key_date_uid(start, title)}",
        f"DTSTAMP:{dtstamp}",
        f"DTSTART;VALUE=DATE:{start:%Y%m%d}",
        f"DTEND;VALUE=DATE:{end_inclusive + timedelta(days=1):%Y%m%d}",
        f"SUMMARY:{escape(title)}",
        "DESCRIPTION:Source: DCU Registry\\, Academic Calendar 2026/27.",
        "CATEGORIES:Key dates",
        "TRANSP:TRANSPARENT",
        f"SEQUENCE:{sequence}",
        "END:VEVENT",
    ]


def wrap_calendar(body: list[str], name: str, ttl_hours: int) -> str:
    header = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:{PRODID}",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"NAME:{escape(name)}",
        f"X-WR-CALNAME:{escape(name)}",
        f"X-WR-CALDESC:{escape(CALENDAR_DESC)}",
        f"X-WR-TIMEZONE:{TZID}",
        f"REFRESH-INTERVAL;VALUE=DURATION:PT{ttl_hours}H",
        f"X-PUBLISHED-TTL:PT{ttl_hours}H",
        *VTIMEZONE.split("\n"),
    ]
    return "\r\n".join(fold(l) for l in header + body + ["END:VCALENDAR"]) + "\r\n"


# --------------------------------------------------------------------------- #
# Reading a previously generated file (used by --diff and for DTSTAMP/SEQUENCE
# stability across runs)                                                       #
# --------------------------------------------------------------------------- #

TRACKED_FIELDS = {
    "UID", "SUMMARY", "LOCATION", "DESCRIPTION", "DTSTART", "DTEND",
    "DTSTAMP", "SEQUENCE",
}


def read_events_from_lines(lines: list[str]) -> dict[str, dict]:
    events: dict[str, dict] = {}
    current: dict | None = None
    nested = 0  # depth inside a sub-component (VALARM, ...) of the current VEVENT
    for line in lines:
        if line == "BEGIN:VEVENT":
            current, nested = {}, 0
            continue
        if line == "END:VEVENT" and current is not None:
            events[current.get("UID", "")] = current
            current = None
            continue
        if current is None:
            continue
        # VALARM (etc.) has its own DESCRIPTION — skip its contents entirely so
        # it can't clobber the VEVENT's own fields of the same name.
        if line.startswith("BEGIN:"):
            nested += 1
            continue
        if line.startswith("END:"):
            nested -= 1
            continue
        if nested:
            continue
        name, _, value = line.partition(":")
        key = name.split(";")[0]
        if key in TRACKED_FIELDS:
            current[key] = value
    return events


def read_events(path: Path) -> dict[str, dict]:
    return read_events_from_lines(unfold(path.read_text(encoding="utf-8")))


def stamp_for(uid: str, snapshot: dict[str, str], previous: dict[str, dict],
              now_stamp: str) -> tuple[str, int]:
    """DTSTAMP/SEQUENCE for one event, reusing the previous run's values when
    `snapshot` (the fields a subscriber would notice) is unchanged, so that
    regenerating the feed without a real timetable change leaves the .ics
    byte-identical."""
    old = previous.get(uid)
    if not old:
        return now_stamp, 0
    if all(old.get(k) == v for k, v in snapshot.items()):
        return old.get("DTSTAMP") or now_stamp, int(old.get("SEQUENCE") or 0)
    return now_stamp, int(old.get("SEQUENCE") or 0) + 1


def print_diff(old: Path, new_body: list[str]) -> None:
    if not old.exists():
        print(f"diff: {old} does not exist yet — nothing to compare")
        return
    before = read_events(old)
    after = read_events_from_lines(new_body)

    added = sorted(set(after) - set(before))
    removed = sorted(set(before) - set(after))
    changed = [
        uid for uid in set(before) & set(after)
        if any(before[uid].get(f) != after[uid].get(f)
               for f in ("SUMMARY", "LOCATION", "DTSTART", "DTEND", "DESCRIPTION"))
    ]

    if not (added or removed or changed):
        print("diff: no change")
        return
    for uid in removed:
        e = before[uid]
        print(f"  - REMOVED  {e.get('DTSTART','?')}  {e.get('SUMMARY','?')}")
    for uid in added:
        e = after[uid]
        print(f"  + ADDED    {e.get('DTSTART','?')}  {e.get('SUMMARY','?')}")
    for uid in sorted(changed):
        b, a = before[uid], after[uid]
        print(f"  ~ CHANGED  {a.get('DTSTART','?')}  {a.get('SUMMARY','?')}")
        for field in ("DTSTART", "DTEND", "LOCATION", "SUMMARY"):
            if b.get(field) != a.get(field):
                print(f"             {field}: {b.get(field)} -> {a.get(field)}")
        if b.get("DESCRIPTION") != a.get("DESCRIPTION"):
            print("             DESCRIPTION: changed (room/lecturer/notes/map link)")
    print(f"diff: {len(added)} added, {len(removed)} removed, {len(changed)} changed")


# --------------------------------------------------------------------------- #
# Main                                                                         #
# --------------------------------------------------------------------------- #

def main() -> int:
    # Windows consoles often default to a non-UTF-8 code page, which mangles
    # the em dashes in the run report. Best-effort; never fatal.
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass

    parser = argparse.ArgumentParser(
        description="Build an Apple Calendar .ics from a DCU MyTimetable Excel export"
    )
    parser.add_argument("--xlsx", type=Path, default=Path("Timetables.xlsx"))
    parser.add_argument("--overrides", type=Path, default=Path("overrides.csv"))
    parser.add_argument("--config", type=Path, default=Path("config.json"),
                        help="semester settings; ignored if the file is absent")
    parser.add_argument("--buildings", type=Path, default=Path("buildings.json"),
                        help="campus building names and coordinates")
    parser.add_argument("--out", type=Path, default=Path("DCU_Semester1_2026.ics"))
    parser.add_argument("--calendar-name", type=str, default=None,
                        help="override the calendar's display name (X-WR-CALNAME); "
                             "useful for building several feeds (e.g. one per group) "
                             "from the same config.json")
    parser.add_argument("--alarm", type=int, default=15,
                        help="minutes before the event; 0 disables alarms")
    parser.add_argument("--no-key-dates", action="store_true")
    parser.add_argument("--merge-adjacent", action="store_true")
    parser.add_argument("--split", action="store_true",
                        help="also write one .ics per module (one colour each)")
    parser.add_argument("--keep-closures", action="store_true",
                        help="do not drop sessions falling on university closure days")
    parser.add_argument("--ttl", type=int, default=6)
    parser.add_argument("--diff", type=Path,
                        help="compare against a previously generated .ics and exit")
    args = parser.parse_args()

    if not args.xlsx.exists():
        sys.exit(f"error: {args.xlsx} not found")

    config_notes = load_config(args.config) + load_buildings(args.buildings)

    overrides = load_overrides(args.overrides)
    used_overrides: set = set()
    sessions = load_sessions(args.xlsx, overrides, used_overrides)
    if args.merge_adjacent:
        sessions = merge_adjacent(sessions)

    kept, dropped = [], []
    for session in sessions:
        if session["date"] in CLOSURES and not args.keep_closures:
            dropped.append(session)
        else:
            kept.append(session)

    alarm = args.alarm if args.alarm > 0 else None

    if args.diff:
        stamp = stamp_utc()
        body: list[str] = []
        for session in kept:
            body += build_event(session, stamp, alarm)
        if not args.no_key_dates:
            for start, end, title in KEY_DATES:
                body += build_key_date(start, end, title, stamp)
        print_diff(args.diff, body)
        return 0

    # DTSTAMP/SEQUENCE stability: reuse the previous run's values for any event
    # whose visible content (time, room, title, description) hasn't changed,
    # so `git diff` on the .ics only ever shows real timetable changes.
    previous = read_events(args.out) if args.out.exists() else {}
    now_stamp = stamp_utc()

    body: list[str] = []
    for session in kept:
        f = event_fields(session)
        snapshot = {
            "SUMMARY": escape(f["summary"]),
            "LOCATION": escape(f["location"]),
            "DESCRIPTION": escape(f["description"]),
            "DTSTART": f["dtstart"],
            "DTEND": f["dtend"],
        }
        dtstamp, sequence = stamp_for(f["uid"], snapshot, previous, now_stamp)
        session["_dtstamp"], session["_sequence"] = dtstamp, sequence
        body += build_event(session, dtstamp, alarm, sequence)
    if not args.no_key_dates:
        for start, end, title in KEY_DATES:
            uid = key_date_uid(start, title)
            snapshot = {
                "SUMMARY": escape(title),
                "DTSTART": f"{start:%Y%m%d}",
                "DTEND": f"{end + timedelta(days=1):%Y%m%d}",
            }
            dtstamp, sequence = stamp_for(uid, snapshot, previous, now_stamp)
            body += build_key_date(start, end, title, dtstamp, sequence)

    calendar_name = args.calendar_name or CALENDAR_NAME
    args.out.write_text(wrap_calendar(body, calendar_name, args.ttl),
                         encoding="utf-8", newline="")

    # ---- report -----------------------------------------------------------
    per_module: dict[str, int] = defaultdict(int)
    per_week: dict[int, int] = defaultdict(int)
    for session in kept:
        per_module[session["code"]] += 1
        per_week[week_number(session["date"])] += 1

    print(f"{args.out}: {len(kept)} class events, {len(per_module)} modules")
    for note in config_notes:
        print(f"  {note}")
    for code in sorted(per_module):
        print(f"  {code:<9} {per_module[code]:>3}")
    if per_week:
        span = f"weeks {min(per_week)}–{max(per_week)}"
        print(f"  coverage: {span}, "
              f"{min(per_week.values())}–{max(per_week.values())} sessions per week")
    else:
        print("  coverage: no sessions in output")
    for session in dropped:
        reason = CLOSURES[session["date"]]
        print(f"  dropped  {session['date']:%d/%m} {session['start']} "
              f"{session['code']} — {reason}")
    unresolved = sorted({
        c.strip() for s in kept for c in s["rooms"].split(",")
        if c.strip() and parse_room(c) is None
    })
    if unresolved:
        print(f"  room code not in buildings.json: {', '.join(unresolved)}")
    no_geo = sorted({
        info["building_code"] for s in kept for c in s["rooms"].split(",")
        if (info := parse_room(c)) and info["lat"] is None
    })
    if no_geo:
        print(f"  no coordinates for building(s) {', '.join(no_geo)} "
              f"— falling back to the campus centre")

    missing = sorted({
        f"{s['code']} {s['weekday'][:3]} {s['start']}"
        for s in kept if not s["kind"]
    })
    if missing:
        print(f"  no session type set for {len(missing)} activities: "
              + ", ".join(missing))

    unused = sorted(set(overrides) - used_overrides)
    if unused:
        print(f"  overrides.csv: {len(unused)} row(s) never matched a session "
              f"(stale after a re-export?): "
              + ", ".join(f"{code} {day} {start}" for code, day, start in unused))

    if args.split:
        by_code: dict[str, list[str]] = defaultdict(list)
        for session in kept:
            by_code[session["code"]] += build_event(
                session, session["_dtstamp"], alarm, session["_sequence"])
        for code, lines in sorted(by_code.items()):
            path = args.out.with_name(f"{args.out.stem}_{code}{args.out.suffix}")
            path.write_text(wrap_calendar(lines, f"DCU {code}", args.ttl),
                            encoding="utf-8", newline="")
            print(f"{path}: {per_module[code]} events")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
