# DCU Semester 1 2026/27 — Apple Calendar feed

[![Rebuild status](https://github.com/Gaspard-SOLUS/dcu-calendar/actions/workflows/build.yml/badge.svg)](https://github.com/Gaspard-SOLUS/dcu-calendar/actions/workflows/build.yml)

Built directly from the MyTimetable Excel export, so a refresh is a re-export
plus one command — no hand-editing of the schedule itself.

`index.html` carries a floating "Signaler une erreur" button and a link in the
"Si un cours change" section — both `mailto:` to gaspard.solus2@mail.dcu.ie
with a pre-filled template (module / day+time / what's wrong). It also has a
"Copier" button next to the feed URL, an emoji favicon, Open Graph/Twitter
meta tags for a proper preview card when the link is shared in WhatsApp/
Messages, and a dark theme that follows the OS/browser preference
automatically (`prefers-color-scheme`, no toggle).

| File | Role |
|---|---|
| `Timetables.xlsx` | The MyTimetable export. Replace it wholesale on every refresh. |
| `overrides.csv` | Your enrichment layer: session type, lecturer, notes. Survives re-exports. |
| `exceptions.csv` | Dated one-off changes from lecturers' own schedules: cancel, shorten, re-room, give a topic (Lab 1: Logic, Loop Quiz 1). |
| `deadlines.csv` | Coursework deadlines, published as their own events. |
| `config.json` | Semester dates, closures, key dates, campus address. |
| `buildings.json` | Campus building names and per-building GPS coordinates. |
| `dcu_to_ics.py` | Generator. Needs `pandas` + `openpyxl`. |
| `DCU_Semester1_2026.ics` | Output: 129 class events + 8 deadlines + 5 all-day academic key dates. |
| `tests/` | `pytest` unit tests for the parsing/escaping/UID logic. |
| `.github/workflows/build.yml` | Rebuilds and commits the feed automatically on push. |

```
EEG1011  37    EEN1018  25    EEN1022  43    EEN1083  12    ESL1009  12
```

Weeks 1–12 covered, 8 to 13 sessions per week, plus the ESL1009 Assessment 3
session added in week 13.

---

## 1. How the two inputs combine

The Excel export lists every occurrence with its real date, so the generator does
no week arithmetic — it copies dates straight across. That means per-week
exceptions come through automatically:

- EEN1083 sits in `GLA.SA105` in week 2 and `GLA.FT307` everywhere else.
- EEG1011 Wednesday moves from `GLA.S144` to `GLA.S143` in week 12.
- ESL1009 has no class in week 7 — SALIS observes the reading week (19–25/10),
  Engineering & Computing does not.
- EEN1018 Wednesday practical runs in even weeks only (2, 4, 6, 8).

`overrides.csv` is keyed on `module_code | day | start`, never on a date, so it
keeps matching after a re-export. Adding a row changes the event title from
`EEN1018 Circuits Analysis Techniques` to
`EEN1018 Circuits Analysis Techniques (Practical)`.

Three activities still have no `kind` set — the Excel export drops the activity
code that carries it. Read them off MyTimetable and fill them in:

| Activity | Code to look for |
|---|---|
| EEG1011 Tue 10:00 (week 6 only) | |
| EEG1011 Thu 14:00 (week 7 only) | |
| EEN1083 Tue 14:00 | |

MyTimetable's naming convention: `OC/L…` = Lecture, `OC/P…` = Practical,
`OC/T…` = Tutorial. The trailing digits are the group number.

### Lecturers' own schedules: `exceptions.csv` and `deadlines.csv`

Some lecturers publish a module schedule that is more precise than
MyTimetable: no lab in week 1, kit distribution 2–3pm only, "Lab 3
(Transistor)", tutorials instead of lectures in week 12, Loop quizzes on a
given Wednesday. MyTimetable never shows any of that, and `overrides.csv`
can't express it because it has no date.

`exceptions.csv` is keyed on `date | module_code | start` — `start` being the
start time *as the export shows it*, never changed, so the event keeps its UID
and updates in place. Per row: `action` (empty = adjust, `cancel` = remove,
`add` = create a session the export doesn't have, e.g. an assessment in week
13 — needs `end` and `rooms`), `end`, `rooms`, `kind`, `topic` (appended to
the title after an em dash and shown as a `Topic:` line) and `notes` (appended
to the description). Every run reports cancelled and added sessions and lists
any row that **never matched a session** — a wrong date or start time fails
loudly there, not silently. An `add` on a slot the export already has is
ignored and reported the same way, never duplicated.

`deadlines.csv` (`date, time, module_code, title, notes`) adds events of its
own, category `Deadlines` plus the module code, free (`TRANSP:TRANSPARENT`).
A deadline is **always a precise moment** — `time` defaults to 23:59, and
there is deliberately no all-day or multi-day form. Each one gets two
reminders, 30 h and 6 h before (about 18:00 the day before and 18:00 on the
day). When a lecturer only gives the week, pick a day and say so in `notes`
(the EEN1018 assignments use the Friday). Something that happens during a
class (a lab, an in-class test, a quiz) belongs on the session as a `topic`
instead.

Both files are optional and `#` comments are allowed. A value containing a
comma must be quoted — a line with more fields than the header is an error.
With `--split`, each module's deadlines go into that module's file.
`--no-deadlines` leaves them out.

Sources currently encoded: EEN1022 and EEN1018 module schedules from the
lecturers, the EEG1011 Loop-quiz email, the ESL1009 assessment email.

## 2. Closures

MyTimetable still shows Monday classes on **26/10/2026**, but DCU states that
scheduled classes do not take place on public holidays, and the Registry calendar
lists that date as "University Closed". Three sessions are dropped and reported on
every run. `--keep-closures` puts them back; `CLOSURES` at the top of the script
is where you add more.

## 3. Commands

```bash
pip install -r requirements.txt                         # once (pandas + openpyxl)

python3 dcu_to_ics.py                                   # default build
python3 dcu_to_ics.py --alarm 25                        # 25-minute reminder
python3 dcu_to_ics.py --alarm 0                         # no reminders
python3 dcu_to_ics.py --merge-adjacent                  # fuse back-to-back sessions
python3 dcu_to_ics.py --split                           # one .ics per module
python3 dcu_to_ics.py --no-key-dates                    # classes only
python3 dcu_to_ics.py --no-deadlines                    # no deadlines.csv events
python3 dcu_to_ics.py --keep-closures                   # keep bank-holiday sessions
python3 dcu_to_ics.py --ttl 2                           # suggest a 2-hour refresh
python3 dcu_to_ics.py --diff DCU_Semester1_2026.ics     # dry run: what would change
```

Regenerating with unchanged inputs produces a byte-identical file: each event
keeps its `DTSTAMP`/`SEQUENCE` from the last run unless its time, room, title
or description actually changed, so `git diff` on the `.ics` only ever shows
real timetable changes, never rebuild noise. This relies on the `.ics` keeping
the exact `\r\n` line endings the generator writes — `.gitattributes` marks
`*.ics -text` so Git (notably on Windows, where `core.autocrlf` would otherwise
double every line ending to `\r\r\n` and corrupt the feed) never touches them.

`--split` matters if you want colours: Apple assigns one colour per calendar,
never per event. Five files, five subscriptions, five colours.

### Tests

```bash
pip install -r requirements-dev.txt
pytest -q
```

Covers escaping/line-folding (including UTF-8 characters like the — in every
room description), room-code parsing, UID stability, and the DTSTAMP/SEQUENCE
stability logic above.

## 4. Weekly refresh

```bash
# 1. MyTimetable -> refresh icon -> Multiple weeks -> select weeks 1-12 -> EXCEL
#    Save over Timetables.xlsx.

# 2. See what moved, before changing anything:
python3 dcu_to_ics.py --diff DCU_Semester1_2026.ics

# 3. Rebuild and publish:
python3 dcu_to_ics.py
git add -A && git commit -m "Timetable refresh $(date +%F)" && git push
```

`--diff` prints added / removed / changed sessions with the old and new values.
Run it before every rebuild — it is the only thing that tells you a room moved.

Steps 2–3 also run automatically: `.github/workflows/build.yml` rebuilds and
commits the feed on every push that touches `Timetables.xlsx`, `overrides.csv`,
`exceptions.csv`, `deadlines.csv`, `config.json`, `buildings.json` or
`dcu_to_ics.py`. So the manual commands above are for
previewing the diff locally before you push — pushing alone is enough to
publish.

UIDs are `date + module code + weekday + start time` (SHA-1 digest of that
fingerprint). Room, session type, lecturer and notes are deliberately left out
of the fingerprint, so editing any of those updates the event in place instead
of deleting and recreating it — and a session dropped from a re-export simply
disappears from the feed rather than lingering.

## 5. Publishing and subscribing

DCU does not support calendar subscriptions to MyTimetable — the option existed at
one point but was never officially supported and has been removed. Self-hosting is
the only route.

```bash
git init dcu-calendar && cd dcu-calendar
cp ../{Timetables.xlsx,overrides.csv,dcu_to_ics.py,DCU_Semester1_2026.ics,README.md} .
git add -A && git commit -m "Initial DCU timetable feed"
git branch -M main
git remote add origin git@github.com:<your-user>/dcu-calendar.git
git push -u origin main
```

**Settings → Pages → Deploy from a branch → main / (root)**, then the feed lives at:

```
https://<your-user>.github.io/dcu-calendar/DCU_Semester1_2026.ics
```

**macOS** — Calendar → File → New Calendar Subscription…, paste the URL, set
*Auto-refresh* to Every hour.

**iPhone** — Settings → Apps → Calendar → Accounts → Add Account → Other → Add
Subscribed Calendar. Swapping `https://` for `webcal://` makes a tapped link open
the subscription dialog directly.

Subscribe on **one** device only and let iCloud propagate, or you get duplicates.
A subscribed calendar is read-only on your devices — keep a separate local calendar
for deadlines and study blocks.

If you would rather not make the repository public, a private GitHub Gist gives an
unguessable `gist.githubusercontent.com` raw URL that works the same way. It is
obscure rather than protected: anyone with the link can read your schedule.

## 6. Customisation points

In `dcu_to_ics.py` (or `config.json`/`buildings.json` where noted):

- `CAMPUS_SUFFIX`, `CAMPUS_LAT`, `CAMPUS_LON` (`config.json`) — the room code stays
  first in `LOCATION`; the campus address is appended so Apple Maps can geocode the
  event and offer "Time to Leave". Set `campus_suffix` to `""` for room codes only.
- `buildings.json` — per-building GPS coordinates. Every event carries the
  coordinates of its own building (falling back to the campus centre only when a
  room code isn't in the file), used for:
  - `X-APPLE-STRUCTURED-LOCATION` — drives Apple Calendar's "Time to Leave" alerts.
  - `GEO` — the standard RFC 5545 property, for any client that reads it.
  - `URL` + a `Map:` line in the description — a direct
    `https://maps.apple.com/...` link to the building, one tap from the event.
- `VALARM` — one display alert at `--alarm` minutes. A second block per event gives
  two-stage reminders.
- `CATEGORIES` — set to the module code; searchable and filterable.
- `TRANSP` — classes busy, key dates free.
- `KEY_DATES` — all-day academic milestones.
- `WEEK1_MONDAY` — only labels events ("Teaching week 7") and drives the coverage
  report; it no longer affects scheduling.

Worth adding later: a second `URL:` (or a line in the description) pointing at the
module's Loop page, and a separate deadlines-only feed subscribed in its own
colour (today deadlines live in the main feed).

## 7. Two English-group feeds (EN1 / EN2)

Today one export feeds the whole calendar — from a student in English Group 2.
Group 1 currently has to mentally adjust (see the warning on the landing page).
This is already prepared end to end for the day you want a feed per group:

1. Get a `Timetables.xlsx` export from someone in Group 1 and someone in
   Group 2 (each exports their own MyTimetable, same as today). Rename them
   `Timetables_EN1.xlsx` and `Timetables_EN2.xlsx`, drop them at the repo
   root next to `Timetables.xlsx`.
2. `git add -A`, commit, push. The Action already watches both filenames and
   already has the two build steps — they're no-ops until the files exist,
   nothing to change in `.github/workflows/build.yml`. It builds
   `DCU_Semester1_2026_EN1.ics` and `DCU_Semester1_2026_EN2.ics`
   automatically, each named `DCU — Semestre 1 2026/27 (Anglais groupe N)`
   via `--calendar-name`.
3. In `index.html`, the "S'abonner" section has an HTML comment block
   starting `<!-- Le jour où l'anglais est scindé... -->`. Uncomment it
   (remove the `<!--`/`-->`), adjust the wording if needed, and delete the
   "mêmes groupes pour tout le monde" line from the warning box just below —
   it stops being true.
4. Decide whether the current generic `Timetables.xlsx` /
   `DCU_Semester1_2026.ics` should keep existing alongside the two
   group-specific feeds, or retire. Left as your call — nothing here does it
   automatically.

Both feeds share `config.json` and `buildings.json` (semester dates,
closures, buildings — none of that varies by English group), and both go
through the same `--diff`/DTSTAMP-stability/tests machinery as the main feed.
