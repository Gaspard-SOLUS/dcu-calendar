# DCU Semester 1 2026/27 — Apple Calendar feed

Built directly from the MyTimetable Excel export, so a refresh is a re-export
plus one command — no hand-editing of the schedule itself.

| File | Role |
|---|---|
| `Timetables.xlsx` | The MyTimetable export. Replace it wholesale on every refresh. |
| `overrides.csv` | Your enrichment layer: session type, lecturer, notes. Survives re-exports. |
| `dcu_to_ics.py` | Generator. Needs `pandas` + `openpyxl`. |
| `DCU_Semester1_2026.ics` | Output: 132 class events + 5 all-day academic key dates. |

```
EEG1011  37    EEN1018  25    EEN1022  47    EEN1083  12    ESL1009  11
```

Weeks 1–12 covered, 9 to 13 sessions per week.

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

Five activities still have no `kind` set — the Excel export drops the activity
code that carries it. Read them off MyTimetable and fill them in:

| Activity | Code to look for |
|---|---|
| EEG1011 Tue 10:00 (week 6 only) | |
| EEG1011 Thu 14:00 (week 7 only) | |
| EEN1018 Wed 17:00 | |
| EEN1022 Thu 17:00 | |
| EEN1083 Tue 14:00 | |

MyTimetable's naming convention: `OC/L…` = Lecture, `OC/P…` = Practical,
`OC/T…` = Tutorial. The trailing digits are the group number.

## 2. Closures

MyTimetable still shows Monday classes on **26/10/2026**, but DCU states that
scheduled classes do not take place on public holidays, and the Registry calendar
lists that date as "University Closed". Three sessions are dropped and reported on
every run. `--keep-closures` puts them back; `CLOSURES` at the top of the script
is where you add more.

## 3. Commands

```bash
pip install pandas openpyxl          # once

python3 dcu_to_ics.py                                   # default build
python3 dcu_to_ics.py --alarm 25                        # 25-minute reminder
python3 dcu_to_ics.py --alarm 0                         # no reminders
python3 dcu_to_ics.py --merge-adjacent                  # fuse back-to-back sessions
python3 dcu_to_ics.py --split                           # one .ics per module
python3 dcu_to_ics.py --no-key-dates                    # classes only
python3 dcu_to_ics.py --keep-closures                   # keep bank-holiday sessions
python3 dcu_to_ics.py --ttl 2                           # suggest a 2-hour refresh
python3 dcu_to_ics.py --diff DCU_Semester1_2026.ics     # dry run: what would change
```

`--split` matters if you want colours: Apple assigns one colour per calendar,
never per event. Five files, five subscriptions, five colours.

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

UIDs are `date + module + weekday + start + type`, so an edited session updates in
place instead of duplicating, and a cancelled one disappears from the feed.

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

In `dcu_to_ics.py`:

- `CAMPUS_SUFFIX`, `CAMPUS_LAT`, `CAMPUS_LON` — the room code stays first in
  `LOCATION`; the campus address is appended so Apple Maps can geocode the event
  and offer "Time to Leave". Set `CAMPUS_SUFFIX = ""` for room codes only.
- `X-APPLE-STRUCTURED-LOCATION` — the coordinate that actually drives travel-time
  alerts. Per-building coordinates would be more precise than the campus centroid.
- `VALARM` — one display alert at `--alarm` minutes. A second block per event gives
  two-stage reminders.
- `CATEGORIES` — set to the module code; searchable and filterable.
- `TRANSP` — classes busy, key dates free.
- `KEY_DATES` — all-day academic milestones.
- `WEEK1_MONDAY` — only labels events ("Teaching week 7") and drives the coverage
  report; it no longer affects scheduling.

Worth adding later: a `URL:` property per event pointing at the module's Loop page,
and a second feed for coursework deadlines subscribed in its own colour.
