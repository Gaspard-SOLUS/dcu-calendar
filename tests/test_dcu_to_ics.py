"""Unit tests for the pure, easy-to-get-wrong parts of dcu_to_ics.py:
escaping/folding, room-code parsing, UID stability, and the DTSTAMP/SEQUENCE
stability logic that keeps `git diff` clean across regenerations.

Run with: pytest  (from the repo root, after `pip install -r requirements-dev.txt`)
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import dcu_to_ics as dcu  # noqa: E402


# --------------------------------------------------------------------------- #
# escape / fold / unfold                                                      #
# --------------------------------------------------------------------------- #

def test_escape_backslash_semicolon_comma_newline():
    assert dcu.escape("a\\b;c,d\ne") == "a\\\\b\\;c\\,d\\ne"


def test_fold_short_line_untouched():
    line = "SUMMARY:short"
    assert dcu.fold(line) == line


def test_fold_unfold_roundtrip():
    long_value = "x" * 200
    line = f"SUMMARY:{long_value}"
    folded = dcu.fold(line)
    assert "\r\n " in folded  # actually wrapped
    doc = f"BEGIN:VEVENT\r\n{folded}\r\nEND:VEVENT"
    unfolded = dcu.unfold(doc)
    assert unfolded == ["BEGIN:VEVENT", line, "END:VEVENT"]


def test_fold_keeps_utf8_multibyte_chars_intact():
    # The em dash (—) is 3 bytes in UTF-8; folding must never split it.
    line = "SUMMARY:" + ("—" * 40)
    folded = dcu.fold(line)
    doc = f"BEGIN:VEVENT\r\n{folded}\r\nEND:VEVENT"
    assert dcu.unfold(doc) == ["BEGIN:VEVENT", line, "END:VEVENT"]


# --------------------------------------------------------------------------- #
# param()                                                                      #
# --------------------------------------------------------------------------- #

def test_param_quotes_and_strips_unsafe_chars():
    assert dcu.param('He said "hi"') == '"He said \'hi\'"'
    assert dcu.param("line1\nline2\r") == '"line1 line2 "'


# --------------------------------------------------------------------------- #
# parse_room() — needs the real buildings.json loaded                         #
# --------------------------------------------------------------------------- #

def setup_module(_module):
    dcu.load_buildings(ROOT / "buildings.json")


def test_parse_room_two_char_building_code():
    info = dcu.parse_room("GLA.SA301")
    assert info["building"] == "Stokes Extension"
    assert info["floor"] == "3rd floor"
    assert info["room"] == "01"


def test_parse_room_falls_back_to_one_char_when_two_char_is_not_a_building():
    # 'SG' isn't a known building, so this must resolve as 'S' + floor 'G',
    # not fail — this is the case the two-step split_at() exists for.
    info = dcu.parse_room("GLA.SG16")
    assert info["building"] == "Stokes Building"
    assert info["floor"] == "ground floor"
    assert info["room"] == "16"


def test_parse_room_unknown_code_returns_none():
    assert dcu.parse_room("GLA.ZZ999") is None


def test_parse_room_malformed_code_returns_none():
    assert dcu.parse_room("not-a-room-code") is None


# --------------------------------------------------------------------------- #
# make_uid() — identity, not presentation                                     #
# --------------------------------------------------------------------------- #

def _session(**overrides) -> dict:
    base = {
        "code": "EEG1011", "name": "Engineering Mathematics IV",
        "date": date(2026, 9, 7), "weekday": "Monday", "start": "09:00",
        "end": "10:00", "rooms": "GLA.S143", "weeks": "1-12",
        "kind": "", "lecturer": "", "notes": "", "topic": "",
    }
    base.update(overrides)
    return base


def test_uid_stable_across_room_kind_lecturer_notes_changes():
    a = dcu.make_uid(_session())
    b = dcu.make_uid(_session(rooms="GLA.S999", kind="Lecture",
                               lecturer="Dr X", notes="moved room"))
    assert a == b


def test_uid_changes_with_start_time():
    a = dcu.make_uid(_session())
    b = dcu.make_uid(_session(start="10:00"))
    assert a != b


def test_uid_changes_with_module_code():
    a = dcu.make_uid(_session())
    b = dcu.make_uid(_session(code="EEN1022"))
    assert a != b


# --------------------------------------------------------------------------- #
# stamp_for() — DTSTAMP/SEQUENCE stability                                    #
# --------------------------------------------------------------------------- #

def test_stamp_for_reuses_dtstamp_when_content_unchanged():
    previous = {"uid-1": {"DTSTAMP": "20260101T000000Z", "SEQUENCE": "3",
                           "SUMMARY": "same"}}
    dtstamp, sequence = dcu.stamp_for("uid-1", {"SUMMARY": "same"}, previous,
                                       "20260914T120000Z")
    assert dtstamp == "20260101T000000Z"
    assert sequence == 3


def test_stamp_for_bumps_sequence_when_content_changed():
    previous = {"uid-1": {"DTSTAMP": "20260101T000000Z", "SEQUENCE": "3",
                           "SUMMARY": "old"}}
    dtstamp, sequence = dcu.stamp_for("uid-1", {"SUMMARY": "new"}, previous,
                                       "20260914T120000Z")
    assert dtstamp == "20260914T120000Z"
    assert sequence == 4


def test_stamp_for_new_uid_starts_at_sequence_zero():
    dtstamp, sequence = dcu.stamp_for("uid-new", {"SUMMARY": "x"}, {},
                                       "20260914T120000Z")
    assert dtstamp == "20260914T120000Z"
    assert sequence == 0


# --------------------------------------------------------------------------- #
# read_events_from_lines() — must not let VALARM fields leak into the VEVENT  #
# --------------------------------------------------------------------------- #

def test_read_events_ignores_nested_valarm_fields():
    lines = [
        "BEGIN:VEVENT",
        "UID:e1@dcu.timetable",
        "SUMMARY:Real summary",
        "DESCRIPTION:Real description",
        "BEGIN:VALARM",
        "ACTION:DISPLAY",
        "DESCRIPTION:Real summary",  # VALARM's own DESCRIPTION — must not win
        "TRIGGER:-PT15M",
        "END:VALARM",
        "END:VEVENT",
    ]
    events = dcu.read_events_from_lines(lines)
    assert events["e1@dcu.timetable"]["DESCRIPTION"] == "Real description"


# --------------------------------------------------------------------------- #
# exceptions.csv — dated one-off changes                                      #
# --------------------------------------------------------------------------- #

def _write(tmp_path, name: str, text: str) -> Path:
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return path


EXC_HEADER = "date,module_code,start,action,end,rooms,kind,topic,notes\n"


def test_apply_exceptions_cancel_and_adjust(tmp_path):
    path = _write(tmp_path, "exceptions.csv", EXC_HEADER
                  + "# comment line\n"
                  + "2026-09-07,EEG1011,09:00,cancel,,,,,No class\n"
                  + "2026-09-14,EEG1011,09:00,,09:30,GLA.SG16,Quiz,Loop Quiz 1,On campus\n")
    table = dcu.load_exceptions(path)
    sessions = [_session(), _session(date=date(2026, 9, 14), notes="Base note"),
                _session(date=date(2026, 9, 21))]
    used: set = set()
    kept, cancelled, added = dcu.apply_exceptions(sessions, table, used)

    assert added == []
    assert [s["date"] for s, _ in cancelled] == [date(2026, 9, 7)]
    assert cancelled[0][1] == "No class"
    assert len(kept) == 2 and len(used) == 2
    adjusted = kept[0]
    assert adjusted["end"] == "09:30"
    assert adjusted["rooms"] == "GLA.SG16"
    assert adjusted["kind"] == "Quiz"
    assert adjusted["topic"] == "Loop Quiz 1"
    assert adjusted["notes"] == "Base note\nOn campus"
    assert kept[1]["end"] == "10:00"  # untouched week


def test_exception_keeps_uid_and_shows_topic_in_title():
    before = _session(topic="")
    after = _session(topic="Loop Quiz 1", end="09:30", rooms="GLA.SG16")
    assert dcu.make_uid(before) == dcu.make_uid(after)
    fields = dcu.event_fields(after)
    assert fields["summary"].endswith("— Loop Quiz 1")
    assert "Topic: Loop Quiz 1" in fields["description"]


def test_load_exceptions_rejects_unknown_action(tmp_path):
    path = _write(tmp_path, "exceptions.csv",
                  EXC_HEADER + "2026-09-07,EEG1011,09:00,delete,,,,,\n")
    with pytest.raises(SystemExit):
        dcu.load_exceptions(path)


def test_load_exceptions_rejects_duplicates(tmp_path):
    row = "2026-09-07,EEG1011,09:00,cancel,,,,,\n"
    path = _write(tmp_path, "exceptions.csv", EXC_HEADER + row + row)
    with pytest.raises(SystemExit):
        dcu.load_exceptions(path)


def test_load_exceptions_rejects_unquoted_comma(tmp_path):
    path = _write(tmp_path, "exceptions.csv",
                  EXC_HEADER + "2026-09-07,EEG1011,09:00,,,,,,Week 13, usual slot\n")
    with pytest.raises(SystemExit):
        dcu.load_exceptions(path)


def test_apply_exceptions_add_creates_session_outside_export(tmp_path):
    path = _write(tmp_path, "exceptions.csv", EXC_HEADER
                  + "2026-12-03,ESL1009,10:00,add,12:00,GLA.CG03,Lecture,Assessment 3,\n")
    sessions = [_session(code="ESL1009", name="English Language in Use",
                         date=date(2026, 11, 26), weekday="Thursday", start="10:00")]
    used: set = set()
    kept, cancelled, added = dcu.apply_exceptions(
        sessions, dcu.load_exceptions(path), used)
    assert cancelled == [] and len(added) == 1 and len(kept) == 2
    new = added[0]
    assert new["weekday"] == "Thursday"
    assert new["name"] == "English Language in Use"
    assert (new["start"], new["end"], new["rooms"]) == ("10:00", "12:00", "GLA.CG03")
    assert kept[-1] is new  # sorted by date
    assert dcu.event_fields(new)["summary"].endswith("(Lecture) — Assessment 3")


def test_add_on_existing_slot_is_not_duplicated(tmp_path):
    path = _write(tmp_path, "exceptions.csv", EXC_HEADER
                  + "2026-09-07,EEG1011,09:00,add,10:00,GLA.S143,,,\n")
    used: set = set()
    kept, _, added = dcu.apply_exceptions([_session()], dcu.load_exceptions(path), used)
    assert added == [] and len(kept) == 1
    assert used == set()  # reported as "never matched"


def test_load_exceptions_add_needs_end_and_rooms(tmp_path):
    path = _write(tmp_path, "exceptions.csv",
                  EXC_HEADER + "2026-12-03,ESL1009,10:00,add,,,,,\n")
    with pytest.raises(SystemExit):
        dcu.load_exceptions(path)


# --------------------------------------------------------------------------- #
# deadlines.csv — always a precise moment, 23:59 by default                   #
# --------------------------------------------------------------------------- #

DL_HEADER = "date,time,module_code,title,notes\n"


def test_deadline_defaults_to_2359_and_is_never_all_day(tmp_path):
    path = _write(tmp_path, "deadlines.csv", DL_HEADER
                  + '2026-10-05,,EEN1022,Lab 1 due,"a, b"\n')
    (item,) = dcu.load_deadlines(path)
    f = dcu.deadline_fields(item, {"EEN1022": "Digital & Analogue Electronics I"})
    assert (f["dtstart"], f["dtend"]) == ("20261005T235900", "20261005T235900")
    assert f["summary"] == "EEN1022 — Lab 1 due"
    assert "Digital & Analogue Electronics I" in f["description"]
    assert "a, b" in f["description"]
    lines = dcu.build_deadline(item, {}, "20260101T000000Z", 15)
    assert "DTSTART;TZID=Europe/Dublin:20261005T235900" in lines
    assert not any("VALUE=DATE" in line for line in lines)
    for trigger in dcu.DEADLINE_ALARMS:
        assert f"TRIGGER:{trigger}" in lines


def test_deadline_explicit_time(tmp_path):
    path = _write(tmp_path, "deadlines.csv", DL_HEADER + "2026-10-05,17:00,EEN1022,Lab 1 due,\n")
    (item,) = dcu.load_deadlines(path)
    assert dcu.deadline_fields(item, {})["dtstart"] == "20261005T170000"


def test_deadline_no_alarm_when_disabled(tmp_path):
    path = _write(tmp_path, "deadlines.csv", DL_HEADER + "2026-10-05,,EEN1022,Lab 1 due,\n")
    (item,) = dcu.load_deadlines(path)
    assert "BEGIN:VALARM" not in dcu.build_deadline(item, {}, "20260101T000000Z", None)


def test_deadline_uid_ignores_notes_and_time():
    a = {"date": date(2026, 10, 5), "time": "23:59", "code": "EEN1022",
         "title": "Lab 1 due", "notes": ""}
    b = dict(a, notes="moved to Loop", time="17:00")
    assert dcu.deadline_fields(a, {})["uid"] == dcu.deadline_fields(b, {})["uid"]


def test_load_deadlines_rejects_bad_time(tmp_path):
    path = _write(tmp_path, "deadlines.csv", DL_HEADER + "2026-10-05,midnight,EEN1022,Lab 1 due,\n")
    with pytest.raises(SystemExit):
        dcu.load_deadlines(path)
