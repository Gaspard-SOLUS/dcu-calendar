"""Unit tests for the pure, easy-to-get-wrong parts of dcu_to_ics.py:
escaping/folding, room-code parsing, UID stability, and the DTSTAMP/SEQUENCE
stability logic that keeps `git diff` clean across regenerations.

Run with: pytest  (from the repo root, after `pip install -r requirements-dev.txt`)
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

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
        "kind": "", "lecturer": "", "notes": "",
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
