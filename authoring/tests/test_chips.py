"""Round-trip and edge-case tests for the chip <-> grammar compiler."""
from __future__ import annotations

import pytest

from authoring import chips


# Every token here is canonical grammar the resolver accepts; each must
# decompile to a chip and recompile to the identical string.
ROUND_TRIP = [
    "serve",
    "serve+5d",
    "serve+2w",
    "serve-1d",
    "serve+3bd",
    "serve+1m",
    "serve+1y",
    "@signing",
    "@signing+2w",
    "@signing+2w@09:00",
    "next:THU",
    "next:THU@14:00",
    "this:MON",
    "this:FRI@17:30",
    "nth:2,TUE,+0m",
    "nth:last,FRI,+1m",
    "dom:15,+0m",
    "dom:1,+1m",
    "month:+0m",
    "month:+1m",
    "week_of:(next:MON)",
    "week_of:(@signing+1w)",
]


@pytest.mark.parametrize("expr", ROUND_TRIP)
def test_round_trip(expr):
    chip = chips.parse_token(expr)
    assert chip["base"]["kind"] != "raw", f"{expr} fell back to raw"
    assert chips.compile_chip(chip) == expr


def test_raw_escape_hatch_passthrough():
    # Something the curated palette doesn't model (stacked offsets) -> raw, but
    # still compiles back to the exact original.
    expr = "next:THU+1w+2d"
    chip = chips.parse_token(expr)
    assert chip["base"]["kind"] == "raw"
    assert chips.compile_chip(chip) == expr


def test_unparseable_falls_back_to_raw_without_crashing():
    chip = chips.parse_token("totally not grammar !!!")
    assert chip["base"]["kind"] == "raw"


def test_body_token_plain():
    chip = chips.parse_token("next:THU@14:00")
    assert chips.compile_body_token(chip) == "{next:THU@14:00}"


def test_body_token_emission():
    chip = {"base": {"kind": "serve"}, "offset": {"amount": 5, "unit": "days"},
            "time": None, "emit_as": "signing"}
    assert chips.compile_body_token(chip) == "{!signing = serve+5d}"


def test_friendly_unit_names_compile():
    chip = {"base": {"kind": "serve"}, "offset": {"amount": 2, "unit": "weeks"}, "time": None}
    assert chips.compile_chip(chip) == "serve+2w"
    chip2 = {"base": {"kind": "serve"}, "offset": {"amount": 3, "unit": "business-days"}, "time": None}
    assert chips.compile_chip(chip2) == "serve+3bd"


def test_negative_offset():
    chip = {"base": {"kind": "anchor", "name": "due"}, "offset": {"amount": -2, "unit": "days"}, "time": None}
    assert chips.compile_chip(chip) == "@due-2d"


def test_bad_weekday_raises():
    with pytest.raises(chips.ChipError):
        chips.compile_chip({"base": {"kind": "next_weekday", "weekday": "XYZ"}})


def test_emit_as_survives_parse():
    chip = chips.parse_token("serve+5d", emit_as="signing")
    assert chip["emit_as"] == "signing"
    assert chips.compile_body_token(chip) == "{!signing = serve+5d}"
