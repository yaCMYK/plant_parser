#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Production plant catalog parser.

Goals:
- semantic (synonym/context) mapping instead of exact header matching;
- strict per-field acceptance rules;
- preserve every raw observation;
- keep multiple values per target cell;
- normalize measurable values while retaining raw value, min/max and source unit;
- avoid inventing missing data;
- deduplicate plants by normalized botanical name;
- export a black-and-white XLSX with Plants, Raw evidence, Field rules, Crawl log, Summary.

Run:
    pip install -r requirements.txt
    python plant_parser_production.py \
        --url "https://example.com/plants" \
        --output "plants.xlsx" \
        --max-pages 500

For JS-heavy sites:
    pip install playwright
    playwright install chromium
    python plant_parser_production.py --url ... --browser playwright

This parser respects robots.txt by default. Use --ignore-robots only when you have permission.
"""
from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import hashlib
import html
import json
import logging
import math
import re
import sys
import time
import unicodedata
from collections import defaultdict, deque
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable, Iterable, Optional
from urllib.parse import urljoin, urlparse, urldefrag
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup, Tag
import xlsxwriter

LOGGER = logging.getLogger("plant_parser")

USER_AGENT = "PlantCatalogResearchBot/1.0 (+data-quality; contact=site-owner-if-required)"

# -----------------------------------------------------------------------------
# Data model
# -----------------------------------------------------------------------------

@dataclass(frozen=True)
class ParsedValue:
    raw_value: str
    normalized_value: str = ""
    min_value: Optional[float] = None
    max_value: Optional[float] = None
    source_unit: str = ""
    canonical_unit: str = ""
    extra: dict[str, Any] = field(default_factory=dict)


@dataclass
class Observation:
    plant_id: str
    field_key: str
    field_label: str
    raw_label: str
    raw_value: str
    normalized_value: str
    min_value: Optional[float]
    max_value: Optional[float]
    source_unit: str
    canonical_unit: str
    confidence: float
    extraction_method: str
    evidence_text: str
    source_url: str
    scraped_at: str
    accepted_to_main: bool = True
    rejection_reason: str = ""

    def display_value(self) -> str:
        return self.normalized_value.strip() or self.raw_value.strip()


@dataclass
class Candidate:
    label: str
    value: str
    context: str
    method: str


@dataclass
class FieldRule:
    key: str
    label: str
    value_type: str
    aliases: tuple[str, ...] = ()
    exclude_terms: tuple[str, ...] = ()
    required_context_any: tuple[str, ...] = ()
    free_text_patterns: tuple[str, ...] = ()
    canonical_unit: str = ""
    enum_map: dict[str, str] = field(default_factory=dict)
    min_confidence: float = 0.86
    allow_fuzzy: bool = True
    notes: str = ""
    parser_name: str = "text"


@dataclass
class CrawlEntry:
    url: str
    status: str
    http_status: Optional[int] = None
    page_type: str = ""
    title: str = ""
    error: str = ""


@dataclass
class PlantRecord:
    plant_id: str
    page_url: str
    page_title: str = ""
    observations: dict[str, list[Observation]] = field(default_factory=lambda: defaultdict(list))
    unclassified_images: list[str] = field(default_factory=list)

    def add(self, observation: Observation) -> None:
        existing = self.observations[observation.field_key]
        fingerprint = (
            normalize_space(observation.raw_value).casefold(),
            normalize_space(observation.source_url).casefold(),
            round(observation.confidence, 3),
        )
        for item in existing:
            fp2 = (
                normalize_space(item.raw_value).casefold(),
                normalize_space(item.source_url).casefold(),
                round(item.confidence, 3),
            )
            if fingerprint == fp2:
                return
        existing.append(observation)


# -----------------------------------------------------------------------------
# Normalization helpers
# -----------------------------------------------------------------------------

DASHES = "\u2010\u2011\u2012\u2013\u2014\u2212"


def normalize_space(value: str) -> str:
    value = html.unescape(value or "")
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def norm_text(value: str) -> str:
    value = normalize_space(value).casefold().replace("ё", "е")
    value = unicodedata.normalize("NFKC", value)
    value = re.sub(f"[{re.escape(DASHES)}]", "-", value)
    value = re.sub(r"[^a-zа-я0-9+×x%°./'\" -]+", " ", value, flags=re.I)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def unique_preserve(seq: Iterable[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for item in seq:
        k = norm_text(item)
        if item and k not in seen:
            seen.add(k)
            out.append(item)
    return out


def safe_float(raw: str) -> Optional[float]:
    try:
        return float(raw.replace(",", "."))
    except (TypeError, ValueError):
        return None


def format_num(value: float) -> str:
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.4f}".rstrip("0").rstrip(".")


# -----------------------------------------------------------------------------
# Measurement parsing. Canonical length unit is meter.
# Raw value is NEVER discarded.
# -----------------------------------------------------------------------------

LENGTH_FACTORS_TO_M = {
    "mm": 0.001,
    "мм": 0.001,
    "cm": 0.01,
    "см": 0.01,
    "m": 1.0,
    "м": 1.0,
    "meter": 1.0,
    "metre": 1.0,
    "meters": 1.0,
    "metres": 1.0,
    "ft": 0.3048,
    "feet": 0.3048,
    "foot": 0.3048,
    "in": 0.0254,
    "inch": 0.0254,
    "inches": 0.0254,
    "\"": 0.0254,
    "'": 0.3048,
}

VOLUME_FACTORS_TO_L = {
    "l": 1.0,
    "л": 1.0,
    "liter": 1.0,
    "litre": 1.0,
    "liters": 1.0,
    "litres": 1.0,
    "ml": 0.001,
    "мл": 0.001,
}

NUMBER = r"(?P<a>\d+(?:[\.,]\d+)?)"
NUMBER2 = r"(?P<b>\d+(?:[\.,]\d+)?)"
RANGE_SEP = r"(?:\s*(?:-|–|—|до|to)\s*)"


def parse_length_values(raw: str) -> list[ParsedValue]:
    """Parse every explicit length/range in raw and normalize to meters.

    Preserves one-sided semantics:
      "до 2 m" -> max=2, normalized "≤2 m"
      "от 1.5 m" -> min=1.5, normalized "≥1.5 m"
    """
    text = normalize_space(raw)
    low = norm_text(text)
    results: list[ParsedValue] = []
    occupied: list[tuple[int, int]] = []
    unit_alt = "|".join(sorted(map(re.escape, LENGTH_FACTORS_TO_M), key=len, reverse=True))

    def overlaps(span: tuple[int, int]) -> bool:
        return any(max(span[0], a) < min(span[1], b) for a, b in occupied)

    # One-sided maximum/minimum first, so generic matcher does not erase semantics.
    one_sided = re.compile(
        rf"(?P<dir>до|up to|maximum|max\.?|не более|at most|от|from|minimum|min\.?|не менее|at least)\s*"
        rf"(?P<a>\d+(?:[\.,]\d+)?)\s*(?P<unit>{unit_alt})(?=\b|\s|$|[,;/)])",
        re.I,
    )
    for m in one_sided.finditer(low):
        a = safe_float(m.group("a"))
        unit = m.group("unit").casefold()
        factor = LENGTH_FACTORS_TO_M.get(unit)
        if a is None or factor is None:
            continue
        value = a * factor
        direction = norm_text(m.group("dir"))
        is_max = direction in {"до", "up to", "maximum", "max", "не более", "at most"}
        min_v = None if is_max else value
        max_v = value if is_max else None
        normalized = f"≤{format_num(value)} m" if is_max else f"≥{format_num(value)} m"
        results.append(ParsedValue(m.group(0), normalized, min_v, max_v, unit, "m", {"bound": "max" if is_max else "min"}))
        occupied.append(m.span())

    pattern = re.compile(
        rf"{NUMBER}(?:{RANGE_SEP}{NUMBER2})?\s*(?P<unit>{unit_alt})(?=\b|\s|$|[,;/)])",
        re.I,
    )
    for m in pattern.finditer(low):
        if overlaps(m.span()):
            continue
        a = safe_float(m.group("a"))
        b = safe_float(m.group("b")) if m.groupdict().get("b") else None
        unit = m.group("unit").casefold()
        factor = LENGTH_FACTORS_TO_M.get(unit)
        if a is None or factor is None:
            continue
        min_v = a * factor
        max_v = (b * factor) if b is not None else min_v
        if max_v < min_v:
            min_v, max_v = max_v, min_v
        normalized = (
            f"{format_num(min_v)}–{format_num(max_v)} m"
            if abs(max_v - min_v) > 1e-12
            else f"{format_num(min_v)} m"
        )
        results.append(
            ParsedValue(
                raw_value=m.group(0),
                normalized_value=normalized,
                min_value=min_v,
                max_value=max_v,
                source_unit=unit,
                canonical_unit="m",
            )
        )
    return dedupe_parsed(results)

def parse_volume_values(raw: str) -> list[ParsedValue]:
    text = normalize_space(raw)
    low = norm_text(text)
    results: list[ParsedValue] = []
    unit_alt = "|".join(sorted(map(re.escape, VOLUME_FACTORS_TO_L), key=len, reverse=True))
    pattern = re.compile(
        rf"{NUMBER}(?:{RANGE_SEP}{NUMBER2})?\s*(?P<unit>{unit_alt})(?=\b|\s|$|[,;/)])",
        re.I,
    )
    for m in pattern.finditer(low):
        a = safe_float(m.group("a"))
        b = safe_float(m.group("b")) if m.groupdict().get("b") else None
        unit = m.group("unit").casefold()
        factor = VOLUME_FACTORS_TO_L.get(unit)
        if a is None or factor is None:
            continue
        min_v = a * factor
        max_v = (b * factor) if b is not None else min_v
        if max_v < min_v:
            min_v, max_v = max_v, min_v
        normalized = (
            f"{format_num(min_v)}–{format_num(max_v)} L"
            if abs(max_v - min_v) > 1e-12
            else f"{format_num(min_v)} L"
        )
        results.append(ParsedValue(m.group(0), normalized, min_v, max_v, unit, "L"))
    return dedupe_parsed(results)


def dedupe_parsed(values: list[ParsedValue]) -> list[ParsedValue]:
    seen: set[tuple[Any, ...]] = set()
    out: list[ParsedValue] = []
    for v in values:
        key = (v.normalized_value, v.min_value, v.max_value, v.source_unit)
        if key not in seen:
            seen.add(key)
            out.append(v)
    return out


def parse_length(raw: str) -> list[ParsedValue]:
    return parse_length_values(raw)


def parse_container_size(raw: str) -> list[ParsedValue]:
    # Container/rootball can legitimately contain BOTH dimensions and volume.
    out = dedupe_parsed(parse_length_values(raw) + parse_volume_values(raw))
    text = norm_text(raw)

    # Nursery container codes such as C5, C20, P9 are meaningful even without a unit.
    for m in re.finditer(r"\b(?:c|p)\s*\d+(?:[\.,]\d+)?\b", text, re.I):
        code = re.sub(r"\s+", "", m.group(0)).upper().replace(",", ".")
        out.append(ParsedValue(m.group(0), code, source_unit="nursery code", canonical_unit="nursery code"))

    # Dimension chains with one shared unit, e.g. 60×60 cm or 40 x 40 x 35 cm.
    unit_alt = "|".join(sorted(map(re.escape, LENGTH_FACTORS_TO_M), key=len, reverse=True))
    dim_re = re.compile(rf"(?P<dims>\d+(?:[\.,]\d+)?(?:\s*[x×]\s*\d+(?:[\.,]\d+)?)+)\s*(?P<unit>{unit_alt})", re.I)
    for m in dim_re.finditer(text):
        unit = m.group("unit").casefold()
        factor = LENGTH_FACTORS_TO_M.get(unit)
        if factor is None:
            continue
        nums = [safe_float(x) for x in re.findall(r"\d+(?:[\.,]\d+)?", m.group("dims"))]
        vals = [x * factor for x in nums if x is not None]
        if len(vals) >= 2:
            normalized = "×".join(format_num(x) for x in vals) + " m"
            out.append(ParsedValue(m.group(0), normalized, min(vals), max(vals), unit, "m", {"dimensions_m": vals}))

    if not out and normalize_space(raw):
        # A labelled container/rootball value may be a standard code/site-specific notation.
        # Preserve it rather than invent units.
        out.append(ParsedValue(raw, normalize_space(raw)))
    return dedupe_parsed(out)


def parse_ph(raw: str) -> list[ParsedValue]:
    text = norm_text(raw)
    # Strict: only numeric pH / clearly labelled acidic-neutral-alkaline text.
    vals: list[ParsedValue] = []
    m = re.search(r"(?:ph\s*)?(\d(?:[\.,]\d+)?)(?:\s*(?:-|–|—|до|to)\s*(\d(?:[\.,]\d+)?))?", text, re.I)
    if m:
        a = safe_float(m.group(1))
        b = safe_float(m.group(2)) if m.group(2) else a
        if a is not None and b is not None and 0 <= a <= 14 and 0 <= b <= 14:
            lo, hi = sorted((a, b))
            normalized = f"pH {format_num(lo)}–{format_num(hi)}" if hi != lo else f"pH {format_num(lo)}"
            vals.append(ParsedValue(raw, normalized, lo, hi, "pH", "pH"))
    descriptors = []
    if re.search(r"\b(acidic|acid|кисл\w*)\b", text):
        descriptors.append("кислая")
    if re.search(r"\b(neutral|нейтрал\w*)\b", text):
        descriptors.append("нейтральная")
    if re.search(r"\b(alkaline|alkalic|щелоч\w*)\b", text):
        descriptors.append("щелочная")
    for d in unique_preserve(descriptors):
        vals.append(ParsedValue(raw, d))
    return dedupe_parsed(vals)


def parse_temperature(raw: str) -> list[ParsedValue]:
    text = norm_text(raw)
    out: list[ParsedValue] = []
    for m in re.finditer(r"(-?\d+(?:[\.,]\d+)?)\s*°?\s*(c|с|f)\b", text, re.I):
        value = safe_float(m.group(1))
        unit = m.group(2).casefold()
        if value is None:
            continue
        if unit == "f":
            c = (value - 32.0) * 5.0 / 9.0
        else:
            c = value
        out.append(ParsedValue(m.group(0), f"{format_num(c)} °C", c, c, unit, "°C"))
    return out


def parse_hardiness(raw: str) -> list[ParsedValue]:
    """Do NOT convert generic 'hardy' to USDA. Keep explicit systems separate."""
    text = norm_text(raw)
    out: list[ParsedValue] = []
    zone = re.search(r"(?:usda\s*)?(?:zone|зона)\s*(\d{1,2}[ab]?)(?:\s*(?:-|–|—|to|до)\s*(\d{1,2}[ab]?))?", text, re.I)
    if zone:
        z1, z2 = zone.group(1).upper(), (zone.group(2) or zone.group(1)).upper()
        normalized = f"USDA {z1}–{z2}" if z1 != z2 else f"USDA {z1}"
        out.append(ParsedValue(raw, normalized, source_unit="USDA zone", canonical_unit="USDA zone"))
    temps = parse_temperature(raw)
    out.extend(temps)
    # Explicit qualitative wording only; no numeric inference.
    if re.search(r"\b(very hardy|extremely hardy|очень зимостойк\w*|высокая зимостойкость)\b", text):
        out.append(ParsedValue(raw, "высокая зимостойкость"))
    elif re.search(r"\b(hardy|winter hardy|зимостойк\w*|морозостойк\w*)\b", text):
        out.append(ParsedValue(raw, "зимостойкое"))
    if re.search(r"\b(not hardy|tender|незимостойк\w*|не зимует|не переносит мороз)\b", text):
        out.append(ParsedValue(raw, "низкая зимостойкость / требует защиты"))
    return dedupe_parsed(out)


MONTH_MAP = {
    "january": "январь", "jan": "январь", "январ": "январь",
    "february": "февраль", "feb": "февраль", "феврал": "февраль",
    "march": "март", "mar": "март", "март": "март",
    "april": "апрель", "apr": "апрель", "апрел": "апрель",
    "may": "май", "май": "май",
    "june": "июнь", "jun": "июнь", "июн": "июнь",
    "july": "июль", "jul": "июль", "июл": "июль",
    "august": "август", "aug": "август", "август": "август",
    "september": "сентябрь", "sep": "сентябрь", "sept": "сентябрь", "сентябр": "сентябрь",
    "october": "октябрь", "oct": "октябрь", "октябр": "октябрь",
    "november": "ноябрь", "nov": "ноябрь", "ноябр": "ноябрь",
    "december": "декабрь", "dec": "декабрь", "декабр": "декабрь",
}

SEASON_MAP = {
    "spring": "весна", "весн": "весна",
    "summer": "лето", "летн": "лето",
    "autumn": "осень", "fall": "осень", "осен": "осень",
    "winter": "зима", "зим": "зима",
    "year round": "круглый год", "all year": "круглый год", "круглый год": "круглый год",
}


def parse_period(raw: str) -> list[ParsedValue]:
    text = norm_text(raw)
    found: list[str] = []
    for token, canonical in MONTH_MAP.items():
        if re.search(rf"\b{re.escape(token)}\w*\b", text):
            found.append(canonical)
    for token, canonical in SEASON_MAP.items():
        if token in text:
            found.append(canonical)
    if not found:
        return [ParsedValue(raw, normalize_space(raw))] if normalize_space(raw) else []
    return [ParsedValue(raw, " / ".join(unique_preserve(found)))]


def parse_enum(raw: str, enum_map: dict[str, str]) -> list[ParsedValue]:
    text = norm_text(raw)
    found: list[str] = []
    for phrase, canonical in enum_map.items():
        phrase_n = norm_text(phrase)
        if phrase_n and (phrase_n in text or re.search(rf"\b{re.escape(phrase_n)}\b", text)):
            found.append(canonical)
    return [ParsedValue(raw, x) for x in unique_preserve(found)]


def parse_explicit_tolerance(raw: str, kind: str) -> list[ParsedValue]:
    """Strictly distinguish preference from tolerance."""
    text = norm_text(raw)
    if kind == "drought":
        high = [r"drought tolerant", r"drought resistant", r"засухоустойчив", r"устойчив\w* к засух"]
        negative = [r"not drought tolerant", r"не переносит засух", r"чувствител\w* к засух"]
    else:
        high = [r"waterlog(?:ging)? tolerant", r"tolerates? waterlog", r"flood tolerant", r"tolerates? flooding", r"переносит переувлажнен", r"устойчив\w* к переувлажнен", r"переносит затоплен"]
        negative = [r"not waterlog", r"does not tolerate waterlog", r"не переносит переувлажнен", r"не переносит застой воды"]
    if any(re.search(p, text) for p in negative):
        return [ParsedValue(raw, "неустойчиво / не переносит")]
    if any(re.search(p, text) for p in high):
        return [ParsedValue(raw, "устойчиво / переносит")]
    return []


def parse_bool_suitability(raw: str, target_terms: tuple[str, ...]) -> list[ParsedValue]:
    text = norm_text(raw)
    target_hit = any(norm_text(t) in text for t in target_terms)
    if not target_hit:
        return []
    negative = re.search(r"\b(not suitable|unsuitable|не подходит|не пригод|не рекомендуется|cannot be grown)\b", text)
    positive = re.search(r"\b(suitable|recommended|can be grown|grows well|подходит|пригод|рекомендуется|можно выращивать|используется)\b", text)
    if negative:
        return [ParsedValue(raw, "нет")]
    if positive:
        return [ParsedValue(raw, "да")]
    return []


def parse_currency(raw: str) -> list[ParsedValue]:
    text = normalize_space(raw)
    # Keep original currency; never convert exchange rates.
    currency_patterns = [
        (r"(?:₽|руб\.?|rub)\s*(\d[\d\s]*(?:[\.,]\d+)?)|(?:\d[\d\s]*(?:[\.,]\d+)?)\s*(?:₽|руб\.?|rub)", "RUB"),
        (r"(?:€|eur)\s*(\d[\d\s]*(?:[\.,]\d+)?)|(?:\d[\d\s]*(?:[\.,]\d+)?)\s*(?:€|eur)", "EUR"),
        (r"(?:\$|usd)\s*(\d[\d\s]*(?:[\.,]\d+)?)|(?:\d[\d\s]*(?:[\.,]\d+)?)\s*(?:\$|usd)", "USD"),
        (r"(?:£|gbp)\s*(\d[\d\s]*(?:[\.,]\d+)?)|(?:\d[\d\s]*(?:[\.,]\d+)?)\s*(?:£|gbp)", "GBP"),
    ]
    out: list[ParsedValue] = []
    low = text.casefold()
    for pattern, cur in currency_patterns:
        for m in re.finditer(pattern, low, re.I):
            num_match = re.search(r"\d[\d\s]*(?:[\.,]\d+)?", m.group(0))
            if not num_match:
                continue
            num_raw = num_match.group(0).replace(" ", "")
            v = safe_float(num_raw)
            if v is None:
                continue
            out.append(ParsedValue(m.group(0), f"{format_num(v)} {cur}", v, v, cur, cur))
    return dedupe_parsed(out)


def parse_text(raw: str) -> list[ParsedValue]:
    value = normalize_space(raw)
    return [ParsedValue(raw, value)] if value else []


def parse_list_text(raw: str) -> list[ParsedValue]:
    raw = normalize_space(raw)
    parts = [normalize_space(x) for x in re.split(r"[;|•\n]+", raw) if normalize_space(x)]
    if len(parts) <= 1:
        return parse_text(raw)
    return [ParsedValue(raw, p) for p in unique_preserve(parts)]


def parse_lifespan(raw: str) -> list[ParsedValue]:
    text = norm_text(raw)
    m = re.search(r"(\d+(?:[\.,]\d+)?)(?:\s*(?:-|–|—|до|to)\s*(\d+(?:[\.,]\d+)?))?\s*(?:years?|yrs?|лет|года?|год)", text, re.I)
    if not m:
        return parse_text(raw)
    a = safe_float(m.group(1))
    b = safe_float(m.group(2)) if m.group(2) else a
    if a is None or b is None:
        return parse_text(raw)
    lo, hi = sorted((a, b))
    normalized = f"{format_num(lo)}–{format_num(hi)} years" if hi != lo else f"{format_num(lo)} years"
    return [ParsedValue(raw, normalized, lo, hi, "years", "years")]


def parse_growth_rate(raw: str) -> list[ParsedValue]:
    out = parse_enum(raw, GROWTH_RATE_ENUM)
    text = norm_text(raw)
    if re.search(r"(?:per year|/year|/yr|в год|за год|annual|годов)", text, re.I):
        for value in parse_length_values(raw):
            if value.canonical_unit == "m":
                if value.min_value is not None and value.max_value is not None:
                    normalized = (
                        f"{format_num(value.min_value)}–{format_num(value.max_value)} m/year"
                        if value.min_value != value.max_value
                        else f"{format_num(value.min_value)} m/year"
                    )
                elif value.max_value is not None:
                    normalized = f"≤{format_num(value.max_value)} m/year"
                elif value.min_value is not None:
                    normalized = f"≥{format_num(value.min_value)} m/year"
                else:
                    normalized = value.normalized_value + "/year"
                out.append(ParsedValue(value.raw_value, normalized, value.min_value, value.max_value, value.source_unit, "m/year"))
    if not out:
        return parse_text(raw)
    return dedupe_parsed(out)


PARSERS: dict[str, Callable[..., list[ParsedValue]]] = {
    "text": parse_text,
    "list_text": parse_list_text,
    "length": parse_length,
    "container_size": parse_container_size,
    "ph": parse_ph,
    "hardiness": parse_hardiness,
    "period": parse_period,
    "currency": parse_currency,
    "lifespan": parse_lifespan,
    "growth_rate": parse_growth_rate,
}


# -----------------------------------------------------------------------------
# Strict semantic dictionary for EVERY target column.
# Technical fields have explicit generation rules instead of web aliases.
# -----------------------------------------------------------------------------

LIGHT_ENUM = {
    "full sun": "солнце", "sun": "солнце", "солнечное место": "солнце", "солнце": "солнце",
    "partial shade": "полутень", "part shade": "полутень", "полутень": "полутень",
    "full shade": "тень", "shade": "тень", "тень": "тень",
}

LIFE_FORM_ENUM = {
    "conifer tree": "хвойное дерево", "хвойное дерево": "хвойное дерево",
    "deciduous tree": "лиственное дерево", "лиственное дерево": "лиственное дерево",
    "evergreen tree": "вечнозелёное дерево", "вечнозеленое дерево": "вечнозелёное дерево",
    "tree": "дерево", "дерево": "дерево",
    "conifer shrub": "хвойный кустарник", "хвойный кустарник": "хвойный кустарник",
    "deciduous shrub": "лиственный кустарник", "лиственный кустарник": "лиственный кустарник",
    "evergreen shrub": "вечнозелёный кустарник", "вечнозеленый кустарник": "вечнозелёный кустарник",
    "shrub": "кустарник", "кустарник": "кустарник",
    "perennial": "многолетник", "herbaceous perennial": "многолетник", "многолетник": "многолетник",
    "annual": "однолетник", "однолетник": "однолетник",
    "biennial": "двулетник", "двулетник": "двулетник",
    "grass": "злак", "ornamental grass": "декоративный злак", "злак": "злак",
    "vine": "лиана", "climber": "лиана", "лиана": "лиана",
    "groundcover": "почвопокровное", "ground cover": "почвопокровное", "почвопокровное": "почвопокровное",
    "fern": "папоротник", "папоротник": "папоротник",
    "bulb": "луковичное", "bulbous": "луковичное", "луковичное": "луковичное",
}

GROWTH_RATE_ENUM = {
    "very slow": "очень медленная", "очень медлен": "очень медленная",
    "slow": "медленная", "медлен": "медленная",
    "moderate": "средняя", "medium": "средняя", "средн": "средняя",
    "fast": "быстрая", "rapid": "быстрая", "быстр": "быстрая",
}

PLANTING_TYPE_ENUM = {
    "specimen": "солитер", "solitary": "солитер", "солитер": "солитер",
    "mass planting": "массив", "mass": "массив", "массив": "массив",
    "group planting": "группа", "group": "группа", "группа": "группа",
    "hedge": "живая изгородь", "живой изгород": "живая изгородь", "живая изгородь": "живая изгородь",
    "border": "бордюр", "бордюр": "бордюр",
    "edge": "опушка", "woodland edge": "опушка", "опушка": "опушка",
    "groundcover": "почвопокровная посадка", "ground cover": "почвопокровная посадка",
    "container": "контейнерная посадка", "кашпо": "контейнерная посадка",
}

SPECIAL_PROPERTY_ENUM = {
    "edible": "съедобное", "edible fruit": "съедобные плоды", "съедобн": "съедобное",
    "medicinal": "лекарственное", "лекарствен": "лекарственное",
    "honey plant": "медонос", "bee plant": "медонос", "nectar plant": "медонос", "медонос": "медонос",
    "pollinator friendly": "привлекает опылителей", "attracts pollinators": "привлекает опылителей",
    "bird friendly": "привлекает птиц", "attracts birds": "привлекает птиц",
    "toxic": "ядовитое", "poisonous": "ядовитое", "ядовит": "ядовитое",
}

SOIL_FERTILITY_ENUM = {
    "poor soil": "бедная", "low fertility": "низкая", "бедн": "бедная",
    "moderately fertile": "средняя", "medium fertility": "средняя", "умеренно плодород": "средняя",
    "fertile": "плодородная", "rich soil": "плодородная", "плодород": "плодородная",
}


def R(key: str, label: str, value_type: str, aliases=(), **kwargs: Any) -> FieldRule:
    return FieldRule(key=key, label=label, value_type=value_type, aliases=tuple(aliases), **kwargs)


FIELD_RULES: list[FieldRule] = [
    R("id", "ID", "generated", notes="Детерминированный ID: ботаническое имя + сорт; fallback — URL."),
    R("life_form", "Жизненная форма", "enum", ("жизненная форма", "тип растения", "plant type", "life form", "habit type"), enum_map=LIFE_FORM_ENUM, notes="Только явный тип растения; не выводить из одного лишь рода."),
    R("ru_genus", "Русское название — род", "taxonomy_ru", ("русское название", "название на русском", "common name ru", "наименование"), required_context_any=("рус", "назван", "наимен"), notes="Берётся из явно русского ботанического названия; первая таксономическая часть."),
    R("ru_species", "Русское название — вид", "taxonomy_ru", ("русское название", "название на русском", "common name ru", "наименование"), required_context_any=("рус", "назван", "наимен"), notes="Вторая часть русского ботанического названия, только если структура однозначна."),
    R("ru_cultivar", "Русское название — сорт", "taxonomy_ru", ("сорт", "cultivar", "cultivar name", "русское название"), exclude_terms=("группа сортов",), notes="Только явно указанный сорт / текст в кавычках при русском названии."),
    R("latin_genus", "Латинское название — род", "taxonomy_latin", ("латинское название", "ботаническое название", "scientific name", "botanical name", "latin name"), notes="Genus из валидного латинского биномиального/триномиального имени."),
    R("latin_species", "Латинское название — вид", "taxonomy_latin", ("латинское название", "ботаническое название", "scientific name", "botanical name", "latin name"), notes="Species/гибридный эпитет из валидного ботанического имени."),
    R("latin_cultivar", "Латинское название — сорт", "taxonomy_latin", ("cultivar", "cultivar name", "сорт", "botanical name", "scientific name"), exclude_terms=("cultivar group", "группа сортов"), notes="Только cultivar в кавычках или явном поле сорта."),
    R("alternative_names", "Альтернативные названия", "list_text", ("синонимы", "альтернативные названия", "другие названия", "synonym", "synonyms", "also known as", "common names"), notes="Список явных синонимов/альтернативных названий."),
    R("cultivar_group", "Группа сортов", "text", ("группа сортов", "cultivar group", "series", "сортогруппа"), notes="Только явно обозначенная группа/серия сортов."),
    R("plant_form", "Форма растения", "text", ("форма растения", "growth habit", "habit", "габитус", "характер роста"), exclude_terms=("форма кроны", "crown shape"), notes="Габитус: прямостоячий, стелющийся, раскидистый, куртинный и т.п."),
    R("crown_form", "Форма кроны", "text", ("форма кроны", "crown shape", "crown form", "форма кроны дерева"), notes="Только характеристика кроны: шаровидная, колонновидная, пирамидальная и т.п."),
    R("root_system", "Корневая система", "text", ("корневая система", "roots", "root system", "root type"), notes="Тип/характер корневой системы, только явные данные."),
    R("lifespan", "Продолжительность жизни", "lifespan", ("продолжительность жизни", "долговечность", "lifespan", "longevity", "life expectancy"), parser_name="lifespan", notes="Сохранять возраст/категорию в raw; не оценивать самостоятельно."),
    R("growth_rate", "Скорость роста", "growth_rate", ("скорость роста", "темп роста", "growth rate", "annual growth", "прирост"), enum_map=GROWTH_RATE_ENUM, parser_name="growth_rate", notes="Категория + явный годовой прирост, если указан; оба значения сохраняются."),
    R("mature_height", "Высота во взрослом виде", "length", ("высота взрослого растения", "высота во взрослом виде", "mature height", "ultimate height", "final height", "adult height", "высота"), exclude_terms=("сажен", "container", "контейнер", "pot", "горш", "shipping size"), canonical_unit="m", parser_name="length", notes="Только взрослый размер. Все найденные диапазоны сохраняются; canonical=m."),
    R("mature_diameter", "Диаметр во взрослом виде", "length", ("диаметр взрослого растения", "ширина взрослого растения", "диаметр кроны", "mature width", "mature spread", "ultimate spread", "spread", "crown width", "ширина"), exclude_terms=("container", "контейнер", "pot", "горш", "сажен", "rootball", "ком"), canonical_unit="m", parser_name="length", notes="Горизонтальный взрослый размер/диаметр/ширина; canonical=m."),
    R("mature_trunk_girth", "Обхват ствола во взрослом виде", "length", ("обхват ствола взрослого", "обхват ствола", "trunk girth", "mature trunk girth", "trunk circumference"), exclude_terms=("сажен", "nursery", "поставка"), canonical_unit="m", parser_name="length", notes="Только взрослый ствол; не путать с калибром поставки."),
    R("seedling_height_business", "Высота саженца — бизнес", "length", ("высота саженца бизнес", "business class seedling height"), required_context_any=("бизнес", "business"), canonical_unit="m", parser_name="length", notes="Заполнять только при явной привязке к классу бизнес."),
    R("seedling_height_business_plus", "Высота саженца — бизнес+", "length", ("высота саженца бизнес+", "business+ seedling height", "business plus seedling height"), required_context_any=("бизнес+", "business+", "business plus"), canonical_unit="m", parser_name="length", notes="Только явная привязка к бизнес+."),
    R("seedling_height_premium", "Высота саженца — премиум", "length", ("высота саженца премиум", "premium seedling height"), required_context_any=("премиум", "premium"), canonical_unit="m", parser_name="length", notes="Только явная привязка к премиум."),
    R("seedling_diameter_business", "Диаметр саженца — бизнес", "length", ("диаметр саженца бизнес", "ширина саженца бизнес", "business class seedling diameter"), required_context_any=("бизнес", "business"), canonical_unit="m", parser_name="length", notes="Только явная привязка к классу бизнес."),
    R("seedling_diameter_business_plus", "Диаметр саженца — бизнес+", "length", ("диаметр саженца бизнес+", "business+ seedling diameter", "business plus seedling diameter"), required_context_any=("бизнес+", "business+", "business plus"), canonical_unit="m", parser_name="length"),
    R("seedling_diameter_premium", "Диаметр саженца — премиум", "length", ("диаметр саженца премиум", "premium seedling diameter"), required_context_any=("премиум", "premium"), canonical_unit="m", parser_name="length"),
    R("seedling_girth_business", "Обхват ствола саженца — бизнес", "length", ("обхват ствола саженца бизнес", "business seedling girth"), required_context_any=("бизнес", "business"), canonical_unit="m", parser_name="length"),
    R("seedling_girth_business_plus", "Обхват ствола саженца — бизнес+", "length", ("обхват ствола саженца бизнес+", "business+ seedling girth", "business plus seedling girth"), required_context_any=("бизнес+", "business+", "business plus"), canonical_unit="m", parser_name="length"),
    R("seedling_girth_premium", "Обхват ствола саженца — премиум", "length", ("обхват ствола саженца премиум", "premium seedling girth"), required_context_any=("премиум", "premium"), canonical_unit="m", parser_name="length"),
    R("rootball_container_size", "Размер кома / контейнера", "container_size", ("размер кома", "корневой ком", "root ball", "rootball", "container size", "pot size", "размер контейнера", "размер горшка", "container"), canonical_unit="m / L", parser_name="container_size", notes="Может содержать несколько размерностей: длины canonical=m, объём canonical=L; всё хранится отдельно в evidence."),
    R("light", "Световой режим", "enum", ("световой режим", "освещение", "экспозиция", "light", "light requirements", "sun exposure", "aspect"), enum_map=LIGHT_ENUM, notes="Нормализуется только в солнце/полутень/тень; допускается несколько значений."),
    R("planting_time", "Посадка / рекомендуемый срок", "period", ("срок посадки", "время посадки", "посадка", "planting time", "planting season", "when to plant"), parser_name="period", notes="Месяцы/сезоны; ограничения сохранять как отдельный raw evidence."),
    R("care_watering", "Уход — полив", "text", ("полив", "watering", "water requirements", "irrigation"), exclude_terms=("waterlogging", "flood", "затоп", "переувлаж"), notes="Только рекомендации по поливу; не путать с переносимостью переувлажнения."),
    R("care_feeding", "Уход — подкормки", "text", ("подкормка", "удобрение", "fertilizing", "fertilising", "feeding", "fertilizer", "fertiliser"), notes="Явные рекомендации по подкормкам."),
    R("care_pruning", "Уход — обрезка", "text", ("обрезка", "подрезка", "pruning", "trimming", "cut back"), notes="Явные рекомендации по обрезке/формовке."),
    R("hardiness", "Зимостойкость", "hardiness", ("зимостойкость", "морозостойкость", "hardiness", "winter hardiness", "usda zone", "hardiness zone", "minimum temperature"), parser_name="hardiness", notes="USDA, температура и качественное описание сохраняются раздельно; не конвертировать одно в другое без явных данных."),
    R("decorative_season", "Сезон декоративности", "period", ("сезон декоративности", "декоративный период", "season of interest", "seasonal interest", "interest"), parser_name="period", notes="Месяцы/сезоны общего декоративного эффекта."),
    R("decorative_peak", "Пик декоративности", "period", ("пик декоративности", "peak interest", "best season", "maximum ornamental effect"), parser_name="period", notes="Только явно обозначенный пик/лучший период."),
    R("decorative_properties", "Декоративные свойства", "list_text", ("декоративные свойства", "декоративность", "ornamental features", "key features", "features", "garden interest"), notes="Только явно перечисленные декоративные качества; не копировать весь маркетинговый текст."),
    R("branches", "Ветви — окраска коры / форма", "text", ("ветви", "ветки", "кора", "branches", "branching", "bark", "stem colour", "stem color"), notes="Окраска коры/побегов и форма ветвей; не листва."),
    R("leaf_form", "Листья / хвоя — форма", "text", ("форма листьев", "форма хвои", "leaf shape", "foliage shape", "needle shape", "листья", "хвоя"), exclude_terms=("цвет", "colour", "color", "окраска"), notes="Форма/тип листовой пластинки или хвои."),
    R("leaf_summer_color", "Листья / хвоя — летняя окраска", "text", ("летняя окраска листьев", "летняя окраска хвои", "summer foliage", "summer leaf colour", "summer leaf color", "foliage colour", "foliage color"), exclude_terms=("autumn", "fall", "осен"), notes="Только вегетационная/летняя окраска; осень исключается."),
    R("leaf_autumn_color", "Листья / хвоя — осенняя окраска", "text", ("осенняя окраска", "autumn colour", "autumn color", "fall color", "fall colour", "autumn foliage"), notes="Только осенняя окраска."),
    R("flower_color", "Цветки — цветовая гамма", "text", ("цвет цветков", "окраска цветков", "flower colour", "flower color", "bloom color", "bloom colour", "цветки"), exclude_terms=("flowering time", "срок цветения", "период цветения"), notes="Цвет/гамма цветков, не сроки."),
    R("flowering_time", "Цветки — сроки цветения", "period", ("сроки цветения", "период цветения", "flowering time", "flowering period", "bloom time", "bloom season"), parser_name="period", notes="Месяцы/сезоны цветения."),
    R("inflorescences", "Соцветия", "text", ("соцветия", "inflorescence", "flower clusters", "flower heads"), notes="Тип/форма/размер соцветий; не общий цвет цветков."),
    R("fruits", "Плоды", "text", ("плоды", "fruit", "berries", "berry", "cones", "шишки", "семенные коробочки"), notes="Описание плодов/ягод/шишек, включая цвет/период при явном указании."),
    R("drought_tolerance", "Устойчивость к засухе", "explicit_tolerance", ("засухоустойчивость", "drought tolerance", "drought tolerant", "drought resistance"), notes="Заполнять только при явном утверждении переносимости/непереносимости засухи; 'любит сухую почву' само по себе недостаточно."),
    R("waterlogging_tolerance", "Устойчивость к обводнению", "explicit_tolerance", ("устойчивость к переувлажнению", "переносит переувлажнение", "waterlogging tolerance", "waterlogging tolerant", "flood tolerance", "flood tolerant", "standing water"), notes="Только явная переносимость/непереносимость переувлажнения/затопления. 'Moist/wet soil' не считается доказательством."),
    R("soil_ph", "Почва — pH", "ph", ("ph", "soil ph", "кислотность почвы", "реакция почвы", "soil reaction"), parser_name="ph", notes="Числовой pH 0–14 и/или явные acidic/neutral/alkaline."),
    R("soil_composition", "Почва — состав", "text", ("тип почвы", "состав почвы", "почва", "soil type", "soil texture", "soil", "loam", "clay", "sand"), exclude_terms=("ph", "fertility", "плодород", "moisture", "влаж"), notes="Текстура/тип: песчаная, суглинок, глина и т.п.; не смешивать с pH, влажностью, плодородием."),
    R("soil_special", "Почва — специальные требования", "text", ("требования к почве", "soil requirements", "special soil requirements", "drainage", "дренаж", "well drained", "известков", "lime free"), exclude_terms=("ph",), notes="Дренаж, известь, специфический субстрат и др.; pH хранить отдельно."),
    R("soil_fertility", "Почва — плодородность", "enum_or_text", ("плодородность", "fertility", "soil fertility", "rich soil", "poor soil"), enum_map=SOIL_FERTILITY_ENUM, notes="Нормализованная категория + исходное описание."),
    R("on_parking", "Рост — подземный паркинг", "suitability", ("подземный паркинг", "над паркингом", "parking deck", "underground parking", "podium over parking"), notes="Да/нет только при явной пригодности/непригодности, не по одному упоминанию объекта."),
    R("on_stylobate", "Рост — стилобат", "suitability", ("стилобат", "эксплуатируемая кровля", "podium", "roof garden", "green roof", "rooftop"), notes="Да/нет только при явной пригодности/непригодности."),
    R("in_ground", "Рост — грунт", "suitability", ("открытый грунт", "в грунте", "open ground", "in ground", "landscape planting"), notes="Да/нет только при явной формулировке."),
    R("in_container", "Рост — кашпо", "suitability", ("в кашпо", "в контейнере", "container growing", "container planting", "pot culture", "patio container"), exclude_terms=("container size", "pot size", "размер контейнера", "размер горшка"), notes="Не путать пригодность к выращиванию в кашпо с размером торгового контейнера."),
    R("substrate_depth", "Глубина субстрата", "length", ("глубина субстрата", "substrate depth", "soil depth", "minimum soil depth", "rooting depth"), canonical_unit="m", parser_name="length", notes="Явная требуемая/минимальная глубина; canonical=m."),
    R("features", "Особенности", "list_text", ("особенности", "характеристики", "special features", "characteristics", "traits"), notes="Общие подтверждённые особенности; не дублировать специальные свойства, если можно классифицировать точнее."),
    R("special_properties", "Особые свойства", "enum_or_text", ("особые свойства", "medicinal", "edible", "honey plant", "pollinator", "лекарственные свойства", "съедобные плоды", "медонос"), enum_map=SPECIAL_PROPERTY_ENUM, notes="Лекарственное, съедобное, медонос, опылители, токсичность и т.п.; только явные утверждения."),
    R("planting_type", "Тип насаждений", "enum_or_text", ("использование", "ландшафтное использование", "тип посадки", "garden use", "landscape use", "uses", "planting type"), enum_map=PLANTING_TYPE_ENUM, notes="Солитер, массив, группа, изгородь и т.п.; только явное применение."),
    R("pioneer_projects", "Применялось на проектах PIONEER", "list_text", ("pioneer", "проекты pioneer", "применено на проектах", "реализованные проекты"), required_context_any=("pioneer", "пионер"), notes="Только конкретно подтверждённые проекты PIONEER; на внешних сайтах обычно пусто."),
    R("cost_material", "Стоимость материала", "currency", ("цена растения", "стоимость растения", "цена саженца", "plant price", "material cost", "price"), exclude_terms=("planting", "посадк", "delivery", "достав", "care", "уход", "warranty", "гарант"), parser_name="currency", notes="Только стоимость самого посадочного материала. Валюту не конвертировать."),
    R("cost_planting", "Стоимость посадки", "currency", ("стоимость посадки", "цена посадки", "planting cost", "installation cost"), parser_name="currency", notes="Только явно отделённая стоимость посадочных работ."),
    R("cost_warranty", "Стоимость гарантии", "currency", ("стоимость гарантии", "warranty cost", "guarantee cost"), parser_name="currency", notes="Только явно отделённая стоимость гарантии."),
    R("cost_establishment_care", "Стоимость ухода в период приживания", "currency", ("уход в период приживания", "establishment care cost", "establishment maintenance cost"), required_context_any=("приж", "establishment"), parser_name="currency", notes="Только период приживания."),
    R("cost_establishment_pruning", "Стоимость формовки в период приживания", "currency", ("формовка в период приживания", "establishment pruning cost", "formative pruning cost"), required_context_any=("приж", "establishment", "formative"), parser_name="currency", notes="Только формовка/обрезка в период приживания."),
    R("cost_operation", "Стоимость ухода в эксплуатации", "currency", ("стоимость эксплуатации", "стоимость ухода", "maintenance cost", "operational maintenance", "opex"), exclude_terms=("приж", "establishment"), parser_name="currency", notes="Эксплуатационный уход после периода приживания; не смешивать с establishment."),
    R("short_description", "Краткое описание", "text", ("краткое описание", "short description", "summary", "overview"), notes="Короткое описание из явного summary/overview; если только один большой description, не дублировать автоматически в оба поля."),
    R("long_description", "Подробное описание", "text", ("подробное описание", "full description", "detailed description", "description", "about the plant"), notes="Основной развернутый текст карточки; маркетинговые блоки вне карточки исключать."),
    R("photo_decorative", "Фото — декоративный период", "image", ("декоративный период", "flowering", "bloom", "autumn colour", "осенняя окраска", "цветение", "fruiting", "плодоношение"), notes="URL изображения только если alt/caption/context явно указывает декоративный период."),
    R("photo_background", "Фото — фоновый период", "image", ("summer habit", "foliage", "общий вид", "habit", "фоновой период"), exclude_terms=("winter", "зима", "flower", "bloom", "цветен", "autumn", "осен"), notes="URL изображения общего вида без выраженного декоративного события; нужна явная подпись/context."),
    R("photo_winter", "Фото — зимой", "image", ("winter", "зима", "winter habit", "winter interest", "snow"), notes="URL изображения только при явном зимнем alt/caption/context."),
    R("source_page_url", "Источник — URL карточки", "generated", notes="Фактический URL карточки, на которой извлечены данные."),
    R("source_image_url", "Источник — URL изображения", "generated", notes="Все уникальные URL изображений, которые были классифицированы в фото-поля."),
    R("scrape_date", "Дата сбора", "generated", notes="ISO дата/время сбора."),
    R("comment", "Комментарий / неоднозначность", "generated", notes="Автоматический список неоднозначностей, отклонённых кандидатов и конфликтов."),
    R("confidence", "Уверенность сопоставления", "generated", notes="Средняя/минимальная уверенность принятых observations; уверенность также хранится по каждому evidence."),
]

RULE_BY_KEY = {r.key: r for r in FIELD_RULES}


# -----------------------------------------------------------------------------
# Taxonomy parsing
# -----------------------------------------------------------------------------

LATIN_NAME_RE = re.compile(
    r"\b([A-Z][a-z]{2,})\s+(×\s*)?([a-z][a-z-]{2,})(?:\s+(?:subsp\.|ssp\.|var\.|f\.)\s+([a-z][a-z-]{2,}))?(?:\s+['‘’\"]([^'‘’\"]{1,80})['‘’\"])?"
)


def extract_latin_name(text: str) -> Optional[dict[str, str]]:
    # Preserve capitalization: use original text, normalized spacing only.
    for m in LATIN_NAME_RE.finditer(normalize_space(text)):
        genus = m.group(1)
        species = ("× " if m.group(2) else "") + m.group(3)
        infra = m.group(4) or ""
        cultivar = m.group(5) or ""
        return {"genus": genus, "species": species, "infra": infra, "cultivar": cultivar}
    return None


def extract_cultivar(text: str) -> str:
    m = re.search(r"['‘’\"]([^'‘’\"]{1,80})['‘’\"]", normalize_space(text))
    return normalize_space(m.group(1)) if m else ""


def split_ru_name(text: str) -> dict[str, str]:
    # Conservative Russian split. Only first two words + quoted cultivar.
    text = normalize_space(text)
    cultivar = extract_cultivar(text)
    bare = re.sub(r"['‘’\"][^'‘’\"]+['‘’\"]", "", text)
    tokens = re.findall(r"[А-Яа-яЁё-]+", bare)
    return {
        "genus": tokens[0] if len(tokens) >= 1 else "",
        "species": tokens[1] if len(tokens) >= 2 else "",
        "cultivar": cultivar,
    }


def canonical_taxon_key(record: PlantRecord) -> str:
    def best(key: str) -> str:
        obs = [o for o in record.observations.get(key, []) if o.accepted_to_main]
        if not obs:
            return ""
        obs.sort(key=lambda o: o.confidence, reverse=True)
        return obs[0].display_value()

    latin_parts = [best("latin_genus"), best("latin_species"), best("latin_cultivar")]
    latin = " ".join(x for x in latin_parts if x).strip()
    if latin:
        return "latin:" + norm_text(latin)
    ru_parts = [best("ru_genus"), best("ru_species"), best("ru_cultivar")]
    ru = " ".join(x for x in ru_parts if x).strip()
    if ru:
        return "ru:" + norm_text(ru)
    return "url:" + norm_text(record.page_url)


def make_id(key: str) -> str:
    return "PLT-" + hashlib.sha1(key.encode("utf-8")).hexdigest()[:12].upper()


# -----------------------------------------------------------------------------
# Candidate extraction from HTML
# -----------------------------------------------------------------------------

GENERIC_NOISE_LABELS = {"menu", "navigation", "cookie", "cookies", "footer", "header", "share", "social"}


def get_visible_text(node: Tag) -> str:
    return normalize_space(node.get_text(" ", strip=True))


def extract_label_value_candidates(soup: BeautifulSoup) -> list[Candidate]:
    out: list[Candidate] = []

    # Definition lists.
    for dt_tag in soup.find_all("dt"):
        dd = dt_tag.find_next_sibling("dd")
        if dd:
            out.append(Candidate(get_visible_text(dt_tag), get_visible_text(dd), get_visible_text(dt_tag.parent), "dl"))

    # Two-column table rows.
    for tr in soup.find_all("tr"):
        cells = tr.find_all(["th", "td"], recursive=False)
        if len(cells) >= 2:
            label = get_visible_text(cells[0])
            value = get_visible_text(cells[1])
            if label and value and label != value:
                out.append(Candidate(label, value, get_visible_text(tr), "table"))

    # Generic key: value list items / paragraphs.
    for node in soup.find_all(["li", "p"]):
        text = get_visible_text(node)
        if 3 <= len(text) <= 800 and ":" in text:
            label, value = text.split(":", 1)
            if 1 <= len(label) <= 120 and value.strip():
                out.append(Candidate(label, value, text, "colon"))

    # Common property/specification cards. Use class name + descendant text.
    for node in soup.find_all(["div", "section"]):
        classes = " ".join(node.get("class", []))
        if not re.search(r"(property|attribute|spec|characteristic|feature|parameter|detail|характер|параметр|свойств)", classes, re.I):
            continue
        children = [c for c in node.find_all(["span", "div", "p", "strong", "b"], recursive=False) if get_visible_text(c)]
        if len(children) == 2:
            label, value = get_visible_text(children[0]), get_visible_text(children[1])
            if label and value and len(label) <= 120:
                out.append(Candidate(label, value, get_visible_text(node), "property-card"))

    # De-duplicate.
    seen: set[tuple[str, str]] = set()
    deduped: list[Candidate] = []
    for c in out:
        key = (norm_text(c.label), norm_text(c.value))
        if not key[0] or key[0] in GENERIC_NOISE_LABELS or not key[1] or key in seen:
            continue
        seen.add(key)
        deduped.append(c)
    return deduped


def extract_jsonld_candidates(soup: BeautifulSoup) -> tuple[list[Candidate], list[tuple[str, str]]]:
    """Extract structured Product/Plant-like data without assuming exact schema."""
    candidates: list[Candidate] = []
    images: list[tuple[str, str]] = []

    def walk(obj: Any) -> None:
        if isinstance(obj, list):
            for item in obj:
                walk(item)
            return
        if not isinstance(obj, dict):
            return
        if "@graph" in obj:
            walk(obj.get("@graph"))
        name = normalize_space(str(obj.get("name", "")))
        desc = normalize_space(str(obj.get("description", "")))
        if desc:
            candidates.append(Candidate("description", desc, desc[:1000], "json-ld"))
        additional = obj.get("additionalProperty") or obj.get("additionalProperties")
        if isinstance(additional, dict):
            additional = [additional]
        if isinstance(additional, list):
            for prop in additional:
                if isinstance(prop, dict):
                    label = normalize_space(str(prop.get("name") or prop.get("propertyID") or ""))
                    value = normalize_space(str(prop.get("value") or prop.get("valueReference") or ""))
                    if label and value:
                        candidates.append(Candidate(label, value, f"{label}: {value}", "json-ld-additionalProperty"))
        offers = obj.get("offers")
        offers_list = offers if isinstance(offers, list) else ([offers] if isinstance(offers, dict) else [])
        for offer in offers_list:
            price = offer.get("price")
            currency = offer.get("priceCurrency")
            if price not in (None, "") and currency:
                candidates.append(Candidate("price", f"{price} {currency}", f"JSON-LD offer for {name}", "json-ld-offer"))
        image_obj = obj.get("image")
        img_list = image_obj if isinstance(image_obj, list) else ([image_obj] if image_obj else [])
        for img in img_list:
            if isinstance(img, str):
                images.append((img, name))
            elif isinstance(img, dict) and img.get("url"):
                images.append((str(img["url"]), name))
        # Recurse into nested dicts that may contain data.
        for key, value in obj.items():
            if key not in {"@graph", "additionalProperty", "additionalProperties", "offers", "image"} and isinstance(value, (dict, list)):
                walk(value)

    for script in soup.find_all("script", attrs={"type": "application/ld+json"}):
        raw = script.string or script.get_text("", strip=True)
        if not raw:
            continue
        try:
            walk(json.loads(raw))
        except Exception:
            continue
    return candidates, images


def extract_page_title(soup: BeautifulSoup) -> str:
    h1 = soup.find("h1")
    if h1 and get_visible_text(h1):
        return get_visible_text(h1)
    if soup.title:
        return get_visible_text(soup.title)
    return ""


def extract_main_description_candidates(soup: BeautifulSoup) -> list[Candidate]:
    # Only likely article/product main content; avoid footer/nav.
    selectors = ["main", "article", "[itemprop='description']", ".description", ".product-description", ".plant-description"]
    out: list[Candidate] = []
    for selector in selectors:
        for node in soup.select(selector):
            text = get_visible_text(node)
            if 80 <= len(text) <= 8000:
                out.append(Candidate("description", text, text[:1000], "main-description"))
    return out[:3]


def extract_images(soup: BeautifulSoup, page_url: str) -> list[tuple[str, str]]:
    """Return (image_url, context). No classification yet."""
    images: list[tuple[str, str]] = []
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or img.get("data-lazy-src") or ""
        srcset = img.get("srcset") or img.get("data-srcset") or ""
        if not src and srcset:
            first = srcset.split(",")[0].strip().split(" ")[0]
            src = first
        if not src:
            continue
        url = urljoin(page_url, src)
        alt = img.get("alt", "")
        title = img.get("title", "")
        fig = img.find_parent("figure")
        caption = ""
        if fig:
            cap = fig.find("figcaption")
            if cap:
                caption = get_visible_text(cap)
        context = normalize_space(" | ".join(x for x in [alt, title, caption] if x))
        # Skip icons/logos by URL/context heuristics.
        if re.search(r"(logo|icon|sprite|avatar|flag|payment|social)", (url + " " + context), re.I):
            continue
        images.append((url, context))
    return images


# -----------------------------------------------------------------------------
# Semantic matching and strict validation
# -----------------------------------------------------------------------------


def contains_any(text: str, terms: Iterable[str]) -> bool:
    t = norm_text(text)
    return any(norm_text(term) in t for term in terms if term)


def token_set(text: str) -> set[str]:
    return {t for t in norm_text(text).split() if len(t) > 2}


def label_score(label: str, rule: FieldRule) -> float:
    l = norm_text(label)
    if not l or rule.value_type in {"generated", "image"}:
        return 0.0
    if contains_any(l, rule.exclude_terms):
        return 0.0
    best = 0.0
    ltokens = token_set(l)
    for alias in rule.aliases:
        a = norm_text(alias)
        if not a:
            continue
        if l == a:
            best = max(best, 1.0)
            continue
        if a in l:
            best = max(best, 0.96)
            continue
        if l in a and len(l) >= 5:
            best = max(best, 0.90)
            continue
        if rule.allow_fuzzy:
            atokens = token_set(a)
            overlap = len(ltokens & atokens) / max(1, len(ltokens | atokens))
            seq = SequenceMatcher(None, l, a).ratio()
            # Fuzzy never reaches exact-confidence; used only with parser validation.
            if overlap >= 0.50:
                best = max(best, min(0.88, 0.70 + 0.18 * overlap))
            elif seq >= 0.88:
                best = max(best, 0.82)
    return best


def context_requirements_ok(candidate: Candidate, rule: FieldRule) -> bool:
    combined = f"{candidate.label} {candidate.context}"
    if rule.required_context_any and not contains_any(combined, rule.required_context_any):
        return False
    if rule.exclude_terms and contains_any(combined, rule.exclude_terms):
        # Exclude terms in value/context matter for confusion-prone fields.
        # Exact label alias still cannot override forbidden context.
        return False
    return True


def parse_rule_value(rule: FieldRule, raw: str, semantic_context: str = "") -> list[ParsedValue]:
    combined = normalize_space(f"{semantic_context} {raw}")
    if rule.value_type == "enum":
        return parse_enum(raw, rule.enum_map)
    if rule.value_type == "growth_rate":
        return parse_growth_rate(raw)
    if rule.value_type == "enum_or_text":
        enum_vals = parse_enum(raw, rule.enum_map)
        if enum_vals:
            # Preserve raw textual value too only if it includes numeric/detail beyond enum.
            if re.search(r"\d", raw):
                enum_vals.extend(parse_text(raw))
            return dedupe_parsed(enum_vals)
        return parse_text(raw)
    if rule.value_type == "explicit_tolerance":
        kind = "drought" if rule.key == "drought_tolerance" else "waterlogging"
        explicit = parse_explicit_tolerance(combined, kind)
        if explicit:
            return explicit
        # Strict categorical value is allowed only because the label/context already
        # identified the tolerance dimension. This handles e.g. "Drought tolerance: High".
        v = norm_text(raw)
        if re.fullmatch(r"(?:yes|да|true|устойчиво|переносит)", v):
            return [ParsedValue(raw, "устойчиво / переносит")]
        if re.fullmatch(r"(?:no|нет|false|неустойчиво|не переносит)", v):
            return [ParsedValue(raw, "неустойчиво / не переносит")]
        if re.fullmatch(r"(?:high|высокая|высокий|strong)", v):
            return [ParsedValue(raw, "высокая")]
        if re.fullmatch(r"(?:medium|moderate|средняя|средний)", v):
            return [ParsedValue(raw, "средняя")]
        if re.fullmatch(r"(?:low|низкая|низкий|weak)", v):
            return [ParsedValue(raw, "низкая")]
        return []
    if rule.value_type == "suitability":
        # The target object is identified by the label/context; the value may simply be yes/no.
        v = norm_text(raw)
        if re.fullmatch(r"(?:yes|да|true|suitable|recommended|подходит|пригодно|рекомендуется)", v):
            return [ParsedValue(raw, "да")]
        if re.fullmatch(r"(?:no|нет|false|unsuitable|not suitable|не подходит|непригодно|не рекомендуется)", v):
            return [ParsedValue(raw, "нет")]
        return parse_bool_suitability(combined, rule.aliases)
    parser = PARSERS.get(rule.parser_name or rule.value_type, parse_text)
    return parser(raw)


def strict_value_validation(rule: FieldRule, parsed: list[ParsedValue], raw: str) -> tuple[bool, str]:
    if not parsed:
        return False, "value parser found no admissible value"
    if rule.value_type == "length":
        # Reject obviously implausible generic numeric captures; length parser already needs units.
        if not any(v.canonical_unit == "m" for v in parsed):
            return False, "no explicit length unit"
        # Avoid zero/negative dimensions.
        if any((v.min_value is not None and v.min_value < 0) for v in parsed):
            return False, "negative dimension"
    if rule.value_type == "currency" and not any(v.canonical_unit in {"RUB", "EUR", "USD", "GBP"} for v in parsed):
        return False, "currency not explicit"
    if rule.key in {"waterlogging_tolerance", "drought_tolerance"}:
        allowed = {"устойчиво / переносит", "неустойчиво / не переносит", "высокая", "средняя", "низкая"}
        if not any(v.normalized_value in allowed for v in parsed):
            return False, "tolerance must be explicitly stated or supplied as a strict categorical value"
    if rule.value_type == "ph" and not parse_ph(raw):
        return False, "not an explicit pH value/descriptor"
    return True, ""


def map_candidates_to_observations(
    plant_id: str,
    candidates: list[Candidate],
    source_url: str,
    scraped_at: str,
) -> tuple[list[Observation], list[Observation]]:
    accepted: list[Observation] = []
    rejected: list[Observation] = []

    for candidate in candidates:
        scores: list[tuple[float, FieldRule]] = []
        for rule in FIELD_RULES:
            if rule.value_type in {"generated", "image", "taxonomy_latin", "taxonomy_ru"}:
                continue
            s = label_score(candidate.label, rule)
            if s > 0:
                scores.append((s, rule))
        scores.sort(key=lambda x: x[0], reverse=True)
        if not scores:
            continue

        top_score, rule = scores[0]
        # Ambiguity guard: if two distinct fields are almost tied, reject to raw evidence.
        second_score = scores[1][0] if len(scores) > 1 else 0.0
        ambiguous = second_score >= 0.90 and abs(top_score - second_score) < 0.04
        if ambiguous:
            rejected.append(Observation(
                plant_id, "_ambiguous", "Неоднозначный кандидат", candidate.label, candidate.value,
                "", None, None, "", "", top_score, candidate.method,
                candidate.context, source_url, scraped_at, False,
                f"ambiguous field mapping: {scores[0][1].label} vs {scores[1][1].label}",
            ))
            continue

        if top_score < rule.min_confidence:
            rejected.append(Observation(
                plant_id, rule.key, rule.label, candidate.label, candidate.value,
                "", None, None, "", rule.canonical_unit, top_score, candidate.method,
                candidate.context, source_url, scraped_at, False,
                f"label confidence {top_score:.2f} below threshold {rule.min_confidence:.2f}",
            ))
            continue

        if not context_requirements_ok(candidate, rule):
            rejected.append(Observation(
                plant_id, rule.key, rule.label, candidate.label, candidate.value,
                "", None, None, "", rule.canonical_unit, top_score, candidate.method,
                candidate.context, source_url, scraped_at, False,
                "required context missing or forbidden context present",
            ))
            continue

        parsed_values = parse_rule_value(rule, candidate.value, f"{candidate.label} {candidate.context}")
        valid, reason = strict_value_validation(rule, parsed_values, candidate.value)
        if not valid:
            rejected.append(Observation(
                plant_id, rule.key, rule.label, candidate.label, candidate.value,
                "", None, None, "", rule.canonical_unit, top_score, candidate.method,
                candidate.context, source_url, scraped_at, False, reason,
            ))
            continue

        for pv in parsed_values:
            accepted.append(Observation(
                plant_id=plant_id,
                field_key=rule.key,
                field_label=rule.label,
                raw_label=candidate.label,
                raw_value=normalize_space(candidate.value),
                normalized_value=pv.normalized_value,
                min_value=pv.min_value,
                max_value=pv.max_value,
                source_unit=pv.source_unit,
                canonical_unit=pv.canonical_unit or rule.canonical_unit,
                confidence=top_score,
                extraction_method=candidate.method,
                evidence_text=normalize_space(candidate.context)[:2000],
                source_url=source_url,
                scraped_at=scraped_at,
                accepted_to_main=True,
            ))

    return accepted, rejected


# -----------------------------------------------------------------------------
# Direct taxonomy extraction (page title + explicit fields)
# -----------------------------------------------------------------------------


def taxonomy_observations(plant_id: str, soup: BeautifulSoup, candidates: list[Candidate], url: str, scraped_at: str) -> list[Observation]:
    out: list[Observation] = []
    title = extract_page_title(soup)
    latin_sources: list[tuple[str, str, float]] = [(title, "h1/title", 0.93)]
    ru_sources: list[tuple[str, str, float]] = []

    for c in candidates:
        lab = norm_text(c.label)
        if any(norm_text(a) in lab for a in ("latin name", "scientific name", "botanical name", "латинское название", "ботаническое название")):
            latin_sources.append((c.value, c.method, 0.99))
        if any(norm_text(a) in lab for a in ("русское название", "название на русском")):
            ru_sources.append((c.value, c.method, 0.99))

    # Latin.
    best_latin: Optional[tuple[dict[str, str], str, str, float]] = None
    for text, method, conf in latin_sources:
        parsed = extract_latin_name(text)
        if parsed:
            if best_latin is None or conf > best_latin[3]:
                best_latin = (parsed, text, method, conf)
    if best_latin:
        parsed, raw, method, conf = best_latin
        mapping = [("latin_genus", parsed["genus"]), ("latin_species", parsed["species"]), ("latin_cultivar", parsed["cultivar"])]
        for key, value in mapping:
            if value:
                out.append(Observation(plant_id, key, RULE_BY_KEY[key].label, "botanical name", raw, value, None, None, "", "", conf, method, raw, url, scraped_at))

    # Russian explicit only. If no explicit field, derive cautiously from title after removing Latin.
    if not ru_sources and title:
        title_wo_latin = LATIN_NAME_RE.sub("", title)
        if re.search(r"[А-Яа-яЁё]", title_wo_latin):
            ru_sources.append((title_wo_latin, "h1/title", 0.78))
    for raw, method, conf in ru_sources[:1]:
        parsed = split_ru_name(raw)
        if parsed["genus"]:
            out.append(Observation(plant_id, "ru_genus", RULE_BY_KEY["ru_genus"].label, "русское название", raw, parsed["genus"], None, None, "", "", conf, method, raw, url, scraped_at, conf >= 0.86, "" if conf >= 0.86 else "Russian taxonomy from title is low-confidence"))
        if parsed["species"]:
            out.append(Observation(plant_id, "ru_species", RULE_BY_KEY["ru_species"].label, "русское название", raw, parsed["species"], None, None, "", "", conf, method, raw, url, scraped_at, conf >= 0.86, "" if conf >= 0.86 else "Russian taxonomy from title is low-confidence"))
        if parsed["cultivar"]:
            out.append(Observation(plant_id, "ru_cultivar", RULE_BY_KEY["ru_cultivar"].label, "русское название", raw, parsed["cultivar"], None, None, "", "", conf, method, raw, url, scraped_at, conf >= 0.86, "" if conf >= 0.86 else "Russian cultivar from title is low-confidence"))
    return out


# -----------------------------------------------------------------------------
# Images: strict context classification only.
# -----------------------------------------------------------------------------


def image_observations(plant_id: str, soup: BeautifulSoup, url: str, scraped_at: str) -> tuple[list[Observation], list[str]]:
    out: list[Observation] = []
    unclassified: list[str] = []
    rules = [RULE_BY_KEY["photo_winter"], RULE_BY_KEY["photo_decorative"], RULE_BY_KEY["photo_background"]]
    for image_url, context in extract_images(soup, url):
        if not context:
            unclassified.append(image_url)
            continue
        matched = False
        for rule in rules:
            ctx = norm_text(context)
            if contains_any(ctx, rule.exclude_terms):
                continue
            if contains_any(ctx, rule.aliases):
                out.append(Observation(
                    plant_id, rule.key, rule.label, "image context", image_url, image_url,
                    None, None, "URL", "URL", 0.92, "image-alt-caption", context,
                    url, scraped_at, True, "",
                ))
                matched = True
                break
        if not matched:
            unclassified.append(image_url)
    return out, unique_preserve(unclassified)


# -----------------------------------------------------------------------------
# Page heuristics and crawling
# -----------------------------------------------------------------------------

PLANT_HINTS = (
    "высота", "зимостойкость", "цветение", "почва", "свет", "растение", "сорт",
    "height", "hardiness", "flowering", "soil", "sun", "plant", "cultivar",
)


def is_probable_plant_page(soup: BeautifulSoup) -> bool:
    title = extract_page_title(soup)
    if extract_latin_name(title):
        return True
    candidates = extract_label_value_candidates(soup)
    labels = " ".join(norm_text(c.label) for c in candidates)
    hits = sum(1 for hint in PLANT_HINTS if hint in labels)
    has_latin_anywhere = bool(extract_latin_name(get_visible_text(soup.find("main") or soup.body or soup)[:4000]))
    return hits >= 3 or (hits >= 1 and has_latin_anywhere)


def normalize_url(url: str) -> str:
    clean, _ = urldefrag(url)
    p = urlparse(clean)
    # strip tracking query parameters conservatively
    if not p.query:
        return clean
    keep_parts = []
    for part in p.query.split("&"):
        key = part.split("=", 1)[0].casefold()
        if key.startswith("utm_") or key in {"fbclid", "gclid", "yclid"}:
            continue
        keep_parts.append(part)
    query = "&".join(keep_parts)
    return p._replace(query=query).geturl()


def same_host(a: str, b: str) -> bool:
    return urlparse(a).netloc.casefold() == urlparse(b).netloc.casefold()


def extract_internal_links(soup: BeautifulSoup, page_url: str, root_url: str) -> list[str]:
    out: list[str] = []
    for a in soup.find_all("a", href=True):
        href = normalize_space(a.get("href", ""))
        if not href or href.startswith(("mailto:", "tel:", "javascript:")):
            continue
        url = normalize_url(urljoin(page_url, href))
        if same_host(url, root_url):
            out.append(url)
    return unique_preserve(out)


class Fetcher:
    def __init__(self, browser: str, timeout: float, delay: float, retries: int):
        self.browser = browser
        self.timeout = timeout
        self.delay = delay
        self.retries = retries
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "ru,en;q=0.8"})
        self._playwright = None
        self._browser_obj = None

    def close(self) -> None:
        if self._browser_obj:
            self._browser_obj.close()
        if self._playwright:
            self._playwright.stop()
        self.session.close()

    def _fetch_requests(self, url: str) -> tuple[str, int]:
        last_exc: Optional[Exception] = None
        for attempt in range(self.retries + 1):
            try:
                resp = self.session.get(url, timeout=self.timeout)
                if resp.status_code in {429, 500, 502, 503, 504} and attempt < self.retries:
                    time.sleep(self.delay * (2 ** attempt))
                    continue
                resp.raise_for_status()
                ctype = resp.headers.get("content-type", "")
                if "text/html" not in ctype and "application/xhtml" not in ctype and ctype:
                    raise ValueError(f"Unsupported content type: {ctype}")
                return resp.text, resp.status_code
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
                if attempt < self.retries:
                    time.sleep(self.delay * (2 ** attempt))
        raise RuntimeError(f"request failed: {last_exc}")

    def _ensure_playwright(self) -> None:
        if self._browser_obj:
            return
        try:
            from playwright.sync_api import sync_playwright  # type: ignore
        except ImportError as exc:
            raise RuntimeError("Playwright is not installed. pip install playwright && playwright install chromium") from exc
        self._playwright = sync_playwright().start()
        self._browser_obj = self._playwright.chromium.launch(headless=True)

    def _fetch_playwright(self, url: str) -> tuple[str, int]:
        self._ensure_playwright()
        page = self._browser_obj.new_page(user_agent=USER_AGENT)
        try:
            response = page.goto(url, wait_until="networkidle", timeout=int(self.timeout * 1000))
            status = response.status if response else 200
            return page.content(), status
        finally:
            page.close()

    def fetch(self, url: str) -> tuple[str, int]:
        time.sleep(max(0.0, self.delay))
        if self.browser == "playwright":
            return self._fetch_playwright(url)
        html_text, status = self._fetch_requests(url)
        if self.browser == "auto":
            # If almost no visible content but JS app scripts exist, retry via Playwright if available.
            soup = BeautifulSoup(html_text, "lxml")
            visible = normalize_space(soup.get_text(" ", strip=True))
            if len(visible) < 300 and len(soup.find_all("script")) >= 5:
                try:
                    return self._fetch_playwright(url)
                except Exception as exc:  # noqa: BLE001
                    LOGGER.warning("Playwright fallback unavailable for %s: %s", url, exc)
        return html_text, status


class RobotsGuard:
    def __init__(self, root_url: str, ignore: bool, session: requests.Session):
        self.ignore = ignore
        self.parser: Optional[RobotFileParser] = None
        if ignore:
            return
        root = urlparse(root_url)
        robots_url = f"{root.scheme}://{root.netloc}/robots.txt"
        try:
            rp = RobotFileParser()
            rp.set_url(robots_url)
            resp = session.get(robots_url, timeout=10)
            if resp.ok:
                rp.parse(resp.text.splitlines())
                self.parser = rp
        except Exception as exc:  # noqa: BLE001
            LOGGER.warning("robots.txt could not be read; proceeding conservatively: %s", exc)

    def allowed(self, url: str) -> bool:
        if self.ignore or self.parser is None:
            return True
        return self.parser.can_fetch(USER_AGENT, url)


# -----------------------------------------------------------------------------
# Page -> PlantRecord
# -----------------------------------------------------------------------------


def build_plant_record(url: str, soup: BeautifulSoup, scraped_at: str) -> tuple[PlantRecord, list[Observation]]:
    provisional_id = make_id("url:" + norm_text(url))
    record = PlantRecord(provisional_id, url, extract_page_title(soup))
    candidates = extract_label_value_candidates(soup)
    candidates.extend(extract_main_description_candidates(soup))
    jsonld_candidates, jsonld_images = extract_jsonld_candidates(soup)
    candidates.extend(jsonld_candidates)

    accepted, rejected = map_candidates_to_observations(provisional_id, candidates, url, scraped_at)
    tax_obs = taxonomy_observations(provisional_id, soup, candidates, url, scraped_at)
    image_obs, unclassified = image_observations(provisional_id, soup, url, scraped_at)
    for image_url, image_context in jsonld_images:
        full_url = urljoin(url, image_url)
        classified = False
        for image_rule in (RULE_BY_KEY["photo_winter"], RULE_BY_KEY["photo_decorative"], RULE_BY_KEY["photo_background"]):
            if image_context and contains_any(image_context, image_rule.aliases) and not contains_any(image_context, image_rule.exclude_terms):
                image_obs.append(Observation(provisional_id, image_rule.key, image_rule.label, "json-ld image", full_url, full_url, None, None, "URL", "URL", 0.88, "json-ld-image", image_context, url, scraped_at))
                classified = True
                break
        if not classified:
            unclassified.append(full_url)

    for obs in accepted + tax_obs + image_obs:
        if obs.accepted_to_main:
            record.add(obs)
        else:
            rejected.append(obs)
    record.unclassified_images = unclassified

    # Generated technical observations.
    technical = [
        Observation(provisional_id, "source_page_url", RULE_BY_KEY["source_page_url"].label, "generated", url, url, None, None, "URL", "URL", 1.0, "generated", "", url, scraped_at),
        Observation(provisional_id, "scrape_date", RULE_BY_KEY["scrape_date"].label, "generated", scraped_at, scraped_at, None, None, "ISO-8601", "ISO-8601", 1.0, "generated", "", url, scraped_at),
    ]
    for obs in technical:
        record.add(obs)

    classified_images = [o.normalized_value for k in ("photo_decorative", "photo_background", "photo_winter") for o in record.observations.get(k, [])]
    for image_url in unique_preserve(classified_images):
        record.add(Observation(provisional_id, "source_image_url", RULE_BY_KEY["source_image_url"].label, "generated", image_url, image_url, None, None, "URL", "URL", 1.0, "generated", "", url, scraped_at))

    if rejected:
        reason_counts: dict[str, int] = defaultdict(int)
        for item in rejected:
            reason_counts[item.rejection_reason or "rejected"] += 1
        summary = "; ".join(f"{reason}: {count}" for reason, count in sorted(reason_counts.items()))
        record.add(Observation(
            provisional_id, "comment", RULE_BY_KEY["comment"].label, "generated", summary,
            f"Отклонено кандидатов: {len(rejected)}. {summary}", None, None, "", "", 1.0,
            "generated", "", url, scraped_at, True, ""
        ))

    # Re-key based on best taxonomy and update observation IDs.
    final_id = make_id(canonical_taxon_key(record))
    record.plant_id = final_id
    for obs_list in record.observations.values():
        for obs in obs_list:
            obs.plant_id = final_id
    for obs in rejected:
        obs.plant_id = final_id

    return record, rejected


# -----------------------------------------------------------------------------
# Deduplication: merge records, never delete evidence.
# -----------------------------------------------------------------------------


def merge_records(records: list[PlantRecord]) -> list[PlantRecord]:
    groups: dict[str, PlantRecord] = {}
    for record in records:
        key = canonical_taxon_key(record)
        if key not in groups:
            groups[key] = record
            continue
        target = groups[key]
        for field_key, obs_list in record.observations.items():
            for obs in obs_list:
                obs.plant_id = target.plant_id
                target.add(obs)
        target.unclassified_images = unique_preserve(target.unclassified_images + record.unclassified_images)
    return list(groups.values())


# -----------------------------------------------------------------------------
# Crawl
# -----------------------------------------------------------------------------


def crawl_site(
    root_url: str,
    max_pages: int,
    browser: str,
    timeout: float,
    delay: float,
    retries: int,
    include_regex: str,
    exclude_regex: str,
    ignore_robots: bool,
) -> tuple[list[PlantRecord], list[Observation], list[CrawlEntry]]:
    root_url = normalize_url(root_url)
    fetcher = Fetcher(browser, timeout, delay, retries)
    robots = RobotsGuard(root_url, ignore_robots, fetcher.session)
    queue = deque([root_url])
    visited: set[str] = set()
    records: list[PlantRecord] = []
    rejected: list[Observation] = []
    log: list[CrawlEntry] = []
    include_re = re.compile(include_regex, re.I) if include_regex else None
    exclude_re = re.compile(exclude_regex, re.I) if exclude_regex else None

    try:
        while queue and len(visited) < max_pages:
            url = queue.popleft()
            if url in visited:
                continue
            visited.add(url)
            if not same_host(url, root_url):
                continue
            if include_re and not include_re.search(url) and url != root_url:
                continue
            if exclude_re and exclude_re.search(url):
                continue
            if not robots.allowed(url):
                log.append(CrawlEntry(url, "blocked_by_robots"))
                continue

            LOGGER.info("[%d/%d] %s", len(visited), max_pages, url)
            try:
                html_text, status = fetcher.fetch(url)
                soup = BeautifulSoup(html_text, "lxml")
                title = extract_page_title(soup)
                page_type = "plant" if is_probable_plant_page(soup) else "catalog/other"
                log.append(CrawlEntry(url, "ok", status, page_type, title))

                if page_type == "plant":
                    scraped_at = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
                    record, rejected_obs = build_plant_record(url, soup, scraped_at)
                    records.append(record)
                    rejected.extend(rejected_obs)

                for link in extract_internal_links(soup, url, root_url):
                    if link not in visited:
                        queue.append(link)
            except Exception as exc:  # noqa: BLE001
                LOGGER.exception("Failed %s", url)
                log.append(CrawlEntry(url, "error", error=str(exc)))
    finally:
        fetcher.close()

    return merge_records(records), rejected, log


# -----------------------------------------------------------------------------
# XLSX export — strict black/white style.
# -----------------------------------------------------------------------------


def convert_m_value(value_m: float, unit: str) -> float:
    factors = {"m": 1.0, "cm": 100.0, "mm": 1000.0}
    return value_m * factors[unit]


def display_observation(o: Observation, display_length_unit: str) -> str:
    if o.canonical_unit in {"m", "m/year"}:
        suffix = f"{display_length_unit}/year" if o.canonical_unit == "m/year" else display_length_unit
        if o.min_value is not None and o.max_value is not None:
            lo = convert_m_value(o.min_value, display_length_unit)
            hi = convert_m_value(o.max_value, display_length_unit)
            if abs(hi - lo) > 1e-12:
                return f"{format_num(lo)}–{format_num(hi)} {suffix}"
            return f"{format_num(lo)} {suffix}"
        if o.max_value is not None:
            return f"≤{format_num(convert_m_value(o.max_value, display_length_unit))} {suffix}"
        if o.min_value is not None:
            return f"≥{format_num(convert_m_value(o.min_value, display_length_unit))} {suffix}"
    return o.display_value()


def cell_values(record: PlantRecord, field_key: str, display_length_unit: str = "m") -> list[str]:
    if field_key == "id":
        return [record.plant_id]
    if field_key == "comment":
        comments = [o.display_value() for o in record.observations.get("comment", []) if o.accepted_to_main]
        if record.unclassified_images:
            comments.append(f"Неклассифицированных изображений: {len(record.unclassified_images)}")
        return unique_preserve(comments)
    if field_key == "confidence":
        obs = [o for values in record.observations.values() for o in values if o.accepted_to_main and o.field_key not in {"source_page_url", "source_image_url", "scrape_date", "comment"}]
        if not obs:
            return []
        avg = sum(o.confidence for o in obs) / len(obs)
        mn = min(o.confidence for o in obs)
        return [f"avg {avg:.0%}; min {mn:.0%}"]
    obs = [o for o in record.observations.get(field_key, []) if o.accepted_to_main]
    values = [display_observation(o, display_length_unit) for o in sorted(obs, key=lambda x: (-x.confidence, x.display_value()))]
    return unique_preserve(values)


def aggregate_cell(record: PlantRecord, field_key: str, display_length_unit: str = "m") -> str:
    values = cell_values(record, field_key, display_length_unit)
    return "\n".join(values)


def completeness(records: list[PlantRecord]) -> list[tuple[str, str, int, float]]:
    rows = []
    n = len(records)
    for rule in FIELD_RULES:
        if rule.key in {"id", "comment", "confidence", "source_page_url", "source_image_url", "scrape_date"}:
            continue
        count = sum(1 for r in records if aggregate_cell(r, rule.key))
        pct = count / n if n else 0.0
        rows.append((rule.key, rule.label, count, pct))
    return rows


def export_xlsx(
    output_path: str,
    records: list[PlantRecord],
    rejected: list[Observation],
    crawl_log: list[CrawlEntry],
    root_url: str,
    display_length_unit: str = "m",
) -> None:
    workbook = xlsxwriter.Workbook(output_path, {"constant_memory": False})
    workbook.set_properties({"title": "База растений", "comments": "Generated by plant_parser_production.py"})

    fmt_header = workbook.add_format({
        "bold": True, "font_color": "white", "bg_color": "black", "border": 1,
        "align": "center", "valign": "vcenter", "text_wrap": True,
    })
    fmt_body = workbook.add_format({"font_color": "black", "bg_color": "white", "border": 1, "valign": "top", "text_wrap": True})
    fmt_pct = workbook.add_format({"font_color": "black", "bg_color": "white", "border": 1, "num_format": "0.0%"})
    fmt_num = workbook.add_format({"font_color": "black", "bg_color": "white", "border": 1, "num_format": "0.####"})
    fmt_title = workbook.add_format({"bold": True, "font_color": "black", "bg_color": "white", "font_size": 14})

    # Plants.
    ws = workbook.add_worksheet("База растений")
    ws.freeze_panes(1, 2)
    ws.set_row(0, 48)
    ws.autofilter(0, 0, max(1, len(records)), len(FIELD_RULES) - 1)
    for col, rule in enumerate(FIELD_RULES):
        ws.write(0, col, rule.label, fmt_header)
        width = 14 if rule.key == "id" else 22
        if rule.key in {"short_description", "long_description", "comment"}:
            width = 40
        elif "photo" in rule.key or "source" in rule.key:
            width = 34
        ws.set_column(col, col, width)
    for row, record in enumerate(records, start=1):
        for col, rule in enumerate(FIELD_RULES):
            ws.write(row, col, aggregate_cell(record, rule.key, display_length_unit), fmt_body)

    # Raw evidence: accepted + rejected + unclassified images.
    ev = workbook.add_worksheet("Raw evidence")
    evidence_headers = [
        "Plant ID", "Accepted to main", "Field key", "Target column", "Raw label", "Raw value",
        "Normalized value", "Min", "Max", "Source unit", "Canonical unit", "Confidence",
        "Extraction method", "Evidence/context", "Source URL", "Scraped at", "Rejection reason",
    ]
    for c, h in enumerate(evidence_headers):
        ev.write(0, c, h, fmt_header)
    ev.freeze_panes(1, 0)
    ev.set_column(0, 0, 18)
    ev.set_column(1, 3, 18)
    ev.set_column(4, 6, 30)
    ev.set_column(7, 11, 14)
    ev.set_column(12, 13, 28)
    ev.set_column(14, 16, 38)
    row = 1
    all_accepted = [o for r in records for obs in r.observations.values() for o in obs]
    for o in all_accepted + rejected:
        values = [
            o.plant_id, "YES" if o.accepted_to_main else "NO", o.field_key, o.field_label, o.raw_label,
            o.raw_value, o.normalized_value, o.min_value, o.max_value, o.source_unit, o.canonical_unit,
            o.confidence, o.extraction_method, o.evidence_text, o.source_url, o.scraped_at, o.rejection_reason,
        ]
        for c, value in enumerate(values):
            if c in {7, 8} and isinstance(value, (int, float)):
                ev.write_number(row, c, value, fmt_num)
            elif c == 11 and isinstance(value, (int, float)):
                ev.write_number(row, c, value, fmt_pct)
            else:
                ev.write(row, c, value if value is not None else "", fmt_body)
        row += 1
    for r in records:
        for image_url in r.unclassified_images:
            vals = [r.plant_id, "NO", "image_unclassified", "Неклассифицированное изображение", "image", image_url, "", "", "", "URL", "URL", 0.0, "image", "No explicit seasonal context", r.page_url, "", "Image context insufficient"]
            for c, value in enumerate(vals):
                ev.write(row, c, value, fmt_body if c != 11 else fmt_pct)
            row += 1
    if row > 1:
        ev.autofilter(0, 0, row - 1, len(evidence_headers) - 1)

    # Field rules — every target column.
    fr = workbook.add_worksheet("Правила полей")
    fr_headers = ["Field key", "Колонка", "Тип", "Синонимы/смысловые маркеры", "Запрещающий контекст", "Обязательный контекст", "Canonical unit", "Min confidence", "Строгое правило"]
    for c, h in enumerate(fr_headers):
        fr.write(0, c, h, fmt_header)
    fr.freeze_panes(1, 0)
    fr.set_column(0, 2, 24)
    fr.set_column(3, 5, 50)
    fr.set_column(6, 7, 16)
    fr.set_column(8, 8, 75)
    for row_i, rule in enumerate(FIELD_RULES, start=1):
        vals = [
            rule.key, rule.label, rule.value_type, "; ".join(rule.aliases), "; ".join(rule.exclude_terms),
            "; ".join(rule.required_context_any), rule.canonical_unit, rule.min_confidence, rule.notes,
        ]
        for c, value in enumerate(vals):
            if c == 7:
                fr.write_number(row_i, c, float(value), fmt_pct)
            else:
                fr.write(row_i, c, value, fmt_body)
    fr.autofilter(0, 0, len(FIELD_RULES), len(fr_headers) - 1)

    # Crawl log.
    cl = workbook.add_worksheet("Crawl log")
    cl_headers = ["URL", "Status", "HTTP", "Page type", "Title", "Error"]
    for c, h in enumerate(cl_headers):
        cl.write(0, c, h, fmt_header)
    cl.set_column(0, 0, 50)
    cl.set_column(1, 3, 18)
    cl.set_column(4, 5, 45)
    for i, entry in enumerate(crawl_log, start=1):
        vals = [entry.url, entry.status, entry.http_status or "", entry.page_type, entry.title, entry.error]
        for c, value in enumerate(vals):
            cl.write(i, c, value, fmt_body)
    if crawl_log:
        cl.autofilter(0, 0, len(crawl_log), len(cl_headers) - 1)

    # Summary.
    sm = workbook.add_worksheet("Отчёт парсинга")
    sm.write("A1", "Отчёт парсинга", fmt_title)
    metrics = [
        ("URL сайта", root_url),
        ("Дата формирования", dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")),
        ("Просканировано URL", len(crawl_log)),
        ("Уникальных растений", len(records)),
        ("Ошибок обхода", sum(1 for x in crawl_log if x.status == "error")),
        ("Отклонённых/неоднозначных evidence", len(rejected)),
        ("Правило", "Пустое поле означает: подтверждённых данных на источнике не найдено."),
        ("Raw policy", "Каждое исходное значение хранится в Raw evidence; нормализация не удаляет raw."),
        ("Units policy", f"Длины canonical=m; основной лист отображает {display_length_unit}; объём контейнера=L; валюты не конвертируются; min/max сохраняются."),
    ]
    for r, (k, v) in enumerate(metrics, start=2):
        sm.write(r, 0, k, fmt_header)
        sm.write(r, 1, v, fmt_body)
    start = 13
    comp_headers = ["Field key", "Колонка", "Заполнено растений", "% заполнения"]
    for c, h in enumerate(comp_headers):
        sm.write(start, c, h, fmt_header)
    for i, (key, label, count, pct) in enumerate(completeness(records), start=start + 1):
        sm.write(i, 0, key, fmt_body)
        sm.write(i, 1, label, fmt_body)
        sm.write_number(i, 2, count, fmt_num)
        sm.write_number(i, 3, pct, fmt_pct)
    sm.set_column(0, 0, 28)
    sm.set_column(1, 1, 65)
    sm.set_column(2, 3, 18)

    workbook.close()


# -----------------------------------------------------------------------------
# Config overrides: aliases and crawl options can be extended without editing core.
# -----------------------------------------------------------------------------


def apply_rule_overrides(path: Optional[str]) -> None:
    if not path:
        return
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    for key, overrides in data.get("fields", {}).items():
        rule = RULE_BY_KEY.get(key)
        if not rule:
            raise ValueError(f"Unknown field key in overrides: {key}")
        if "add_aliases" in overrides:
            rule.aliases = tuple(unique_preserve(list(rule.aliases) + list(overrides["add_aliases"])))
        if "add_exclude_terms" in overrides:
            rule.exclude_terms = tuple(unique_preserve(list(rule.exclude_terms) + list(overrides["add_exclude_terms"])))
        if "min_confidence" in overrides:
            rule.min_confidence = float(overrides["min_confidence"])


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Strict semantic plant catalog parser -> XLSX")
    p.add_argument("--url", required=True, help="Catalog or site start URL")
    p.add_argument("--output", default="plants.xlsx", help="Output XLSX")
    p.add_argument("--max-pages", type=int, default=500, help="Maximum same-host pages to visit")
    p.add_argument("--browser", choices=["requests", "playwright", "auto"], default="auto")
    p.add_argument("--timeout", type=float, default=25.0)
    p.add_argument("--delay", type=float, default=0.5, help="Delay between requests, seconds")
    p.add_argument("--retries", type=int, default=2)
    p.add_argument("--include-regex", default="", help="Only crawl URLs matching regex (root URL is always allowed)")
    p.add_argument("--exclude-regex", default=r"/(cart|checkout|login|account|privacy|terms)(/|$)")
    p.add_argument("--ignore-robots", action="store_true", help="Use only when you have permission")
    p.add_argument("--rules-overrides", default="", help="JSON file with site-specific aliases/exclusions")
    p.add_argument("--display-length-unit", choices=["m", "cm", "mm"], default="m", help="Display unit for length cells in main sheet; canonical storage remains meters")
    p.add_argument("--log-level", choices=["DEBUG", "INFO", "WARNING", "ERROR"], default="INFO")
    return p


def main(argv: Optional[list[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)s %(message)s",
    )
    apply_rule_overrides(args.rules_overrides or None)

    records, rejected, crawl_log = crawl_site(
        root_url=args.url,
        max_pages=max(1, args.max_pages),
        browser=args.browser,
        timeout=max(1.0, args.timeout),
        delay=max(0.0, args.delay),
        retries=max(0, args.retries),
        include_regex=args.include_regex,
        exclude_regex=args.exclude_regex,
        ignore_robots=args.ignore_robots,
    )
    export_xlsx(args.output, records, rejected, crawl_log, args.url, args.display_length_unit)

    print(json.dumps({
        "output": str(Path(args.output).resolve()),
        "unique_plants": len(records),
        "crawled_urls": len(crawl_log),
        "rejected_evidence": len(rejected),
        "errors": sum(1 for x in crawl_log if x.status == "error"),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
