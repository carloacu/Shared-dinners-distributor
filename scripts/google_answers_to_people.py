#!/usr/bin/env python3
"""Convert a Google Forms export into the people CSV format used by this repo."""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import unicodedata
import urllib.error
import urllib.request
from collections import defaultdict
from pathlib import Path


PEOPLE_HEADER = [
    "ID",
    "name",
    "gender",
    "year_of_birth",
    "postal_address",
    "postal_code",
    "city",
    "recieving_for_drinks",
    "number_max_recieving_for_drinks",
    "recieving_for_dinner",
    "number_max_recieving_for_dinner",
    "can_host_pmr",
    "can_host_both_events",
]

OPENAI_RESPONSES_URL = "https://api.openai.com/v1/responses"
DEFAULT_OPENAI_MODEL = "gpt-5-mini"


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return " ".join(value.split())


def collapse_spaces(value: str) -> str:
    return " ".join((value or "").strip().split())


def normalize_lookup_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize_text(value))


def pick_header(fieldnames: list[str], needle: str) -> str:
    normalized_needle = normalize_text(needle)
    exact_matches = [name for name in fieldnames if normalize_text(name) == normalized_needle]
    if len(exact_matches) == 1:
        return exact_matches[0]

    prefix_matches = [name for name in fieldnames if normalize_text(name).startswith(normalized_needle)]
    if len(prefix_matches) == 1:
        return prefix_matches[0]

    matches = [name for name in fieldnames if normalized_needle in normalize_text(name)]
    if len(matches) != 1:
        raise KeyError(f"Could not uniquely resolve column for: {needle!r}. Matches: {matches}")
    return matches[0]


def parse_year(value: str) -> int:
    cleaned = collapse_spaces(value)
    if not cleaned:
        return 0
    direct = re.fullmatch(r"\d{4}", cleaned)
    if direct:
        return int(cleaned)
    match = re.search(r"(19|20)\d{2}", cleaned)
    if match:
        return int(match.group(0))
    return 0


def parse_yes_no(value: str) -> str:
    normalized = normalize_text(value)
    if normalized.startswith("oui") or normalized in {"yes", "true", "1"}:
        return "yes"
    return "no"


def parse_capacity(value: str) -> int:
    match = re.search(r"\d+", value or "")
    return int(match.group(0)) if match else 0


def clean_postal_code(value: str) -> str:
    cleaned = collapse_spaces(value)
    match = re.search(r"\b\d{5}\b", cleaned)
    if match:
        return match.group(0)
    return cleaned


def load_gender_lookup(people_dir: Path, ignored_output: Path | None) -> tuple[dict[str, str], dict[str, str]]:
    by_full_name: dict[str, set[str]] = defaultdict(set)
    by_first_name: dict[str, set[str]] = defaultdict(set)

    for path in sorted(people_dir.glob("*.csv")):
        if ignored_output is not None and path.resolve() == ignored_output.resolve():
            continue
        with path.open(newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                gender = collapse_spaces(row.get("gender", "")).upper()
                name = collapse_spaces(row.get("name", ""))
                if gender not in {"M", "F"} or not name:
                    continue
                full_key = normalize_lookup_key(name)
                by_full_name[full_key].add(gender)

                first_name = collapse_spaces(name.split()[0])
                if first_name:
                    by_first_name[normalize_lookup_key(first_name)].add(gender)

    resolved_full = {key: next(iter(values)) for key, values in by_full_name.items() if len(values) == 1}
    resolved_first = {key: next(iter(values)) for key, values in by_first_name.items() if len(values) == 1}
    return resolved_full, resolved_first


def load_gender_cache(cache_path: Path) -> tuple[dict[str, str], dict[str, str], set[str], set[str]]:
    if not cache_path.exists():
        return {}, {}, set(), set()

    with cache_path.open(encoding="utf-8") as handle:
        payload = json.load(handle)

    full_name_cache = {
        str(key): str(value).upper()
        for key, value in (payload.get("full_name") or {}).items()
        if str(value).upper() in {"M", "F", "UNKNOWN"}
    }
    first_name_cache = {
        str(key): str(value).upper()
        for key, value in (payload.get("first_name") or {}).items()
        if str(value).upper() in {"M", "F", "UNKNOWN"}
    }
    negative_full = {key for key, value in full_name_cache.items() if value == "UNKNOWN"}
    negative_first = {key for key, value in first_name_cache.items() if value == "UNKNOWN"}
    resolved_full = {key: value for key, value in full_name_cache.items() if value in {"M", "F"}}
    resolved_first = {key: value for key, value in first_name_cache.items() if value in {"M", "F"}}
    return resolved_full, resolved_first, negative_full, negative_first


class GenderResolver:
    def __init__(self, people_dir: Path, ignored_output: Path | None, cache_path: Path) -> None:
        csv_full_lookup, csv_first_lookup = load_gender_lookup(people_dir, ignored_output)
        cache_full_lookup, cache_first_lookup, negative_full, negative_first = load_gender_cache(cache_path)

        self.cache_path = cache_path
        self.cache_payload = {
            "full_name": {**cache_full_lookup},
            "first_name": {**cache_first_lookup},
        }
        self.full_name_lookup = {**cache_full_lookup, **csv_full_lookup}
        self.first_name_lookup = {**cache_first_lookup, **csv_first_lookup}
        self.negative_full = set(negative_full)
        self.negative_first = set(negative_first)
        self.cache_dirty = False
        self.api_key = os.getenv("OPENAI_API_KEY", "").strip()
        self.model = (
            os.getenv("OPENAI_GENDER_MODEL", "").strip()
            or os.getenv("OPENAI_MODEL", "").strip()
            or DEFAULT_OPENAI_MODEL
        )

    def infer_gender(self, name: str) -> str:
        full_key = normalize_lookup_key(name)
        if full_key in self.full_name_lookup:
            return self.full_name_lookup[full_key]
        if full_key in self.negative_full and not self.api_key:
            return ""

        first_name = collapse_spaces(name.split()[0] if name else "")
        first_key = normalize_lookup_key(first_name)
        if first_key in self.first_name_lookup:
            return self.first_name_lookup[first_key]
        if first_key in self.negative_first and not self.api_key:
            return ""

        inferred = self.fetch_gender_from_openai(name, first_name)
        self.update_cache(full_key, first_key, inferred)
        return inferred if inferred in {"M", "F"} else ""

    def update_cache(self, full_key: str, first_key: str, gender: str) -> None:
        cached_value = gender if gender in {"M", "F"} else "UNKNOWN"
        self.cache_payload["full_name"][full_key] = cached_value
        if first_key:
            self.cache_payload["first_name"][first_key] = cached_value
        self.cache_dirty = True

        if cached_value in {"M", "F"}:
            self.full_name_lookup[full_key] = cached_value
            if first_key:
                self.first_name_lookup[first_key] = cached_value
            self.negative_full.discard(full_key)
            self.negative_first.discard(first_key)
        else:
            self.negative_full.add(full_key)
            if first_key:
                self.negative_first.add(first_key)

    def fetch_gender_from_openai(self, full_name: str, first_name: str) -> str:
        if not self.api_key:
            raise RuntimeError(
                "Gender cache miss for "
                f"{full_name!r} and OPENAI_API_KEY is not set. "
                "Set OPENAI_API_KEY or pre-populate data/cache/gender_inference_cache.json."
            )

        prompt = (
            "Infer the most likely binary gender marker for this participant name. "
            "Return exactly one token: M or F. "
            "Use common real-world first-name usage across likely languages and cultures as a best-effort guess. "
            "Even if uncertain, choose the more likely of M or F. Do not return UNKNOWN.\n\n"
            f"Full name: {full_name}\n"
            f"First name: {first_name}"
        )
        body = {
            "model": self.model,
            "input": [
                {
                    "role": "system",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "You classify likely gender markers for CSV cleanup. "
                                "Return a best-effort guess from the name."
                            ),
                        }
                    ],
                },
                {
                    "role": "user",
                    "content": [{"type": "input_text", "text": prompt}],
                },
            ],
            "text": {
                "format": {
                    "type": "json_schema",
                    "name": "gender_guess",
                    "strict": True,
                    "schema": {
                        "type": "object",
                        "properties": {
                            "gender": {
                                "type": "string",
                                "enum": ["M", "F"],
                            }
                        },
                        "required": ["gender"],
                        "additionalProperties": False,
                    },
                }
            },
            "max_output_tokens": 128,
        }
        reasoning = reasoning_options_for_model(self.model)
        if reasoning is not None:
            body["reasoning"] = reasoning
        payload = self.call_openai(body, full_name)

        gender = extract_gender_from_response(payload)
        if gender in {"M", "F"}:
            return gender

        fallback_body = {
            "model": self.model,
            "input": (
                "Classify the likely gender marker for this name. "
                "Reply with exactly one token: M or F.\n\n"
                f"Full name: {full_name}\n"
                f"First name: {first_name}"
            ),
            "max_output_tokens": 64,
        }
        if reasoning is not None:
            fallback_body["reasoning"] = reasoning
        fallback_payload = self.call_openai(fallback_body, full_name)
        fallback_gender = normalize_gender_label(extract_response_text(fallback_payload))
        if fallback_gender in {"M", "F"}:
            return fallback_gender

        raise RuntimeError(
            f"OpenAI gender inference returned an unusable value for {full_name!r}: "
            f"structured={extract_response_text(payload)!r}, "
            f"fallback={extract_response_text(fallback_payload)!r}"
        )

    def call_openai(self, body: dict, full_name: str) -> dict:
        payload = self._perform_openai_request(body, full_name)
        if response_has_usable_text(payload):
            return payload

        if payload.get("status") == "incomplete" or not extract_response_text(payload):
            retry_body = dict(body)
            retry_body["max_output_tokens"] = max(int(body.get("max_output_tokens", 64)) * 4, 256)
            payload = self._perform_openai_request(retry_body, full_name)
        return payload

    def _perform_openai_request(self, body: dict, full_name: str) -> dict:
        request = urllib.request.Request(
            OPENAI_RESPONSES_URL,
            data=json.dumps(body).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )

        try:
            with urllib.request.urlopen(request) as response:
                return json.load(response)
        except urllib.error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"OpenAI gender inference failed for {full_name!r}: {details}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"OpenAI gender inference network error for {full_name!r}: {exc}") from exc

    def save_cache(self) -> None:
        if not self.cache_dirty:
            return
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "full_name": dict(sorted(self.cache_payload["full_name"].items())),
            "first_name": dict(sorted(self.cache_payload["first_name"].items())),
        }
        with self.cache_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2, sort_keys=True)
            handle.write("\n")


def extract_response_text(payload: dict) -> str:
    output_text = payload.get("output_text")
    if isinstance(output_text, str) and output_text.strip():
        return output_text.strip()

    texts: list[str] = []
    for item in payload.get("output", []):
        if item.get("type") != "message":
            continue
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                texts.append(str(content.get("text", "")))
            elif content.get("type") == "text":
                text_value = content.get("text", "")
                if isinstance(text_value, dict):
                    texts.append(str(text_value.get("value", "")))
                else:
                    texts.append(str(text_value))
            elif content.get("type") == "refusal":
                texts.append(str(content.get("refusal", "")))
    return "\n".join(texts).strip()


def normalize_gender_label(value: str) -> str:
    normalized = normalize_text(value).replace(" ", "")
    if normalized in {"m", "male", "man", "homme"}:
        return "M"
    if normalized in {"f", "female", "woman", "femme"}:
        return "F"
    return "UNKNOWN"


def extract_gender_from_response(payload: dict) -> str:
    raw_text = extract_response_text(payload)
    if raw_text:
        try:
            parsed = json.loads(raw_text)
            if isinstance(parsed, dict):
                gender = normalize_gender_label(str(parsed.get("gender", "")))
                if gender in {"M", "F"}:
                    return gender
        except json.JSONDecodeError:
            gender = normalize_gender_label(raw_text)
            if gender in {"M", "F"}:
                return gender

    for item in payload.get("output", []):
        if item.get("type") != "message":
            continue
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                text = str(content.get("text", "")).strip()
                if not text:
                    continue
                try:
                    parsed = json.loads(text)
                    if isinstance(parsed, dict):
                        gender = normalize_gender_label(str(parsed.get("gender", "")))
                        if gender in {"M", "F"}:
                            return gender
                except json.JSONDecodeError:
                    gender = normalize_gender_label(text)
                    if gender in {"M", "F"}:
                        return gender

    return "UNKNOWN"


def response_has_usable_text(payload: dict) -> bool:
    if extract_response_text(payload):
        return True
    return extract_gender_from_response(payload) in {"M", "F"}


def reasoning_options_for_model(model: str) -> dict | None:
    normalized = (model or "").strip().lower()
    if normalized.startswith("gpt-5") or normalized.startswith("o"):
        return {"effort": "minimal"}
    return None


def same_household(primary: dict[str, str | int], secondary: dict[str, str | int]) -> bool:
    return all(
        normalize_text(str(primary[field])) == normalize_text(str(secondary[field])) and normalize_text(str(primary[field]))
        for field in ("postal_address", "postal_code", "city")
    )


def build_person(
    *,
    group_id: int,
    first_name: str,
    last_name: str,
    year: str,
    postal_address: str,
    postal_code: str,
    city: str,
    drinks: str,
    max_drinks: int,
    dinner: str,
    max_dinner: int,
    can_host_pmr: str,
    can_host_both_events: str,
    gender_resolver: GenderResolver,
) -> dict[str, str | int]:
    name = collapse_spaces(f"{collapse_spaces(first_name)} {collapse_spaces(last_name)}")
    return {
        "ID": group_id,
        "name": name,
        "gender": gender_resolver.infer_gender(name),
        "year_of_birth": parse_year(year),
        "postal_address": collapse_spaces(postal_address),
        "postal_code": clean_postal_code(postal_code),
        "city": collapse_spaces(city),
        "recieving_for_drinks": drinks,
        "number_max_recieving_for_drinks": max_drinks,
        "recieving_for_dinner": dinner,
        "number_max_recieving_for_dinner": max_dinner,
        "can_host_pmr": can_host_pmr,
        "can_host_both_events": can_host_both_events,
    }


def convert(input_path: Path, output_path: Path) -> int:
    repo_root = Path(__file__).resolve().parents[1]
    people_dir = repo_root / "data" / "input" / "people"
    gender_cache_path = repo_root / "data" / "cache" / "gender_inference_cache.json"
    gender_resolver = GenderResolver(people_dir, output_path, gender_cache_path)

    with input_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        fieldnames = reader.fieldnames or []
        columns = {
            "primary_last_name": pick_header(fieldnames, "Ton nom"),
            "primary_first_name": pick_header(fieldnames, "Ton prénom"),
            "primary_year": pick_header(fieldnames, "Ton année de naissance"),
            "primary_address": pick_header(fieldnames, "Ton adresse postale"),
            "primary_postal_code": pick_header(fieldnames, "Ton code postal"),
            "primary_city": pick_header(fieldnames, "Ta ville"),
            "secondary_last_name": pick_header(fieldnames, "Nom du participant 2"),
            "secondary_first_name": pick_header(fieldnames, "Prénom du participant 2"),
            "secondary_year": pick_header(fieldnames, "Année de naissance du participant 2"),
            "secondary_address": pick_header(fieldnames, "Adresse postale du participant 2"),
            "secondary_postal_code": pick_header(fieldnames, "Code postale du participant 2"),
            "secondary_city": pick_header(fieldnames, "Ville du participant 2"),
            "same_place": pick_header(
                fieldnames,
                "Souhaites-tu aller au même lieu pour l’apéro et le dîner que le participant 2",
            ),
            "apero_accept": pick_header(fieldnames, "recevoir (au moins 5 personnes) pour l'apéritif"),
            "apero_count": pick_header(fieldnames, "combien de personnes peux-tu recevoir pour l'apéritif"),
            "dinner_accept": pick_header(fieldnames, "recevoir (au moins 5 personnes) pour le dîner"),
            "dinner_count": pick_header(fieldnames, "combien de personnes peux-tu recevoir pour le dîner"),
            "pmr": pick_header(fieldnames, "accès PMR"),
        }

        rows: list[dict[str, str | int]] = []
        next_group_id = 1

        for raw_row in reader:
            primary_name = collapse_spaces(
                f"{raw_row[columns['primary_first_name']]} {raw_row[columns['primary_last_name']]}"
            )
            if not primary_name:
                continue

            drinks = parse_yes_no(raw_row[columns["apero_accept"]])
            dinner = parse_yes_no(raw_row[columns["dinner_accept"]])
            max_drinks = parse_capacity(raw_row[columns["apero_count"]]) if drinks == "yes" else 0
            max_dinner = parse_capacity(raw_row[columns["dinner_count"]]) if dinner == "yes" else 0
            can_host_pmr = parse_yes_no(raw_row[columns["pmr"]]) if drinks == "yes" or dinner == "yes" else "no"

            primary_group_id = next_group_id
            next_group_id += 1

            primary = build_person(
                group_id=primary_group_id,
                first_name=raw_row[columns["primary_first_name"]],
                last_name=raw_row[columns["primary_last_name"]],
                year=raw_row[columns["primary_year"]],
                postal_address=raw_row[columns["primary_address"]],
                postal_code=raw_row[columns["primary_postal_code"]],
                city=raw_row[columns["primary_city"]],
                drinks=drinks,
                max_drinks=max_drinks,
                dinner=dinner,
                max_dinner=max_dinner,
                can_host_pmr=can_host_pmr,
                can_host_both_events="no",
                gender_resolver=gender_resolver,
            )
            rows.append(primary)

            secondary_name = collapse_spaces(
                f"{raw_row[columns['secondary_first_name']]} {raw_row[columns['secondary_last_name']]}"
            )
            if not secondary_name:
                continue

            same_place_answer = normalize_text(raw_row[columns["same_place"]])
            same_group = same_place_answer.startswith("oui")
            secondary_group_id = primary_group_id if same_group else next_group_id
            if not same_group:
                next_group_id += 1

            secondary = build_person(
                group_id=secondary_group_id,
                first_name=raw_row[columns["secondary_first_name"]],
                last_name=raw_row[columns["secondary_last_name"]],
                year=raw_row[columns["secondary_year"]],
                postal_address=raw_row[columns["secondary_address"]],
                postal_code=raw_row[columns["secondary_postal_code"]],
                city=raw_row[columns["secondary_city"]],
                drinks="no",
                max_drinks=0,
                dinner="no",
                max_dinner=0,
                can_host_pmr="no",
                can_host_both_events="no",
                gender_resolver=gender_resolver,
            )

            if same_household(primary, secondary):
                secondary["recieving_for_drinks"] = drinks
                secondary["number_max_recieving_for_drinks"] = max_drinks
                secondary["recieving_for_dinner"] = dinner
                secondary["number_max_recieving_for_dinner"] = max_dinner
                secondary["can_host_pmr"] = can_host_pmr
                secondary["can_host_both_events"] = "no"

            rows.append(secondary)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=PEOPLE_HEADER, lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    gender_resolver.save_cache()
    return len(rows)


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_csv", help="Path to the Google Forms CSV export.")
    parser.add_argument(
        "-o",
        "--output",
        help="Output people CSV path. Defaults to data/input/people/<input filename>.",
    )
    args = parser.parse_args()

    input_path = Path(args.input_csv)
    if not input_path.is_absolute():
        input_path = repo_root / input_path
    input_path = input_path.resolve()

    if args.output:
        output_path = Path(args.output)
        if not output_path.is_absolute():
            output_path = repo_root / output_path
    else:
        output_path = repo_root / "data" / "input" / "people" / input_path.name
    output_path = output_path.resolve()

    row_count = convert(input_path, output_path)
    print(f"Wrote {row_count} rows to {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
