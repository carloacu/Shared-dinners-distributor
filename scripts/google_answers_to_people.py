#!/usr/bin/env python3
"""Convert a Google Forms export into the people CSV format used by this repo."""

from __future__ import annotations

import argparse
import csv
import re
import unicodedata
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
]


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


def infer_gender(name: str, full_name_lookup: dict[str, str], first_name_lookup: dict[str, str]) -> str:
    full_key = normalize_lookup_key(name)
    if full_key in full_name_lookup:
        return full_name_lookup[full_key]
    first_name = collapse_spaces(name.split()[0] if name else "")
    if first_name:
        first_key = normalize_lookup_key(first_name)
        if first_key in first_name_lookup:
            return first_name_lookup[first_key]
    return ""


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
    full_name_lookup: dict[str, str],
    first_name_lookup: dict[str, str],
) -> dict[str, str | int]:
    name = collapse_spaces(f"{collapse_spaces(first_name)} {collapse_spaces(last_name)}")
    return {
        "ID": group_id,
        "name": name,
        "gender": infer_gender(name, full_name_lookup, first_name_lookup),
        "year_of_birth": parse_year(year),
        "postal_address": collapse_spaces(postal_address),
        "postal_code": clean_postal_code(postal_code),
        "city": collapse_spaces(city),
        "recieving_for_drinks": drinks,
        "number_max_recieving_for_drinks": max_drinks,
        "recieving_for_dinner": dinner,
        "number_max_recieving_for_dinner": max_dinner,
        "can_host_pmr": can_host_pmr,
    }


def convert(input_path: Path, output_path: Path) -> int:
    repo_root = Path(__file__).resolve().parents[1]
    people_dir = repo_root / "data" / "input" / "people"
    full_name_lookup, first_name_lookup = load_gender_lookup(people_dir, output_path)

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
                full_name_lookup=full_name_lookup,
                first_name_lookup=first_name_lookup,
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
                full_name_lookup=full_name_lookup,
                first_name_lookup=first_name_lookup,
            )

            if same_household(primary, secondary):
                secondary["recieving_for_drinks"] = drinks
                secondary["number_max_recieving_for_drinks"] = max_drinks
                secondary["recieving_for_dinner"] = dinner
                secondary["number_max_recieving_for_dinner"] = max_dinner
                secondary["can_host_pmr"] = can_host_pmr

            rows.append(secondary)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=PEOPLE_HEADER, lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

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
