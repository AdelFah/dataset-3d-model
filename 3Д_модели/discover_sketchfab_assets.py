#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ищет 3D-модели на Sketchfab по списку автомобилей из датасета.
Для каждой машины делает поиск по названию и сохраняет подходящие варианты
в таблицу кандидатов для дальнейшего скачивания.
"""

from __future__ import annotations

import argparse
import csv
import re
import time
from pathlib import Path
from urllib.parse import urlparse

import requests


ROOT = Path(__file__).parent.resolve()
ASSET_CANDIDATES = ROOT / "cars_dataset" / "metadata" / "asset_candidates.csv"
SEARCH_URL = "https://api.sketchfab.com/v3/search"

CYRILLIC_ALIASES = {
    "веста": "vesta",
    "гранта": "granta",
    "калина": "kalina",
    "ларгус": "largus",
    "приора": "priora",
    "нива": "niva",
    "нива тревел": "niva travel",
    "2107": "2107",
    "2110": "2110",
    "2114": "2114",
}

MARK_ALIASES = {
    "lada": ["lada", "vaz", "ваз"],
    "volkswagen": ["volkswagen", "vw"],
    "mercedes-benz": ["mercedes-benz", "mercedes", "benz"],
    "chevrolet": ["chevrolet", "chevy"],
    "toyota": ["toyota"],
    "nissan": ["nissan", "datsun"],
    "honda": ["honda"],
    "mazda": ["mazda"],
    "mitsubishi": ["mitsubishi"],
    "hyundai": ["hyundai"],
    "kia": ["kia"],
    "renault": ["renault"],
    "peugeot": ["peugeot"],
    "citroen": ["citroen", "citroën"],
    "bmw": ["bmw"],
    "audi": ["audi"],
    "ford": ["ford"],
    "opel": ["opel"],
    "skoda": ["skoda", "škoda"],
    "geely": ["geely"],
    "chery": ["chery"],
    "lexus": ["lexus"],
    "infiniti": ["infiniti"],
    "jaguar": ["jaguar"],
    "suzuki": ["suzuki"],
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Find Sketchfab candidates for search_needed rows")
    parser.add_argument("--limit-combos", type=int, default=0, help="Limit unique mark/model searches; 0 = all")
    parser.add_argument(
        "--min-restyling",
        type=int,
        default=-1,
        help="Only search rows with restyling >= N (-1 = all search_needed)",
    )
    parser.add_argument("--count", type=int, default=24, help="Sketchfab results to inspect per query")
    parser.add_argument("--sleep", type=float, default=0.6, help="Delay between vehicle searches")
    parser.add_argument("--checkpoint-every", type=int, default=25, help="Save CSV every N processed combos")
    parser.add_argument("--min-score", type=int, default=4, help="Minimum match score (strict pass)")
    parser.add_argument(
        "--fallback-min-score",
        type=int,
        default=2,
        help="Relaxed score when strict pass finds nothing (model token only)",
    )
    return parser.parse_args()


def norm_text(value: str) -> str:
    value = value.lower().replace("ё", "е")
    value = re.sub(r"[^0-9a-zа-я]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def model_tokens(model: str) -> list[str]:
    normalized = norm_text(model)
    tokens = [t for t in normalized.split() if t not in {"class", "series"}]
    if not tokens and normalized:
        tokens = [normalized]
    return tokens


def aliases_for(mark: str, model: str) -> tuple[list[str], list[str]]:
    mark_n = norm_text(mark)
    mark_aliases = MARK_ALIASES.get(mark_n, [mark_n])
    model_aliases = [norm_text(model)]
    model_lower = norm_text(model)
    if model_lower in CYRILLIC_ALIASES:
        model_aliases.append(CYRILLIC_ALIASES[model_lower])
    return mark_aliases, [a for a in model_aliases if a]


def search_queries(
    mark: str,
    model: str,
    generation: str = "",
    restyling: str = "",
    year_from: str = "",
    year_to: str = "",
) -> list[str]:
    mark_aliases, model_aliases = aliases_for(mark, model)
    queries: list[str] = []
    for model_alias in model_aliases:
        for mark_alias in mark_aliases[:3]:
            version_bits = []
            if generation:
                version_bits.append(f"generation {generation}")
            if restyling not in {"",}:
                version_bits.append(f"restyling {restyling}")
                if restyling == "0":
                    version_bits.append("pre facelift")
                elif restyling == "1":
                    version_bits.append("facelift")
            if year_from or year_to:
                version_bits.append(f"{year_from}-{year_to}".strip("-"))
            version = " ".join(version_bits)
            if version:
                queries.append(f"{mark_alias} {model_alias} {version} car")
                queries.append(f"{mark_alias} {model_alias} {year_from} car".strip())
            queries.append(f"{mark_alias} {model_alias} car")
            queries.append(f"{mark_alias} {model_alias}")
        queries.append(f"{model_alias} car")
        queries.append(model_alias)
        if model_alias == "niva":
            queries.extend(["lada niva car", "lada niva", "chevy niva", "chevrolet niva"])
        if model_alias == "niva travel":
            queries.extend(["lada niva travel", "niva travel", "chevrolet niva travel"])
        if model_alias == "largus":
            queries.extend(["lada largus", "renault largus", "dacia lodgy"])
    deduped: list[str] = []
    for query in queries:
        if query not in deduped:
            deduped.append(query)
    return deduped


def sketchfab_uid(url: str) -> str | None:
    if "sketchfab.com" not in url:
        return None
    path = urlparse(url).path.strip("/")
    parts = path.split("/")
    if len(parts) < 2 or parts[0] != "3d-models":
        return None
    match = re.search(r"([0-9a-f]{32})$", parts[1])
    return match.group(1) if match else None


def result_url(result: dict) -> str:
    url = result.get("viewerUrl") or ""
    uid = result.get("uid") or ""
    if url:
        return url
    return f"https://sketchfab.com/3d-models/model-{uid}" if uid else ""


AUTOMOTIVE_HINTS = {
    "car", "auto", "automobile", "vehicle", "sedan", "hatchback", "wagon", "suv",
    "crossover", "coupe", "pickup", "truck", "van", "minivan", "roadster",
    "авто", "машина", "седан", "внедорожник", "универсал", "хэтчбек",
}


def is_automotive_name(name: str) -> bool:
    hay = norm_text(name)
    return any(hint in hay for hint in AUTOMOTIVE_HINTS)


def score_result(
    mark: str,
    model: str,
    result: dict,
    restyling: str = "",
    year_from: str = "",
    year_to: str = "",
) -> int:
    url = result_url(result)
    name = result.get("name", "")
    hay = norm_text(f"{name} {url}")
    mark_aliases, model_aliases = aliases_for(mark, model)
    tokens = model_tokens(model)

    score = 0
    mark_hit = any(alias and alias in hay for alias in mark_aliases)
    model_hit = any(alias and alias in hay for alias in model_aliases)
    if mark_hit:
        score += 2
    if model_hit:
        score += 4

    token_hits = sum(1 for token in tokens if token in hay)
    score += token_hits * 2
    if tokens and token_hits < max(1, min(len(tokens), 2)):
        score -= 3

    if is_automotive_name(name):
        score += 2
    elif not mark_hit:
        score -= 4

    # Restyling / year hints in model title (facelift, model year).
    rst = str(restyling or "").strip()
    if rst == "1" and any(x in hay for x in ("facelift", "restyling", "updated", "new")):
        score += 2
    if rst == "0" and any(x in hay for x in ("pre facelift", "prefacelift", "old")):
        score += 1
    for year in (year_from, year_to):
        if year and str(year) in hay:
            score += 2

    # Penalize obvious non-car matches on fallback.
    bad_hints = {
        "keyboard", "sculpture", "tamagotchi", "crossbow", "game", "cat", "box",
        "matrix", "cards", "spades", "infantry", "carrier", "tank", "aircraft",
    }
    if any(bad in hay for bad in bad_hints):
        score -= 6

    return score


def quality_from_archives(result: dict) -> tuple[str, str]:
    archives = result.get("archives") or {}
    face_count = 0
    for item in archives.values():
        if isinstance(item, dict) and isinstance(item.get("faceCount"), int):
            face_count = max(face_count, item["faceCount"])
    if face_count >= 250_000:
        quality = "high"
    elif face_count >= 25_000:
        quality = "medium"
    elif face_count > 0:
        quality = "low"
    else:
        quality = "unknown"
    triangles = f"{face_count:,}".replace(",", " ") if face_count else ""
    return quality, triangles


def sketchfab_search(params: dict, max_retries: int = 6) -> dict:
    for attempt in range(max_retries):
        resp = requests.get(SEARCH_URL, params=params, timeout=45)
        if resp.status_code == 429:
            wait_s = min(60, 2 ** attempt)
            time.sleep(wait_s)
            continue
        resp.raise_for_status()
        return resp.json()
    resp.raise_for_status()
    return {}


def apply_results(rows: list[dict], combo_to_result: dict[str, dict]) -> int:
    updated = 0
    for row in rows:
        if row.get("candidate_status") != "search_needed":
            continue
        result = combo_to_result.get(row.get("vehicle_id", ""))
        if not result:
            continue
        url = result_url(result)
        if not url:
            continue
        license_info = result.get("license") or {}
        quality, triangles = quality_from_archives(result)
        row["candidate_status"] = "found_candidate"
        row["candidate_source"] = "Sketchfab"
        row["candidate_url"] = url
        row["license_status"] = license_info.get("label", "check")
        row["asset_quality"] = quality
        row["triangles"] = triangles
        row["local_asset_path"] = ""
        row["scale_check_status"] = "not_imported"
        row["note"] = f"Auto-discovered via Sketchfab search: {result.get('name', '')}".strip()
        updated += 1
    return updated


def save_rows(fields: list[str], rows: list[dict]) -> None:
    with ASSET_CANDIDATES.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def search_combo(
    mark: str,
    model: str,
    generation: str,
    restyling: str,
    year_from: str,
    year_to: str,
    rejected_uids: set[str],
    used_uids: set[str],
    count: int,
    min_score: int,
    fallback_min_score: int,
) -> dict | None:
    best: tuple[int, dict] | None = None
    fallback: tuple[int, dict] | None = None
    seen_uids: set[str] = set()
    mark_aliases, model_aliases = aliases_for(mark, model)
    for query in search_queries(mark, model, generation, restyling, year_from, year_to):
        payload = sketchfab_search(
            {
                "type": "models",
                "q": query,
                "downloadable": "true",
                "sort_by": "-likeCount",
                "count": count,
            }
        )
        for result in payload.get("results", []):
            uid = result.get("uid")
            if not uid or uid in seen_uids:
                continue
            if uid in rejected_uids or uid in used_uids:
                continue
            seen_uids.add(uid)
            score = score_result(mark, model, result, restyling, year_from, year_to)
            if score >= min_score and (best is None or score > best[0]):
                best = (score, result)
            name = norm_text(result.get("name", ""))
            mark_ok = any(alias and alias in name for alias in mark_aliases)
            model_ok = any(alias and alias in name for alias in model_aliases)
            auto_ok = is_automotive_name(result.get("name", ""))
            if (
                score >= fallback_min_score
                and mark_ok
                and model_ok
                and auto_ok
                and (fallback is None or score > fallback[0])
            ):
                fallback = (score, result)
    if best:
        return best[1]
    return fallback[1] if fallback else None


def main() -> None:
    args = parse_args()
    if not ASSET_CANDIDATES.exists():
        raise FileNotFoundError(ASSET_CANDIDATES)

    rows = list(csv.DictReader(ASSET_CANDIDATES.open(encoding="utf-8-sig")))
    if not rows:
        return
    fields = list(rows[0].keys())
    combos: list[tuple[str, str, str, str, str, str, str]] = []
    for row in rows:
        if row.get("candidate_status") != "search_needed":
            continue
        if args.min_restyling >= 0:
            try:
                if int(str(row.get("restyling", "0") or "0")) < args.min_restyling:
                    continue
            except ValueError:
                continue
        combo = (
            row.get("vehicle_id", ""),
            row.get("mark", ""),
            row.get("model", ""),
            row.get("generation", ""),
            row.get("restyling", ""),
            row.get("year_from", ""),
            row.get("year_to", ""),
        )
        if combo not in combos:
            combos.append(combo)
    if args.limit_combos > 0:
        combos = combos[: args.limit_combos]

    rejected_by_vehicle: dict[str, set[str]] = {}
    for row in rows:
        uid = sketchfab_uid(row.get("rejected_candidate_url", ""))
        if uid:
            rejected_by_vehicle.setdefault(row.get("vehicle_id", ""), set()).add(uid)

    used_uids_by_model: dict[tuple[str, str], set[str]] = {}
    for row in rows:
        if row.get("candidate_status") != "found_candidate":
            continue
        uid = sketchfab_uid(row.get("candidate_url", ""))
        if uid:
            used_uids_by_model.setdefault((row.get("mark", ""), row.get("model", "")), set()).add(uid)

    combo_to_result: dict[str, dict] = {}
    print(f"Searching Sketchfab for {len(combos)} unique vehicle/version combos")
    for i, (vehicle_id, mark, model, generation, restyling, year_from, year_to) in enumerate(combos, 1):
        try:
            result = search_combo(
                mark,
                model,
                generation,
                restyling,
                year_from,
                year_to,
                rejected_by_vehicle.get(vehicle_id, set()),
                used_uids_by_model.setdefault((mark, model), set()),
                args.count,
                args.min_score,
                args.fallback_min_score,
            )
            if result:
                uid = result.get("uid")
                if uid:
                    used_uids_by_model[(mark, model)].add(uid)
                combo_to_result[vehicle_id] = result
                print(f"[{i}/{len(combos)}] FOUND {mark} {model} g{generation} r{restyling}: {result.get('name')}")
            else:
                print(f"[{i}/{len(combos)}] miss {mark} {model} g{generation} r{restyling}")
        except Exception as exc:
            print(f"[{i}/{len(combos)}] error {mark} {model} g{generation} r{restyling}: {exc}")
        if args.checkpoint_every > 0 and i % args.checkpoint_every == 0:
            updated_now = apply_results(rows, combo_to_result)
            save_rows(fields, rows)
            print(f"Checkpoint saved after {i}/{len(combos)} combos ({updated_now} rows updated so far)")
        time.sleep(args.sleep)

    updated = apply_results(rows, combo_to_result)
    save_rows(fields, rows)

    print(f"Updated rows: {updated}")
    print(f"Saved: {ASSET_CANDIDATES}")


if __name__ == "__main__":
    main()
