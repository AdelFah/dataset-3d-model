#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Выбирает лучшую 3D-модель для каждого автомобиля из нескольких кандидатов.
Если для одной машины нашлось несколько моделей на Sketchfab,
скрипт выбирает наиболее подходящую по качеству и детализации.

PowerShell:
    python select_best_external_assets.py
"""

from __future__ import annotations

import csv
import os
import re
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
DATASET_ROOT = Path(os.environ.get("CARS_DATASET_ROOT", ROOT / "cars_dataset")).resolve()
ASSET_CANDIDATES = DATASET_ROOT / "metadata" / "asset_candidates.csv"
MODELS_INDEX = DATASET_ROOT / "metadata" / "models_index.csv"
FITTED_DIR = DATASET_ROOT / "external_fitted"
OUT_CSV = DATASET_ROOT / "metadata" / "best_external_models.csv"

QUALITY_SCORE = {"high": 300_000, "medium": 100_000, "low": 10_000, "unknown": 1_000}
FITTED_PATTERN = re.compile(r"^(.+)__([0-9a-f]{8})\.blend$", re.IGNORECASE)
UID_PATTERN = re.compile(r"([0-9a-f]{32})")


def parse_triangles(value: str) -> int:
    text = str(value or "").strip().lower().replace(" ", "").replace(",", "")
    if not text:
        return 0
    multiplier = 1
    if text.endswith("k"):
        multiplier = 1_000
        text = text[:-1]
    elif text.endswith("m"):
        multiplier = 1_000_000
        text = text[:-1]
    try:
        return int(float(text) * multiplier)
    except ValueError:
        return 0


def sketchfab_uid(url: str) -> str | None:
    match = UID_PATTERN.search(url or "")
    return match.group(1) if match else None


def load_candidate_index() -> dict[tuple[str, str], dict]:
    """Map (vehicle_id, uid_prefix8) -> candidate row."""
    index: dict[tuple[str, str], dict] = {}
    if not ASSET_CANDIDATES.exists():
        return index
    for row in csv.DictReader(ASSET_CANDIDATES.open(encoding="utf-8-sig")):
        if row.get("candidate_status") != "found_candidate":
            continue
        uid = sketchfab_uid(row.get("candidate_url", ""))
        if not uid:
            continue
        key = (row["vehicle_id"], uid[:8].lower())
        index[key] = row
    return index


def collect_fitted_files() -> dict[str, list[tuple[str, Path]]]:
    by_vehicle: dict[str, list[tuple[str, Path]]] = {}
    if not FITTED_DIR.exists():
        return by_vehicle
    for blend in sorted(FITTED_DIR.glob("*.blend")):
        match = FITTED_PATTERN.match(blend.name)
        if not match:
            continue
        vehicle_id, uid8 = match.group(1), match.group(2).lower()
        by_vehicle.setdefault(vehicle_id, []).append((uid8, blend))
    return by_vehicle


def load_vehicle_metadata() -> dict[str, dict]:
    if not MODELS_INDEX.exists():
        return {}
    return {
        row["vehicle_id"]: row
        for row in csv.DictReader(MODELS_INDEX.open(encoding="utf-8-sig"))
    }


def score_option(row: dict | None, blend: Path) -> float:
    quality = (row or {}).get("asset_quality", "unknown")
    triangles = parse_triangles((row or {}).get("triangles", ""))
    base = QUALITY_SCORE.get(quality, QUALITY_SCORE["unknown"])
    size_bonus = blend.stat().st_size / 1024 if blend.exists() else 0
    return base + triangles + size_bonus


def main() -> None:
    candidate_index = load_candidate_index()
    vehicle_metadata = load_vehicle_metadata()
    fitted = collect_fitted_files()

    claimed_uid_by_model: dict[tuple[str, str, str], str] = {}

    out_rows: list[dict[str, str]] = []
    for vehicle_id in sorted(fitted):
        options: list[tuple[float, str, Path, dict | None]] = []
        for uid8, blend in fitted[vehicle_id]:
            row = candidate_index.get((vehicle_id, uid8))
            if row is None:
                continue
            options.append((score_option(row, blend), uid8, blend, row))
        if not options:
            continue
        options.sort(key=lambda item: item[0], reverse=True)

        best_score, best_uid8, best_blend, best_row = None, None, None, None
        for score, uid8, blend, row in options:
            mark = (row or {}).get("mark", "")
            model = (row or {}).get("model", "")
            claim_key = (mark, model, uid8)
            if claim_key in claimed_uid_by_model and claimed_uid_by_model[claim_key] != vehicle_id:
                continue
            best_score, best_uid8, best_blend, best_row = score, uid8, blend, row
            break
        if best_row is None:
            continue
        meta = vehicle_metadata.get(vehicle_id, {})
        mark = (best_row or {}).get("mark", "")
        model = (best_row or {}).get("model", "")
        claimed_uid_by_model[(mark, model, best_uid8)] = vehicle_id
        out_rows.append(
            {
                "vehicle_id": vehicle_id,
                "mark": mark,
                "model": model,
                "generation": meta.get("generation", (best_row or {}).get("generation", "")),
                "restyling": meta.get("restyling", (best_row or {}).get("restyling", "")),
                "year_from": meta.get("year_from", (best_row or {}).get("year_from", "")),
                "year_to": meta.get("year_to", (best_row or {}).get("year_to", "")),
                "best_blend_path": str(best_blend),
                "sketchfab_uid8": best_uid8,
                "asset_quality": (best_row or {}).get("asset_quality", "unknown"),
                "triangles": (best_row or {}).get("triangles", ""),
                "candidate_url": (best_row or {}).get("candidate_url", ""),
                "license_status": (best_row or {}).get("license_status", ""),
                "alternatives_count": str(len(options) - 1),
                "selection_score": f"{best_score:.0f}",
                "note": (best_row or {}).get("note", ""),
            }
        )

    fields = [
        "vehicle_id",
        "mark",
        "model",
        "generation",
        "restyling",
        "year_from",
        "year_to",
        "best_blend_path",
        "sketchfab_uid8",
        "asset_quality",
        "triangles",
        "candidate_url",
        "license_status",
        "alternatives_count",
        "selection_score",
        "note",
    ]
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with OUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(out_rows)

    multi = sum(1 for r in out_rows if int(r["alternatives_count"]) > 0)
    print(f"Selected best external model for {len(out_rows)} vehicles")
    print(f"Vehicles with alternatives: {multi}")
    print(f"Saved: {OUT_CSV}")


if __name__ == "__main__":
    main()
