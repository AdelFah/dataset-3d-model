#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Собирает финальный индекс всех 3D-моделей датасета.
Для каждого автомобиля записывает какая модель с Sketchfab ему соответствует
и какие габариты используются. Результат сохраняется в final_models_index.csv.

PowerShell:
    python select_best_external_assets.py
    python build_final_models_index.py
"""

from __future__ import annotations

import csv
import os
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).parent.resolve()
DATASET_ROOT = Path(os.environ.get("CARS_DATASET_ROOT", ROOT / "cars_dataset")).resolve()
MODELS_INDEX = DATASET_ROOT / "metadata" / "models_index.csv"
BEST_EXTERNAL = DATASET_ROOT / "metadata" / "best_external_models.csv"
OUT_CSV = DATASET_ROOT / "metadata" / "final_models_index.csv"
PROCEDURAL_DIR = DATASET_ROOT / "models"


def version_key(row: dict) -> tuple[str, str, str, str]:
    return (
        row.get("generation", ""),
        row.get("restyling", ""),
        row.get("year_from", ""),
        row.get("year_to", ""),
    )


def reused_visual_version_ids(best_rows: list[dict]) -> set[str]:
    grouped: dict[tuple[str, str, str], list[dict]] = defaultdict(list)
    for row in best_rows:
        uid = row.get("sketchfab_uid8", "")
        if not uid:
            continue
        grouped[(row.get("mark", ""), row.get("model", ""), uid)].append(row)

    invalid: set[str] = set()
    for rows in grouped.values():
        if len(rows) < 2:
            continue
        versions = {version_key(row) for row in rows}
        if len(versions) < 2:
            continue
        # Keep one best row per shared UID; reject the rest.
        rows_sorted = sorted(rows, key=lambda r: (r.get("generation", ""), r.get("restyling", "")))
        invalid.update(row["vehicle_id"] for row in rows_sorted[1:])
    return invalid


def main() -> None:
    if not MODELS_INDEX.exists():
        raise FileNotFoundError(MODELS_INDEX)

    best_by_vehicle: dict[str, dict] = {}
    best_rows: list[dict] = []
    if BEST_EXTERNAL.exists():
        best_rows = list(csv.DictReader(BEST_EXTERNAL.open(encoding="utf-8-sig")))
        for row in best_rows:
            best_by_vehicle[row["vehicle_id"]] = row
    invalid_reuse = reused_visual_version_ids(best_rows)

    out_rows: list[dict[str, str]] = []
    for row in csv.DictReader(MODELS_INDEX.open(encoding="utf-8-sig")):
        vehicle_id = row["vehicle_id"]
        external = best_by_vehicle.get(vehicle_id)
        procedural = PROCEDURAL_DIR / f"{vehicle_id}.blend"

        if external and Path(external["best_blend_path"]).exists() and vehicle_id not in invalid_reuse:
            out_rows.append(
                {
                    "vehicle_id": vehicle_id,
                    "mark": row.get("mark", external.get("mark", "")),
                    "model": row.get("model", external.get("model", "")),
                    "generation": row.get("generation", ""),
                    "restyling": row.get("restyling", ""),
                    "year_from": row.get("year_from", ""),
                    "year_to": row.get("year_to", ""),
                    "body_class": row.get("body_class", ""),
                    "length_mm": row.get("length_mm", ""),
                    "width_mm": row.get("width_mm", ""),
                    "height_mm": row.get("height_mm", ""),
                    "source_type": "external",
                    "model_path": external["best_blend_path"],
                    "asset_quality": external.get("asset_quality", ""),
                    "sketchfab_uid8": external.get("sketchfab_uid8", ""),
                    "visual_match_status": "unique_version_asset",
                    "status": "ok",
                    "fallback_path": "",
                }
            )
        else:
            status = "needs_generation_specific_asset" if vehicle_id in invalid_reuse else "search_needed"
            out_rows.append(
                {
                    "vehicle_id": vehicle_id,
                    "mark": row.get("mark", ""),
                    "model": row.get("model", ""),
                    "generation": row.get("generation", ""),
                    "restyling": row.get("restyling", ""),
                    "year_from": row.get("year_from", ""),
                    "year_to": row.get("year_to", ""),
                    "body_class": row.get("body_class", ""),
                    "length_mm": row.get("length_mm", ""),
                    "width_mm": row.get("width_mm", ""),
                    "height_mm": row.get("height_mm", ""),
                    "source_type": "missing",
                    "model_path": "",
                    "asset_quality": "",
                    "sketchfab_uid8": external.get("sketchfab_uid8", "") if external else "",
                    "visual_match_status": "reused_across_versions" if vehicle_id in invalid_reuse else "",
                    "status": status,
                    "fallback_path": "",
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
        "body_class",
        "length_mm",
        "width_mm",
        "height_mm",
        "source_type",
        "model_path",
        "asset_quality",
        "sketchfab_uid8",
        "visual_match_status",
        "status",
        "fallback_path",
    ]
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with OUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(out_rows)

    external = sum(1 for r in out_rows if r["source_type"] == "external")
    needs_specific = sum(1 for r in out_rows if r["status"] == "needs_generation_specific_asset")
    missing = sum(1 for r in out_rows if r["status"] == "search_needed")
    ok_by_rst: dict[str, int] = {}
    for r in out_rows:
        if r["status"] == "ok":
            key = str(r.get("restyling", ""))
            ok_by_rst[key] = ok_by_rst.get(key, 0) + 1
    print(f"Final registry: {len(out_rows)} vehicles (per generation + restyling)")
    print(f"  external ok: {external}")
    print(f"  needs_generation_specific_asset: {needs_specific}")
    print(f"  search_needed: {missing}")
    if ok_by_rst:
        print(f"  ok by restyling: {dict(sorted(ok_by_rst.items(), key=lambda x: int(x[0]) if x[0].lstrip('-').isdigit() else 999))}")
    print(f"Saved: {OUT_CSV}")


if __name__ == "__main__":
    main()
