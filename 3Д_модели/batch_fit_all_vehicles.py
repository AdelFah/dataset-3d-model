#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Подгоняет скачанные 3D-модели под реальные габариты каждого автомобиля.
Берёт модель с Sketchfab и масштабирует её до нужной длины/ширины/высоты
из датасета. Сохраняет результат в формате .blend для Blender.

PowerShell:
    $env:CARS_DATASET_ROOT = "$PWD\cars_dataset"
    python batch_fit_all_vehicles.py
    python batch_fit_all_vehicles.py --skip-existing
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
ROOT = Path(os.environ.get("CARS_DATASET_ROOT", SCRIPT_DIR / "cars_dataset")).resolve()
MODELS_INDEX = ROOT / "metadata" / "models_index.csv"
ASSET_CANDIDATES = ROOT / "metadata" / "asset_candidates.csv"
DOWNLOAD_ROOT = ROOT / "external_assets" / "sketchfab"
FITTED_DIR = ROOT / "external_fitted"
FIT_SCRIPT = SCRIPT_DIR / "fit_external_asset_blender.py"

QUALITY_RANK = {"high": 3, "medium": 2, "low": 1, "unknown": 0}
UID_RE = re.compile(r"([0-9a-f]{32})")
FITTED_RE = re.compile(r"^(.+)__([0-9a-f]{8})\.blend$", re.IGNORECASE)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--skip-existing", action="store_true", help="Skip vehicle_ids with any fitted .blend")
    parser.add_argument("--blender", default="")
    return parser.parse_args()


def find_blender(explicit: str) -> str:
    if explicit:
        return explicit
    candidates = [
        "blender",
        r"C:\Program Files\Blender Foundation\Blender 5.1\blender.exe",
        r"C:\Program Files\Blender Foundation\Blender 4.2\blender.exe",
        r"C:\Program Files\Blender Foundation\Blender 4.1\blender.exe",
        r"C:\Program Files\Blender Foundation\Blender 4.0\blender.exe",
        r"C:\Program Files\Blender Foundation\Blender 3.6\blender.exe",
    ]
    for candidate in candidates:
        if candidate == "blender":
            from shutil import which

            found = which("blender")
            if found:
                return found
        elif Path(candidate).exists():
            return candidate
    raise FileNotFoundError("Blender executable not found. Pass --blender PATH.")


def parse_triangles(value: str) -> int:
    text = str(value or "").strip().lower().replace(" ", "").replace(",", "")
    if not text:
        return 0
    mult = 1
    if text.endswith("k"):
        mult = 1_000
        text = text[:-1]
    elif text.endswith("m"):
        mult = 1_000_000
        text = text[:-1]
    try:
        return int(float(text) * mult)
    except ValueError:
        return 0


def sketchfab_uid(url: str) -> str | None:
    match = UID_RE.search(url or "")
    return match.group(1) if match else None


def asset_path_for_uid(uid: str) -> Path | None:
    marker = DOWNLOAD_ROOT / uid / "IMPORTABLE_PATH.txt"
    if not marker.exists():
        return None
    path = Path(marker.read_text(encoding="utf-8").strip())
    return path if path.exists() else None


def row_score(row: dict) -> int:
    q = QUALITY_RANK.get(row.get("asset_quality", "unknown"), 0)
    tris = parse_triangles(row.get("triangles", ""))
    has_local = 1 if row.get("local_asset_path") and Path(row["local_asset_path"]).exists() else 0
    return q * 1_000_000 + tris + has_local * 10_000


def best_candidate_for_vehicle(vehicle_id: str, candidate_rows: list[dict]) -> tuple[str, Path] | None:
    options: list[tuple[int, str, Path]] = []
    for row in candidate_rows:
        if row.get("candidate_status") != "found_candidate":
            continue
        uid = sketchfab_uid(row.get("candidate_url", ""))
        if not uid:
            continue
        local = row.get("local_asset_path", "")
        asset_path = Path(local) if local and Path(local).exists() else asset_path_for_uid(uid)
        if not asset_path:
            continue
        options.append((row_score(row), uid, asset_path))
    if not options:
        return None
    options.sort(key=lambda item: item[0], reverse=True)
    _, uid, asset_path = options[0]
    return uid, asset_path


def existing_fitted(vehicle_id: str) -> bool:
    prefix = f"{vehicle_id}__"
    return any(p.name.startswith(prefix) for p in FITTED_DIR.glob("*.blend"))


def load_candidates_by_vehicle() -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = {}
    if not ASSET_CANDIDATES.exists():
        return grouped
    for row in csv.DictReader(ASSET_CANDIDATES.open(encoding="utf-8-sig")):
        grouped.setdefault(row["vehicle_id"], []).append(row)
    return grouped


def main() -> None:
    args = parse_args()
    blender = find_blender(args.blender)
    if not MODELS_INDEX.exists():
        raise FileNotFoundError(MODELS_INDEX)

    candidates_by_vehicle = load_candidates_by_vehicle()
    vehicles = list(csv.DictReader(MODELS_INDEX.open(encoding="utf-8-sig")))
    if args.limit > 0:
        vehicles = vehicles[: args.limit]

    jobs: list[tuple[str, str, Path]] = []
    skipped = 0
    no_asset = 0
    for vehicle in vehicles:
        vehicle_id = vehicle["vehicle_id"]
        if args.skip_existing and existing_fitted(vehicle_id):
            skipped += 1
            continue
        picked = best_candidate_for_vehicle(vehicle_id, candidates_by_vehicle.get(vehicle_id, []))
        if not picked:
            no_asset += 1
            continue
        uid, asset_path = picked
        jobs.append((vehicle_id, uid, asset_path))

    print(f"Vehicles in index: {len(vehicles)}")
    print(f"Skip existing: {skipped}, no downloaded asset: {no_asset}")
    print(f"Fitting {len(jobs)} vehicles with Blender")

    ok = 0
    failed: list[tuple[str, str]] = []
    env = os.environ.copy()
    env["CARS_DATASET_ROOT"] = str(ROOT)
    FITTED_DIR.mkdir(parents=True, exist_ok=True)

    for i, (vehicle_id, uid, asset_path) in enumerate(jobs, 1):
        suffix = f"__{uid[:8]}"
        out_file = FITTED_DIR / f"{vehicle_id}{suffix}.blend"
        print(f"[{i}/{len(jobs)}] {vehicle_id} <- {asset_path.name} ({uid[:8]})")
        cmd = [
            blender,
            "--background",
            "--python",
            str(FIT_SCRIPT),
            "--",
            "--vehicle-id",
            vehicle_id,
            "--asset-file",
            str(asset_path),
            "--output-suffix",
            suffix,
        ]
        try:
            result = subprocess.run(cmd, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
            combined = f"{result.stdout or ''}\n{result.stderr or ''}"
            if result.returncode != 0 or "Saved fitted asset:" not in combined or not out_file.exists():
                err = combined.strip().splitlines()
                failed.append((vehicle_id, err[-1] if err else "no output .blend created"))
                print(f"  FAIL: {failed[-1][1]}")
            else:
                ok += 1
                print(f"  Saved: {out_file.name}")
        except Exception as exc:
            failed.append((vehicle_id, str(exc)))
            print(f"  FAIL: {exc}")

    print(f"Done: {ok} ok, {len(failed)} failed, {skipped} skipped, {no_asset} no asset")
    if failed:
        for vehicle_id, reason in failed[:20]:
            print(f"  {vehicle_id}: {reason}")
        if len(failed) > 20:
            print(f"  ... and {len(failed) - 20} more")
        sys.exit(1)


if __name__ == "__main__":
    main()
