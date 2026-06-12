#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Формирует список кандидатов 3D-моделей для каждого автомобиля в датасете.
Записывает ссылки на найденные модели и их характеристики
для последующего скачивания и отбора.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path
from urllib.parse import quote_plus


ROOT = Path(__file__).parent.resolve()
DATASET_ROOT = ROOT / "cars_dataset"
MODELS_INDEX = DATASET_ROOT / "metadata" / "models_index.csv"
OUT_CSV = DATASET_ROOT / "metadata" / "asset_candidates.csv"
DOWNLOAD_DIR = DATASET_ROOT / "external_assets"


def norm(value: str) -> str:
    value = str(value).strip().lower().replace("ё", "е")
    return re.sub(r"\s+", " ", value)


def version_query(mark: str, model: str, row: dict[str, str]) -> str:
    generation = row.get("generation", "")
    restyling = row.get("restyling", "")
    year_from = row.get("year_from", "")
    year_to = row.get("year_to", "")
    parts = [mark, model]
    if generation:
        parts.append(f"generation {generation}")
    if restyling not in ("",):
        parts.append(f"restyling {restyling}")
    if year_from or year_to:
        parts.append(f"{year_from}-{year_to}".strip("-"))
    parts.extend(["car", "downloadable", "3d model"])
    return " ".join(str(part) for part in parts if part)


def sketchfab_search_url(mark: str, model: str, row: dict[str, str] | None = None) -> str:
    q = quote_plus(version_query(mark, model, row or {}))
    return f"https://sketchfab.com/search?features=downloadable&type=models&q={q}"


def beamng_search_url(mark: str, model: str, row: dict[str, str] | None = None) -> str:
    q = quote_plus(version_query(mark, model, row or {}))
    return f"https://www.beamng.com/resources/?query={q}"


# Seeded from manual web search. Licenses must be rechecked on download day.
KNOWN_CANDIDATES = {
    ("lada", "приора"): [
        {
            "url": "https://sketchfab.com/3d-models/lada-priora-vaz-sedan-b6522aae5ada4911acbf5b446d6bcc51",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "12.3k",
            "note": "Low-poly game-ready sedan.",
        },
        {
            "url": "https://sketchfab.com/3d-models/lada-priora-vaz-2170-4c67aa9408564f989460878c374dcb0f",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "12.3k",
            "note": "VAZ-2170 candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/vaz-priora-free-8e04c24b03fc4a8fb01c6c064dbd6074",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "high",
            "triangles": "1M",
            "note": "Very detailed and heavy.",
        },
    ],
    ("lada", "2107"): [
        {
            "url": "https://sketchfab.com/3d-models/low-poly-lada-2107-4d5aac74abed45eaafad6ff73f85501b",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "5.5k",
            "note": "Low-poly; page lists real dimensions.",
        },
        {
            "url": "https://sketchfab.com/3d-models/lada-2107-55bd23c141284e758e4eacc01df14037",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "107k",
            "note": "Detailed candidate.",
        },
    ],
    ("lada", "4x4 2121 нива"): [
        {
            "url": "https://sketchfab.com/3d-models/lada-niva-5a83c0cc63df418fa5bb9a2bf98dc9a6",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "high",
            "triangles": "134.2k",
            "note": "Lada Niva candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/low-poly-lada-niva-b67df18b55ab497ba30cdbd6be7dee19",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "5.8k",
            "note": "Low-poly fallback.",
        },
    ],
    ("toyota", "corolla"): [
        {
            "url": "https://sketchfab.com/3d-models/toyota-corolla-e170-2017-d87af59849814c5cb35a0c233f2c2f19",
            "source": "Sketchfab",
            "license": "CC Attribution-ShareAlike",
            "quality": "high",
            "triangles": "755.9k",
            "note": "Detailed E170 2017 candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/toyota-corolla-ae86-trueno-fe02fba6302e450ea8424591493341ea",
            "source": "Sketchfab",
            "license": "CC Attribution / NoAI",
            "quality": "medium",
            "triangles": "27.2k",
            "note": "AE86 only; NoAI restriction.",
        },
    ],
    ("toyota", "camry"): [
        {
            "url": "https://sketchfab.com/3d-models/toyota-camry-2020-236a5a6e2fa6420fbdf641f4800cd544",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "high",
            "triangles": "176.7k",
            "note": "Modern Camry candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/toyota-camry-40-560c174d073f4a10bb153ded41ada8c9",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "18.4k",
            "note": "Camry 40 candidate.",
        },
    ],
    ("toyota", "rav4"): [
        {
            "url": "https://sketchfab.com/3d-models/toyota-rav4-awd-2015-28e93488e507472ea0f555aa880d4865",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "high",
            "triangles": "2.7M",
            "note": "Very heavy RAV4 candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/2023-toyota-rav4-hybrid-ed155ad0cb7d447085a519eaff9aa2df",
            "source": "Sketchfab",
            "license": "check",
            "quality": "high",
            "triangles": "399.2k",
            "note": "2023 RAV4 Hybrid candidate.",
        },
    ],
    ("ford", "focus"): [
        {
            "url": "https://sketchfab.com/3d-models/2016-ford-focus-rs-27e25fcead154e62a1dd8e92ccedb691",
            "source": "Sketchfab",
            "license": "CC Attribution-NonCommercial-ShareAlike",
            "quality": "high",
            "triangles": "90.1k",
            "note": "NC license; check before dataset inclusion.",
        },
        {
            "url": "https://sketchfab.com/3d-models/ford-focus-hatchback-2012-0adceae5166e40a58501d8d8e3e9b9a9",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "high",
            "triangles": "1.4M",
            "note": "Heavy hatchback candidate.",
        },
    ],
    ("volkswagen", "polo"): [
        {
            "url": "https://sketchfab.com/3d-models/volkswagen-polo-2017-0e6da82f444949538104141f51905246",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "126.6k",
            "note": "2017 Polo candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/2017-volkswagen-polo-sedan-290a03e4dbfc4e5ea351efa4f89b38b7",
            "source": "Sketchfab",
            "license": "Free Standard / NoAI",
            "quality": "medium",
            "triangles": "132.5k",
            "note": "Russian-market sedan; NoAI restriction.",
        },
    ],
    ("mitsubishi", "lancer"): [
        {
            "url": "https://sketchfab.com/3d-models/mitsubishi-lancer-evolution-6-wwwvecarzcom-c3d5dcd8ff724bc88c46760d92fc5188",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "80.9k",
            "note": "Evolution VI candidate, not base Lancer.",
        },
        {
            "url": "https://sketchfab.com/3d-models/2022-mitsubishi-lancer-evolution-x-asphalt-8-963b8275e8c940d390ccff5abeb0ce7e",
            "source": "Sketchfab",
            "license": "CC Attribution-NonCommercial-ShareAlike",
            "quality": "medium",
            "triangles": "34.2k",
            "note": "Evolution X; NC-SA license.",
        },
    ],
    ("audi", "80"): [
        {
            "url": "https://sketchfab.com/3d-models/audi-80-lowpoly-b3e18613e61a41b295052cf2b4b70099",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "medium",
            "triangles": "4k",
            "note": "Lightweight low-poly candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/audi-80-b3-1986-1991-5781d39624d340ffb5e06ce7005ca51b",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "high",
            "triangles": "102.2k",
            "note": "Audi 80 B3 candidate.",
        },
        {
            "url": "https://sketchfab.com/3d-models/audi-80-b4-1991-2000-1cf6b54bd461481eb9aa3eec5a442b47",
            "source": "Sketchfab",
            "license": "CC Attribution",
            "quality": "high",
            "triangles": "260.7k",
            "note": "Audi 80 B4 candidate.",
        },
    ],
}


def candidates_for(mark: str, model: str) -> list[dict[str, str]]:
    mark_n = norm(mark)
    model_n = norm(model)

    exact = KNOWN_CANDIDATES.get((mark_n, model_n))
    if exact:
        return exact

    # Some dataset model names contain trim/body prefixes.
    for (candidate_mark, candidate_model), candidates in KNOWN_CANDIDATES.items():
        if mark_n == candidate_mark and (candidate_model in model_n or model_n in candidate_model):
            return candidates
    return []


def main() -> None:
    if not MODELS_INDEX.exists():
        raise FileNotFoundError(f"Not found: {MODELS_INDEX}. Run prepare_for_blender.py first.")

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)

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
        "candidate_status",
        "candidate_source",
        "candidate_url",
        "license_status",
        "asset_quality",
        "triangles",
        "local_asset_path",
        "scale_check_status",
        "note",
        "sketchfab_search_url",
        "beamng_search_url",
    ]

    rows_out = []
    rows = list(csv.DictReader(MODELS_INDEX.open(encoding="utf-8-sig")))
    for row in rows:
        mark = row["mark"]
        model = row["model"]
        # Known seed candidates are mark/model level only — do not reuse across restylings.
        found = candidates_for(mark, model) if str(row.get("restyling", "0")) in {"0", ""} else []
        if found:
            for candidate in found:
                rows_out.append(
                    {
                        "vehicle_id": row["vehicle_id"],
                        "mark": mark,
                        "model": model,
                        "generation": row["generation"],
                        "restyling": row.get("restyling", ""),
                        "year_from": row.get("year_from", ""),
                        "year_to": row.get("year_to", ""),
                        "body_class": row["body_class"],
                        "length_mm": row["length_mm"],
                        "width_mm": row["width_mm"],
                        "height_mm": row["height_mm"],
                        "candidate_status": "found_candidate",
                        "candidate_source": candidate["source"],
                        "candidate_url": candidate["url"],
                        "license_status": candidate["license"],
                        "asset_quality": candidate["quality"],
                        "triangles": candidate["triangles"],
                        "local_asset_path": "",
                        "scale_check_status": "not_imported",
                        "note": candidate["note"],
                        "sketchfab_search_url": sketchfab_search_url(mark, model, row),
                        "beamng_search_url": beamng_search_url(mark, model, row),
                    }
                )
        else:
            rows_out.append(
                {
                    "vehicle_id": row["vehicle_id"],
                    "mark": mark,
                    "model": model,
                    "generation": row["generation"],
                    "restyling": row.get("restyling", ""),
                    "year_from": row.get("year_from", ""),
                    "year_to": row.get("year_to", ""),
                    "body_class": row["body_class"],
                    "length_mm": row["length_mm"],
                    "width_mm": row["width_mm"],
                    "height_mm": row["height_mm"],
                    "candidate_status": "search_needed",
                    "candidate_source": "",
                    "candidate_url": "",
                    "license_status": "unknown",
                    "asset_quality": "unknown",
                    "triangles": "",
                    "local_asset_path": "",
                    "scale_check_status": "not_imported",
                    "note": "Use search URLs, verify license, download manually, then validate in Blender.",
                    "sketchfab_search_url": sketchfab_search_url(mark, model, row),
                    "beamng_search_url": beamng_search_url(mark, model, row),
                }
            )

    with OUT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows_out)

    found_count = sum(1 for row in rows_out if row["candidate_status"] == "found_candidate")
    search_count = sum(1 for row in rows_out if row["candidate_status"] == "search_needed")
    print(f"Saved: {OUT_CSV}")
    print(f"Rows: {len(rows_out)}")
    print(f"Found candidate rows: {found_count}")
    print(f"Search-needed rows: {search_count}")
    print(f"Manual downloads folder: {DOWNLOAD_DIR}")


if __name__ == "__main__":
    main()
