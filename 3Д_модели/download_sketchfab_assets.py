#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скачивает 3D-модели с Sketchfab по списку кандидатов.
Нужен API-токен: sketchfab.com → Settings → Password & API.

PowerShell:
    $env:SKETCHFAB_TOKEN = "ВАШ_ТОКЕН"
    python download_sketchfab_assets.py --limit 20

The script:
  - reads cars_dataset/metadata/asset_candidates.csv;
  - downloads each unique Sketchfab model once;
  - extracts archives into cars_dataset/external_assets/sketchfab/<uid>/;
  - writes local paths back to asset_candidates.csv.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import sys
import time
import zipfile
from pathlib import Path
from urllib.parse import urlparse

import requests


ROOT = Path(__file__).parent.resolve()
DATASET_ROOT = ROOT / "cars_dataset"
ASSET_CANDIDATES = DATASET_ROOT / "metadata" / "asset_candidates.csv"
DOWNLOAD_ROOT = DATASET_ROOT / "external_assets" / "sketchfab"
API_ROOT = "https://api.sketchfab.com/v3"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Download Sketchfab candidates from asset_candidates.csv")
    parser.add_argument("--limit", type=int, default=0, help="Max unique models to download; 0 = all")
    parser.add_argument(
        "--prefer",
        choices=["source", "gltf"],
        default="source",
        help="Prefer source archive when available, otherwise glTF.",
    )
    parser.add_argument("--sleep", type=float, default=0.5, help="Delay between downloads (0 = none)")
    parser.add_argument(
        "--missing-only",
        action="store_true",
        help="Only attempt UIDs without IMPORTABLE_PATH.txt",
    )
    parser.add_argument("--batch-size", type=int, default=0, help="Pause after N downloads; 0 = disabled")
    parser.add_argument("--batch-pause", type=float, default=0.0, help="Seconds to pause after each batch")
    return parser.parse_args()


def sketchfab_uid(url: str) -> str | None:
    """Extract Sketchfab model UID from /3d-models/slug-uid URL."""
    if "sketchfab.com" not in url:
        return None
    path = urlparse(url).path.strip("/")
    parts = path.split("/")
    if len(parts) < 2 or parts[0] != "3d-models":
        return None
    slug_uid = parts[1]
    match = re.search(r"([0-9a-f]{32})$", slug_uid)
    return match.group(1) if match else None


def auth_headers(token: str) -> dict[str, str]:
    # Account API token from sketchfab.com/settings/password uses "Token".
    # OAuth access tokens from the login flow use "Bearer".
    scheme = os.environ.get("SKETCHFAB_AUTH_SCHEME", "Token").strip() or "Token"
    return {"Authorization": f"{scheme} {token}"}


def get_download_info(uid: str, token: str, max_retries: int = 4) -> dict | None:
    """Return download JSON or None if API rate-limited (429)."""
    for attempt in range(max_retries):
        resp = requests.get(f"{API_ROOT}/models/{uid}/download", headers=auth_headers(token), timeout=60)
        if resp.status_code == 429:
            retry_after = int(resp.headers.get("Retry-After", 0) or 0)
            wait_s = retry_after if retry_after > 0 else min(180, 45 + attempt * 45)
            print(f"  лимит API Sketchfab, пауза {wait_s}с (попытка {attempt + 1}/{max_retries})")
            time.sleep(wait_s)
            continue
        if resp.status_code == 401:
            raise RuntimeError(
                "Sketchfab token rejected (401). Use an API token from sketchfab.com/settings/password "
                "(default auth: Token). For OAuth access tokens set SKETCHFAB_AUTH_SCHEME=Bearer."
            )
        if resp.status_code == 403:
            raise RuntimeError(f"Model {uid} is not downloadable for this token (403).")
        if resp.status_code == 404:
            raise RuntimeError(f"Model {uid} not found or not downloadable (404).")
        resp.raise_for_status()
        return resp.json()
    return None


def pick_download(download_info: dict, prefer: str) -> tuple[str, str]:
    """Return (kind, url). Download API usually exposes source/gltf/usdz entries."""
    ordered = [prefer, "gltf", "source", "usdz"]
    seen = set()
    for kind in ordered:
        if kind in seen:
            continue
        seen.add(kind)
        item = download_info.get(kind)
        if isinstance(item, dict) and item.get("url"):
            return kind, item["url"]
    raise RuntimeError("No downloadable format in API response.")


def download_file(url: str, out_file: Path, max_retries: int = 4) -> None:
    out_file.parent.mkdir(parents=True, exist_ok=True)
    last_exc: Exception | None = None
    for attempt in range(max_retries):
        try:
            with requests.get(url, stream=True, timeout=600) as resp:
                resp.raise_for_status()
                with out_file.open("wb") as f:
                    for chunk in resp.iter_content(chunk_size=1024 * 1024):
                        if chunk:
                            f.write(chunk)
            return
        except (requests.RequestException, OSError) as exc:
            last_exc = exc
            if out_file.exists():
                out_file.unlink()
            time.sleep(min(60, 2 ** attempt))
    if last_exc:
        raise last_exc


def safe_extract_zip(zf: zipfile.ZipFile, target_dir: Path) -> None:
    """Extract zip safely on Windows (trim spaces, skip bad paths)."""
    root = target_dir.resolve()
    for member in zf.infolist():
        name = member.filename.replace(":", "_").strip().strip("/")
        if not name:
            continue
        dest = (target_dir / name).resolve()
        if not str(dest).startswith(str(root)):
            continue
        try:
            if member.is_dir() or name.endswith("/"):
                dest.mkdir(parents=True, exist_ok=True)
                continue
            dest.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(member) as src, dest.open("wb") as dst:
                shutil.copyfileobj(src, dst)
        except OSError:
            continue


def unpack_if_zip(path: Path, target_dir: Path) -> Path:
    if path.suffix.lower() != ".zip":
        return path
    extract_dir = target_dir / "extracted"
    if extract_dir.exists():
        shutil.rmtree(extract_dir, ignore_errors=True)
    extract_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path) as zf:
        safe_extract_zip(zf, extract_dir)
    return extract_dir


def unpack_nested_archives(root: Path) -> None:
    while True:
        unpacked_any = False
        for archive in list(root.rglob("*.zip")):
            target = archive.parent / archive.stem.strip()
            if target.exists() and any(target.iterdir()):
                continue
            target.mkdir(parents=True, exist_ok=True)
            try:
                with zipfile.ZipFile(archive) as zf:
                    safe_extract_zip(zf, target)
                unpacked_any = True
            except (OSError, zipfile.BadZipFile):
                continue
        if not unpacked_any:
            break


def first_importable_path(root: Path) -> Path:
    if root.is_file():
        return root
    unpack_nested_archives(root)
    preferred_exts = [".blend", ".fbx", ".glb", ".gltf", ".obj", ".usdc", ".usd", ".dae"]
    files = [p for p in root.rglob("*") if p.is_file()]
    for ext in preferred_exts:
        matches = [p for p in files if p.suffix.lower() == ext]
        if matches:
            return matches[0]
    return root


def main() -> None:
    args = parse_args()
    token = os.environ.get("SKETCHFAB_TOKEN", "").strip()
    if not token:
        print("SKETCHFAB_TOKEN is not set.")
        print("Create/get a Sketchfab API token, then run:")
        print('  $env:SKETCHFAB_TOKEN = "YOUR_TOKEN"')
        print("  python download_sketchfab_assets.py --limit 5")
        sys.exit(2)

    if not ASSET_CANDIDATES.exists():
        raise FileNotFoundError(f"Not found: {ASSET_CANDIDATES}. Run build_asset_candidates.py first.")

    rows = list(csv.DictReader(ASSET_CANDIDATES.open(encoding="utf-8-sig")))
    if not rows:
        print("No rows in asset_candidates.csv")
        return

    fieldnames = list(rows[0].keys())
    unique: dict[str, str] = {}
    for row in rows:
        if row.get("candidate_source") != "Sketchfab":
            continue
        if row.get("candidate_status") != "found_candidate":
            continue
        url = row.get("candidate_url", "")
        uid = sketchfab_uid(url)
        if uid:
            unique[uid] = url

    uids = list(unique.keys())
    if args.missing_only:
        uids = [
            uid
            for uid in uids
            if not (DOWNLOAD_ROOT / uid / "IMPORTABLE_PATH.txt").exists()
        ]
    if args.limit > 0:
        uids = uids[: args.limit]

    print(f"Unique Sketchfab candidates to download: {len(uids)}")
    uid_to_local_path: dict[str, str] = {}
    failures: dict[str, str] = {}
    deferred: list[str] = []
    ok_since_pause = 0

    def try_download(uid: str, label: str) -> str:
        target_dir = DOWNLOAD_ROOT / uid
        importable_marker = target_dir / "IMPORTABLE_PATH.txt"
        if importable_marker.exists():
            local_path = importable_marker.read_text(encoding="utf-8").strip()
            if local_path:
                uid_to_local_path[uid] = local_path
                print(f"{label} cached {uid}")
                return "cached"
        try:
            print(f"{label} API {uid}")
            info = get_download_info(uid, token)
            if info is None:
                print(f"  отложено (лимит API) — попробуем позже")
                return "deferred"
            kind, download_url = pick_download(info, args.prefer)
            archive_ext = ".zip" if kind in ("source", "gltf") else ".bin"
            archive_path = target_dir / f"{kind}_{uid}{archive_ext}"
            print(f"  скачивание {kind} -> {archive_path.name}")
            download_file(download_url, archive_path)
            unpacked = unpack_if_zip(archive_path, target_dir)
            importable = first_importable_path(unpacked)
            importable_marker.write_text(str(importable), encoding="utf-8")
            uid_to_local_path[uid] = str(importable)
            return "ok"
        except Exception as exc:
            failures[uid] = str(exc)
            print(f"  ошибка: {exc}")
            return "fail"

    for i, uid in enumerate(uids, 1):
        status = try_download(uid, f"[{i}/{len(uids)}]")
        if status == "deferred":
            deferred.append(uid)
            if args.sleep > 0:
                time.sleep(args.sleep)
            continue
        if status == "ok":
            ok_since_pause += 1
            if args.sleep > 0:
                time.sleep(args.sleep)
            if args.batch_size > 0 and ok_since_pause >= args.batch_size and args.batch_pause > 0:
                print(f"Пакет из {args.batch_size} скачан — пауза {int(args.batch_pause)}с (лимит API)")
                time.sleep(args.batch_pause)
                ok_since_pause = 0

    if deferred:
        print(f"\nПовтор отложенных: {len(deferred)} моделей")
        if args.batch_pause > 0:
            time.sleep(min(args.batch_pause, 60))
        still_deferred: list[str] = []
        for j, uid in enumerate(deferred, 1):
            status = try_download(uid, f"[deferred {j}/{len(deferred)}]")
            if status == "deferred":
                still_deferred.append(uid)
            elif status == "ok" and args.sleep > 0:
                time.sleep(args.sleep)
        if still_deferred:
            failures.update({uid: "rate_limited_deferred" for uid in still_deferred})
            print(f"Всё ещё под лимитом API: {len(still_deferred)} — повторите через 1-2 часа")

    for row in rows:
        uid = sketchfab_uid(row.get("candidate_url", ""))
        if not uid:
            continue
        if uid in uid_to_local_path:
            row["local_asset_path"] = uid_to_local_path[uid]
            row["scale_check_status"] = "downloaded_not_fitted"
            row["note"] = (row.get("note", "") + " Downloaded automatically.").strip()
        elif uid in failures:
            row["note"] = (row.get("note", "") + f" Download failed: {failures[uid]}").strip()

    with ASSET_CANDIDATES.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Updated: {ASSET_CANDIDATES}")
    print(f"Downloaded unique models: {len(uid_to_local_path)}")
    print(f"Failures: {len(failures)}")
    if failures:
        print("Some failures are expected for NoAI/restricted/deleted models or token permissions.")


if __name__ == "__main__":
    main()
