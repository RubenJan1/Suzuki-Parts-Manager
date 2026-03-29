from __future__ import annotations

import json
import urllib.request
import urllib.error
from dataclasses import dataclass
from typing import Optional


@dataclass
class UpdateInfo:
    update_available: bool
    current_version: str
    latest_version: str = ""
    release_url: str = ""
    download_url: str = ""
    asset_name: str = ""
    error: str = ""


def _normalize_version(v: str) -> str:
    return (v or "").strip().lower().removeprefix("v")


def _parse_version(v: str) -> tuple[int, ...]:
    """
    '1.2.3' -> (1, 2, 3)
    Niet-numerieke stukjes worden genegeerd.
    """
    cleaned = _normalize_version(v)
    parts = []
    for p in cleaned.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def _pick_best_asset(assets: list[dict]) -> tuple[str, str]:
    """
    Kies liefst een zip of exe.
    Retourneert (asset_name, browser_download_url).
    """
    if not assets:
        return "", ""

    preferred = []
    fallback = []

    for a in assets:
        name = str(a.get("name", "") or "")
        url = str(a.get("browser_download_url", "") or "")
        if not url:
            continue

        lower = name.lower()
        if lower.endswith(".zip") or lower.endswith(".exe"):
            preferred.append((name, url))
        else:
            fallback.append((name, url))

    if preferred:
        return preferred[0]
    if fallback:
        return fallback[0]
    return "", ""


def check_github_release(
    *,
    current_version: str,
    github_owner: str,
    github_repo: str,
    timeout_seconds: int = 3,
) -> UpdateInfo:
    """
    Checkt de latest GitHub release.
    Bij fout: geen crash, maar UpdateInfo met error.
    """
    api_url = f"https://api.github.com/repos/{github_owner}/{github_repo}/releases/latest"

    req = urllib.request.Request(
        api_url,
        headers={
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "Suzuki-Parts-Manager",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout_seconds) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error=f"HTTP {e.code}",
        )
    except urllib.error.URLError as e:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error=f"Netwerkfout: {e.reason}",
        )
    except Exception as e:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error=str(e),
        )

    latest_tag = str(data.get("tag_name", "") or "")
    release_url = str(data.get("html_url", "") or "")
    assets = data.get("assets", []) or []

    asset_name, download_url = _pick_best_asset(assets)

    if not latest_tag:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error="Geen tag_name gevonden in latest release.",
        )

    current_parsed = _parse_version(current_version)
    latest_parsed = _parse_version(latest_tag)

    return UpdateInfo(
        update_available=(latest_parsed > current_parsed),
        current_version=current_version,
        latest_version=latest_tag,
        release_url=release_url,
        download_url=download_url,
        asset_name=asset_name,
        error="",
    )