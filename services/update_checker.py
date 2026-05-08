from __future__ import annotations

import json
import urllib.request
import urllib.error
import ssl
from dataclasses import dataclass
from typing import Optional

import certifi


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
    cleaned = _normalize_version(v)
    parts = []
    for p in cleaned.split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def _pick_best_asset(assets: list[dict]) -> tuple[str, str]:
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


def _fetch_json(url: str, timeout: int) -> dict | list | None:
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "Suzuki-Parts-Manager",
        },
    )
    context = ssl.create_default_context(cafile=certifi.where())
    with urllib.request.urlopen(req, timeout=timeout, context=context) as resp:
        return json.loads(resp.read().decode("utf-8"))


def check_github_release(
    *,
    current_version: str,
    github_owner: str,
    github_repo: str,
    timeout_seconds: int = 8,
) -> UpdateInfo:
    base = f"https://api.github.com/repos/{github_owner}/{github_repo}"

    # Probeer eerst /releases/latest (snelste pad)
    try:
        data = _fetch_json(f"{base}/releases/latest", timeout_seconds)
        releases = [data] if isinstance(data, dict) else []
    except urllib.error.HTTPError as e:
        if e.code == 404:
            # Geen "latest" release (bijv. alleen prereleases) — haal lijst op
            releases = []
        else:
            return UpdateInfo(
                update_available=False,
                current_version=current_version,
                error=f"HTTP {e.code}",
            )
    except urllib.error.URLError as e:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error=f"Geen verbinding: {e.reason}",
        )
    except Exception as e:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error=str(e),
        )

    # Fallback: haal de releases lijst op en pak de eerste niet-draft release
    if not releases:
        try:
            data = _fetch_json(f"{base}/releases?per_page=10", timeout_seconds)
            if isinstance(data, list):
                for r in data:
                    if not r.get("draft", False):
                        releases = [r]
                        break
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
                error=f"Geen verbinding: {e.reason}",
            )
        except Exception as e:
            return UpdateInfo(
                update_available=False,
                current_version=current_version,
                error=str(e),
            )

    if not releases:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error="Geen releases gevonden op GitHub.",
        )

    release = releases[0]
    latest_tag = str(release.get("tag_name", "") or "")
    release_url = str(release.get("html_url", "") or "")
    assets = release.get("assets", []) or []

    asset_name, download_url = _pick_best_asset(assets)

    if not latest_tag:
        return UpdateInfo(
            update_available=False,
            current_version=current_version,
            error="Geen versienummer gevonden in de release.",
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
