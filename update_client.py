# -*- coding: utf-8 -*-
import hashlib
import base64
import json
import os
import re
import tempfile
import urllib.request
from pathlib import Path
from urllib.parse import urlparse


MAX_MANIFEST_BYTES = 1024 * 1024
DEFAULT_USER_AGENT = "NaeilERPUpdater/1.0"


class UpdateError(Exception):
    pass


def parse_version(version):
    text = str(version or "").strip().lower()
    text = text[1:] if text.startswith("v") else text
    numbers = [int(part) for part in re.findall(r"\d+", text)]
    return tuple(numbers or [0])


def is_newer_version(latest_version, current_version):
    latest = parse_version(latest_version)
    current = parse_version(current_version)
    size = max(len(latest), len(current))
    latest += (0,) * (size - len(latest))
    current += (0,) * (size - len(current))
    return latest > current


def _read_url(url, timeout):
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": DEFAULT_USER_AGENT,
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        content = response.read(MAX_MANIFEST_BYTES + 1)
    if len(content) > MAX_MANIFEST_BYTES:
        raise UpdateError("latest.json 파일이 너무 큽니다.")
    return content


def _decode_manifest_json(raw):
    payload = json.loads(raw.decode("utf-8-sig"))
    if (
        isinstance(payload, dict)
        and payload.get("encoding") == "base64"
        and isinstance(payload.get("content"), str)
    ):
        content = payload["content"].replace("\n", "")
        decoded = base64.b64decode(content)
        return json.loads(decoded.decode("utf-8-sig"))
    return payload


def fetch_latest_manifest(latest_url, timeout=8):
    if not latest_url:
        raise UpdateError("업데이트 확인 URL이 비어 있습니다.")

    try:
        raw = _read_url(latest_url, timeout)
        manifest = _decode_manifest_json(raw)
    except UpdateError:
        raise
    except Exception as exc:
        raise UpdateError(f"latest.json 확인 실패: {exc}") from exc

    if not isinstance(manifest, dict):
        raise UpdateError("latest.json 형식이 올바르지 않습니다.")

    version = str(manifest.get("version", "")).strip()
    download_url = str(
        manifest.get("download_url") or manifest.get("url") or ""
    ).strip()
    sha256 = str(manifest.get("sha256", "")).strip().lower()

    if not version:
        raise UpdateError("latest.json에 version 값이 없습니다.")
    if not download_url:
        raise UpdateError("latest.json에 download_url 값이 없습니다.")
    if not re.fullmatch(r"[0-9a-fA-F]{64}", sha256):
        raise UpdateError("latest.json의 sha256 값이 올바르지 않습니다.")

    manifest["version"] = version
    manifest["download_url"] = download_url
    manifest["sha256"] = sha256
    return manifest


def fetch_available_update(latest_url, current_version, timeout=8):
    manifest = fetch_latest_manifest(latest_url, timeout=timeout)
    if is_newer_version(manifest["version"], current_version):
        return manifest
    return None


def format_release_notes(manifest, max_chars=600):
    notes = manifest.get("notes") or manifest.get("description") or ""
    if isinstance(notes, list):
        notes = "\n".join(f"- {item}" for item in notes)
    notes = str(notes).strip()
    if len(notes) > max_chars:
        notes = notes[: max_chars - 3].rstrip() + "..."
    return notes


def _installer_filename(download_url, version):
    path_name = Path(urlparse(download_url).path).name
    if path_name and path_name.lower().endswith(".exe"):
        return path_name
    safe_version = re.sub(r"[^0-9A-Za-z_.-]+", "_", str(version or "latest"))
    return f"NaeilERPUpdater_Setup_{safe_version}.exe"


def download_installer(manifest, timeout=20, progress_callback=None):
    download_url = manifest["download_url"]
    expected_sha256 = manifest["sha256"].lower()
    file_name = _installer_filename(download_url, manifest.get("version"))
    target_dir = Path(tempfile.mkdtemp(prefix="NaeilERPUpdaterDownload_"))
    target_path = target_dir / file_name

    request = urllib.request.Request(
        download_url,
        headers={"User-Agent": DEFAULT_USER_AGENT},
    )
    digest = hashlib.sha256()
    downloaded = 0

    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            total_header = response.headers.get("Content-Length")
            total = int(total_header) if total_header and total_header.isdigit() else None
            with target_path.open("wb") as out_file:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    out_file.write(chunk)
                    digest.update(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        progress_callback(downloaded, total)
    except Exception as exc:
        raise UpdateError(f"설치본 다운로드 실패: {exc}") from exc

    actual_sha256 = digest.hexdigest().lower()
    if actual_sha256 != expected_sha256:
        try:
            os.remove(target_path)
        except Exception:
            pass
        raise UpdateError(
            "설치본 SHA256 검증 실패\n"
            f"기대값: {expected_sha256}\n"
            f"실제값: {actual_sha256}"
        )

    return target_path
