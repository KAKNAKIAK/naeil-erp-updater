# -*- coding: utf-8 -*-
import argparse
import hashlib
import os
import subprocess
import sys
import time
from pathlib import Path


APP_NAME = "NaeilERPUpdater"
LOG_NAME = "NaeilERPUpdater_update.log"


def log(message):
    try:
        log_path = Path(os.environ.get("TEMP", ".")) / LOG_NAME
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except Exception:
        pass


def sha256_file(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest().lower()


def verify_sha256(path, expected):
    expected = (expected or "").strip().lower()
    if not expected:
        return
    actual = sha256_file(path)
    if actual != expected:
        raise RuntimeError(f"SHA256 mismatch: expected={expected}, actual={actual}")


def wait_for_process_exit(pid, timeout=90):
    if not pid:
        return True
    deadline = time.time() + timeout

    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.windll.kernel32
            synchronize = 0x00100000
            handle = kernel32.OpenProcess(synchronize, False, int(pid))
            if not handle:
                return True

            remaining = max(1, int((deadline - time.time()) * 1000))
            result = kernel32.WaitForSingleObject(handle, remaining)
            kernel32.CloseHandle(handle)
            return result == 0
        except Exception as exc:
            log(f"Windows process wait fallback: {exc}")

    while time.time() < deadline:
        try:
            os.kill(int(pid), 0)
        except Exception:
            return True
        time.sleep(0.5)
    return False


def run_installer(installer, install_dir):
    cmd = [
        str(installer),
        "--silent",
        "--install-dir",
        str(install_dir),
        "--no-shortcuts",
    ]
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    result = subprocess.run(cmd, creationflags=creationflags)
    if result.returncode != 0:
        raise RuntimeError(f"installer exited with code {result.returncode}")


def relaunch_app(app_exe, install_dir):
    if not Path(app_exe).exists():
        raise FileNotFoundError(f"app exe not found: {app_exe}")
    subprocess.Popen([str(app_exe)], cwd=str(install_dir))


def parse_args():
    parser = argparse.ArgumentParser(description="Naeil ERP Updater helper.")
    parser.add_argument("--installer", required=True)
    parser.add_argument("--sha256", default="")
    parser.add_argument("--install-dir", required=True)
    parser.add_argument("--app-exe", required=True)
    parser.add_argument("--wait-pid", type=int, default=0)
    parser.add_argument("--wait-timeout", type=int, default=90)
    return parser.parse_args()


def main():
    args = parse_args()
    installer = Path(args.installer)
    install_dir = Path(args.install_dir)
    app_exe = Path(args.app_exe)

    try:
        log("update helper started")
        verify_sha256(installer, args.sha256)
        if not wait_for_process_exit(args.wait_pid, timeout=args.wait_timeout):
            raise TimeoutError(f"{APP_NAME}.exe did not exit in time")
        run_installer(installer, install_dir)
        relaunch_app(app_exe, install_dir)
        log("update helper finished")
        return 0
    except Exception as exc:
        log(f"update helper failed: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
