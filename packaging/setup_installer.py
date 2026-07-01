# -*- coding: utf-8 -*-
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import threading

try:
    import tkinter as tk
    from tkinter import messagebox, ttk, filedialog
except Exception:
    tk = None
    messagebox = None
    ttk = None
    filedialog = None

try:
    import winreg
except Exception:
    winreg = None


APP_NAME = "NaeilERPUpdater"
SHORTCUT_NAME = "Naeil ERP Fare Updater"
APP_VERSION = "v5.0.11"

PUBLISHER = "Naeil Tour"
UNINSTALLER_NAME = "Uninstall.exe"
# Windows '앱 및 기능' 등록 키 (현재 사용자 범위)
UNINSTALL_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\NaeilERPUpdater"

DEFAULT_SELECTORS = {
    "login_id": "input[name='userId']",
    "login_pw": "input[name='userPw']",
    "login_btn": "#btnLogin",
    "search_date_input": "#searchStDate",
    "search_date_end_input": "#searchEnDate",
    "airline_select": "#air2Cd",
    "price_desc_input": "#priceDesc",
    "hotel_name_input": "#hotelKorNm",
    "hotel_seq_input": "#hotelSeq",
    "event_modify_button": "#eventModify",
    "bulk_update_save_button": "#appSave",
    "progress_status_checkbox": "#procCdChk",
    "progress_status_select": "#procCd",
    "search_button": "#gridMain_r",
    "header_all_checkbox": "td.aui-grid-row-check-header input",
    "update_button": "#priceUpdate",
    "adult_air_input": "#addAir01",
    "adult_hotel_input": "#addHotel11",
    "adult_land_input": "#addLand21",
    "adult_tour_input": "#addExpense40",
    "adult_profit_input": "#addProfit41",
    "child_fare_input": "#addChild90",
    "infant_fare_input": "#addInfant91",
    "save_button": "#priceSave",
    "cancel_button": "#popCloseBtn",
    "date_cell_in_row": ".aui-grid-default-column",
}


def resource_path(name):
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_dir / name


def default_install_dir():
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "NaeilTour" / "ERPUpdater"
    return Path.home() / "AppData" / "Local" / "NaeilTour" / "ERPUpdater"


def merge_v2_selectors(config_path):
    if not config_path.exists():
        return
    try:
        config = json.loads(config_path.read_text(encoding="utf-8-sig"))
        if not isinstance(config, dict):
            return
        selectors = config.get("selectors")
        if not isinstance(selectors, dict):
            selectors = {}

        merged = dict(DEFAULT_SELECTORS)
        for key, value in selectors.items():
            if key == "adult_fare_input":
                if "adult_air_input" not in selectors:
                    merged["adult_air_input"] = value
                continue
            if key == "child_fare_input" and value == "#addProfit41" and "adult_profit_input" not in selectors:
                merged["adult_profit_input"] = value
                continue
            merged[key] = value

        config["selectors"] = merged
        config_path.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def ps_quote(value):
    return "'" + str(value).replace("'", "''") + "'"


def create_shortcuts(install_dir):
    exe_path = install_dir / f"{APP_NAME}.exe"
    script = f"""
$ErrorActionPreference = "Stop"
$WshShell = New-Object -ComObject WScript.Shell

$DesktopShortcutPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "{SHORTCUT_NAME}.lnk"
$DesktopShortcut = $WshShell.CreateShortcut($DesktopShortcutPath)
$DesktopShortcut.TargetPath = {ps_quote(exe_path)}
$DesktopShortcut.WorkingDirectory = {ps_quote(install_dir)}
$DesktopShortcut.IconLocation = {ps_quote(exe_path)}
$DesktopShortcut.Save()

$StartMenuDir = Join-Path ([Environment]::GetFolderPath("Programs")) "Naeil Tour"
New-Item -ItemType Directory -Force -Path $StartMenuDir | Out-Null
$StartMenuShortcutPath = Join-Path $StartMenuDir "{SHORTCUT_NAME}.lnk"
$StartMenuShortcut = $WshShell.CreateShortcut($StartMenuShortcutPath)
$StartMenuShortcut.TargetPath = {ps_quote(exe_path)}
$StartMenuShortcut.WorkingDirectory = {ps_quote(install_dir)}
$StartMenuShortcut.IconLocation = {ps_quote(exe_path)}
$StartMenuShortcut.Save()
"""

    with tempfile.TemporaryDirectory(prefix="NaeilERPUpdaterShortcut_") as temp_dir:
        script_path = Path(temp_dir) / "create_shortcuts.ps1"
        # utf-8-sig(BOM)로 저장해야 PowerShell 5.1이 인코딩을 올바르게 인식한다.
        script_path.write_text(script, encoding="utf-8-sig")
        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if result.returncode != 0:
            detail = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(
                f"바로가기 생성 실패 (exit {result.returncode})"
                + (f"\n{detail}" if detail else "")
            )


def install_payload(install_dir, create_shortcut=True):
    payload_path = resource_path("payload.zip")
    if not payload_path.exists():
        raise FileNotFoundError(f"payload.zip was not found: {payload_path}")

    install_dir.mkdir(parents=True, exist_ok=True)
    config_path = install_dir / "config.json"
    config_backup = None

    with tempfile.TemporaryDirectory(prefix="NaeilERPUpdaterPayload_") as temp_dir:
        temp_root = Path(temp_dir)

        if config_path.exists():
            config_backup = temp_root / "config.backup.json"
            shutil.copy2(config_path, config_backup)

        with zipfile.ZipFile(payload_path, "r") as archive:
            archive.extractall(temp_root)

        for item in temp_root.iterdir():
            if item == config_backup:
                continue
            dest = install_dir / item.name
            if item.is_dir():
                shutil.copytree(item, dest, dirs_exist_ok=True)
            else:
                shutil.copy2(item, dest)

        if config_backup and config_backup.exists():
            shutil.copy2(config_backup, config_path)

    merge_v2_selectors(config_path)

    (install_dir / "logs" / "screenshots").mkdir(parents=True, exist_ok=True)
    (install_dir / "ChromeProfile").mkdir(parents=True, exist_ok=True)

    exe_path = install_dir / f"{APP_NAME}.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"{APP_NAME}.exe was not installed correctly.")

    # 제거기 복사 + Windows '앱 및 기능' 등록 (실패해도 설치는 성공으로 간주).
    register_uninstall(install_dir)

    # 바로가기 생성은 실패해도 설치 자체는 성공으로 간주한다.
    # (파일 복사는 이미 끝났으므로 프로그램은 실행 가능하다.)
    shortcut_warning = None
    if create_shortcut:
        try:
            create_shortcuts(install_dir)
        except Exception as exc:
            shortcut_warning = str(exc)

    return exe_path, shortcut_warning


def get_setup_exe_path():
    """동결(frozen) 상태에서 현재 실행 중인 셋업 EXE 경로를 반환. 소스 실행이면 None."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable)
    return None


def register_uninstall(install_dir):
    """설치 폴더에 제거기(Uninstall.exe)를 두고 레지스트리에 제거 항목을 등록한다.

    - frozen 셋업 EXE를 install_dir/Uninstall.exe 로 복사한다.
    - HKCU\\...\\Uninstall\\NaeilERPUpdater 키에 DisplayName/UninstallString 등을 쓴다.
    모두 실패해도 예외를 올리지 않는다(설치 자체를 막지 않기 위함).
    """
    try:
        setup_exe = get_setup_exe_path()
        uninstaller = None
        if setup_exe and setup_exe.exists():
            uninstaller = install_dir / UNINSTALLER_NAME
            try:
                shutil.copy2(setup_exe, uninstaller)
            except Exception:
                uninstaller = None

        if winreg is None:
            return

        app_exe = install_dir / f"{APP_NAME}.exe"
        if uninstaller is not None:
            uninstall_cmd = f'"{uninstaller}" --uninstall'
            quiet_cmd = f'"{uninstaller}" --uninstall --silent'
        else:
            # 제거기 복사 실패 시: 설치 폴더 경로만 넘겨 자기 자신으로 제거 가능하게.
            uninstall_cmd = f'"{app_exe}" --uninstall'
            quiet_cmd = uninstall_cmd

        key = winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER, UNINSTALL_REG_KEY, 0, winreg.KEY_WRITE
        )
        try:
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, SHORTCUT_NAME)
            winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
            winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(install_dir))
            winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, str(app_exe))
            winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, uninstall_cmd)
            winreg.SetValueEx(key, "QuietUninstallString", 0, winreg.REG_SZ, quiet_cmd)
            winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
        finally:
            winreg.CloseKey(key)
    except Exception:
        # 등록 실패는 무시 (제거 모드 자체는 install_dir 인자로도 동작한다).
        pass


def read_install_location():
    """레지스트리에서 설치 위치를 읽는다. 없으면 None."""
    if winreg is None:
        return None
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_KEY, 0, winreg.KEY_READ)
        try:
            value, _ = winreg.QueryValueEx(key, "InstallLocation")
            if value:
                return Path(value)
        finally:
            winreg.CloseKey(key)
    except FileNotFoundError:
        return None
    except Exception:
        return None
    return None


def remove_uninstall_registry():
    if winreg is None:
        return
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, UNINSTALL_REG_KEY)
    except FileNotFoundError:
        pass
    except Exception:
        pass


def remove_shortcuts():
    """바탕화면 + 시작메뉴 바로가기와 시작메뉴 폴더를 삭제한다."""
    script = f"""
$ErrorActionPreference = "SilentlyContinue"
$Desktop = Join-Path ([Environment]::GetFolderPath("Desktop")) "{SHORTCUT_NAME}.lnk"
if (Test-Path $Desktop) {{ Remove-Item $Desktop -Force }}
$StartMenuDir = Join-Path ([Environment]::GetFolderPath("Programs")) "Naeil Tour"
$StartShortcut = Join-Path $StartMenuDir "{SHORTCUT_NAME}.lnk"
if (Test-Path $StartShortcut) {{ Remove-Item $StartShortcut -Force }}
if (Test-Path $StartMenuDir) {{
    $remaining = Get-ChildItem $StartMenuDir -Force
    if (-not $remaining) {{ Remove-Item $StartMenuDir -Force -Recurse }}
}}
"""
    try:
        with tempfile.TemporaryDirectory(prefix="NaeilERPUpdaterUninstall_") as temp_dir:
            script_path = Path(temp_dir) / "remove_shortcuts.ps1"
            script_path.write_text(script, encoding="utf-8-sig")
            subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script_path),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
    except Exception:
        pass


def _schedule_delete(install_dir):
    """실행 중인 제거기(Uninstall.exe) 자신과 설치 폴더를, 프로세스 종료 후 지운다.

    Windows는 실행 중인 EXE를 삭제할 수 없으므로, 제거기 프로세스(PID)가 종료될 때까지
    대기했다가 폴더를 지우는 분리(detached) 배치를 띄운다. 배치는 끝나면 자신을 삭제한다.
    """
    detached = getattr(subprocess, "DETACHED_PROCESS", 0x00000008)
    no_window = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    pid = os.getpid()
    
    # timeout 대신 ping을 지연시간으로 사용하여 콘솔창 생성을 방지하고,
    # tasklist로 제거기 프로세스가 죽을 때까지 대기(최대 120초)한 후 rmdir을 실행합니다.
    batch = (
        "@echo off\r\n"
        "set /a n=0\r\n"
        ":loop\r\n"
        "set /a n+=1\r\n"
        "ping 127.0.0.1 -n 2 >nul\r\n"
        f"tasklist /fi \"PID eq {pid}\" 2>nul | findstr /i \"{pid}\" >nul\r\n"
        "if %errorlevel% equ 0 (\r\n"
        "    if %n% GEQ 120 goto done\r\n"
        "    goto loop\r\n"
        ")\r\n"
        f'rmdir /s /q "{install_dir}" 2>nul\r\n'
        f'if not exist "{install_dir}" goto done\r\n'
        "if %n% GEQ 120 goto done\r\n"
        "goto loop\r\n"
        ":done\r\n"
        'del "%~f0"\r\n'
    )
    try:
        fd, bat_path = tempfile.mkstemp(prefix="NaeilERPUninstall_", suffix=".bat")
        # .bat는 콘솔 코드페이지로 해석되므로 Windows ANSI(mbcs)로 저장(한글 경로 대응).
        with os.fdopen(fd, "w", encoding="mbcs", errors="replace") as f:
            f.write(batch)
        subprocess.Popen(
            ["cmd", "/c", bat_path],
            creationflags=detached | no_window,
            close_fds=True,
        )
    except Exception:
        pass


def uninstall_app(install_dir):
    """바로가기/레지스트리/설치 폴더를 제거한다.

    설치 폴더 안에서 실행 중(제거기 자신이 그 안에 있음)이면, 폴더 전체 삭제는
    프로세스 종료 후로 예약한다. 반환값: 실제 제거를 수행했으면 True.
    """
    remove_shortcuts()
    remove_uninstall_registry()

    install_dir = Path(install_dir)
    if not install_dir.exists():
        return False

    setup_exe = get_setup_exe_path()
    running_inside = False
    if setup_exe is not None:
        try:
            running_inside = install_dir in setup_exe.parents or setup_exe.parent == install_dir
        except Exception:
            running_inside = False

    if running_inside:
        # 실행 중인 제거기 자신은 못 지우므로, 나머지부터 지우고 폴더는 종료 후 예약.
        for child in install_dir.iterdir():
            if child == setup_exe:
                continue
            try:
                if child.is_dir():
                    shutil.rmtree(child, ignore_errors=True)
                else:
                    child.unlink(missing_ok=True)
            except Exception:
                pass
        _schedule_delete(install_dir)
    else:
        shutil.rmtree(install_dir, ignore_errors=True)

    return True


def show_message(title, message, is_error=False):
    if messagebox and tk:
        root = tk.Tk()
        root.withdraw()
        if is_error:
            messagebox.showerror(title, message)
        else:
            messagebox.showinfo(title, message)
        root.destroy()
    else:
        print(f"{title}: {message}")


def _close_pyi_splash():
    """onefile 부트로더 스플래시가 떠 있으면 닫는다(진행 창으로 전환)."""
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def run_gui(default_dir, create_shortcut, allow_choose=True, mode="auto"):
    """설치/제거를 한 창에서 단계별로 진행한다.

    mode:
      - "install"   : 설치(폴더 선택 → 진행 → 결과). allow_choose=False면 폴더 선택 생략.
      - "uninstall" : 제거(확인 → 진행 → 결과).
      - "auto"      : 이미 설치돼 있으면 모드 선택(다시 설치/제거), 아니면 설치.
    반환값: 오류 없이 끝나면 True.
    """
    BG = "#15161c"
    CARD = "#21232c"
    FG = "#f3f4f6"
    MUTED = "#969aa6"
    GREEN = "#22c55e"
    RED = "#ef4444"
    ACCENT = "#3b82f6"
    BORDER = "#2c303a"

    root = tk.Tk()
    root.title("Naeil ERP Fare Updater 설치")
    root.configure(bg=BG)
    root.resizable(False, False)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(
        "Inst.Horizontal.TProgressbar",
        troughcolor="#0e0f14",
        background=GREEN,
        bordercolor="#0e0f14",
        lightcolor=GREEN,
        darkcolor=GREEN,
        thickness=14,
    )

    tk.Frame(root, bg=ACCENT, height=4).pack(fill="x", side="top")

    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True, padx=26, pady=22)

    state = {"dir": Path(default_dir)}

    def _center():
        root.update_idletasks()
        w = root.winfo_reqwidth()
        h = root.winfo_reqheight()
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        x = max((sw - w) // 2, 0)
        y = max((sh - h) // 3, 0)
        root.geometry(f"{w}x{h}+{x}+{y}")

    def _clear_body():
        for child in body.winfo_children():
            child.destroy()

    def _title(text):
        tk.Label(
            body, text=text, font=("맑은 고딕", 14, "bold"), bg=BG, fg=FG
        ).pack(anchor="w")

    # ---------- 1단계: 설치 폴더 선택 ----------
    def build_chooser():
        _clear_body()
        _title("ERP 항공요금 업데이트 설치")
        tk.Label(
            body,
            text="프로그램을 설치할 폴더를 선택하세요.",
            font=("맑은 고딕", 10),
            bg=BG,
            fg=MUTED,
        ).pack(anchor="w", pady=(10, 14))

        tk.Label(
            body, text="설치 폴더", font=("맑은 고딕", 9, "bold"), bg=BG, fg=MUTED
        ).pack(anchor="w")

        path_var = tk.StringVar(value=str(state["dir"]))
        row = tk.Frame(body, bg=BG)
        row.pack(anchor="w", fill="x", pady=(4, 0))

        entry = tk.Entry(
            row,
            textvariable=path_var,
            font=("맑은 고딕", 10),
            width=42,
            bg=CARD,
            fg=FG,
            insertbackground=FG,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        entry.pack(side="left", ipady=5, fill="x", expand=True)
        entry.xview_moveto(1.0)

        def on_browse():
            initial = state["dir"].parent if state["dir"].parent.exists() else Path.home()
            picked = filedialog.askdirectory(
                title="설치 폴더 선택", initialdir=str(initial), parent=root
            )
            if picked:
                p = Path(picked)
                # 선택한 폴더 안에 항상 앱 전용 폴더를 둔다(엉뚱한 곳에 흩어지지 않도록).
                if p.name.lower() not in ("erpupdater", "naeiltour"):
                    p = p / "ERPUpdater"
                path_var.set(str(p))
                entry.xview_moveto(1.0)

        browse_btn = tk.Button(
            row,
            text="찾아보기",
            command=on_browse,
            font=("맑은 고딕", 9),
            bg=CARD,
            fg=FG,
            activebackground="#2a2d38",
            activeforeground=FG,
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=12,
            pady=6,
        )
        browse_btn.pack(side="left", padx=(8, 0))

        tk.Label(
            body,
            text="선택한 폴더 안에 프로그램 파일이 설치됩니다.",
            font=("맑은 고딕", 9),
            bg=BG,
            fg=MUTED,
        ).pack(anchor="w", pady=(8, 0))

        btn_row = tk.Frame(body, bg=BG)
        btn_row.pack(anchor="e", pady=(20, 0))

        def on_install():
            chosen = path_var.get().strip()
            if not chosen:
                return
            state["dir"] = Path(chosen)
            build_progress()

        cancel_btn = tk.Button(
            btn_row,
            text="취소",
            command=root.destroy,
            font=("맑은 고딕", 10),
            bg=CARD,
            fg=FG,
            activebackground="#2a2d38",
            activeforeground=FG,
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=14,
            pady=6,
        )
        cancel_btn.pack(side="right", padx=(8, 0))

        install_btn = tk.Button(
            btn_row,
            text="설치",
            command=on_install,
            font=("맑은 고딕", 10, "bold"),
            bg=ACCENT,
            fg="white",
            activebackground="#2f6fe0",
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            bd=0,
            padx=22,
            pady=6,
        )
        install_btn.pack(side="right")

        _center()

    # ---------- 2단계: 설치 진행 ----------
    def build_progress():
        _clear_body()
        _title("ERP 항공요금 업데이트 설치")

        status_lbl = tk.Label(
            body,
            text="파일을 설치하고 있습니다…",
            font=("맑은 고딕", 10),
            bg=BG,
            fg=MUTED,
            justify="left",
            wraplength=420,
        )
        status_lbl.pack(anchor="w", pady=(12, 16))

        pb = ttk.Progressbar(
            body,
            mode="indeterminate",
            style="Inst.Horizontal.TProgressbar",
            length=420,
        )
        pb.pack(anchor="w")
        pb.start(12)
        _center()

        install_dir = state["dir"]

        def worker():
            try:
                exe_path, warning = install_payload(
                    install_dir, create_shortcut=create_shortcut
                )
                state["exe"] = exe_path
                state["warning"] = warning
            except Exception as exc:
                state["error"] = exc

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

        def poll():
            if thread.is_alive():
                root.after(120, poll)
            else:
                pb.stop()
                pb.destroy()
                show_result(status_lbl, install_dir)

        poll()

    # ---------- 3단계: 결과 ----------
    def show_result(status_lbl, install_dir):
        if "error" in state:
            status_lbl.config(
                text=(
                    "설치 중 오류가 발생했습니다.\n\n"
                    f"{state['error']}\n\n"
                    "프로그램이 실행 중이라면 종료한 뒤 다시 설치해 주세요."
                ),
                fg=RED,
            )
        elif state.get("warning"):
            status_lbl.config(
                text=(
                    "설치가 완료되었습니다.\n"
                    "다만 바탕화면 바로가기 생성에는 실패했어요.\n"
                    "아래 위치의 실행 파일을 직접 실행해 주세요.\n\n"
                    f"{state.get('exe')}"
                ),
                fg=FG,
            )
        else:
            status_lbl.config(
                text=(
                    "설치가 완료되었습니다.\n"
                    "바탕화면 바로가기에서 실행해 주세요.\n\n"
                    f"설치 위치: {install_dir}"
                ),
                fg=FG,
            )

        btn = tk.Button(
            body,
            text="확인",
            command=root.destroy,
            font=("맑은 고딕", 10, "bold"),
            bg=ACCENT,
            fg="white",
            activebackground="#2f6fe0",
            activeforeground="white",
            relief="flat",
            cursor="hand2",
            width=10,
            bd=0,
            padx=8,
            pady=6,
        )
        btn.pack(anchor="e", pady=(18, 0))
        _center()

    # ---------- 모드 선택(이미 설치된 경우) ----------
    def build_mode_select():
        _clear_body()
        _title("ERP 항공요금 업데이트")
        tk.Label(
            body,
            text=(
                "이미 설치되어 있습니다.\n"
                f"설치 위치: {state['dir']}"
            ),
            font=("맑은 고딕", 10),
            bg=BG,
            fg=MUTED,
            justify="left",
            wraplength=420,
        ).pack(anchor="w", pady=(10, 18))

        btn_row = tk.Frame(body, bg=BG)
        btn_row.pack(anchor="e")

        tk.Button(
            btn_row, text="취소", command=root.destroy,
            font=("맑은 고딕", 10), bg=CARD, fg=FG,
            activebackground="#2a2d38", activeforeground=FG,
            relief="flat", cursor="hand2", bd=0, padx=14, pady=6,
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            btn_row, text="제거", command=build_uninstall_confirm,
            font=("맑은 고딕", 10, "bold"), bg="#ef4444", fg="white",
            activebackground="#dc2626", activeforeground="white",
            relief="flat", cursor="hand2", bd=0, padx=18, pady=6,
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            btn_row, text="다시 설치", command=build_chooser,
            font=("맑은 고딕", 10, "bold"), bg=ACCENT, fg="white",
            activebackground="#2f6fe0", activeforeground="white",
            relief="flat", cursor="hand2", bd=0, padx=18, pady=6,
        ).pack(side="right")

        _center()

    # ---------- 제거 확인 ----------
    def build_uninstall_confirm():
        _clear_body()
        _title("ERP 항공요금 업데이트 제거")
        tk.Label(
            body,
            text=(
                "프로그램을 제거하시겠습니까?\n"
                "바탕화면·시작메뉴 바로가기와 설치 폴더가 삭제됩니다.\n\n"
                f"설치 위치: {state['dir']}"
            ),
            font=("맑은 고딕", 10),
            bg=BG,
            fg=MUTED,
            justify="left",
            wraplength=420,
        ).pack(anchor="w", pady=(10, 18))

        btn_row = tk.Frame(body, bg=BG)
        btn_row.pack(anchor="e")

        tk.Button(
            btn_row, text="취소", command=root.destroy,
            font=("맑은 고딕", 10), bg=CARD, fg=FG,
            activebackground="#2a2d38", activeforeground=FG,
            relief="flat", cursor="hand2", bd=0, padx=14, pady=6,
        ).pack(side="right", padx=(8, 0))

        tk.Button(
            btn_row, text="제거", command=build_uninstall_progress,
            font=("맑은 고딕", 10, "bold"), bg="#ef4444", fg="white",
            activebackground="#dc2626", activeforeground="white",
            relief="flat", cursor="hand2", bd=0, padx=22, pady=6,
        ).pack(side="right")

        _center()

    # ---------- 제거 진행 ----------
    def build_uninstall_progress():
        _clear_body()
        _title("ERP 항공요금 업데이트 제거")

        status_lbl = tk.Label(
            body,
            text="프로그램을 제거하고 있습니다…",
            font=("맑은 고딕", 10),
            bg=BG, fg=MUTED, justify="left", wraplength=420,
        )
        status_lbl.pack(anchor="w", pady=(12, 16))

        pb = ttk.Progressbar(
            body, mode="indeterminate",
            style="Inst.Horizontal.TProgressbar", length=420,
        )
        pb.pack(anchor="w")
        pb.start(12)
        _center()

        install_dir = state["dir"]

        def worker():
            try:
                uninstall_app(install_dir)
            except Exception as exc:
                state["error"] = exc

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

        def poll():
            if thread.is_alive():
                root.after(120, poll)
            else:
                pb.stop()
                pb.destroy()
                show_uninstall_result(status_lbl)

        poll()

    def show_uninstall_result(status_lbl):
        if "error" in state:
            status_lbl.config(
                text=(
                    "제거 중 오류가 발생했습니다.\n\n"
                    f"{state['error']}\n\n"
                    "프로그램이 실행 중이라면 종료한 뒤 다시 시도해 주세요."
                ),
                fg=RED,
            )
        else:
            status_lbl.config(text="제거가 완료되었습니다.", fg=FG)

        tk.Button(
            body, text="확인", command=root.destroy,
            font=("맑은 고딕", 10, "bold"), bg=ACCENT, fg="white",
            activebackground="#2f6fe0", activeforeground="white",
            relief="flat", cursor="hand2", width=10, bd=0, padx=8, pady=6,
        ).pack(anchor="e", pady=(18, 0))
        _center()

    _close_pyi_splash()
    if mode == "uninstall":
        root.title("Naeil ERP Fare Updater 제거")
        build_uninstall_confirm()
    elif mode == "install":
        if allow_choose and filedialog is not None:
            build_chooser()
        else:
            build_progress()
    else:  # auto
        installed = (state["dir"] / f"{APP_NAME}.exe").exists()
        if installed:
            build_mode_select()
        elif allow_choose and filedialog is not None:
            build_chooser()
        else:
            build_progress()
    root.mainloop()

    return "error" not in state


def parse_args():
    parser = argparse.ArgumentParser(description="Install/Uninstall Naeil ERP Fare Updater.")
    parser.add_argument("--silent", action="store_true", help="Do not show message boxes.")
    parser.add_argument("--install-dir", default=None, help="Override install directory.")
    parser.add_argument("--no-shortcuts", action="store_true", help="Skip desktop/start menu shortcuts.")
    parser.add_argument("--uninstall", action="store_true", help="Uninstall instead of install.")
    return parser.parse_args()


def _resolve_uninstall_dir(args):
    """제거 대상 설치 폴더를 결정: --install-dir > 레지스트리 > 기본값."""
    if args.install_dir:
        return Path(args.install_dir)
    reg_dir = read_install_location()
    if reg_dir:
        return reg_dir
    return default_install_dir()


def _run_uninstall_text(install_dir):
    """silent/GUI 불가 환경의 제거 처리(텍스트)."""
    _close_pyi_splash()
    try:
        uninstall_app(install_dir)
    except Exception as exc:
        print(f"제거 중 오류가 발생했습니다.\n{exc}")
        return 1
    print(f"제거가 완료되었습니다.\n설치 위치: {install_dir}")
    return 0


def _run_install_text(install_dir, create_shortcut):
    """silent/GUI 불가 환경의 설치 처리(텍스트)."""
    _close_pyi_splash()
    try:
        exe_path, shortcut_warning = install_payload(
            install_dir, create_shortcut=create_shortcut
        )
    except Exception as exc:
        message = (
            "설치 중 오류가 발생했습니다.\n\n"
            f"{exc}\n\n"
            "프로그램이 실행 중이라면 종료한 뒤 다시 설치해 주세요."
        )
        print(message)
        return 1

    if shortcut_warning:
        message = (
            "설치가 완료되었습니다.\n\n"
            f"설치 위치: {install_dir}\n\n"
            "다만 바탕화면 바로가기 생성에는 실패했습니다.\n"
            f"실행 파일: {exe_path}\n"
            f"[참고] {shortcut_warning}"
        )
    else:
        message = (
            "설치가 완료되었습니다.\n\n"
            f"설치 위치: {install_dir}\n"
            "바탕화면 바로가기에서 실행해 주세요."
        )
    print(message)
    return 0


def main():
    args = parse_args()
    gui_available = (not args.silent) and tk is not None and ttk is not None

    # ----- 제거 모드 -----
    if args.uninstall:
        install_dir = _resolve_uninstall_dir(args)
        if gui_available:
            ok = run_gui(install_dir, create_shortcut=False, mode="uninstall")
            return 0 if ok else 1
        return _run_uninstall_text(install_dir)

    # ----- 설치 모드 -----
    install_dir = Path(args.install_dir) if args.install_dir else default_install_dir()
    # 사용자 지정이 없고 이미 설치돼 있으면(레지스트리) 그 위치를 기준으로 한다.
    if not args.install_dir:
        existing = read_install_location()
        if existing and (existing / f"{APP_NAME}.exe").exists():
            install_dir = existing
    if gui_available:
        ok = run_gui(
            install_dir,
            create_shortcut=not args.no_shortcuts,
            allow_choose=args.install_dir is None,
            mode="auto",
        )
        return 0 if ok else 1

    return _run_install_text(install_dir, create_shortcut=not args.no_shortcuts)


if __name__ == "__main__":
    raise SystemExit(main())
