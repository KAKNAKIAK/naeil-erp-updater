# -*- coding: utf-8 -*-
"""
Naeil Tour ERP 요금 업데이트 RPA — v4.0.0

주요 기능:
  - 앱에 내장된 스프레드시트(셀) 그리드에 요금 직접 입력/수정.
  - 엑셀에서 복사한 표를 셀에 그대로 붙여넣기(Ctrl+V) 가능.
  - 기존 엑셀 파일은 '요금불러오기' 버튼으로 그리드에 가져오기 가능.
  - ERP 요금수정 탭과 TOPAS AC1 연속 조회 탭을 분리.
"""
import os
import sys
import json
import time
import csv
import ast
import operator
import threading
import queue
import shutil
import subprocess
import tempfile
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.chrome.options
import selenium.webdriver.chrome.service
import selenium.webdriver.chrome.webdriver
from webdriver_manager.chrome import ChromeDriverManager
from tksheet import Sheet
import pandas as pd
import excel_loader
import update_client
from topas.availability import parse_availability_text

APP_VERSION = "v4.0.0"
UPDATER_EXE_NAME = "UpdateHelper.exe"

# 그리드 컬럼 정의
COL_DATE = 0
COL_ADULT_AIR = 1
COL_ADULT_HOTEL = 2
COL_ADULT_LAND = 3
COL_ADULT_TOUR = 4
COL_ADULT_PROFIT = 5
COL_CHILD = 6
COL_INFANT = 7

SHEET_HEADERS = ["날짜", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아요금", "유아요금"]
SHORT_COL_NAMES = ["날짜", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"]
INITIAL_BLANK_ROWS = 12


def get_app_dir():
    """실행 파일 또는 소스 파일 옆의 쓰기 가능한 폴더 경로를 반환합니다."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(rel_path):
    """번들된 리소스(이미지/GIF 등)의 실제 경로를 반환한다."""
    base = getattr(sys, '_MEIPASS', None)
    if base:
        cand = os.path.join(base, rel_path)
        if os.path.exists(cand):
            return cand
    return os.path.join(get_app_dir(), rel_path)


def normalize_date(raw):
    """다양한 표기의 날짜 문자열을 YYYY-MM-DD로 정규화한다.
    실패하면 None을 반환한다.
    허용: 2026-06-07 / 2026.6.7 / 2026/06/07 / 20260607 / '2026-06-07 00:00:00'"""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    # 'YYYY-MM-DD HH:MM:SS' 같은 경우 앞부분만
    s = s.split(' ')[0].split('T')[0]
    t = s.replace('/', '-').replace('.', '-').strip('-')
    if len(t) == 8 and t.isdigit():
        return f"{t[:4]}-{t[4:6]}-{t[6:]}"
    parts = t.split('-')
    if len(parts) == 3:
        try:
            y, m, d = parts
            if len(y) == 4:
                return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
        except (ValueError, TypeError):
            return None
    return None


# 셀에서 허용할 사칙연산(+ - * / 와 거듭제곱/나머지, 단항 부호)
_ALLOWED_OPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Mod: operator.mod,
    ast.Pow: operator.pow,
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
}


def _eval_ast_node(node, resolver):
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return node.value
    if isinstance(node, ast.BinOp) and type(node.op) in _ALLOWED_OPS:
        return _ALLOWED_OPS[type(node.op)](_eval_ast_node(node.left, resolver), _eval_ast_node(node.right, resolver))
    if isinstance(node, ast.UnaryOp) and type(node.op) in _ALLOWED_OPS:
        return _ALLOWED_OPS[type(node.op)](_eval_ast_node(node.operand, resolver))
    if isinstance(node, ast.Name):
        # 칸 참조(앞의열/항공비/알선수익/A·B·C 등) → resolver가 숫자로 풀어준다
        if resolver is not None:
            val = resolver(node.id)
            if val is not None:
                return val
        raise ValueError(f"알 수 없는 참조: {node.id}")
    raise ValueError("허용되지 않은 식")


def eval_arithmetic(expr, resolver=None):
    """'='로 시작하는 사칙연산 식을 안전하게 계산한다(함수 호출 차단).
    resolver(name)가 주어지면 식 안의 칸 참조(앞의열/항공비 등)를 숫자로 풀어준다.
    실패하면 None을 반환한다. 예: '=320000+32000' -> 352000, '=앞의열*0.1'"""
    s = str(expr).strip()
    if s.startswith('='):
        s = s[1:]
    s = s.replace(',', '').strip()
    if not s:
        return None
    try:
        return _eval_ast_node(ast.parse(s, mode='eval').body, resolver)
    except Exception:
        return None


class NullStream:
    encoding = 'utf-8'
    def write(self, _string):
        pass
    def flush(self):
        pass


class GUIConsoleRedirector:
    def __init__(self, text_widget, root):
        self.text_widget = text_widget
        self.root = root
        self.main_thread_id = threading.get_ident()
        self._queue = queue.Queue()
        self._closed = False
        self._poll_after_id = self.root.after(50, self._drain_queue)

    def write(self, string):
        if not string or self._closed:
            return
        if threading.get_ident() == self.main_thread_id:
            self._append_text(string)
        else:
            self._queue.put(string)

    def flush(self):
        pass

    def close(self):
        self._drain_queue_once()
        self._closed = True
        if self._poll_after_id is not None:
            try:
                self.root.after_cancel(self._poll_after_id)
            except Exception:
                pass
            self._poll_after_id = None

    def _append_text(self, string):
        try:
            if self.text_widget.winfo_exists():
                self.text_widget.insert(tk.END, string)
                self.text_widget.see(tk.END)
        except Exception:
            pass

    def _drain_queue_once(self):
        chunks = []
        while True:
            try:
                chunks.append(self._queue.get_nowait())
            except queue.Empty:
                break
        if chunks:
            self._append_text(''.join(chunks))

    def _drain_queue(self):
        if self._closed:
            return
        self._drain_queue_once()
        try:
            self._poll_after_id = self.root.after(50, self._drain_queue)
        except Exception:
            self._closed = True


class RpaGuiApp:
    FARE_SITE_URL = 'https://fare-calculator-2026.web.app/'
    GUIDE_SITE_URL = 'https://kaknakiak.github.io/naeil-erp-updater/사용가이드.html'
    ERP_LOGIN_URL = 'https://erp.naeiltour.co.kr/erp/login'
    TOPAS_LOGIN_URL = 'https://www.topassellconnect.com/login'
    TOPAS_PROMPT_INPUT = '#cryptics1_cmd_shellbridge_shellWindow_top_left_modeString_cmdPromptInput'
    TOPAS_SHELL_ROOT = '#cryptics1_cmd_shellbridge_shellWindow_top_left'

    def __init__(self, root):
        self.root = root
        self.root.title(f'Naeil Tour ERP 요금 업데이트 RPA ({APP_VERSION})')

        # 오른쪽 '요금 입력표' 펼침 패널 크기 (기본 접힘 상태로 시작)
        self.win_height = 820
        self.panel_width = 750
        self.collapsed_width = 860
        self.expanded_width = self.collapsed_width + self.panel_width
        self.panel_expanded = False
        self.root.geometry(f'{self.collapsed_width}x{self.win_height}')
        self.root.minsize(self.collapsed_width, 700)

        # --- 색상 팔레트 (다크 테마 정돈) ---
        self.bg_color = '#15161c'
        self.card_color = '#21232c'
        self.card_hover = '#2a2d38'
        self.fg_color = '#f3f4f6'
        self.fg_muted = '#969aa6'
        self.accent_color = '#3b82f6'
        self.accent_hover = '#2f6fe0'
        self.accent_green = '#22c55e'
        self.accent_green_hover = '#1ba34d'
        self.accent_orange = '#f59e0b'
        self.accent_orange_hover = '#d98806'
        self.accent_red = '#ef4444'
        self.accent_red_hover = '#dc2626'
        self.border_color = '#2c303a'

        self.root.configure(bg=self.bg_color)
        self._setup_styles()

        self.config = self.load_config()
        self.fares_data = []
        self.rpa_thread = None
        self.is_running = False
        self.is_paused = False
        self.is_user_stopped = False
        self.driver = None
        self.console_redirector = None
        self.toolbar_buttons = []
        self.topas_thread = None
        self.topas_results_raw = []
        self.topas_stop_requested = False
        self._topas_shell_el = None
        self._topas_prompt_el = None
        self.current_main_tab = 'fare'
        self.main_tab_controls = {}

        # 수식 입력줄(formula bar) / 클릭 참조 상태
        self.fb_var = tk.StringVar()
        self._active_cell = (0, 0)
        self._ref_mode = False
        self._loading_fb = False
        # 수식 보관소: 셀에는 계산 결과를 표시하고, 식은 여기에 (row,col)->'=...' 로 보관
        self.formulas = {}
        self._results = {}        # (row,col)->마지막 계산 결과 문자열 (수동 수정 감지용)
        self._recalc_busy = False  # 재계산 중 재진입 방지

        # 날짜 필터 관리 변수
        self.filter_mode = tk.StringVar(value="ALL")  # ALL, FROM_DATE, SPECIFIC, DATE_RANGE
        self.filter_value = tk.StringVar(value="")
        self.filter_value_end = tk.StringVar(value="")
        self.filter_mode.trace_add("write", self._on_filter_mode_change)
        self.filter_value.trace_add("write", self._on_filter_value_change)
        self.filter_value_end.trace_add("write", self._on_filter_value_change)

        # 로딩 애니메이션 상태
        self._loading_frames = []
        self._loading_idx = 0
        self._loading_after_id = None
        self._loading_overlay = None

        self.build_ui()
        self._on_filter_mode_change()
        self._load_loading_frames()
        self.refresh_count()
        self.set_status("오른쪽 위 ‘요금직접입력하기 ▶’ 버튼으로 요금을 입력하세요", self.fg_muted)
        self.root.after(800, self.check_for_updates_on_startup)

    # ------------------------------------------------------------------
    # 설정 로드/저장
    # ------------------------------------------------------------------
    def load_config(self):
        base_dir = get_app_dir()
        config_path = os.path.join(base_dir, 'config.json')
        default_selectors = {
            'login_id': "input[name='userId']",
            'login_pw': "input[name='userPw']",
            'login_btn': '#btnLogin',
            'search_date_input': '#searchStDate',
            'search_date_end_input': '#searchEnDate',
            'search_button': '#gridMain_r',
            'header_all_checkbox': 'td.aui-grid-row-check-header input',
            'update_button': '#priceUpdate',
            'adult_air_input': '#addAir01',
            'adult_hotel_input': '#addHotel11',
            'adult_land_input': '#addLand21',
            'adult_tour_input': '#addExpense40',
            'adult_profit_input': '#addProfit41',
            'child_fare_input': '#addChild90',
            'infant_fare_input': '#addInfant91',
            'save_button': '#priceSave',
            'cancel_button': '#popCloseBtn',
            'date_cell_in_row': '.aui-grid-default-column'
        }
        default_config = {
            'headless': False,
            'dry_run': False,
            'timeout': 25,
            'login_timeout': 180,
            'use_existing_browser': True,
            'debugger_address': '127.0.0.1:9222',
            'erp_url': 'http://erp.naeiltour.co.kr',
            'history_log_path': 'logs/update_history.csv',
            'screenshot_dir': 'logs/screenshots',
            'update_enabled': True,
            'update_latest_url': 'https://api.github.com/repos/KAKNAKIAK/naeil-erp-updater/contents/latest.json?ref=main',
            'update_check_timeout': 8,
            'update_download_timeout': 30,
            'erp_poll_interval': 0.25,
            'erp_short_pause': 0.2,
            'erp_grid_ready_stable_seconds': 1.0,
            'topas_poll_interval': 0.04,
            'topas_response_stable_wait': 0.2,
            'topas_batch_timeout': 120,
            'topas_parse_tail_chars': 80000,
            'topas_batch_size': 10,
            'selectors': default_selectors
        }
        if not os.path.exists(config_path):
            return default_config
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                user_conf = json.load(f)
            for k, v in default_config.items():
                if k not in user_conf:
                    user_conf[k] = v
            user_selectors = user_conf.get('selectors')
            if not isinstance(user_selectors, dict):
                user_selectors = {}
            merged_selectors = dict(default_selectors)
            for key, value in user_selectors.items():
                if key == 'adult_fare_input':
                    if 'adult_air_input' not in user_selectors:
                        merged_selectors['adult_air_input'] = value
                    continue
                if key == 'child_fare_input' and value == '#addProfit41' and 'adult_profit_input' not in user_selectors:
                    merged_selectors['adult_profit_input'] = value
                    continue
                merged_selectors[key] = value
            user_conf['selectors'] = merged_selectors
            return user_conf
        except Exception as e:
            print(f"[경고] 설정 로드 실패, 기본값 사용: {str(e)}")
            return default_config

    def _float_config(self, key, default, min_value=None, max_value=None):
        try:
            value = float(self.config.get(key, default))
        except (TypeError, ValueError):
            value = float(default)
        if min_value is not None:
            value = max(float(min_value), value)
        if max_value is not None:
            value = min(float(max_value), value)
        return value

    def _int_config(self, key, default, min_value=None, max_value=None):
        try:
            value = int(self.config.get(key, default))
        except (TypeError, ValueError):
            value = int(default)
        if min_value is not None:
            value = max(int(min_value), value)
        if max_value is not None:
            value = min(int(max_value), value)
        return value

    def check_for_updates_on_startup(self):
        if not self.config.get('update_enabled', True):
            return

        latest_url = (
            os.environ.get('NAEIL_ERP_UPDATE_LATEST_URL')
            or self.config.get('update_latest_url', '')
        ).strip()
        if not latest_url:
            return

        timeout = int(self.config.get('update_check_timeout', 8))
        thread = threading.Thread(
            target=self._check_for_updates_worker,
            args=(latest_url, timeout),
            daemon=True,
        )
        thread.start()

    def _check_for_updates_worker(self, latest_url, timeout):
        try:
            manifest = update_client.fetch_available_update(
                latest_url,
                APP_VERSION,
                timeout=timeout,
            )
        except Exception as exc:
            print(f"[업데이트 확인] 건너뜀: {exc}")
            return

        if manifest:
            self.root.after(0, lambda: self.prompt_for_update(manifest))

    def prompt_for_update(self, manifest):
        if self.is_running:
            return

        latest_version = manifest.get('version', '')
        notes = update_client.format_release_notes(manifest)
        message = (
            "새 버전이 있습니다.\n\n"
            f"현재 버전: {APP_VERSION}\n"
            f"새 버전: {latest_version}\n\n"
            "지금 업데이트를 설치하시겠습니까?"
        )
        if notes:
            message += f"\n\n[변경 내용]\n{notes}"

        if messagebox.askyesno("업데이트 확인", message):
            self.download_and_apply_update(manifest)

    def download_and_apply_update(self, manifest):
        if self.is_running:
            messagebox.showwarning(
                "업데이트 대기",
                "요금 수정 작업이 진행 중입니다. 작업을 마친 뒤 다시 실행해 주세요.",
            )
            return

        self.show_loading('업데이트 설치본을 다운로드하고 있어요...')
        self.set_status('업데이트 다운로드 중', self.accent_orange)

        thread = threading.Thread(
            target=self._download_update_worker,
            args=(manifest,),
            daemon=True,
        )
        thread.start()

    def _download_update_worker(self, manifest):
        timeout = int(self.config.get('update_download_timeout', 30))

        def progress(downloaded, total):
            if total:
                pct = int((downloaded / total) * 100)
                self.root.after(
                    0,
                    lambda pct=pct: self.set_status(
                        f'업데이트 다운로드 중 ({pct}%)',
                        self.accent_orange,
                    ),
                )

        try:
            installer_path = update_client.download_installer(
                manifest,
                timeout=timeout,
                progress_callback=progress,
            )
        except Exception as exc:
            self.root.after(0, lambda exc=exc: self._show_update_error(exc))
            return

        self.root.after(
            0,
            lambda: self.launch_updater_and_exit(installer_path, manifest),
        )

    def _show_update_error(self, exc):
        self.hide_loading()
        self.set_status('업데이트 실패', self.accent_red)
        messagebox.showerror("업데이트 실패", str(exc))

    def launch_updater_and_exit(self, installer_path, manifest):
        base_dir = get_app_dir()
        updater_src = os.path.join(base_dir, UPDATER_EXE_NAME)
        if not os.path.exists(updater_src):
            self._show_update_error(
                FileNotFoundError(f"{UPDATER_EXE_NAME} 파일을 찾을 수 없습니다.")
            )
            return

        try:
            temp_dir = tempfile.mkdtemp(prefix='NaeilERPUpdaterHelper_')
            updater_dst = os.path.join(temp_dir, UPDATER_EXE_NAME)
            shutil.copy2(updater_src, updater_dst)

            app_exe = (
                sys.executable
                if getattr(sys, 'frozen', False)
                else os.path.join(base_dir, 'NaeilERPUpdater.exe')
            )
            args = [
                updater_dst,
                '--installer',
                str(installer_path),
                '--sha256',
                manifest.get('sha256', ''),
                '--install-dir',
                base_dir,
                '--app-exe',
                app_exe,
                '--wait-pid',
                str(os.getpid()),
            ]
            creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0)
            subprocess.Popen(args, cwd=temp_dir, creationflags=creationflags)
        except Exception as exc:
            self._show_update_error(exc)
            return

        self.hide_loading()
        self.is_running = False
        self.root.destroy()

    # ------------------------------------------------------------------
    # 스타일 / 공통 UI 헬퍼
    # ------------------------------------------------------------------
    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass
        style.configure(
            'Accent.Horizontal.TProgressbar',
            troughcolor='#0e0f14',
            bordercolor='#0e0f14',
            background=self.accent_green,
            lightcolor=self.accent_green,
            darkcolor=self.accent_green,
            thickness=18,
        )
        self._style = style

    def _add_hover(self, widget, normal_bg, hover_bg, normal_fg=None, hover_fg=None):
        def on_enter(_e):
            if str(widget['state']) != 'disabled':
                widget.config(bg=hover_bg)
                if hover_fg is not None:
                    widget.config(fg=hover_fg)
        def on_leave(_e):
            if str(widget['state']) != 'disabled':
                widget.config(bg=normal_bg)
                if normal_fg is not None:
                    widget.config(fg=normal_fg)
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)

    def set_status(self, text, color=None):
        color = color or self.fg_color
        if hasattr(self, 'status_dot'):
            self.status_dot.config(fg=color)
        self.status_lbl.config(text=text, fg=color)

    # ------------------------------------------------------------------
    # 로딩 애니메이션 (캐릭터 GIF)
    # ------------------------------------------------------------------
    def _load_loading_frames(self, size=140):
        try:
            from PIL import Image, ImageTk
            gif_path = get_resource_path(os.path.join('assets', 'loading.gif'))
            if not os.path.exists(gif_path):
                return
            im = Image.open(gif_path)
            frames = []
            durations = []
            idx = 0
            while True:
                try:
                    im.seek(idx)
                except EOFError:
                    break
                durations.append(im.info.get('duration', 70) or 70)
                frame = im.convert('RGBA').resize((size, size), Image.LANCZOS)
                frames.append(ImageTk.PhotoImage(frame))
                idx += 1
            if frames:
                self._loading_frames = frames
                self._loading_durations = durations
        except Exception as e:
            print(f"[안내] 로딩 애니메이션 비활성화: {e}")

    def show_loading(self, text='잠시만 기다려 주세요…'):
        if not self._loading_frames:
            self.set_status(text, self.accent_orange)
            return
        if self._loading_overlay is not None:
            try:
                self._loading_text_lbl.config(text=text)
            except Exception:
                pass
            return
        overlay = tk.Frame(self.root, bg=self.bg_color)
        overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        inner = tk.Frame(overlay, bg=self.bg_color)
        inner.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        self._loading_img_lbl = tk.Label(inner, bg=self.bg_color, image=self._loading_frames[0])
        self._loading_img_lbl.pack()
        self._loading_text_lbl = tk.Label(inner, text=text, font=('맑은 고딕', 12, 'bold'), bg=self.bg_color, fg=self.fg_color)
        self._loading_text_lbl.pack(pady=(10, 0))
        self._loading_overlay = overlay
        self._loading_idx = 0
        self._animate_loading()

    def _animate_loading(self):
        if self._loading_overlay is None or not self._loading_frames:
            return
        frame = self._loading_frames[self._loading_idx]
        try:
            self._loading_img_lbl.config(image=frame)
        except Exception:
            return
        delay = 70
        try:
            delay = self._loading_durations[self._loading_idx]
        except Exception:
            pass
        self._loading_idx = (self._loading_idx + 1) % len(self._loading_frames)
        self._loading_after_id = self.root.after(delay, self._animate_loading)

    def hide_loading(self):
        if self._loading_after_id is not None:
            try:
                self.root.after_cancel(self._loading_after_id)
            except Exception:
                pass
            self._loading_after_id = None
        if self._loading_overlay is not None:
            try:
                self._loading_overlay.destroy()
            except Exception:
                pass
            self._loading_overlay = None

    def open_fare_site(self):
        try:
            import webbrowser
            webbrowser.open(self.FARE_SITE_URL)
        except Exception as e:
            messagebox.showerror('사이트 열기 실패', f'요금조회 사이트를 열지 못했습니다.\n{e}')

    def open_guide_site(self):
        try:
            import webbrowser
            webbrowser.open(self.GUIDE_SITE_URL)
        except Exception as e:
            messagebox.showerror('사이트 열기 실패', f'가이드 매뉴얼을 열지 못했습니다.\n{e}')


    # ------------------------------------------------------------------
    # UI 빌드
    # ------------------------------------------------------------------
    def build_ui(self):
        # 좌우 분할: 왼쪽 메인 컬럼 + 오른쪽 펼침 패널(요금 입력표)
        body = tk.Frame(self.root, bg=self.bg_color)
        body.pack(fill=tk.BOTH, expand=True)

        # 오른쪽 펼침 패널 (기본은 접힌 상태 — pack은 toggle에서)
        self.side_panel = tk.Frame(body, bg=self.bg_color, width=self.panel_width)
        self.side_panel.pack_propagate(False)

        # 왼쪽 메인 컬럼
        main_col = tk.Frame(body, bg=self.bg_color)
        main_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 1. 헤더
        header_frame = tk.Frame(main_col, bg=self.bg_color, pady=14)
        header_frame.pack(fill=tk.X, padx=24)

        self.chrome_launch_btn = tk.Button(
            header_frame, text='브라우저 켜기', width=14, height=2,
            bg=self.accent_color, fg='white', font=('맑은 고딕', 10, 'bold'),
            activebackground=self.accent_hover, activeforeground='white',
            bd=0, relief=tk.FLAT, cursor='hand2', command=self.launch_debug_chrome)
        self.chrome_launch_btn.pack(side=tk.RIGHT, anchor=tk.CENTER)
        self._add_hover(self.chrome_launch_btn, self.accent_color, self.accent_hover)

        title_frame = tk.Frame(header_frame, bg=self.bg_color)
        title_frame.pack(side=tk.LEFT, anchor=tk.CENTER)

        title_row = tk.Frame(title_frame, bg=self.bg_color)
        title_row.pack(anchor=tk.W)
        title_label = tk.Label(title_row, text='Naeil Tour ERP 요금 업데이트', font=('맑은 고딕', 17, 'bold'), bg=self.bg_color, fg=self.fg_color)
        title_label.pack(side=tk.LEFT)
        version_badge = tk.Label(title_row, text=f' {APP_VERSION} ', font=('맑은 고딕', 8, 'bold'), bg=self.card_color, fg=self.accent_orange)
        version_badge.pack(side=tk.LEFT, padx=(8, 0), pady=(4, 0))

        fare_site_btn = tk.Button(
            title_row, text='요금조회 사이트가기 ↗', font=('맑은 고딕', 8, 'bold'),
            bg=self.card_color, fg=self.accent_color,
            activebackground=self.accent_color, activeforeground='white',
            bd=0, relief=tk.FLAT, cursor='hand2', padx=8, pady=2,
            command=self.open_fare_site)
        fare_site_btn.pack(side=tk.LEFT, padx=(10, 0), pady=(4, 0))
        self._add_hover(fare_site_btn, self.card_color, self.accent_color, normal_fg=self.accent_color, hover_fg='white')

        guide_btn = tk.Button(
            title_row, text='가이드 매뉴얼 보기 ↗', font=('맑은 고딕', 8, 'bold'),
            bg=self.card_color, fg=self.accent_orange,
            activebackground=self.accent_orange, activeforeground='white',
            bd=0, relief=tk.FLAT, cursor='hand2', padx=8, pady=2,
            command=self.open_guide_site)
        guide_btn.pack(side=tk.LEFT, padx=(10, 0), pady=(4, 0))
        self._add_hover(guide_btn, self.card_color, self.accent_orange, normal_fg=self.accent_orange, hover_fg='white')



        template_btn = tk.Button(
            title_frame, text='엑셀 양식 다운받기 ↗', font=('맑은 고딕', 8, 'bold'),
            bg=self.card_color, fg=self.accent_orange,
            activebackground=self.accent_orange, activeforeground='white',
            bd=0, relief=tk.FLAT, cursor='hand2', padx=8, pady=2,
            command=self.download_excel_template)
        template_btn.pack(anchor=tk.W, pady=(6, 0))
        self._add_hover(template_btn, self.card_color, self.accent_orange, normal_fg=self.accent_orange, hover_fg='white')

        self._build_main_tab_bar(main_col)
        self.main_content = tk.Frame(main_col, bg=self.bg_color)
        self.main_content.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 0))
        self.fare_tab = tk.Frame(self.main_content, bg=self.bg_color)
        self.topas_tab = tk.Frame(self.main_content, bg=self.bg_color)

        # 2. 요금 입력표 카드 (셀 그리드) — 오른쪽 펼침 패널 안에 배치
        sheet_card = tk.LabelFrame(self.side_panel, text=' 요금 입력표 ', font=('맑은 고딕', 9, 'bold'), bg=self.card_color, fg=self.fg_muted, bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightthickness=1, padx=10, pady=8)
        sheet_card.pack(fill=tk.BOTH, expand=True)

        # 툴바
        toolbar = tk.Frame(sheet_card, bg=self.card_color)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        self._make_toolbar_btn(toolbar, '전체 지우기', self.card_hover, self.border_color, self.clear_sheet)
        self._make_toolbar_btn(toolbar, '실행취소', self.card_hover, self.border_color, self.undo_sheet)
        self._make_toolbar_btn(toolbar, '다시실행', self.card_hover, self.border_color, self.redo_sheet)
        self._make_toolbar_btn(toolbar, '엑셀로 다운받기', self.accent_color, self.accent_hover, self.export_sheet_to_excel)

        self.period_mode_var = tk.BooleanVar(value=False)
        self.period_mode_cb = tk.Checkbutton(
            toolbar, text="시작일/종료일 분리 (기간 입력)",
            variable=self.period_mode_var,
            bg=self.card_color, fg=self.fg_color,
            activebackground=self.card_color, activeforeground=self.fg_color,
            selectcolor=self.bg_color,
            font=('맑은 고딕', 9),
            command=self.on_toggle_period_mode
        )
        self.period_mode_cb.pack(side=tk.RIGHT, padx=(10, 0))

        # 수식 입력줄(formula bar): '=' 식 편집 중 표의 칸을 클릭하면 참조가 삽입된다
        fb_frame = tk.Frame(sheet_card, bg=self.card_color)
        fb_frame.pack(fill=tk.X, pady=(0, 6))
        tk.Label(fb_frame, text='fx', font=('맑은 고딕', 10, 'bold'), bg=self.card_color, fg=self.accent_orange).pack(side=tk.LEFT, padx=(0, 6))
        self.active_lbl = tk.Label(fb_frame, text='', width=13, anchor=tk.W, font=('맑은 고딕', 8), bg=self.card_color, fg=self.fg_muted)
        self.active_lbl.pack(side=tk.LEFT, padx=(0, 6))
        self.formula_entry = tk.Entry(fb_frame, textvariable=self.fb_var, bg=self.bg_color, fg=self.fg_color, insertbackground='white', bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1, font=('Consolas', 10))
        self.formula_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=3)
        self.formula_entry.bind('<Return>', self._commit_formula_bar)
        self.formula_entry.bind('<KP_Enter>', self._commit_formula_bar)
        self.formula_entry.bind('<Escape>', self._cancel_formula_bar)
        self.formula_entry.bind('<KeyRelease>', self._on_fb_key)
        self.formula_entry.bind('<FocusIn>', self._on_fb_focus_in)

        # 그리드 본체
        sheet_wrap = tk.Frame(sheet_card, bg=self.border_color)
        sheet_wrap.pack(fill=tk.BOTH, expand=True)
        self.sheet = Sheet(
            sheet_wrap,
            headers=SHEET_HEADERS,
            data=[["", "", "", "", "", "", "", ""] for _ in range(INITIAL_BLANK_ROWS)],
            show_row_index=True,
            font=("맑은 고딕", 11, "normal"),
            header_font=("맑은 고딕", 10, "bold"),
            index_font=("맑은 고딕", 10, "normal"),
        )
        try:
            self.sheet.change_theme("dark")
        except Exception:
            pass
        # rc_insert_row/rc_delete_row는 비활성(행 추가·삭제는 툴바 버튼으로 → 수식↔행 매핑 유지)
        self.sheet.enable_bindings(
            "single_select", "drag_select", "row_select", "column_select",
            "arrowkeys", "row_height_resize", "column_width_resize",
            "double_click_column_resize", "rc_select",
            "copy", "cut", "paste", "delete", "undo", "edit_cell",
        )
        try:
            # 붙여넣을 데이터가 현재 표보다 길면 행을 자동으로 늘린다.
            # (컬럼은 날짜 + 요금 7개 고정 — 가로 확장은 막는다)
            self.sheet.set_options(paste_can_expand_y=True, paste_can_expand_x=False)
        except Exception:
            pass
        try:
            self.sheet.set_column_widths([100, 80, 80, 80, 80, 80, 80, 80])
        except Exception:
            pass
        self.sheet.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        # 사용자 편집/붙여넣기/삭제/Ctrl+Z 시: 수식 감지 → 결과 재계산 → 건수 갱신
        # (programmatic set_cell_data 에는 발화하지 않아 재귀 위험이 없다)
        # 셀 선택 시: (수식 편집 중이면) 참조 삽입, 아니면 수식줄에 셀 내용 로드
        self.sheet.extra_bindings([
            ('modified', self._on_sheet_modified),
            ('cell_select', self._on_cell_select),
            ('column_select', self._on_cell_select),
            ('begin_edit_cell', self._on_begin_edit),
        ])
        # 인셀 클릭 참조: 셀을 편집하는 중 '='로 시작하면, 다른 칸을 클릭해 참조 삽입(엑셀식)
        #  - 수식 편집 중에는 셀 클릭으로 인한 FocusOut 로 편집기가 닫히지 않게 막고
        #  - ButtonPress-1 을 가로채 클릭한 칸의 참조를 편집기에 끼워넣는다
        try:
            mt_canvas = self.sheet.MT
            self._orig_close_text_editor = mt_canvas.close_text_editor
            mt_canvas.close_text_editor = self._wrapped_close_text_editor
            self._refclick_tag = f'RefClick{id(self)}'
            mt_canvas.bind_class(self._refclick_tag, '<ButtonPress-1>', self._incell_ref_click)
            mt_canvas.bindtags((self._refclick_tag,) + tuple(mt_canvas.bindtags()))
        except Exception as e:
            print(f'[안내] 인셀 클릭 참조 비활성화: {e}')
        self._load_active_into_fb()

        # 2-1. 기능 버튼 줄: 요금불러오기 + 요금직접입력하기
        action_row = tk.Frame(self.fare_tab, bg=self.bg_color)
        action_row.pack(fill=tk.X, padx=6, pady=(10, 8))
        self._make_toolbar_btn(action_row, '요금불러오기', self.accent_color, self.accent_hover, self.import_excel_to_sheet)

        self.toggle_btn = tk.Button(
            action_row, text='요금직접입력하기 ▶',
            bg=self.accent_green, fg='white', font=('맑은 고딕', 9, 'bold'),
            activebackground=self.accent_green_hover, activeforeground='white',
            bd=0, relief=tk.FLAT, cursor='hand2', padx=10, pady=4, command=self.toggle_sheet_panel)
        self.toggle_btn.pack(side=tk.LEFT, padx=(10, 6))
        self._add_hover(self.toggle_btn, self.accent_green, self.accent_green_hover)


        # 3. 날짜 필터 카드
        filter_card = tk.LabelFrame(self.fare_tab, text=' 날짜 필터 (선택) ', font=('맑은 고딕', 9, 'bold'), bg=self.card_color, fg=self.fg_muted, bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightthickness=1, padx=10, pady=6)
        filter_card.pack(fill=tk.X, padx=6, pady=(0, 8))

        filter_modes_frame = tk.Frame(filter_card, bg=self.card_color)
        filter_modes_frame.grid(row=0, column=0, padx=6, pady=4, sticky=tk.W)
        for txt, val in [("전체 대상", "ALL"), ("특정일 이후", "FROM_DATE"), ("특정 날짜 지정", "SPECIFIC"), ("기간 범위 지정", "DATE_RANGE")]:
            rb = tk.Radiobutton(filter_modes_frame, text=txt, variable=self.filter_mode, value=val, bg=self.card_color, fg=self.fg_color, selectcolor=self.card_color, activebackground=self.card_color, activeforeground=self.fg_color)
            rb.pack(side=tk.LEFT, padx=(0, 10))

        self.filter_input_container = tk.Frame(filter_card, bg=self.card_color)
        self.filter_input_container.grid(row=1, column=0, padx=6, pady=(2, 4), sticky=tk.W)
        self.filter_tip_lbl = tk.Label(filter_card, text='', font=('맑은 고딕', 8), bg=self.card_color, fg=self.fg_muted)
        self.filter_tip_lbl.grid(row=1, column=1, padx=(10, 0), sticky=tk.W)

        # 4. 진행 상태 바
        progress_frame = tk.Frame(self.fare_tab, bg=self.bg_color)
        progress_frame.pack(fill=tk.X, padx=6, pady=(2, 4))

        self.status_dot = tk.Label(progress_frame, text='●', font=('맑은 고딕', 11), bg=self.bg_color, fg=self.fg_muted)
        self.status_dot.pack(side=tk.LEFT, padx=(0, 6))
        self.status_lbl = tk.Label(progress_frame, text='대기 중', font=('맑은 고딕', 9, 'bold'), bg=self.bg_color, fg=self.fg_color)
        self.status_lbl.pack(side=tk.LEFT)

        self.progress_lbl = tk.Label(progress_frame, text='0 / 0 (0%)', font=('맑은 고딕', 9), bg=self.bg_color, fg=self.fg_muted)
        self.progress_lbl.pack(side=tk.RIGHT)

        self.progress_bar = ttk.Progressbar(self.fare_tab, orient='horizontal', mode='determinate', style='Accent.Horizontal.TProgressbar')
        self.progress_bar.pack(fill=tk.X, padx=6, pady=(0, 10))

        # 5. 로그 영역
        log_title = tk.Label(self.fare_tab, text='실시간 작업 내용', font=('맑은 고딕', 9, 'bold'), bg=self.bg_color, fg=self.fg_muted)
        log_title.pack(anchor=tk.W, padx=6, pady=(0, 3))

        log_wrap = tk.Frame(self.fare_tab, bg=self.border_color, bd=0)
        log_wrap.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 10))
        self.log_txt = ScrolledText(log_wrap, height=8, bg='#0c0d11', fg='#c9cdd6', insertbackground='white', font=('Consolas', 9), bd=0, relief=tk.FLAT, padx=10, pady=8)
        self.log_txt.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # 6. 컨트롤 버튼
        control_frame = tk.Frame(self.fare_tab, bg=self.bg_color)
        control_frame.pack(fill=tk.X, padx=6, pady=(2, 14))

        self.start_btn = tk.Button(control_frame, text='▶  요금수정 시작', width=18, height=2, bg=self.accent_green, fg='white', font=('맑은 고딕', 10, 'bold'), activebackground=self.accent_green_hover, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', command=self.start_rpa)
        self.start_btn.pack(side=tk.LEFT)
        self._add_hover(self.start_btn, self.accent_green, self.accent_green_hover)

        self.pause_btn = tk.Button(control_frame, text='‖  일시 중지', width=13, height=2, bg=self.accent_orange, fg='white', font=('맑은 고딕', 10, 'bold'), activebackground=self.accent_orange_hover, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', command=self.toggle_pause)
        self.pause_btn.pack(side=tk.LEFT, padx=12)
        self.pause_btn.config(state=tk.DISABLED)
        self._add_hover(self.pause_btn, self.accent_orange, self.accent_orange_hover)

        self.stop_btn = tk.Button(control_frame, text='■  중지', width=13, height=2, bg=self.accent_red, fg='white', font=('맑은 고딕', 10, 'bold'), activebackground=self.accent_red_hover, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', command=self.stop_rpa)
        self.stop_btn.pack(side=tk.RIGHT)
        self.stop_btn.config(state=tk.DISABLED)
        self._add_hover(self.stop_btn, self.accent_red, self.accent_red_hover)

        self._build_topas_tab()
        self._select_main_tab('fare')

    def _build_main_tab_bar(self, parent):
        tab_outer = tk.Frame(parent, bg=self.bg_color)
        tab_outer.pack(fill=tk.X, padx=18, pady=(0, 8))

        tab_shell = tk.Frame(
            tab_outer,
            bg=self.card_color,
            highlightbackground=self.border_color,
            highlightthickness=1,
        )
        tab_shell.pack(anchor=tk.W)

        self.main_tab_controls = {}
        for key, label in [('fare', '요금수정'), ('topas', '토파스조회')]:
            slot = tk.Frame(tab_shell, bg=self.card_color)
            slot.pack(side=tk.LEFT)
            btn = tk.Button(
                slot,
                text=label,
                bg=self.card_color,
                fg=self.fg_muted,
                activebackground=self.card_hover,
                activeforeground=self.fg_color,
                font=('맑은 고딕', 10, 'bold'),
                bd=0,
                relief=tk.FLAT,
                cursor='hand2',
                padx=22,
                pady=8,
                command=lambda tab_key=key: self._select_main_tab(tab_key),
            )
            btn.pack(fill=tk.X)
            underline = tk.Frame(slot, bg=self.card_color, height=2)
            underline.pack(fill=tk.X, padx=12)
            self.main_tab_controls[key] = (btn, underline, slot)
            btn.bind('<Enter>', lambda _e, tab_key=key: self._hover_main_tab(tab_key, True))
            btn.bind('<Leave>', lambda _e, tab_key=key: self._hover_main_tab(tab_key, False))

    def _hover_main_tab(self, tab_key, entering):
        if tab_key == self.current_main_tab or tab_key not in self.main_tab_controls:
            return
        btn, _underline, _slot = self.main_tab_controls[tab_key]
        btn.config(
            bg=self.card_hover if entering else self.card_color,
            fg=self.fg_color if entering else self.fg_muted,
        )

    def _select_main_tab(self, tab_key):
        if tab_key == 'topas':
            self._collapse_sheet_panel()

        self.current_main_tab = tab_key

        if hasattr(self, 'fare_tab') and hasattr(self, 'topas_tab'):
            self.fare_tab.pack_forget()
            self.topas_tab.pack_forget()
            target = self.topas_tab if tab_key == 'topas' else self.fare_tab
            target.pack(fill=tk.BOTH, expand=True)

        selected_bg = '#1b1f29'
        for key, controls in self.main_tab_controls.items():
            btn, underline, slot = controls
            selected = key == tab_key
            bg = selected_bg if selected else self.card_color
            fg = 'white' if selected else self.fg_muted
            btn.config(
                bg=bg,
                fg=fg,
                activebackground=bg if selected else self.card_hover,
                activeforeground='white' if selected else self.fg_color,
            )
            slot.config(bg=bg)
            underline.config(bg=self.accent_color if selected else self.card_color)

    def _make_toolbar_btn(self, parent, text, bg, hover, command):
        btn = tk.Button(parent, text=text, bg=bg, fg='white', font=('맑은 고딕', 9, 'bold'), activebackground=hover, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', padx=10, pady=4, command=command)
        btn.pack(side=tk.LEFT, padx=(0, 6))
        self._add_hover(btn, bg, hover)
        self.toolbar_buttons.append(btn)
        return btn

    def _build_topas_tab(self):
        topas_card = tk.LabelFrame(
            self.topas_tab,
            text=' TOPAS 조회 ',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_muted,
            bd=0,
            relief=tk.FLAT,
            highlightbackground=self.border_color,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        topas_card.pack(fill=tk.X, padx=6, pady=(10, 10))

        tk.Label(
            topas_card,
            text='TOPAS 화면에서 첫 날짜 조회를 먼저 실행한 뒤 시작하세요.',
            font=('맑은 고딕', 10, 'bold'),
            bg=self.card_color,
            fg=self.fg_color,
        ).pack(anchor=tk.W)

        tk.Label(
            topas_card,
            text='조회 시작 후 현재 Entry 화면의 마지막 조회 원문부터 보관하고, 입력한 횟수만큼 AC1을 실행합니다.',
            font=('맑은 고딕', 9),
            bg=self.card_color,
            fg=self.fg_muted,
        ).pack(anchor=tk.W, pady=(4, 10))

        topas_actions = tk.Frame(topas_card, bg=self.card_color)
        topas_actions.pack(fill=tk.X)

        self.topas_query_btn = tk.Button(
            topas_actions,
            text='토파스 조회하기',
            width=18,
            height=2,
            bg=self.accent_green,
            fg='white',
            font=('맑은 고딕', 10, 'bold'),
            activebackground=self.accent_green_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=self.prompt_topas_query_count,
        )
        self.topas_query_btn.pack(side=tk.LEFT)
        self._add_hover(self.topas_query_btn, self.accent_green, self.accent_green_hover)

        self.topas_stop_btn = tk.Button(
            topas_actions,
            text='■  중지',
            width=13,
            height=2,
            bg=self.accent_red,
            fg='white',
            font=('맑은 고딕', 10, 'bold'),
            activebackground=self.accent_red_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=self.stop_topas_query,
        )
        self.topas_stop_btn.pack(side=tk.RIGHT)
        self.topas_stop_btn.config(state=tk.DISABLED)
        self._add_hover(self.topas_stop_btn, self.accent_red, self.accent_red_hover)

        topas_progress_frame = tk.Frame(self.topas_tab, bg=self.bg_color)
        topas_progress_frame.pack(fill=tk.X, padx=6, pady=(2, 4))

        self.topas_status_lbl = tk.Label(
            topas_progress_frame,
            text='대기 중',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.bg_color,
            fg=self.fg_muted,
        )
        self.topas_status_lbl.pack(side=tk.LEFT)

        self.topas_progress_lbl = tk.Label(
            topas_progress_frame,
            text='0 / 0 (0%)',
            font=('맑은 고딕', 9),
            bg=self.bg_color,
            fg=self.fg_muted,
        )
        self.topas_progress_lbl.pack(side=tk.RIGHT)

        self.topas_progress_bar = ttk.Progressbar(
            self.topas_tab,
            orient='horizontal',
            mode='determinate',
            style='Accent.Horizontal.TProgressbar',
        )
        self.topas_progress_bar.pack(fill=tk.X, padx=6, pady=(0, 10))

        topas_log_title = tk.Label(
            self.topas_tab,
            text='토파스 조회 작업 내용',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.bg_color,
            fg=self.fg_muted,
        )
        topas_log_title.pack(anchor=tk.W, padx=6, pady=(0, 3))

        topas_log_wrap = tk.Frame(self.topas_tab, bg=self.border_color, bd=0)
        topas_log_wrap.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 10))
        self.topas_log_txt = ScrolledText(
            topas_log_wrap,
            height=8,
            bg='#0c0d11',
            fg='#c9cdd6',
            insertbackground='white',
            font=('Consolas', 9),
            bd=0,
            relief=tk.FLAT,
            padx=10,
            pady=8,
        )
        self.topas_log_txt.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

    def toggle_sheet_panel(self):
        """오른쪽 '요금 입력표' 패널을 펼치거나 접고, 창 너비를 함께 조절한다."""
        if self.panel_expanded:
            self._collapse_sheet_panel()
        else:
            self.side_panel.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 16), pady=16)
            self.panel_expanded = True
            self.toggle_btn.config(text='◀ 입력표 접기')
            self.root.geometry(f'{self.expanded_width}x{self.win_height}')

    def _collapse_sheet_panel(self):
        if not self.panel_expanded:
            return
        self.side_panel.pack_forget()
        self.panel_expanded = False
        if hasattr(self, 'toggle_btn'):
            self.toggle_btn.config(text='요금직접입력하기 ▶')
        self.root.geometry(f'{self.collapsed_width}x{self.win_height}')

    def get_fare_cols(self):
        is_period = self.period_mode_var.get()
        offset = 1 if is_period else 0
        return [1 + offset, 2 + offset, 3 + offset, 4 + offset, 5 + offset, 6 + offset, 7 + offset]

    def is_date_col(self, col):
        is_period = self.period_mode_var.get()
        if is_period:
            return col in (0, 1)
        return col == 0

    def on_toggle_period_mode(self):
        is_period = self.period_mode_var.get()
        try:
            current_data = self.sheet.get_sheet_data()
        except Exception:
            current_data = []

        # Shift formula coordinates
        new_formulas = {}
        if is_period:
            # 8 -> 9 columns: column 0 (Date) -> 0 (Start Date), insert 1 (End Date)
            # Other columns shift right by 1: c -> c + 1 for c >= 1
            for (r, c), f in self.formulas.items():
                if c >= 1:
                    new_formulas[(r, c + 1)] = f
                else:
                    new_formulas[(r, c)] = f
            
            # Update data: duplicate column 0 to column 1
            new_data = []
            for row in current_data:
                row_extended = list(row) + [""] * max(0, 8 - len(row))
                new_row = [row_extended[0], row_extended[0]] + row_extended[1:8]
                new_data.append(new_row)
                
            headers = ["시작일", "종료일", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"]
            col_widths = [100, 100, 75, 75, 75, 75, 75, 70, 70]
        else:
            # 9 -> 8 columns: column 0 -> 0, remove 1
            # Other columns shift left by 1: c -> c - 1 for c >= 2
            for (r, c), f in self.formulas.items():
                if c >= 2:
                    new_formulas[(r, c - 1)] = f
                elif c == 0:
                    new_formulas[(r, c)] = f
                # c == 1 (End Date) is dropped
                
            new_data = []
            for row in current_data:
                row_extended = list(row) + [""] * max(0, 9 - len(row))
                new_row = [row_extended[0]] + row_extended[2:9]
                new_data.append(new_row)
                
            headers = ["날짜", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"]
            col_widths = [100, 80, 80, 80, 80, 80, 80, 80]
            
        self.formulas = new_formulas
        self._results.clear()  # Clear cache to recalculate
        
        self.sheet.headers(headers)
        self.sheet.set_sheet_data(new_data, reset_col_positions=True, reset_row_positions=False)
        try:
            self.sheet.set_column_widths(col_widths)
        except Exception:
            pass
            
        self._recalc_formulas()
        self.refresh_count()
        self._load_active_into_fb()

    # ------------------------------------------------------------------
    # 그리드 조작
    # ------------------------------------------------------------------
    def undo_sheet(self):
        try:
            self.sheet.undo()
        except Exception:
            pass
        self._on_sheet_modified()

    def redo_sheet(self):
        try:
            self.sheet.redo()
        except Exception:
            pass
        self._on_sheet_modified()

    # ------------------------------------------------------------------
    # 수식 엔진: 셀에는 결과 표시, 식은 self.formulas 에 보관 + 자동 재계산
    # ------------------------------------------------------------------
    def _on_sheet_modified(self, event=None):
        if self._recalc_busy:
            return
        self._recalc_busy = True
        try:
            self._sync_formulas_from_cells()
            self._recalc_formulas()
        finally:
            self._recalc_busy = False
        self.refresh_count()

    def _sync_formulas_from_cells(self):
        """셀 내용을 보고 수식 보관소를 갱신한다.
        - '='로 시작하는 셀 → 수식으로 등록
        - 등록된 수식 칸인데 값이 마지막 계산결과와 다르고 '='도 아니면 → 사용자가 직접 고친 것 → 수식 해제"""
        try:
            data = self.sheet.get_sheet_data()
        except Exception:
            return
        for r, row in enumerate(data):
            for col in self.get_fare_cols():
                text = '' if (len(row) <= col or row[col] is None) else str(row[col]).strip()
                key = (r, col)
                if text.startswith('='):
                    self.formulas[key] = text
                elif key in self.formulas:
                    if text != self._results.get(key, ''):
                        self.formulas.pop(key, None)
                        self._results.pop(key, None)

    def _recalc_formulas(self):
        """보관된 수식을 계산해 셀에 결과를 써넣는다. 참조 연쇄를 위해 몇 번 반복."""
        if not self.formulas:
            return
        for _ in range(6):
            try:
                data = self.sheet.get_sheet_data()
            except Exception:
                return
            changed = False
            is_period = self.period_mode_var.get()
            num_cols = 9 if is_period else 8
            for (r, col), formula in list(self.formulas.items()):
                if r >= len(data):
                    self.formulas.pop((r, col), None)
                    self._results.pop((r, col), None)
                    continue
                row_cells = [data[r][i] if i < len(data[r]) else '' for i in range(num_cols)]
                row_cells[col] = formula  # 해당 칸은 식으로 평가
                val = self._eval_cell(row_cells, col)
                if val is None:
                    new_text = formula  # 계산 불가(빈 참조/오류) → 식 그대로 표시
                    self._results.pop((r, col), None)
                else:
                    try:
                        new_text = str(int(round(float(val))))
                    except (ValueError, TypeError):
                        new_text = formula
                        val = None
                    if val is not None:
                        self._results[(r, col)] = new_text
                cur = '' if data[r][col] is None else str(data[r][col])
                if cur != new_text:
                    try:
                        self.sheet.set_cell_data(r, col, new_text, redraw=False)
                        changed = True
                    except Exception:
                        pass
            if not changed:
                break
        try:
            self.sheet.redraw()
        except Exception:
            pass

    def _remap_formulas_after_delete(self, deleted_rows, old_len):
        """행 삭제로 인덱스가 밀린 만큼 수식/결과 보관소 키를 재매핑한다."""
        deleted = set(deleted_rows)
        remap = {}
        new_idx = 0
        for old in range(old_len):
            if old in deleted:
                continue
            remap[old] = new_idx
            new_idx += 1
        self.formulas = {(remap[r], c): f for (r, c), f in self.formulas.items() if r in remap}
        self._results = {(remap[r], c): v for (r, c), v in self._results.items() if r in remap}

    def _on_begin_edit(self, event=None):
        """인셀 편집 시작 훅.
        - 일반 칸: 그대로 편집(반환값이 에디터 텍스트가 됨).
        - 수식이 들어있는 칸: 편집 대신 '수식 전체 적용' 안내 → 편집 취소(None)."""
        try:
            r = event['row']
            c = event['column']
            value = event['value']
        except Exception:
            try:
                r, c, value = event.row, event.column, event.value
            except Exception:
                return None
        if (r, c) in self.formulas:
            self.root.after(1, lambda rr=r, cc=c: self._prompt_apply_formula(rr, cc))
            return None  # 인셀 편집 취소
        return value

    def _last_data_row(self):
        """데이터가 있는 마지막 행 인덱스. 없으면 -1."""
        try:
            data = self.sheet.get_sheet_data()
        except Exception:
            return -1
        last = -1
        is_period = self.period_mode_var.get()
        num_cols = 9 if is_period else 8
        for i, row in enumerate(data):
            for k in range(min(num_cols, len(row))):
                if row[k] is not None and str(row[k]).strip():
                    last = i
                    break
        return last

    def _prompt_apply_formula(self, r, c):
        formula = self.formulas.get((r, c))
        if not formula:
            return
        headers = self.sheet.headers()
        col_name = headers[c] if c < len(headers) else f'{c}열'
        last = self._last_data_row()
        rows_n = last + 1 if last >= 0 else 0
        msg = (f"이 칸의 계산식\n\n    {formula}\n\n"
               f"을(를) '{col_name}' 열 전체(데이터가 있는 {rows_n}개 행)에 똑같이 적용할까요?\n"
               f"(각 행의 같은 줄 값을 기준으로 계산됩니다)")
        if messagebox.askyesno('수식 전체 적용', msg):
            self._apply_formula_to_column(c, formula)

    def _apply_formula_to_column(self, col, formula):
        """formula를 col 열의 0~마지막 데이터 행까지 모두 적용하고 재계산한다."""
        last = self._last_data_row()
        if last < 0:
            return
            
        # 1. 실행 취소(Undo) 지원을 위한 변경 전 셀 값 및 수식 백업
        undo_cells = {}
        for i in range(last + 1):
            try:
                # 보관된 수식 식 문자열이 있으면 식을, 없으면 셀 값을 구함
                old_val = self.formulas.get((i, col))
                if old_val is None:
                    old_val = self.sheet.get_cell_data(i, col)
                undo_cells[(i, col)] = "" if old_val is None else str(old_val)
            except Exception:
                undo_cells[(i, col)] = ""

        # 2. 수식 적용 및 재계산
        for i in range(last + 1):
            self.formulas[(i, col)] = formula
            self._results.pop((i, col), None)
        self._recalc_busy = True
        try:
            self._recalc_formulas()
        finally:
            self._recalc_busy = False

        # 3. tksheet 내부 실행취소 스택에 복구 데이터 밀어넣기
        try:
            undo_item = {
                'name': 'edit_table',
                'data': {
                    'eventname': 'edit_table',
                    'sheetname': str(self.sheet),
                    'cells': {
                        'table': undo_cells,
                        'header': {},
                        'index': {}
                    },
                    'moved': {'rows': {}, 'columns': {}},
                    'added': {'rows': {}, 'columns': {}},
                    'deleted': {'rows': {}, 'columns': {}, 'header': {}, 'index': {}, 'column_widths': {}, 'row_heights': {}},
                    'named_spans': {},
                    'options': {},
                    'selection_boxes': {},
                    'selected': (),
                    'being_selected': (),
                    'data': {},
                    'key': '',
                    'value': None,
                    'loc': (),
                    'row': None,
                    'column': None,
                    'resized': {'rows': {}, 'columns': {}},
                    'sheet_state': {},
                    'treeview': {'nodes': {}, 'renamed': {}, 'text': {}}
                }
            }
            self.sheet.MT.undo_stack.append(undo_item)
            if hasattr(self.sheet.MT, 'redo_stack'):
                self.sheet.MT.redo_stack.clear()
        except Exception as undo_err:
            print(f"[경고] 수식 전체적용 실행취소 등록 실패: {undo_err}")

        self.refresh_count()
        self._load_active_into_fb()
        try:
            col_name = self.sheet.headers()[col]
        except Exception:
            col_name = f'{col}열'
        self.set_status(f"'{col_name}' 열 {last + 1}개 행에 수식을 적용했습니다.", self.accent_green)

    # 칸 참조 이름 → 컬럼 인덱스 매핑
    def _name_to_col(self, name, cur_col):
        # 기간 모드(9열)에서는 시작일/종료일이 0,1번을 차지하므로
        # 요금 열이 1칸씩 밀린다. 이름 매핑은 단일 모드 기준 상수에
        # offset(+1)을 더해 실제 열을 계산한다.
        offset = 1 if self.period_mode_var.get() else 0
        n = str(name).strip()
        low = n.lower()
        if n in ('앞의열', '앞열', '앞칸', '왼칸', '왼쪽', '왼쪽칸'):
            return cur_col - 1
        if n in ('뒤의열', '뒤열', '뒤칸', '오른칸', '오른쪽', '오른쪽칸'):
            return cur_col + 1
        if n in ('항공비', '항공비(성인)', '성인항공비'):
            return COL_ADULT_AIR + offset
        if n in ('호텔비', '호텔비(성인)', '성인호텔비'):
            return COL_ADULT_HOTEL + offset
        if n in ('지상비', '지상비(성인)', '성인지상비'):
            return COL_ADULT_LAND + offset
        if n in ('여행경비', '경비', '여행경비(성인)', '성인여행경비'):
            return COL_ADULT_TOUR + offset
        if n in ('알선수익', '알선수익(성인)', '성인알선수익'):
            return COL_ADULT_PROFIT + offset
        if n in ('소아', '소아요금', '아동', '아동요금'):
            return COL_CHILD + offset
        if n in ('유아', '유아요금'):
            return COL_INFANT + offset
        if n in ('날짜', '시작일', '출발일'):
            return COL_DATE
        if n in ('종료일', '도착일') and offset:
            return 1  # 기간 모드 종료일
        # A,B,C... 는 그리드의 절대 열 문자이므로 모드와 무관하게 그대로 매핑
        if len(low) == 1 and 'a' <= low <= 'i':
            return ord(low) - ord('a')
        return None

    def _eval_cell(self, row_cells, col, _stack=None):
        """한 행(row_cells=[날짜,항공비,알선수익])에서 col 칸의 값을 숫자로 계산한다.
        '='로 시작하면 사칙연산 + 칸 참조(앞의열/항공비 등)를 계산. 실패 시 None."""
        if _stack is None:
            _stack = frozenset()
        if col in _stack or col < 0 or col >= len(row_cells):
            return None  # 순환 참조/범위 밖
        raw = '' if row_cells[col] is None else str(row_cells[col]).strip()
        if not raw:
            return None
        if raw.startswith('='):
            next_stack = _stack | {col}

            def resolver(name):
                target = self._name_to_col(name, col)
                if target is None or self.is_date_col(target):
                    return None  # 날짜 칸은 계산 참조 대상이 아님
                return self._eval_cell(row_cells, target, next_stack)

            return eval_arithmetic(raw, resolver)
        try:
            return float(raw.replace(',', ''))
        except (ValueError, TypeError):
            return None

    def _coerce_fare(self, row_cells, col):
        """_eval_cell 결과를 정수로 반올림해 반환. 실패 시 None."""
        val = self._eval_cell(row_cells, col)
        if val is None:
            return None
        try:
            return int(round(float(val)))
        except (ValueError, TypeError):
            return None

    # ------------------------------------------------------------------
    # 수식 입력줄 + 클릭 참조 (엑셀식)
    # ------------------------------------------------------------------
    def _ref_name_for_col(self, c):
        # 그리드의 실시간 헤더명을 참조 이름으로 사용한다.
        # (단일/기간 모드 전환 시 열 구성이 달라져도 헤더를 그대로 따라간다)
        if c is None or c < 0:
            return None
        try:
            headers = self.sheet.headers()
        except Exception:
            return None
        if c >= len(headers):
            return None
        name = str(headers[c]).strip()
        return name or None

    # ---- 인셀 클릭 참조 (셀 편집 중 '=' 상태에서 다른 칸 클릭 → 참조 삽입) ----
    def _editor_is_formula(self):
        try:
            te = self.sheet.MT.text_editor
            if not getattr(te, 'open', False):
                return False
            txt = te.tktext.get('1.0', 'end-1c')
            return txt.lstrip().startswith('=')
        except Exception:
            return False

    def _wrapped_close_text_editor(self, event=None):
        # 수식 편집 중 셀 클릭 등으로 인한 FocusOut → 닫지 않음 (Enter/Esc/일반편집은 정상)
        if getattr(event, 'keysym', None) == 'FocusOut' and self._editor_is_formula():
            return 'break'
        return self._orig_close_text_editor(event)

    def _incell_ref_click(self, event=None):
        # 수식 편집 중이 아니면 평소대로(반환 없음 → tksheet 기본 동작)
        if not self._editor_is_formula():
            return
        try:
            col = self.sheet.identify_column(event)
        except Exception:
            return
        name = self._ref_name_for_col(col) if col is not None else None
        if not name:
            return
        try:
            te = self.sheet.MT.text_editor
            te.tktext.insert('insert', name)
            te.tktext.focus_set()
        except Exception:
            return
        return 'break'  # 선택 이동/편집 종료를 막고 편집 유지

    def _on_cell_select(self, event=None):
        try:
            sel = event.selected
            r, c = sel.row, sel.column
        except Exception:
            return
        # 수식 편집 중('='로 시작) 칸 클릭 → 편집 대상은 그대로, 참조만 삽입
        if self._ref_mode and self.fb_var.get().lstrip().startswith('='):
            name = self._ref_name_for_col(c)
            if name:
                try:
                    self.formula_entry.insert(tk.INSERT, name)
                except Exception:
                    self.fb_var.set(self.fb_var.get() + name)
            try:
                self.formula_entry.focus_set()
            except Exception:
                pass
            return
        # 일반 선택 → 편집 대상 변경 + 수식줄에 셀 내용 로드
        if r is None or c is None:
            return
        self._active_cell = (r, c)
        self._load_active_into_fb()

    def _load_active_into_fb(self):
        r, c = self._active_cell
        try:
            total = self.sheet.get_total_rows()
        except Exception:
            total = r + 1
        if total and r >= total:
            r = max(0, total - 1)
            self._active_cell = (r, c)
        # 수식이 보관된 칸이면 식을, 아니면 셀 값을 수식줄에 띄운다
        if (r, c) in self.formulas:
            val = self.formulas[(r, c)]
        else:
            try:
                val = self.sheet.get_cell_data(r, c)
            except Exception:
                val = ''
        self._loading_fb = True
        try:
            self.fb_var.set('' if val is None else str(val))
        finally:
            self._loading_fb = False
        try:
            tag = ' (수식)' if (r, c) in self.formulas else ''
            self.active_lbl.config(text=f'행 {r + 1} · {SHORT_COL_NAMES[c]}{tag}')
        except Exception:
            self.active_lbl.config(text=f'행 {r + 1}')
        self._ref_mode = False

    def _on_fb_focus_in(self, event=None):
        self._ref_mode = self.fb_var.get().lstrip().startswith('=')

    def _on_fb_key(self, event=None):
        # <KeyRelease>는 수식줄이 포커스된 상태에서만 발생하므로 내용만 보면 된다
        if self._loading_fb:
            return
        self._ref_mode = self.fb_var.get().lstrip().startswith('=')

    def _commit_formula_bar(self, event=None):
        r, c = self._active_cell
        text = self.fb_var.get()
        # 수식이면 보관소에 등록, 아니면 해제 (날짜 열에는 수식 등록 불가)
        if text.lstrip().startswith('=') and not self.is_date_col(c):
            self.formulas[(r, c)] = text
        else:
            self.formulas.pop((r, c), None)
            self._results.pop((r, c), None)
        self._recalc_busy = True
        try:
            self.sheet.set_cell_data(r, c, text, redraw=False)
            self._recalc_formulas()  # 셀에 계산 결과 표시
        except Exception:
            pass
        finally:
            self._recalc_busy = False
        self._ref_mode = False
        # 엑셀 느낌: 같은 칸의 다음 행으로 이동
        try:
            total = self.sheet.get_total_rows()
        except Exception:
            total = r + 1
        nr = r + 1 if (r + 1) < total else r
        self._active_cell = (nr, c)
        try:
            self.sheet.select_cell(nr, c, run_binding_func=False)
            self.sheet.see(nr, c)
        except Exception:
            pass
        self._load_active_into_fb()
        self.refresh_count()
        return 'break'

    def _cancel_formula_bar(self, event=None):
        self._ref_mode = False
        self._load_active_into_fb()
        try:
            self.sheet.focus_set()
        except Exception:
            pass
        return 'break'

    def add_row(self):
        data = self.sheet.get_sheet_data()
        data.append(["", "", "", "", "", "", "", ""])
        self.sheet.set_sheet_data(data, reset_col_positions=False, reset_row_positions=True)
        self.refresh_count()
        self._load_active_into_fb()

    def delete_selected_rows(self):
        try:
            rows = self.sheet.get_selected_rows(get_cells_as_rows=True)
        except Exception:
            rows = set()
        rows = sorted(rows, reverse=True)
        if not rows:
            messagebox.showinfo('행 삭제', '삭제할 행을 먼저 선택해 주세요. (왼쪽 행 번호를 클릭하면 행 전체가 선택됩니다.)')
            return
        data = self.sheet.get_sheet_data()
        old_len = len(data)
        deleted = [r for r in rows if 0 <= r < old_len]
        for r in rows:
            if 0 <= r < len(data):
                del data[r]
        if not data:
            data = [["", "", "", "", "", "", "", ""]]
        self._remap_formulas_after_delete(deleted, old_len)
        self.sheet.set_sheet_data(data, reset_col_positions=False, reset_row_positions=True)
        self.refresh_count()
        self._load_active_into_fb()

    def clear_sheet(self):
        if not messagebox.askyesno('전체 지우기', '입력한 모든 행을 지울까요?'):
            return
        self.formulas.clear()
        self._results.clear()
        self.sheet.set_sheet_data([["", "", "", "", "", "", "", ""] for _ in range(INITIAL_BLANK_ROWS)], reset_col_positions=False, reset_row_positions=True)
        self.refresh_count()
        self._load_active_into_fb()

    def import_excel_to_sheet(self):
        path = excel_loader.select_excel_file()
        if not path:
            return
        try:
            data_res = excel_loader.load_and_validate_fares(path)
        except Exception as e:
            messagebox.showerror('요금불러오기 실패', f'엑셀을 읽는 중 오류가 발생했습니다.\n{e}')
            return
        if not data_res or not data_res[0]:
            messagebox.showwarning('요금불러오기', '엑셀에서 유효한 요금 행을 찾지 못했습니다.')
            return
            
        data, is_period = data_res
        
        # Set checkbox state based on excel data format
        self.period_mode_var.set(is_period)
        
        # Rebuild layout to match loaded Excel mode
        headers = ["시작일", "종료일", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"] if is_period else ["날짜", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"]
        col_widths = [100, 100, 75, 75, 75, 75, 75, 70, 70] if is_period else [100, 80, 80, 80, 80, 80, 80, 80]
        self.sheet.headers(headers)
        try:
            self.sheet.set_column_widths(col_widths)
        except Exception:
            pass
            
        if is_period:
            grid = [[
                d['date'], 
                d['date_end'],
                str(d['adult_air']), 
                str(d['adult_hotel']), 
                str(d['adult_land']), 
                str(d['adult_tour']), 
                str(d['adult_profit']), 
                str(d['child_fare']), 
                str(d['infant_fare'])
            ] for d in data]
        else:
            grid = [[
                d['date'], 
                str(d['adult_air']), 
                str(d['adult_hotel']), 
                str(d['adult_land']), 
                str(d['adult_tour']), 
                str(d['adult_profit']), 
                str(d['child_fare']), 
                str(d['infant_fare'])
            ] for d in data]
            
        self.formulas.clear()
        self._results.clear()
        self.sheet.set_sheet_data(grid, reset_col_positions=False, reset_row_positions=True)
        # 불러온 결과를 바로 볼 수 있도록 패널을 펼친다
        if not self.panel_expanded:
            self.toggle_sheet_panel()
        self.refresh_count()
        self._load_active_into_fb()
        self.set_status(f'엑셀에서 {len(grid)}건을 표로 불러왔습니다.', self.accent_green)

    def export_sheet_to_excel(self):
        """현재 그리드의 데이터를 엑셀 파일로 내보냅니다."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="요금 데이터 엑셀로 저장",
            initialfile="ERP 요금수정 내역"
        )
        if not file_path:
            return
        
        try:
            raw = self.sheet.get_sheet_data()
            if not raw:
                messagebox.showwarning("저장 실패", "저장할 데이터가 없습니다.")
                return
            
            is_period = self.period_mode_var.get()
            num_cols = 9 if is_period else 8
            headers = ["시작일", "종료일", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아요금", "유아요금"] if is_period else SHEET_HEADERS

            data_list = []
            for r in raw:
                r_padded = [r[k] if k < len(r) else "" for k in range(num_cols)]
                r_cleaned = ["" if x is None else str(x).strip() for x in r_padded]
                if not any(r_cleaned):
                    continue
                data_list.append(r_cleaned)
            
            if not data_list:
                messagebox.showwarning("저장 실패", "저장할 유효한 요금 데이터가 없습니다.")
                return

            df = pd.DataFrame(data_list, columns=headers)
            
            # 요금 관련 컬럼은 정수로 형변환하여 저장
            target_cols = headers[2:] if is_period else headers[1:]
            for col in target_cols:
                def clean_numeric(val):
                    if not val:
                        return ""
                    try:
                        return int(float(str(val).replace(',', '')))
                    except ValueError:
                        return val
                df[col] = df[col].apply(clean_numeric)

            df.to_excel(file_path, index=False)
            messagebox.showinfo("저장 완료", f"엑셀 파일이 성공적으로 저장되었습니다.\n경로: {file_path}")
            print(f"[엑셀 내보내기] 데이터를 엑셀 파일로 저장했습니다: {file_path}")
        except Exception as e:
            messagebox.showerror("저장 실패", f"엑셀 파일 저장 중 오류가 발생했습니다.\n{e}")
            print(f"[오류] 엑셀 내보내기 실패: {str(e)}")

    def _ask_template_format(self):
        """엑셀 양식 종류를 라디오 버튼으로 선택받는다.
        반환: 'single' | 'period' | None(취소)"""
        dlg = tk.Toplevel(self.root)
        dlg.title('엑셀 양식 선택')
        dlg.configure(bg=self.bg_color)
        dlg.transient(self.root)
        dlg.resizable(False, False)

        result = {'value': None}
        # 현재 그리드 모드를 기본 선택으로 둔다
        sel = tk.StringVar(value=('period' if self.period_mode_var.get() else 'single'))

        pad = tk.Frame(dlg, bg=self.bg_color, padx=20, pady=16)
        pad.pack(fill=tk.BOTH, expand=True)

        tk.Label(pad, text='어떤 양식으로 받으시겠어요?', font=('맑은 고딕', 11, 'bold'),
                 bg=self.bg_color, fg=self.fg_color).pack(anchor=tk.W, pady=(0, 12))

        rb_kwargs = dict(bg=self.bg_color, fg=self.fg_color, selectcolor=self.bg_color,
                         activebackground=self.bg_color, activeforeground=self.fg_color,
                         font=('맑은 고딕', 10), anchor=tk.W)
        tk.Radiobutton(pad, text='단일 날짜 양식', variable=sel, value='single', **rb_kwargs).pack(fill=tk.X)
        tk.Label(pad, text='        날짜를 한 칸에 입력 (하루 또는 날짜별 요금)', font=('맑은 고딕', 8),
                 bg=self.bg_color, fg=self.fg_muted).pack(anchor=tk.W, pady=(0, 8))
        tk.Radiobutton(pad, text='기간 입력 양식', variable=sel, value='period', **rb_kwargs).pack(fill=tk.X)
        tk.Label(pad, text='        시작일·종료일로 입력 (같은 요금이 이어지는 구간)', font=('맑은 고딕', 8),
                 bg=self.bg_color, fg=self.fg_muted).pack(anchor=tk.W, pady=(0, 14))

        btn_row = tk.Frame(pad, bg=self.bg_color)
        btn_row.pack(fill=tk.X)

        def on_download():
            result['value'] = sel.get()
            dlg.destroy()

        def on_cancel():
            result['value'] = None
            dlg.destroy()

        cancel_btn = tk.Button(btn_row, text='취소', command=on_cancel,
                               bg=self.card_color, fg=self.fg_color,
                               activebackground=self.card_hover, activeforeground=self.fg_color,
                               bd=0, relief=tk.FLAT, font=('맑은 고딕', 9), padx=14, pady=6, cursor='hand2')
        cancel_btn.pack(side=tk.RIGHT)
        dl_btn = tk.Button(btn_row, text='다운받기', command=on_download,
                           bg=self.accent_color, fg='white',
                           activebackground=self.accent_hover, activeforeground='white',
                           bd=0, relief=tk.FLAT, font=('맑은 고딕', 9, 'bold'), padx=14, pady=6, cursor='hand2')
        dl_btn.pack(side=tk.RIGHT, padx=(0, 8))

        dlg.bind('<Return>', lambda e: on_download())
        dlg.bind('<Escape>', lambda e: on_cancel())
        dlg.protocol('WM_DELETE_WINDOW', on_cancel)

        # 부모 창 중앙에 배치
        dlg.update_idletasks()
        try:
            rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
            rw, rh = self.root.winfo_width(), self.root.winfo_height()
            w, h = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
            x, y = rx + (rw - w) // 2, ry + (rh - h) // 2
            dlg.geometry(f'+{max(x, 0)}+{max(y, 0)}')
        except Exception:
            pass

        dlg.grab_set()
        dl_btn.focus_set()
        self.root.wait_window(dlg)
        return result['value']

    def download_excel_template(self):
        """양식 종류(단일/기간)를 선택받아 빈 엑셀 양식을 다운로드합니다."""
        choice = self._ask_template_format()
        if choice is None:
            return  # 취소

        is_period = (choice == 'period')
        default_name = "ERP 요금수정 양식 (기간)" if is_period else "ERP 요금수정 양식 (단일)"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="엑셀 양식 다운로드",
            initialfile=default_name
        )
        if not file_path:
            return

        try:
            headers = ["시작일", "종료일", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아요금", "유아요금"] if is_period else SHEET_HEADERS
            df = pd.DataFrame(columns=headers)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("다운로드 완료", f"엑셀 양식 파일이 생성되었습니다.\n경로: {file_path}")
            print(f"[양식 다운로드] 엑셀 양식 파일을 저장했습니다: {file_path}")
        except Exception as e:
            messagebox.showerror("다운로드 실패", f"양식 다운로드 중 오류가 발생했습니다.\n{e}")
            print(f"[오류] 양식 다운로드 실패: {str(e)}")

    def read_sheet_data(self):
        """그리드를 읽어 (유효행 리스트, 오류메시지 리스트)를 반환한다.
        유효행: {row_index, date, date_end, adult_air, adult_hotel, adult_land, adult_tour, adult_profit, child_fare, infant_fare}"""
        raw = self.sheet.get_sheet_data()
        rows = []
        errors = []
        is_period = self.period_mode_var.get()
        num_cols = 9 if is_period else 8
        
        # Columns offsets based on mode
        offset = 1 if is_period else 0
        col_date = 0
        col_date_end = 1 if is_period else None
        col_adult_air = 1 + offset
        col_adult_hotel = 2 + offset
        col_adult_land = 3 + offset
        col_adult_tour = 4 + offset
        col_adult_profit = 5 + offset
        col_child = 6 + offset
        col_infant = 7 + offset
        
        for i, r in enumerate(raw):
            r_padded = [r[k] if k < len(r) else None for k in range(num_cols)]
            
            # Helper to check if row is completely empty
            non_empty_cells = [str(x).strip() for x in r_padded if x is not None and str(x).strip() != ""]
            if not non_empty_cells:
                continue
                
            date_cell = str(r_padded[col_date]).strip() if r_padded[col_date] is not None else ""
            date_end_cell = ""
            if is_period and col_date_end is not None:
                date_end_cell = str(r_padded[col_date_end]).strip() if r_padded[col_date_end] is not None else ""
                if not date_end_cell:
                    date_end_cell = date_cell  # Fallback to start date if blank
            else:
                date_end_cell = date_cell
                
            adult_air_cell = str(r_padded[col_adult_air]).strip() if r_padded[col_adult_air] is not None else ""
            adult_hotel_cell = str(r_padded[col_adult_hotel]).strip() if r_padded[col_adult_hotel] is not None else ""
            adult_land_cell = str(r_padded[col_adult_land]).strip() if r_padded[col_adult_land] is not None else ""
            adult_tour_cell = str(r_padded[col_adult_tour]).strip() if r_padded[col_adult_tour] is not None else ""
            adult_profit_cell = str(r_padded[col_adult_profit]).strip() if r_padded[col_adult_profit] is not None else ""
            child_cell = str(r_padded[col_child]).strip() if r_padded[col_child] is not None else ""
            infant_cell = str(r_padded[col_infant]).strip() if r_padded[col_infant] is not None else ""

            line = i + 1
            norm_date = normalize_date(date_cell)
            if not norm_date:
                errors.append(f"{line}행: 시작 날짜 형식 오류 ('{date_cell}')")
                continue
                
            norm_date_end = normalize_date(date_end_cell)
            if not norm_date_end:
                errors.append(f"{line}행: 종료 날짜 형식 오류 ('{date_end_cell}')")
                continue
                
            if norm_date > norm_date_end:
                errors.append(f"{line}행: 시작 날짜('{norm_date}')가 종료 날짜('{norm_date_end}')보다 늦습니다")
                continue

            row_cells = list(r_padded)
            
            def coerce_val(col_idx, label):
                val = self._coerce_fare(row_cells, col_idx)
                if val is None:
                    return 0
                if val < 0:
                    errors.append(f"{line}행: {label} 요금은 음수가 될 수 없습니다")
                    return 0
                return val

            adult_air = coerce_val(col_adult_air, "항공비(성인)")
            adult_hotel = coerce_val(col_adult_hotel, "호텔비(성인)")
            adult_land = coerce_val(col_adult_land, "지상비(성인)")
            adult_tour = coerce_val(col_adult_tour, "여행경비(성인)")
            adult_profit = coerce_val(col_adult_profit, "알선수익(성인)")
            child_fare = coerce_val(col_child, "소아요금")
            infant_fare = coerce_val(col_infant, "유아요금")

            rows.append({
                "row_index": line, 
                "date": norm_date, 
                "date_end": norm_date_end,
                "adult_air": adult_air, 
                "adult_hotel": adult_hotel, 
                "adult_land": adult_land, 
                "adult_tour": adult_tour, 
                "adult_profit": adult_profit, 
                "child_fare": child_fare, 
                "infant_fare": infant_fare
            })
        return rows, errors

    def _apply_date_filter(self, rows):
        """그리드에서 읽은 행 리스트에 날짜 필터를 적용한다(콘솔 출력 없음).
        기간 모드에서는 각 행이 [date ~ date_end] 구간을 가지므로,
        '겹치기만 하면 포함' 기준으로 종료일까지 고려한다.
        단일 모드는 date_end == date 라서 동일 로직이 그대로 들어맞는다."""
        def _end(r):
            # 안전장치: date_end 가 없으면 시작일과 동일하게 취급
            return r.get("date_end") or r["date"]

        mode = self.filter_mode.get()
        if mode == "ALL":
            return list(rows)
        if mode == "FROM_DATE":
            v = normalize_date(self.filter_value.get())
            if not v:
                return list(rows)
            # 기간이 기준일까지 닿으면(종료일 >= 기준일) 포함
            return [r for r in rows if _end(r) >= v]
        if mode == "SPECIFIC":
            toks = set()
            for d in self.filter_value.get().split(','):
                nd = normalize_date(d)
                if nd:
                    toks.add(nd)
            if not toks:
                return list(rows)
            # 지정일 중 하나라도 기간 구간 안에 들어오면 포함
            return [r for r in rows if any(r["date"] <= t <= _end(r) for t in toks)]
        if mode == "DATE_RANGE":
            s = normalize_date(self.filter_value.get())
            e = normalize_date(self.filter_value_end.get())
            if not s or not e:
                return list(rows)
            if s > e:
                s, e = e, s
            # 기간 [date~date_end] 와 [s~e] 가 겹치면 포함
            return [r for r in rows if r["date"] <= e and _end(r) >= s]
        return list(rows)

    def refresh_count(self, *args):
        """그리드/필터 상태를 읽어 대상 건수와 상태표시를 갱신한다."""
        if self.is_running:
            return
        try:
            rows, errors = self.read_sheet_data()
        except Exception:
            return
        filtered = self._apply_date_filter(rows)
        self.fares_data = filtered
        n = len(filtered)
        self.progress_bar['maximum'] = max(n, 1)
        self.progress_bar['value'] = 0
        self.progress_lbl.config(text=f'0 / {n} (0%)')
        if errors:
            self.set_status(f'사용 가능 {n}건 · 입력 오류 {len(errors)}행 (시작 시 확인)', self.accent_orange)
        elif n:
            self.set_status(f'입력 {n}건 준비됨', self.accent_green)
        else:
            self.set_status('입력된 요금 행이 없습니다', self.fg_muted)

    # ------------------------------------------------------------------
    # 날짜 필터 입력 UI (동적)
    # ------------------------------------------------------------------
    def _on_filter_mode_change(self, *args):
        try:
            if not hasattr(self, 'filter_input_container') or not hasattr(self, 'filter_tip_lbl'):
                return
            for widget in self.filter_input_container.winfo_children():
                widget.destroy()

            mode = self.filter_mode.get()
            entry_kwargs = dict(bg=self.bg_color, fg=self.fg_color, insertbackground='white', bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)

            if mode == "ALL":
                self.filter_tip_lbl.config(text="표에 입력된 모든 행을 실행합니다.")
                if self.filter_value.get():
                    self.filter_value.set("")
                if self.filter_value_end.get():
                    self.filter_value_end.set("")
            elif mode == "FROM_DATE":
                self.filter_tip_lbl.config(text="예: 2026-06-07 또는 20260607 (이 날짜 포함, 이후만 실행)")
                tk.Entry(self.filter_input_container, textvariable=self.filter_value, width=40, **entry_kwargs).pack(side=tk.LEFT, ipady=4)
            elif mode == "SPECIFIC":
                self.filter_tip_lbl.config(text="예: 2026-06-07, 2026-06-08 (콤마로 구분한 날짜만 실행)")
                tk.Entry(self.filter_input_container, textvariable=self.filter_value, width=40, **entry_kwargs).pack(side=tk.LEFT, ipady=4)
            elif mode == "DATE_RANGE":
                self.filter_tip_lbl.config(text="시작일 ~ 종료일 범위 안의 행만 실행")
                tk.Entry(self.filter_input_container, textvariable=self.filter_value, width=18, **entry_kwargs).pack(side=tk.LEFT, ipady=4)
                tk.Label(self.filter_input_container, text=" ~ ", font=('맑은 고딕', 10, 'bold'), bg=self.card_color, fg=self.fg_color).pack(side=tk.LEFT, padx=3)
                tk.Entry(self.filter_input_container, textvariable=self.filter_value_end, width=18, **entry_kwargs).pack(side=tk.LEFT, ipady=4)

            self.refresh_count()
        except Exception:
            pass

    def _on_filter_value_change(self, *args):
        try:
            self.refresh_count()
        except Exception:
            pass

    def _set_filter_inputs_state(self, state):
        try:
            if hasattr(self, 'filter_input_container'):
                for w in self.filter_input_container.winfo_children():
                    if isinstance(w, tk.Entry):
                        try:
                            w.config(state=state)
                        except Exception:
                            pass
        except Exception:
            pass

    def _set_inputs_locked(self, locked):
        """RPA 실행 중 입력(그리드/툴바/필터)을 잠그거나 해제한다."""
        state = tk.DISABLED if locked else tk.NORMAL
        for btn in self.toolbar_buttons:
            try:
                btn.config(state=state)
            except Exception:
                pass
        self._set_filter_inputs_state(state)
        try:
            self.formula_entry.config(state=state)
        except Exception:
            pass
        try:
            if locked:
                self.sheet.disable_bindings("edit_cell", "paste", "cut", "delete", "undo")
            else:
                self.sheet.enable_bindings("edit_cell", "paste", "cut", "delete", "undo")
        except Exception:
            pass

    # ------------------------------------------------------------------
    # RPA 실행 제어
    # ------------------------------------------------------------------
    def start_rpa(self):
        if self.is_running:
            messagebox.showwarning('실행 중', '이미 실행 중인 작업이 있습니다.')
            return
        rows, errors = self.read_sheet_data()
        if errors:
            preview = "\n".join(errors[:10])
            if len(errors) > 10:
                preview += f"\n... 외 {len(errors) - 10}건"
            msg = ("표에서 다음 행에 문제가 있습니다:\n\n"
                   f"{preview}\n\n"
                   "문제 행은 건너뛰고 정상 행만 진행할까요?")
            if not messagebox.askyesno('입력 확인', msg):
                return

        filtered = self._apply_date_filter(rows)
        if not filtered:
            messagebox.showwarning('데이터 없음', '실행할 유효한 요금 행이 없습니다.\n표 입력과 날짜 필터를 확인해 주세요.')
            return
        self.fares_data = filtered

        self.is_running = True
        self.is_paused = False
        self.is_user_stopped = False

        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL, text='‖  일시 중지', bg=self.accent_orange)
        self.stop_btn.config(state=tk.NORMAL)
        self._set_inputs_locked(True)

        self.progress_bar['maximum'] = len(self.fares_data)
        self.progress_bar['value'] = 0
        self.progress_lbl.config(text=f'0 / {len(self.fares_data)} (0%)')

        self.log_txt.delete('1.0', tk.END)
        self.old_stdout = sys.stdout
        self.console_redirector = GUIConsoleRedirector(self.log_txt, self.root)
        sys.stdout = self.console_redirector

        self.rpa_thread = threading.Thread(target=self.rpa_worker_loop, daemon=True)
        self.rpa_thread.start()

    def toggle_pause(self):
        if self.is_running:
            if not self.is_paused:
                self.is_paused = True
                self.pause_btn.config(text='▶  다시 시작', bg=self.accent_green)
                self.set_status('일시 중지됨', self.accent_orange)
                print('\n[RPA 일시중지] 현재 진행 중인 날짜 처리가 끝난 후 대기 상태로 들어갑니다.')
            else:
                self.is_paused = False
                self.pause_btn.config(text='‖  일시 중지', bg=self.accent_orange)
                self.set_status('요금 주입 루프 작동 중 (LIVE)', self.accent_green)
                print('\n[RPA 다시시작] 대기를 해제하고 다음 작업을 속행합니다.')

    def stop_rpa(self):
        if self.is_running:
            if messagebox.askyesno('중지 확인', '진행 중인 요금수정을 중지하시겠습니까?'):
                self.is_running = False
                self.is_paused = False
                self.is_user_stopped = True
                self.pause_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.DISABLED, text='중지 중...')
                self.set_status('중지 요청됨 · 현재 단계 정리 중', self.accent_red)
                print('\n[RPA 중단] 사용자 요청에 의해 중지 요청을 보냈습니다. 현재 단계가 정리되면 프로그램 화면으로 돌아옵니다...')

    def clean_up_ui_after_rpa(self):
        self.hide_loading()
        if hasattr(self, 'old_stdout'):
            if self.console_redirector is not None:
                self.console_redirector.close()
                self.console_redirector = None
            sys.stdout = self.old_stdout

        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text='‖  일시 중지', bg=self.accent_orange)
        self.stop_btn.config(state=tk.DISABLED, text='■  중지')
        self.is_running = False
        self.is_paused = False
        self._set_inputs_locked(False)
        self.progress_bar['value'] = 0
        self.progress_lbl.config(text=f'0 / {len(self.fares_data)} (0%)')

    def finish_rpa_on_ui(self, history):
        try:
            self.generate_and_show_report(history)
        finally:
            self.clean_up_ui_after_rpa()

    def _get_debug_browser_target(self):
        if getattr(self, 'current_main_tab', 'fare') == 'topas':
            return self.TOPAS_LOGIN_URL, 'TOPAS'
        return self.ERP_LOGIN_URL, 'ERP'

    # ------------------------------------------------------------------
    # TOPAS 조회 제어
    # ------------------------------------------------------------------
    def prompt_topas_query_count(self):
        if self.is_running:
            messagebox.showwarning('실행 중', '이미 실행 중인 작업이 있습니다.')
            return

        dialog = tk.Toplevel(self.root)
        dialog.title('토파스 조회 횟수')
        dialog.configure(bg=self.bg_color)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        box = tk.Frame(dialog, bg=self.bg_color, padx=18, pady=16)
        box.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            box,
            text='AC1 실행 횟수',
            font=('맑은 고딕', 11, 'bold'),
            bg=self.bg_color,
            fg=self.fg_color,
        ).pack(anchor=tk.W)
        tk.Label(
            box,
            text='첫 조회는 TOPAS 화면에서 직접 실행한 상태여야 합니다.',
            font=('맑은 고딕', 9),
            bg=self.bg_color,
            fg=self.fg_muted,
        ).pack(anchor=tk.W, pady=(4, 10))

        count_var = tk.StringVar(value='')
        count_entry = tk.Entry(
            box,
            textvariable=count_var,
            width=18,
            bg='#0c0d11',
            fg=self.fg_color,
            insertbackground='white',
            relief=tk.FLAT,
            highlightbackground=self.border_color,
            highlightcolor=self.accent_color,
            highlightthickness=1,
            font=('Consolas', 13),
        )
        count_entry.pack(anchor=tk.W, ipady=5)

        tk.Label(
            box,
            text='예: 200 입력 시 AC1을 200번 실행합니다.',
            font=('맑은 고딕', 8),
            bg=self.bg_color,
            fg=self.fg_muted,
        ).pack(anchor=tk.W, pady=(6, 14))

        tk.Label(
            box,
            text='AC1은 10개씩 묶어 전송하고, 10개 응답이 모두 확인된 뒤 다음 묶음을 전송합니다.',
            font=('맑은 고딕', 8),
            bg=self.bg_color,
            fg=self.fg_muted,
        ).pack(anchor=tk.W, pady=(6, 14))

        buttons = tk.Frame(box, bg=self.bg_color)
        buttons.pack(fill=tk.X)

        def run_from_dialog(_event=None):
            raw = count_var.get().strip()
            try:
                count = int(raw)
            except ValueError:
                messagebox.showwarning('입력 확인', '조회 횟수는 숫자로 입력해 주세요.', parent=dialog)
                return
            if count <= 0:
                messagebox.showwarning('입력 확인', '조회 횟수는 1 이상이어야 합니다.', parent=dialog)
                return
            dialog.grab_release()
            dialog.destroy()
            batch_size = self._int_config('topas_batch_size', 10, 10, 10)
            self.start_topas_query(count, batch_size)

        run_btn = tk.Button(
            buttons,
            text='실행',
            width=10,
            bg=self.accent_green,
            fg='white',
            font=('맑은 고딕', 9, 'bold'),
            activebackground=self.accent_green_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=run_from_dialog,
        )
        run_btn.pack(side=tk.LEFT)
        self._add_hover(run_btn, self.accent_green, self.accent_green_hover)

        cancel_btn = tk.Button(
            buttons,
            text='취소',
            width=10,
            bg=self.card_hover,
            fg='white',
            font=('맑은 고딕', 9, 'bold'),
            activebackground=self.border_color,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=lambda: (dialog.grab_release(), dialog.destroy()),
        )
        cancel_btn.pack(side=tk.LEFT, padx=(8, 0))
        self._add_hover(cancel_btn, self.card_hover, self.border_color)

        dialog.bind('<Return>', run_from_dialog)
        count_entry.focus_set()
        self.root.update_idletasks()
        x = self.root.winfo_rootx() + max(0, (self.root.winfo_width() - 320) // 2)
        y = self.root.winfo_rooty() + max(0, (self.root.winfo_height() - 180) // 2)
        dialog.geometry(f'+{x}+{y}')

    def start_topas_query(self, count, batch_size=1):
        if self.is_running:
            messagebox.showwarning('실행 중', '이미 실행 중인 작업이 있습니다.')
            return

        self.is_running = True
        self.topas_stop_requested = False
        self.topas_results_raw = []
        self._topas_reset_element_cache()

        self.topas_query_btn.config(state=tk.DISABLED)
        self.topas_stop_btn.config(state=tk.NORMAL)
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        self._set_inputs_locked(True)

        self.topas_progress_bar['maximum'] = count
        self.topas_progress_bar['value'] = 0
        self.topas_progress_lbl.config(text=f'0 / {count} (0%)')
        self._set_topas_status('TOPAS 연결 준비 중', self.accent_orange)
        self.topas_log_txt.delete('1.0', tk.END)
        if batch_size > 1:
            self._append_topas_log(f'[묶음 전송] AC1을 {batch_size}건씩 전송합니다.\n')

        self.topas_thread = threading.Thread(target=self.topas_worker_loop, args=(count, batch_size), daemon=True)
        self.topas_thread.start()

    def stop_topas_query(self):
        if self.is_running:
            self.topas_stop_requested = True
            self.topas_stop_btn.config(state=tk.DISABLED, text='중지 중...')
            self._set_topas_status('중지 요청됨 · 현재 조회 완료 후 정리', self.accent_red)
            self._append_topas_log('[중지 요청] 현재 AC1 응답을 받은 뒤 멈춥니다.\n')

    def topas_worker_loop(self, count, batch_size=1):
        driver = None
        results = []
        error_message = None
        stopped = False
        started = time.perf_counter()
        batch_size = max(1, int(batch_size or 1))

        try:
            self.root.after(0, lambda: self._append_topas_log('[TOPAS] 디버그 브라우저 연결 중...\n'))
            options = webdriver.ChromeOptions()
            debug_addr = self.config.get("debugger_address", self.config.get("debuggerAddress", "127.0.0.1:9222"))
            options.add_experimental_option("debuggerAddress", debug_addr)
            driver = webdriver.Chrome(options=options)
            self.driver = driver
            self._topas_reset_element_cache()

            first_block = self._topas_get_latest_block(driver)
            if first_block is None:
                if not self._topas_has_entry_shell(driver):
                    location = self._topas_debug_location(driver)
                    raise RuntimeError(
                        '디버그 브라우저(127.0.0.1:9222)에서 TOPAS Entry 화면을 찾지 못했습니다.\n'
                        'TOPAS 조회 탭에서 [브라우저 켜기]로 TOPAS를 연 뒤 로그인하고 첫 날짜 조회까지 실행해 주세요.\n'
                        f'현재 연결된 브라우저: {location}'
                    )
                excerpt = self._topas_shell_excerpt(driver)
                raise RuntimeError(
                    'TOPAS Entry 화면은 찾았지만 조회 결과 원문을 파싱하지 못했습니다.\n'
                    f'화면 원문 일부: {excerpt}'
                )

            results.append(first_block.raw_text)
            previous_block = first_block
            first_desc = first_block.request_command or '현재 조회'
            self.root.after(0, lambda desc=first_desc: self._append_topas_log(f'[초기 조회 보관] {desc}\n'))
            self.root.after(0, lambda: self._set_topas_status('AC1 조회 중', self.accent_green))

            done = 0
            while done < count:
                if self.topas_stop_requested or not self.is_running:
                    stopped = True
                    break

                chunk_size = min(batch_size, count - done)
                step_started = time.perf_counter()
                self._topas_send_ac1_batch(driver, chunk_size)
                blocks = self._topas_wait_next_blocks(driver, previous_block, chunk_size)
                chunk_elapsed = time.perf_counter() - step_started
                if not blocks:
                    raise TimeoutError('TOPAS가 AC1 묶음에 대한 새 응답을 반환하지 않았습니다.')

                for block in blocks:
                    done += 1
                    results.append(block.raw_text)
                    previous_block = block

                    request = block.request_command or f'AC1 #{done}'
                    pct = int(done / count * 100)
                    elapsed_per_item = chunk_elapsed / max(1, len(blocks))
                    self.root.after(
                        0,
                        lambda i=done, c=count, p=pct, req=request, sec=elapsed_per_item: self._topas_progress_update(
                            i, c, p, req, sec
                        ),
                    )

            self.topas_results_raw = results
        except Exception as exc:
            error_message = str(exc)
            self.root.after(0, lambda msg=error_message: self._append_topas_log(f'[오류] {msg}\n'))
        finally:
            total_elapsed = time.perf_counter() - started
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass
            self.driver = None
            self._topas_reset_element_cache()
            self.root.after(
                0,
                lambda res=list(results), err=error_message, stop=stopped, sec=total_elapsed: self.finish_topas_query_on_ui(
                    res, err, stop, sec
                ),
            )

    def _topas_send_ac1(self, driver):
        self._topas_send_ac1_batch(driver, 1)

    def _topas_send_ac1_batch(self, driver, count):
        prompt = self._topas_find_prompt_input(driver)
        if prompt is None:
            raise RuntimeError('TOPAS 명령 입력창을 찾지 못했습니다. Entry 화면이 열려 있는지 확인해 주세요.')
        try:
            driver.execute_script('arguments[0].focus(); arguments[0].value = "";', prompt)
            keys = []
            for _ in range(max(1, int(count))):
                keys.extend(('AC1', Keys.ENTER))
            prompt.send_keys(*keys)
        except Exception:
            self._topas_prompt_el = None
            prompt = self._topas_find_prompt_input(driver)
            if prompt is None:
                raise
            driver.execute_script('arguments[0].focus(); arguments[0].value = "";', prompt)
            keys = []
            for _ in range(max(1, int(count))):
                keys.extend(('AC1', Keys.ENTER))
            prompt.send_keys(*keys)

    def _topas_wait_next_block(self, driver, previous_block, timeout=None):
        return self._topas_wait_next_blocks(driver, previous_block, 1, timeout=timeout)[0]

    def _topas_wait_next_blocks(self, driver, previous_block, expected_count, timeout=None):
        if timeout is None:
            timeout = self._float_config('topas_batch_timeout', 120, 10, 300)
        deadline = time.perf_counter() + timeout
        last_request = None
        last_loaded_count = 0
        candidate_signature = None
        candidate_blocks = []
        candidate_seen_at = None
        stable_wait = self._float_config('topas_response_stable_wait', 0.2, 0.05, 2.0)
        poll_interval = self._float_config('topas_poll_interval', 0.04, 0.02, 0.2)
        expected_count = max(1, int(expected_count))

        while time.perf_counter() < deadline:
            blocks = self._topas_get_blocks_after(driver, previous_block)
            if blocks:
                last_request = blocks[-1].request_command
            new_blocks = self._topas_collect_new_response_blocks(blocks, previous_block)
            last_loaded_count = max(last_loaded_count, len(new_blocks))
            if len(new_blocks) >= expected_count:
                candidate = new_blocks[:expected_count]
                signature = tuple((self._topas_block_key(block), block.raw_text) for block in candidate)
                if signature != candidate_signature:
                    candidate_signature = signature
                    candidate_blocks = candidate
                    candidate_seen_at = time.perf_counter()
                elif candidate_seen_at is not None and self._topas_prompt_is_idle(driver):
                    stable_for = time.perf_counter() - candidate_seen_at
                    if stable_for >= stable_wait and len(candidate_blocks) >= expected_count:
                        return candidate_blocks
            time.sleep(poll_interval)
        raise TimeoutError(
            f'TOPAS 묶음 응답 대기 {int(timeout)}초 초과: '
            f'{expected_count}개 중 {last_loaded_count}개만 확인했습니다. '
            f'마지막 응답: {last_request or "없음"}'
        )

    def _topas_collect_new_response_blocks(self, blocks, previous_block):
        collected = []
        seen = set()
        for block in blocks:
            if not self._topas_is_response_block(block):
                continue
            key = self._topas_block_key(block)
            if key in seen:
                continue
            seen.add(key)
            collected.append(block)
        return collected

    def _topas_is_new_response_block(self, block, previous_block):
        return self._topas_is_response_block(block) and block.raw_text != previous_block.raw_text

    def _topas_is_response_block(self, block):
        if not block.request_command:
            return False

        raw_upper = block.raw_text.upper()
        has_response_marker = (
            'AMADEUS AVAILABILITY' in raw_upper
            or 'NO FLIGHT' in raw_upper
            or 'REQUEST NEW AVAILABILITY' in raw_upper
        )
        return has_response_marker

    def _topas_block_key(self, block):
        return (
            block.request_command or '',
            block.offset_days,
            block.travel_date.isoformat() if block.travel_date else '',
        )

    def _topas_get_latest_block(self, driver):
        blocks = self._topas_get_recent_blocks(driver)
        return blocks[-1] if blocks else None

    def _topas_get_recent_blocks(self, driver):
        shell = self._topas_find_entry_shell(driver)
        if shell is None:
            return []
        text = self._topas_shell_text(driver, shell)
        tail_chars = self._int_config('topas_parse_tail_chars', 80000, 4000, 300000)
        if len(text) > tail_chars:
            text = text[-tail_chars:]
        blocks = parse_availability_text(text, year_hint=datetime.now().year)
        return blocks

    def _topas_get_blocks_after(self, driver, previous_block):
        shell = self._topas_find_entry_shell(driver)
        if shell is None:
            return []
        text = self._topas_shell_text(driver, shell)
        segment = self._topas_text_after_previous_block(text, previous_block)
        blocks = parse_availability_text(segment, year_hint=datetime.now().year)
        return blocks

    def _topas_text_after_previous_block(self, text, previous_block):
        if not text:
            return ''

        raw = (previous_block.raw_text or '').strip()
        if raw:
            idx = text.rfind(raw)
            if idx >= 0:
                return text[idx + len(raw):]

        request = (previous_block.request_command or '').strip()
        if request:
            idx = text.rfind(request)
            if idx >= 0:
                next_prompt = text.find('\n>', idx)
                if next_prompt >= 0:
                    return text[next_prompt + 1:]
                return text[idx + len(request):]

        tail_chars = self._int_config('topas_parse_tail_chars', 80000, 4000, 300000)
        return text[-tail_chars:]

    def _topas_find_entry_shell(self, driver):
        return self._topas_get_cached_element(driver, '_topas_shell_el', self.TOPAS_SHELL_ROOT)

    def _topas_find_prompt_input(self, driver):
        return self._topas_get_cached_element(driver, '_topas_prompt_el', self.TOPAS_PROMPT_INPUT)

    def _topas_reset_element_cache(self):
        self._topas_shell_el = None
        self._topas_prompt_el = None

    def _topas_get_cached_element(self, driver, attr, selector):
        cached = getattr(self, attr, None)
        if cached is not None:
            try:
                _ = cached.tag_name
                return cached
            except Exception:
                setattr(self, attr, None)

        found = self._topas_find_element(driver, selector)
        setattr(self, attr, found)
        return found

    def _topas_find_element(self, driver, selector):
        try:
            original_handle = driver.current_window_handle
        except Exception:
            original_handle = None

        for handle in list(driver.window_handles):
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                found = self._topas_find_element_in_frames(driver, selector)
                if found is not None:
                    return found
            except Exception:
                continue

        if original_handle:
            try:
                driver.switch_to.window(original_handle)
                driver.switch_to.default_content()
            except Exception:
                pass
        return None

    def _topas_find_element_in_frames(self, driver, selector, depth=0, max_depth=4):
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        if elements:
            return elements[0]
        if depth >= max_depth:
            return None

        frames = driver.find_elements(By.CSS_SELECTOR, 'iframe, frame')
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                found = self._topas_find_element_in_frames(driver, selector, depth + 1, max_depth)
                if found is not None:
                    return found
            except Exception:
                pass
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass
        return None

    def _topas_has_entry_shell(self, driver):
        return self._topas_find_entry_shell(driver) is not None

    def _topas_shell_text(self, driver, shell):
        try:
            text = driver.execute_script(
                'return arguments[0].innerText || arguments[0].textContent || "";',
                shell,
            ) or ''
            if text.strip():
                return text
        except Exception:
            pass
        return shell.text or ''

    def _topas_shell_excerpt(self, driver, limit=240):
        shell = self._topas_find_entry_shell(driver)
        if shell is None:
            return ''
        text = self._topas_shell_text(driver, shell).replace('\r', ' ').replace('\n', ' ')
        text = ' '.join(text.split())
        return text[:limit] if text else '(빈 화면)'

    def _topas_debug_location(self, driver):
        locations = []
        try:
            handles = list(driver.window_handles)
        except Exception:
            handles = []
        for idx, handle in enumerate(handles, start=1):
            try:
                driver.switch_to.window(handle)
                title = driver.title or '(제목 없음)'
                url = driver.current_url or '(주소 없음)'
                locations.append(f'{idx}. {title} - {url}')
            except Exception:
                continue
        return ' / '.join(locations) if locations else '(확인 불가)'

    def _topas_prompt_is_idle(self, driver):
        try:
            prompt = self._topas_find_prompt_input(driver)
            if prompt is None:
                return False
            return (prompt.get_attribute('value') or '') == ''
        except Exception:
            return False

    def _topas_progress_update(self, done, total, pct, request, elapsed):
        self.topas_progress_bar['value'] = done
        self.topas_progress_lbl.config(text=f'{done} / {total} ({pct}%)')
        self._set_topas_status(f'조회 중 · {done}/{total}', self.accent_green)
        self._append_topas_log(f'[{done}/{total}] {request} 완료 ({elapsed:.2f}s)\n')

    def finish_topas_query_on_ui(self, results, error_message, stopped, elapsed):
        self.clean_up_ui_after_topas()
        if error_message:
            self._set_topas_status('조회 오류', self.accent_red)
            messagebox.showerror('TOPAS 조회 오류', error_message)
            if results:
                self.show_topas_results_popup(results, elapsed, stopped=True)
            return

        if stopped:
            self._set_topas_status('사용자 중지 · 부분 결과 보관', self.accent_orange)
            self._append_topas_log(f'[중지 완료] {len(results)}개 원문을 보관했습니다. 총 {elapsed:.1f}s\n')
        else:
            self._set_topas_status('조회 완료', self.accent_green)
            self._append_topas_log(f'[조회 완료] {len(results)}개 원문을 보관했습니다. 총 {elapsed:.1f}s\n')

        if results:
            self.show_topas_results_popup(results, elapsed, stopped=stopped)

    def clean_up_ui_after_topas(self):
        self.is_running = False
        self.topas_stop_requested = False
        self.topas_query_btn.config(state=tk.NORMAL)
        self.topas_stop_btn.config(state=tk.DISABLED, text='■  중지')
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED, text='‖  일시 중지', bg=self.accent_orange)
        self.stop_btn.config(state=tk.DISABLED, text='■  중지')
        self._set_inputs_locked(False)

    def show_topas_results_popup(self, results, elapsed, stopped=False):
        title = 'TOPAS 조회 결과'
        popup_bg = '#ffffff'
        popup_fg = '#111827'
        popup_border = '#d1d5db'
        popup = tk.Toplevel(self.root)
        popup.title(title)
        popup.configure(bg=popup_bg)
        popup.geometry('980x680')
        popup.minsize(720, 460)

        header = tk.Frame(popup, bg=popup_bg, padx=14, pady=12)
        header.pack(fill=tk.X)
        status = '부분 결과' if stopped else '전체 결과'
        tk.Label(
            header,
            text=f'{status} · 원문 {len(results)}개 · {elapsed:.1f}s',
            font=('맑은 고딕', 11, 'bold'),
            bg=popup_bg,
            fg=popup_fg,
        ).pack(side=tk.LEFT)

        body = tk.Frame(popup, bg=popup_border)
        body.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 10))

        result_text = '\n\n'.join(results)
        text_widget = ScrolledText(
            body,
            bg=popup_bg,
            fg=popup_fg,
            insertbackground=popup_fg,
            selectbackground='#bfdbfe',
            selectforeground=popup_fg,
            font=('Consolas', 10),
            bd=0,
            relief=tk.FLAT,
            padx=10,
            pady=8,
            wrap=tk.NONE,
        )
        text_widget.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        text_widget.insert('1.0', result_text)

        footer = tk.Frame(popup, bg=popup_bg, padx=14)
        footer.pack(fill=tk.X, pady=(0, 12))

        def copy_all():
            self.root.clipboard_clear()
            self.root.clipboard_append(result_text)
            self.root.update()
            self._append_topas_log('[복사] TOPAS 조회 원문 전체를 클립보드에 복사했습니다.\n')

        copy_btn = tk.Button(
            footer,
            text='전체 복사하기',
            width=14,
            bg=self.accent_color,
            fg='white',
            font=('맑은 고딕', 9, 'bold'),
            activebackground=self.accent_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=copy_all,
        )
        copy_btn.pack(side=tk.LEFT)
        self._add_hover(copy_btn, self.accent_color, self.accent_hover)

        close_btn = tk.Button(
            footer,
            text='닫기',
            width=10,
            bg='#f3f4f6',
            fg=popup_fg,
            font=('맑은 고딕', 9, 'bold'),
            activebackground='#e5e7eb',
            activeforeground=popup_fg,
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=popup.destroy,
        )
        close_btn.pack(side=tk.LEFT, padx=(8, 0))
        self._add_hover(close_btn, '#f3f4f6', '#e5e7eb', normal_fg=popup_fg, hover_fg=popup_fg)

    def _set_topas_status(self, text, color=None):
        if hasattr(self, 'topas_status_lbl'):
            self.topas_status_lbl.config(text=text, fg=color or self.fg_color)

    def _append_topas_log(self, text):
        if not hasattr(self, 'topas_log_txt'):
            return
        self.topas_log_txt.insert(tk.END, text)
        self.topas_log_txt.see(tk.END)

    # ------------------------------------------------------------------
    # ERP 제어 (gui.py의 검증된 로직 그대로)
    # ------------------------------------------------------------------
    def find_and_switch_frame(self, selector):
        by = By.XPATH if selector.startswith('//') or selector.startswith('xpath=') else By.CSS_SELECTOR
        search_val = selector.replace('xpath=', '') if selector.startswith('xpath=') else selector

        try:
            self.driver.switch_to.default_content()
            if len(self.driver.find_elements(by, search_val)) > 0:
                return True
        except Exception:
            pass

        try:
            self.driver.switch_to.default_content()
            iframes = self.driver.find_elements(By.TAG_NAME, 'iframe')
            for iframe in iframes:
                try:
                    self.driver.switch_to.default_content()
                    self.driver.switch_to.frame(iframe)
                    if len(self.driver.find_elements(by, search_val)) > 0:
                        return True
                except Exception:
                    continue
        except Exception:
            pass

        return False

    def is_grid_locked(self):
        mask_selector = ".aui-grid-mask, .aui-grid-loading, .aui-grid-loading-loading"
        for mask in self.driver.find_elements(By.CSS_SELECTOR, mask_selector):
            try:
                if mask.is_displayed():
                    return True
            except Exception:
                continue
        return False

    def wait_until_grid_ready_after_save(self, selectors, timeout):
        timeout = max(1.0, float(timeout))
        poll_interval = self._float_config('erp_poll_interval', 0.25, 0.1, 1.0)
        stable_seconds = self._float_config('erp_grid_ready_stable_seconds', 1.0, poll_interval, 3.0)
        required_unlocked_ticks = max(1, int((stable_seconds / poll_interval) + 0.999))
        unlocked_ticks = 0
        saw_lock = False
        started_at = time.time()

        self.driver.switch_to.default_content()
        self.find_and_switch_frame(selectors["search_date_input"])

        print(f" -> 저장 후 그리드 활성화 대기 중(스마트 폴링, 최대 {int(timeout)}초)...")

        while (time.time() - started_at) < timeout:
            if not self.is_running:
                return False

            time.sleep(poll_interval)
            elapsed = time.time() - started_at

            try:
                is_locked = self.is_grid_locked()
            except Exception as poll_ex:
                print(f" -> [대기 확인 경고] 그리드 차단막 확인 중 예외 발생: {str(poll_ex)}")
                is_locked = False

            if is_locked:
                saw_lock = True
                unlocked_ticks = 0
                print(f" -> [{elapsed:.1f}/{timeout:.0f}초] ERP 처리 중: 그리드 차단막 감지")
                continue

            unlocked_ticks += 1

            if saw_lock:
                print(f" -> [안정화 완료] {elapsed:.1f}초 만에 그리드 차단막이 해제되었습니다.")
                return True

            if unlocked_ticks >= required_unlocked_ticks:
                print(f" -> [안정화 완료] {elapsed:.1f}초 동안 차단막 없이 그리드가 활성 상태입니다.")
                return True

            print(f" -> [{elapsed:.1f}/{timeout:.0f}초] 차단막 없음: 안정 상태 재확인 중")

        print(f" -> [주의] 그리드 차단막 대기 {int(timeout)}초 초과. 현재 상태로 다음 날짜를 진행합니다.")
        return False

    def get_grid_page_state(self):
        """현재 그리드의 페이지 상태를 읽어 (현재페이지, 전체페이지수, 총건수, 다음버튼노출)을 반환한다.

        ERP의 totalPage 변수는 행 바인딩보다 수 초 늦게 채워지므로(실측 5~8초 지연),
        조회 결과 모든 행에 즉시 담기는 totRows로 전체 페이지 수를 직접 계산한다.
        이 방식이 totalPage 지연으로 인한 '2페이지 이후 통째 누락' 버그를 원천 차단한다.
        """
        try:
            page_size = int(self.config.get('grid_page_size', 500)) or 500
        except (TypeError, ValueError):
            page_size = 500
        grid_id = self.config.get('grid_id', '#gridMain')
        js = """
        try {
            var gid = arguments[0];
            var d = (typeof AUIGrid !== 'undefined') ? AUIGrid.getGridData(gid) : null;
            var tot = (d && d.length) ? (d[0].totRows || d.length) : 0;
            var nextBtn = document.querySelector('#paging-button-next');
            return {
                cur: (typeof currentPage !== 'undefined' ? currentPage : 1),
                totRows: tot,
                pageRows: (d ? d.length : 0),
                nextVisible: !!(nextBtn && nextBtn.offsetParent !== null)
            };
        } catch (e) {
            return {cur: 1, totRows: 0, pageRows: 0, nextVisible: false};
        }
        """
        try:
            st = self.driver.execute_script(js, grid_id) or {}
        except Exception as ex:
            print(f" -> [페이징 상태 경고] 그리드 상태 확인 중 예외: {str(ex)}")
            st = {}

        tot = int(st.get('totRows') or 0)
        cur = int(st.get('cur') or 1)
        page_rows = int(st.get('pageRows') or 0)
        next_visible = bool(st.get('nextVisible'))

        if tot > 0:
            total_pages = max(1, -(-tot // page_size))  # ceil(tot / page_size)
        elif page_rows > 0 and next_visible:
            # totRows를 못 읽었는데 다음 버튼이 보이면 멀티페이지로 간주(안전측)
            total_pages = max(2, cur + 1)
        else:
            total_pages = 1

        return cur, total_pages, tot, next_visible

    def navigate_to_grid_page(self, selectors, target_page, timeout):
        """저장 후 1페이지로 리셋된 상태에서 target_page로 직접 재이동한다.

        ERP 내부 moveToPage(goPage,'this','#gridMain_r')를 직접 호출한다.
        next 버튼을 N번 누르는 누적 클릭 방식과 달리, 어느 페이지에서든 목표 페이지로
        한 번에 점프하므로 클릭 누적/락 어긋남으로 인한 페이지 오인이 없다.
        이동 후 currentPage가 목표값과 일치하는지 검증하고, 실패 시 누락을 막기 위해 예외를 던진다.
        """
        rbutton = self.config.get('paging_search_button', '#gridMain_r')
        cur = -1
        for attempt in range(3):
            if not self.is_running:
                return False

            self.driver.switch_to.default_content()
            self.find_and_switch_frame(selectors["search_date_input"])
            self.driver.execute_script(
                "try { moveToPage(arguments[0], 'this', arguments[1]); } catch (e) {}",
                target_page, rbutton
            )
            self.wait_until_grid_ready_after_save(selectors, timeout)

            cur, _, _, _ = self.get_grid_page_state()
            if cur == target_page:
                return True

            print(f" -> [페이지 이동 재시도] 목표 {target_page} / 현재 {cur} (시도 {attempt + 1}/3)")
            time.sleep(0.5)

        raise RuntimeError(
            f"{target_page}페이지로 이동하지 못했습니다(현재 {cur}). "
            f"데이터 누락을 막기 위해 이 기간 처리를 중단합니다."
        )

    def _dismiss_alert_and_modal(self, selectors):
        """떠 있는 네이티브 얼럿(예: 408 서버 지연)과 열린 요금 모달을 정리한다.
        페이지 저장 재시도 전에 화면을 깨끗한 상태로 되돌리기 위해 사용."""
        for _ in range(5):
            try:
                al = self.driver.switch_to.alert
                print(f" -> [얼럿 정리]: {(al.text or '')[:80]}")
                al.accept()
                time.sleep(0.3)
            except Exception:
                break
        try:
            self.driver.switch_to.default_content()
            self.find_and_switch_frame(selectors["cancel_button"])
            cbs = self.driver.find_elements(By.CSS_SELECTOR, selectors["cancel_button"])
            if cbs and cbs[0].is_displayed():
                self.driver.execute_script("arguments[0].click();", cbs[0])
        except Exception:
            pass

    def _save_current_page(self, selectors, wait, inputs_mapping, driver_timeout,
                           erp_short_pause, erp_poll_interval):
        """현재 페이지를 전체선택 → 요금 모달 → 값 주입 → 이중검증 → 저장 → 얼럿 처리 →
        모달 닫기 → 그리드 준비 대기까지 1회 수행한다.

        저장 결과 얼럿이 408/지연/오류성 문구이면 '저장 실패 가능'으로 보고 예외를 던져
        상위 루프가 재시도하도록 한다(서버 지연으로 저장이 누락되는 것을 막기 위함)."""
        # 1) 전체 선택
        header_chk = self.driver.find_element(By.CSS_SELECTOR, selectors["header_all_checkbox"])
        if not header_chk.is_selected():
            self.driver.execute_script("arguments[0].click();", header_chk)
            try:
                WebDriverWait(self.driver, 2, poll_frequency=erp_poll_interval).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, selectors["header_all_checkbox"]).is_selected()
                )
            except Exception:
                time.sleep(erp_short_pause)

        # 2) 요금 모달 열기
        update_btn = self.driver.find_element(By.CSS_SELECTOR, selectors["update_button"])
        self.driver.execute_script("arguments[0].click();", update_btn)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, selectors["adult_air_input"])))

        # 3) 값 주입 (콤마 포매터 때문에 값 한 번에 주입 후 이벤트만 발생)
        for key, val in inputs_mapping.items():
            inp = self.driver.find_element(By.CSS_SELECTOR, selectors[key])
            self.driver.execute_script(
                """
                const el = arguments[0], v = arguments[1];
                el.focus();
                el.value = v;
                el.dispatchEvent(new Event('input',  {bubbles:true}));
                el.dispatchEvent(new Event('keyup',  {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                el.dispatchEvent(new Event('blur',   {bubbles:true}));
                """,
                inp, str(val)
            )
        time.sleep(erp_short_pause)

        # 4) 저장 전 검증
        def _digits_to_int(text):
            digits = "".join(ch for ch in str(text) if ch.isdigit())
            return int(digits) if digits else 0

        mismatches = []
        for key, val in inputs_mapping.items():
            inp = self.driver.find_element(By.CSS_SELECTOR, selectors[key])
            actual_raw = self.driver.execute_script("return arguments[0].value;", inp)
            if _digits_to_int(actual_raw) != _digits_to_int(val):
                mismatches.append(f"{key}: 기대 {_digits_to_int(val)} / 실제 '{actual_raw}'")
        if mismatches:
            try:
                cancel_btns = self.driver.find_elements(By.CSS_SELECTOR, selectors["cancel_button"])
                if cancel_btns:
                    self.driver.execute_script("arguments[0].click();", cancel_btns[0])
            except Exception:
                pass
            raise RuntimeError("입력값 검증 실패(저장 안 함) - " + "; ".join(mismatches))

        # 5) 저장
        save_btn = self.driver.find_element(By.CSS_SELECTOR, selectors["save_button"])
        self.driver.execute_script("arguments[0].click();", save_btn)

        # 6) 저장 확인/결과 얼럿 처리 — 오류성 얼럿이면 저장 실패로 간주
        error_keywords = ('408', '지연', '연결이 원활', '관리자에게 문의', '오류가', '실패')
        save_alert_error = None
        for _ in range(4):
            try:
                alert = self.driver.switch_to.alert
                atext = alert.text or ''
                print(f" -> [얼럿 감지]: {atext}")
                if any(k in atext for k in error_keywords):
                    save_alert_error = atext
                    alert.accept()
                    break
                alert.accept()
                time.sleep(erp_short_pause)
            except Exception:
                break
        if save_alert_error:
            raise RuntimeError(f"서버 응답 지연/오류로 저장이 반영되지 않았을 수 있음: {save_alert_error[:120]}")

        # 7) 모달 닫기
        try:
            self.driver.switch_to.default_content()
            self.find_and_switch_frame(selectors["cancel_button"])
            cancel_btn = self.driver.find_elements(By.CSS_SELECTOR, selectors["cancel_button"])
            if cancel_btn and cancel_btn[0].is_displayed():
                self.driver.execute_script("arguments[0].click();", cancel_btn[0])
        except Exception:
            pass

        # 8) 저장 후 그리드 락 해제 및 재조회(1페이지 리셋) 완료 대기
        self.wait_until_grid_ready_after_save(selectors, driver_timeout)

    def launch_debug_chrome(self):
        base_dir = get_app_dir()
        bat_path = os.path.join(base_dir, 'chrome_debug.bat')
        if not os.path.exists(bat_path):
            messagebox.showerror('파일 없음', 'chrome_debug.bat 파일을 찾을 수 없습니다.')
            return

        target_url, target_name = self._get_debug_browser_target()
        self.chrome_launch_btn.config(state=tk.DISABLED, text='여는 중…')
        if target_name == 'TOPAS':
            self._set_topas_status('TOPAS 브라우저를 여는 중입니다…', self.accent_orange)
        else:
            self.set_status('ERP 브라우저를 여는 중입니다…', self.accent_orange)
        self.root.update_idletasks()

        def _restore():
            self.chrome_launch_btn.config(state=tk.NORMAL, text='브라우저 켜기')
            if target_name == 'TOPAS':
                self._set_topas_status('TOPAS 브라우저 준비 완료 · 로그인 후 첫 조회를 실행하세요', self.accent_green)
            else:
                self.set_status('ERP 브라우저 준비 완료 · 로그인 후 시작하세요', self.accent_green)

        try:
            subprocess.Popen([bat_path, target_url], creationflags=subprocess.CREATE_NEW_CONSOLE)
            print(f'[브라우저 기동] 디버깅용 크롬 브라우저 기동 명령을 전달했습니다. ({target_name})')
            self.root.after(3000, _restore)
        except Exception as e:
            print(f'[오류] 크롬 기동 실패: {str(e)}')
            self.chrome_launch_btn.config(state=tk.NORMAL, text='브라우저 켜기')
            if target_name == 'TOPAS':
                self._set_topas_status('크롬 기동 실패', self.accent_red)
            else:
                self.set_status('크롬 기동 실패', self.accent_red)
            messagebox.showerror('기동 오류', f'크롬 기동에 실패했습니다:\n{str(e)}')

    def rpa_worker_loop(self):
        rpa_history = []

        try:
            self.root.after(0, lambda: self.show_loading('ERP에 연결하고 있어요…'))
            print("[RPA 시작] Selenium 웹드라이버 연결 중...")
            options = webdriver.ChromeOptions()
            debug_addr = self.config.get("debugger_address", self.config.get("debuggerAddress", "127.0.0.1:9222"))
            options.add_experimental_option("debuggerAddress", debug_addr)

            driver_timeout = int(self.config.get("timeout", 25))
            erp_poll_interval = self._float_config('erp_poll_interval', 0.25, 0.1, 1.0)
            erp_short_pause = self._float_config('erp_short_pause', 0.2, 0.05, 0.5)

            try:
                # Selenium Manager가 로컬 크롬 버전에 맞춰 드라이버를 자동 제어하도록 서비스 오버헤드 제거
                self.driver = webdriver.Chrome(options=options)
            except Exception as connect_ex:
                # except 변수는 블록 종료 시 사라지므로 메시지를 미리 캡처해 람다 기본값으로 넘긴다
                err_text = str(connect_ex)
                print(f"[오류] 브라우저 연결 실패: {err_text}")
                print(" -> 크롬 브라우저가 디버깅 모드로 켜져 있는지 확인하고 모든 일반 크롬 창을 닫아주세요.")
                self.root.after(0, lambda msg=err_text: messagebox.showerror("연결 실패", f"디버그 브라우저 연결에 실패했습니다.\n{msg}"))
                return

            self.root.after(0, self.hide_loading)

            selectors = self.config["selectors"]
            required_selector_keys = [
                "search_date_input",
                "search_date_end_input",
                "search_button",
                "header_all_checkbox",
                "update_button",
                "adult_air_input",
                "adult_hotel_input",
                "adult_land_input",
                "adult_tour_input",
                "adult_profit_input",
                "child_fare_input",
                "infant_fare_input",
                "save_button",
                "cancel_button",
                "date_cell_in_row",
            ]
            missing_selector_keys = [key for key in required_selector_keys if key not in selectors]
            if missing_selector_keys:
                err_text = "config.json selectors 누락: " + ", ".join(missing_selector_keys)
                print(f"[오류] {err_text}")
                self.root.after(0, lambda msg=err_text: messagebox.showerror("설정 오류", msg))
                return

            total_items = len(self.fares_data)

            print(f"[RPA 정보] 총 {total_items}개의 날짜 데이터 처리를 시작합니다.")

            for index, row in enumerate(self.fares_data):
                if not self.is_running:
                    print("[RPA 알림] 사용자가 작업을 강제 중단했습니다.")
                    break

                while self.is_paused and self.is_running:
                    time.sleep(0.5)

                if not self.is_running:
                    break

                date_val = str(row["date"]).strip()
                date_end_val = str(row.get("date_end", date_val)).strip()
                date_log_str = f"{date_val} ~ {date_end_val}" if date_val != date_end_val else date_val

                adult_air = str(row.get("adult_air", 0)).strip()
                adult_hotel = str(row.get("adult_hotel", 0)).strip()
                adult_land = str(row.get("adult_land", 0)).strip()
                adult_tour = str(row.get("adult_tour", 0)).strip()
                adult_profit = str(row.get("adult_profit", 0)).strip()
                child_val = str(row.get("child_fare", 0)).strip()
                infant_val = str(row.get("infant_fare", 0)).strip()

                print(f"\n========================================================")
                print(f"[{index+1}/{total_items}] 대상 날짜: {date_log_str}")
                print(f" -> 입력 데이터: 성인(항공={adult_air}, 호텔={adult_hotel}, 지상={adult_land}, 경비={adult_tour}, 수익={adult_profit}), 소아={child_val}, 유아={infant_val}")

                if not self.find_and_switch_frame(selectors["search_date_input"]):
                    err_msg = "출발일자 입력 필드를 찾을 수 없습니다."
                    print(f" -> [오류] {err_msg}")
                    rpa_history.append({"date": date_log_str, "status": "FAIL", "error": err_msg})
                    self.update_progress_ui(index + 1, total_items)
                    continue

                try:
                    st_date_input = self.driver.find_element(By.CSS_SELECTOR, selectors["search_date_input"])

                    # 1) 시작일 입력 후, datepicker 의 '시작일 변경 시 종료일을 시작일로 맞추는'
                    #    자동 리셋 동작을 먼저 발생시킨다(change/blur). 그래야 그 다음에 넣는
                    #    종료일이 리셋에 덮이지 않는다.
                    self.driver.execute_script("arguments[0].value = '';", st_date_input)
                    st_date_input.send_keys(date_val)
                    self.driver.execute_script(
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                        st_date_input)

                    # 2) 종료일을 마지막에 입력하고, 위젯이 값을 무시/리셋하는 경우를 대비해
                    #    실제 표시값을 다시 읽어 일치할 때까지 최대 4회 재입력한다.
                    def _read_end_value():
                        try:
                            el = self.driver.find_element(By.CSS_SELECTOR, selectors["search_date_end_input"])
                            return (el.get_attribute('value') or '').strip().replace('.', '-').replace('/', '-')
                        except Exception:
                            return ''

                    for _ in range(4):
                        en_date_input = self.driver.find_element(By.CSS_SELECTOR, selectors["search_date_end_input"])
                        self.driver.execute_script("arguments[0].value = '';", en_date_input)
                        en_date_input.send_keys(date_end_val)
                        self.driver.execute_script(
                            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));"
                            "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));",
                            en_date_input)
                        if _read_end_value() == date_end_val:
                            break
                        time.sleep(0.2)

                    if _read_end_value() != date_end_val:
                        print(f" -> [경고] 종료일이 ERP에 '{_read_end_value()}'(으)로 남아 기대값({date_end_val})과 다릅니다. "
                              f"해당 기간 일부 상품이 조회/수정에서 누락될 수 있으니 ERP에서 직접 확인해 주세요.")

                    search_btn = self.driver.find_element(By.CSS_SELECTOR, selectors["search_button"])
                    self.driver.execute_script("arguments[0].click();", search_btn)
                    print(" -> 조회 버튼을 클릭했습니다. 조회 데이터 로딩 대기 중...")

                    wait = WebDriverWait(self.driver, driver_timeout, poll_frequency=erp_poll_interval)
                    norm_date = date_val.replace('-', '').replace('.', '').replace('/', '')
                    norm_date_end = date_end_val.replace('-', '').replace('.', '').replace('/', '')
                    # 조회 완료 판정: AUIGrid 데이터 모델의 startDay(출발일)로 판정한다.
                    # (구버전은 .aui-grid-default-column DOM 텍스트에서 날짜 형태를 모두 긁어
                    #  범위 검사했는데, 한 행에 출발일 외 부가 날짜 컬럼이 섞여 있으면 단일일/
                    #  좁은 기간 조회에서 그 부가 날짜가 범위를 벗어나 '조회결과 없음'으로 잘못
                    #  건너뛰는 false negative가 있었다. 데이터 모델의 startDay만 보면 해결되고,
                    #  행 바인딩 지연도 폴링으로 흡수된다.)
                    grid_id = self.config.get('grid_id', '#gridMain')
                    start_day_js = (
                        "try {"
                        "  var a = (typeof AUIGrid!=='undefined') ? AUIGrid.getGridData(arguments[0]) : null;"
                        "  if (!a) return null;"
                        "  var out = [];"
                        "  for (var i=0;i<a.length;i++){ out.push(String(a[i].startDay==null?'':a[i].startDay)); }"
                        "  return out;"
                        "} catch(e){ return null; }"
                    )
                    matched = False
                    deadline = time.time() + driver_timeout

                    while time.time() < deadline:
                        if not self.is_running:
                            break

                        try:
                            start_days = self.driver.execute_script(start_day_js, grid_id)
                        except Exception:
                            start_days = None

                        if start_days:
                            norm_days = [
                                str(s).replace('-', '').replace('.', '').replace('/', '')[:8]
                                for s in start_days if s
                            ]
                            if norm_days and all(norm_date <= d <= norm_date_end for d in norm_days):
                                matched = True
                                break
                        time.sleep(erp_poll_interval)

                    if not matched:
                        print(f" -> [조회결과 없음] {date_log_str} 일자 데이터를 반영하지 못했습니다. ERP에서 직접 날짜 조회를 확인해 주세요.")
                        rpa_history.append({"date": date_log_str, "status": "SKIP", "error": "조회결과 없음"})
                        self.update_progress_ui(index + 1, total_items)
                        continue

                    # === 페이징 처리: 조회 결과 500건 초과 시 모든 페이지를 순회하며 저장 ===
                    # totalPage 변수는 행 바인딩보다 늦게 채워지므로(실측 지연), 행에 즉시 담기는
                    # totRows로 전체 페이지 수를 계산한다. 저장하면 그리드가 1페이지로 리셋되므로
                    # 다음 페이지는 moveToPage 직접 점프로 재이동한다.
                    _cur, total_pages, tot_rows, _next_vis = self.get_grid_page_state()
                    if total_pages > 1:
                        print(f" -> [페이징] 총 {tot_rows}건 · 약 {total_pages}페이지 감지. 페이지별로 순차 저장합니다.")

                    inputs_mapping = {
                        "adult_air_input": adult_air,
                        "adult_hotel_input": adult_hotel,
                        "adult_land_input": adult_land,
                        "adult_tour_input": adult_tour,
                        "adult_profit_input": adult_profit,
                        "child_fare_input": child_val,
                        "infant_fare_input": infant_val
                    }

                    # 페이지별 재시도/페이싱 설정 (서버 일시 지연 408 등 대응)
                    page_max_retries = max(1, int(self.config.get('page_max_retries', 3)))
                    page_pause = self._float_config('erp_page_pause', 1.0, 0.0, 10.0)

                    target_page = 1
                    pages_done = 0
                    pages_failed = []

                    while target_page <= total_pages:
                        if not self.is_running:
                            print(" -> [중단] 사용자 중지로 남은 페이지 처리를 멈춥니다.")
                            break

                        page_ok = False
                        page_err = None
                        for attempt in range(1, page_max_retries + 1):
                            if not self.is_running:
                                break
                            try:
                                # 저장하면 1페이지로 리셋되고, 재시도 시 위치가 틀어졌을 수 있으므로
                                # 2페이지 이후 또는 재시도일 때는 목표 페이지로 직접 재이동한다.
                                if target_page > 1 or attempt > 1:
                                    if attempt == 1:
                                        print(f" -> [페이지 이동] {target_page}/{total_pages}페이지로 재이동합니다...")
                                    self.navigate_to_grid_page(selectors, target_page, driver_timeout)

                                if total_pages > 1:
                                    suffix = f" (재시도 {attempt}/{page_max_retries})" if attempt > 1 else ""
                                    print(f" -> [페이지 {target_page}/{total_pages}] 전체선택 → 요금 입력 → 저장{suffix}")

                                self._save_current_page(
                                    selectors, wait, inputs_mapping, driver_timeout,
                                    erp_short_pause, erp_poll_interval
                                )
                                page_ok = True
                                break
                            except Exception as page_ex:
                                page_err = str(page_ex).replace("\n", " ")
                                print(f" -> [페이지 {target_page} 오류] {page_err[:160]}")
                                # 떠 있을 수 있는 얼럿/모달 정리 후 재시도
                                self._dismiss_alert_and_modal(selectors)
                                if attempt < page_max_retries and self.is_running:
                                    backoff = page_pause + attempt
                                    print(f" -> [재시도 대기] {backoff:.1f}초 후 {target_page}페이지 다시 시도합니다…")
                                    time.sleep(backoff)

                        if page_ok:
                            pages_done += 1
                            if total_pages > 1:
                                print(f" -> [저장 완료] {date_log_str} {target_page}/{total_pages}페이지")
                        else:
                            pages_failed.append(target_page)
                            print(f" -> [페이지 실패] {date_log_str} {target_page}/{total_pages}페이지 — {page_max_retries}회 시도 모두 실패")

                        target_page += 1
                        # 서버 과부하(408) 완화를 위해 페이지 사이에 짧은 텀
                        if page_pause > 0 and self.is_running and target_page <= total_pages:
                            time.sleep(page_pause)

                    if not self.is_running:
                        raise RuntimeError(f"사용자 중단: {date_log_str} {pages_done}/{total_pages}페이지까지 처리 후 멈춤")

                    if pages_failed:
                        failed_str = ', '.join(str(p) for p in pages_failed)
                        msg = (
                            f"{total_pages}페이지 중 {len(pages_failed)}개 페이지 저장 실패"
                            f"(실패 페이지: {failed_str} / 성공 {pages_done}페이지). "
                            f"서버 지연 등 일시 오류일 수 있으니 이 날짜를 다시 실행해 주세요."
                        )
                        print(f" -> [부분 실패] {date_log_str}: {msg}")
                        rpa_history.append({"date": date_log_str, "status": "FAIL", "error": msg})
                    else:
                        if total_pages > 1:
                            print(f" -> [성공] {date_log_str} 요금 업데이트 완료 (전체 {pages_done}/{total_pages}페이지)")
                        else:
                            print(f" -> [성공] {date_log_str} 요금 업데이트 완료")
                        rpa_history.append({"date": date_log_str, "status": "SUCCESS", "error": ""})

                except Exception as row_ex:
                    err_msg = str(row_ex).replace("\n", " ")
                    print(f" -> [실패] 오류 발생: {err_msg}")
                    rpa_history.append({"date": date_log_str, "status": "FAIL", "error": err_msg})

                    try:
                        self.driver.switch_to.alert.accept()
                    except Exception:
                        pass

                self.update_progress_ui(index + 1, total_items)

            print("\n========================================================")
            print("[RPA 종료] 전체 날짜에 대한 작업이 마무리되었습니다.")
            print("========================================================\n")

        except Exception as loop_ex:
            print(f"[치명적 오류] 작업 루프 처리 실패: {str(loop_ex)}")

        finally:
            history_snapshot = list(rpa_history)
            try:
                self.root.after(0, lambda: self.finish_rpa_on_ui(history_snapshot))
            except Exception:
                pass

    def update_progress_ui(self, current, total):
        pct = int((current / total) * 100) if total else 0
        self.root.after(0, lambda: self.set_status(f'진행 중 ({pct}%)', self.accent_orange))
        self.root.after(0, lambda: self.progress_lbl.config(text=f'{current} / {total} ({pct}%)'))
        self.root.after(0, lambda: self.progress_bar.config(value=current))

    def generate_and_show_report(self, history):
        if not history:
            return

        total_cnt = len(history)
        success_cnt = sum(1 for x in history if x["status"] == "SUCCESS")
        fail_cnt = total_cnt - success_cnt

        sorted_history = sorted(history, key=lambda x: x["date"])

        single_failures = []
        consecutive_failures = []

        current_streak = []
        for x in sorted_history:
            if x["status"] != "SUCCESS":
                current_streak.append(x)
            else:
                if current_streak:
                    if len(current_streak) == 1:
                        single_failures.append(current_streak[0])
                    else:
                        consecutive_failures.append(current_streak)
                    current_streak = []
        if current_streak:
            if len(current_streak) == 1:
                single_failures.append(current_streak[0])
            else:
                consecutive_failures.append(current_streak)

        report = []
        report.append("========================================================")
        report.append("                 요금 업데이트 결과 보고서")
        report.append("========================================================")

        report.append("\n[단발성 실패]")
        if single_failures:
            for item in single_failures:
                reason = item["error"] if item["error"] else "조회결과 없음"
                report.append(f"- {item['date']} (오류: {reason})")
        else:
            report.append("- 없음")

        report.append("\n[연속 실패]")
        if consecutive_failures:
            for streak in consecutive_failures:
                start_date = streak[0]["date"]
                end_date = streak[-1]["date"]
                days = len(streak)
                errors = [s["error"] for s in streak if s["error"]]
                rep_err = max(set(errors), key=errors.count) if errors else "조회결과 없음"
                report.append(f"- {start_date} ~ {end_date} ({days}일간 연속 실패 / 대표오류: {rep_err})")
        else:
            report.append("- 없음")

        report.append("\n[요약 통계]")
        report.append(f"- 전체 대상: {total_cnt}일")
        report.append(f"- 성공: {success_cnt}일")
        report.append(f"- 실패/스킵: {fail_cnt}일")
        report.append("========================================================\n")

        print("\n".join(report))

        # 실패/스킵이 있으면 콘솔 요약으로만 끝내지 않고 팝업으로 분명히 경고한다.
        # 요금 미수정은 매출에 직결되는 사안이라, 사용자가 결과를 놓치지 않도록 한다.
        failed_items = [x for x in sorted_history if x["status"] != "SUCCESS"]
        try:
            if failed_items:
                lines = []
                for item in failed_items[:20]:
                    reason = item["error"] if item.get("error") else "조회결과 없음"
                    lines.append(f"• {item['date']}  ({reason})")
                more = len(failed_items) - 20
                if more > 0:
                    lines.append(f"… 외 {more}건")
                msg = (
                    f"전체 {total_cnt}일 중 {fail_cnt}일이 수정되지 않았습니다.\n"
                    f"(성공 {success_cnt}일 / 실패·스킵 {fail_cnt}일)\n\n"
                    "아래 날짜는 요금이 반영되지 않았습니다.\n"
                    "ERP에서 직접 확인하거나 해당 날짜만 다시 실행해 주세요.\n\n"
                    + "\n".join(lines)
                )
                messagebox.showwarning(f"⚠️ 요금 수정 실패 {fail_cnt}건 — 확인 필요", msg)
            else:
                messagebox.showinfo(
                    "요금 수정 완료",
                    f"전체 {total_cnt}일을 모두 정상적으로 수정했습니다."
                )
        except Exception:
            pass


def _acquire_single_instance_lock():
    """중복 실행 방지용 Windows 네임드 뮤텍스를 잡는다.
    테스트 버전은 운영본(gui.py)과 충돌하지 않도록 별도 뮤텍스명을 사용한다."""
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32
        mutex_name = "Global\\NaeilERPUpdaterV2Test_SingleInstance_Mutex"
        handle = kernel32.CreateMutexW(None, wintypes.BOOL(True), mutex_name)
        ERROR_ALREADY_EXISTS = 183
        if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
            return False, handle
        return True, handle
    except Exception:
        return True, None


if __name__ == "__main__":
    try:
        if sys.stdout is not None and getattr(sys.stdout, 'encoding', None) != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    if sys.stdout is None:
        sys.stdout = NullStream()

    acquired, _mutex_handle = _acquire_single_instance_lock()
    if not acquired:
        try:
            _alert = tk.Tk()
            _alert.withdraw()
            messagebox.showinfo("이미 실행 중", "프로그램(테스트 버전)이 이미 실행 중입니다.\n열려 있는 창을 확인해 주세요.")
            _alert.destroy()
        except Exception:
            pass
        sys.exit(0)

    root_win = tk.Tk()
    app = RpaGuiApp(root_win)

    def on_window_close():
        if app.is_running:
            if messagebox.askyesno("앱 종료", "현재 RPA가 작동 중입니다. 중지하고 프로그램을 종료하시겠습니까?"):
                app.is_running = False
                root_win.destroy()
        else:
            root_win.destroy()

    root_win.protocol("WM_DELETE_WINDOW", on_window_close)
    root_win.mainloop()
