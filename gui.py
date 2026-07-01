# -*- coding: utf-8 -*-
"""
Naeil Tour ERP 요금 업데이트 RPA — v4.0.1

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
import copy
import ast
import operator
import re
import threading
import queue
import shutil
import subprocess
import tempfile
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
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
from fare.calculator import calculate_round_trips, js_round
from fare.exporter import export_results_to_excel, results_to_tsv, to_erp_rows
from fare.parser import parse_topas_text, summarize_records
from fare.route_select import filter_routes, infer_route_from_topas_text
from fare.store import load_fare_snapshot
from topas.availability import parse_availability_text
from topas.collector import join_raw_blocks, save_raw_backup

APP_VERSION = "v5.0.9"
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
AIRLINE_EMPTY_LABEL = "_선택_"
RETURN_AC1_SUGGEST_PERCENT = 110
DEFAULT_AIRLINE_CHOICES = [
    ("", AIRLINE_EMPTY_LABEL),
]


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
        self.win_height = 906
        self.panel_width = 878
        self.collapsed_width = 1006
        self.expanded_width = self.collapsed_width + self.panel_width
        self.panel_expanded = False
        self.root.geometry(f'{self.collapsed_width}x{self.win_height}')
        self.root.minsize(self.collapsed_width, 774)

        # --- 색상 팔레트 (V5 다크 디자인 토큰) ---
        self.bg_color = '#0e1117'          # 앱 배경
        self.card_color = '#161b24'        # 카드 배경
        self.card_hover = '#1f2632'        # 카드 호버
        self.fg_color = '#e8ecf4'          # 본문 텍스트
        self.fg_muted = '#8b95a7'          # 보조 텍스트
        self.accent_color = '#4c8dff'      # 주 액션 (블루)
        self.accent_hover = '#3a77e6'
        self.accent_green = '#2fae6f'      # 실행 (그린)
        self.accent_green_hover = '#27955e'
        self.accent_orange = '#e3a23d'     # 주의 (앰버)
        self.accent_orange_hover = '#c98a28'
        self.accent_red = '#e05d56'        # 위험 (레드)
        self.accent_red_hover = '#c64b45'
        self.border_color = '#252d3b'      # 카드 테두리
        # 파생 토큰
        self.input_bg = '#0a0d13'          # 입력창/콘솔 배경
        self.input_fg = '#ccd4e0'          # 콘솔 텍스트
        self.btn_neutral_bg = '#222a37'    # 보조(중립) 버튼
        self.btn_neutral_hover = '#2c3646'
        self.btn_neutral_fg = '#cfd7e4'
        self.tab_selected_bg = '#1c2433'   # 선택된 탭 배경
        self.tree_bg = '#11161f'           # 결과표 배경
        self.tree_row_alt = '#151b26'      # 결과표 줄무늬
        self.tree_head_bg = '#1a212d'      # 결과표/입력표 헤더
        self.disabled_fg = '#4d5666'       # 비활성 텍스트

        self.root.configure(bg=self.bg_color)
        self._setup_styles()

        self.config = self.load_config()
        self.fares_data = []
        self.job_queue = []
        self.rpa_jobs_to_run = []
        self.current_rpa_job = None
        self.editing_job_index = None
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
        self._topas_window_handle = None
        self.topas_current_ac1_count = 0
        self.departure_ac1_count = 0
        self.current_main_tab = 'fare'
        self.main_tab_controls = {}
        self.topas_target_slot = 'departure'
        self.fare_snapshot = None
        self.fare_routes = []
        self.route_var = tk.StringVar(value='')
        self.route_user_modified = False
        self.custom_night_var = tk.StringVar(value='')
        self.airline_var = tk.StringVar(value=AIRLINE_EMPTY_LABEL)
        self.price_desc_var = tk.StringVar(value='')
        self.hotel_name_var = tk.StringVar(value='')
        self.progress_text_var = tk.StringVar(value='')
        self.airline_choices = []
        self.airline_value_by_label = {}
        self.airline_label_by_value = {}
        self.airline_refresh_thread = None
        self.hotel_choices = []
        self.hotel_record_by_label = {}
        self.hotel_record_by_seq = {}
        self.hotel_search_thread = None
        self.hotel_search_after_id = None
        self.hotel_search_request_id = 0
        self.hotel_result_records = []
        self.current_hotel_seq = ''
        self.selected_airline_code = ''
        self.selected_price_desc = ''
        self.selected_hotel_name = ''
        self.selected_hotel_seq = ''
        self.selected_progress_text = ''
        self.night_vars = {}
        self.night_chip_buttons = {}
        self.v5_calculation_result = None
        self.selected_result_night = None
        self.last_topas_backup_paths = []
        self._night_auto_recalc_busy = False
        self._merged_fare_restore_data = None
        self._merged_fare_merged_data = None

        # 수식 입력줄(formula bar) / 클릭 참조 상태
        self.fb_var = tk.StringVar()
        self._active_cell = (0, 0)
        self._ref_mode = False
        self._loading_fb = False
        # 수식 보관소: 셀에는 계산 결과를 표시하고, 식은 여기에 (row,col)->'=...' 로 보관
        self.formulas = {}
        self._results = {}        # (row,col)->마지막 계산 결과 문자열 (수동 수정 감지용)
        self._recalc_busy = False  # 재계산 중 재진입 방지
        self._sheet_undo_stack = []
        self._sheet_redo_stack = []
        self._sheet_last_snapshot = None
        self._applying_sheet_snapshot = False
        self._sheet_undo_limit = 80

        # 날짜 필터 관리 변수
        self.filter_mode = tk.StringVar(value="ALL")  # ALL, FROM_DATE, SPECIFIC, DATE_RANGE
        self.filter_value = tk.StringVar(value="")
        self.filter_value_end = tk.StringVar(value="")
        self.filter_panel_expanded = False
        self.filter_mode.trace_add("write", self._on_filter_mode_change)
        self.filter_value.trace_add("write", self._on_filter_value_change)
        self.filter_value_end.trace_add("write", self._on_filter_value_change)

        # 로딩 애니메이션 상태
        self._loading_frames = []
        self._loading_idx = 0
        self._loading_after_id = None
        self._loading_overlay = None

        self._set_airline_choices(DEFAULT_AIRLINE_CHOICES)
        self.build_ui()
        self._sync_sheet_undo_baseline()
        self._on_filter_mode_change()
        self._load_loading_frames()
        self.refresh_count()
        self.set_status("오른쪽 위 ‘요금직접입력하기 ▶’ 버튼으로 요금을 입력하세요", self.fg_muted)
        self.root.after(1200, lambda: self.refresh_airline_options_from_erp(silent=True))
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
            'airline_select': '#air2Cd',
            'price_desc_input': '#priceDesc',
            'hotel_name_input': '#hotelKorNm',
            'hotel_seq_input': '#hotelSeq',
            'event_modify_button': '#eventModify',
            'bulk_update_save_button': '#appSave',
            'progress_status_checkbox': '#procCdChk',
            'progress_status_select': '#procCd',
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
            'topas_debugger_address': '127.0.0.1:9222',
            'erp_debugger_address': '127.0.0.1:9223',
            'topas_chrome_profile_dir': 'ChromeProfile',
            'erp_chrome_profile_dir': 'ChromeProfile_ERP',
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
            'topas_batch_timeout': 80,
            'topas_parse_tail_chars': 80000,
            'topas_batch_size': 10,
            'topas_raw_dir': 'logs/topas_raw',
            'fare_cache_path': 'cache/fares_snapshot.json',
            'anomaly_y_ratio': [0.2, 1.5],
            'firestore': {
                'project_id': 'fare-calculator-2026',
                'api_key': 'AIzaSyAm-_kZz9Kn4WgeJyVMEyV_TZ3Za2uouFs',
            },
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
        # 진행바: 슬림한 형태
        style.configure(
            'Accent.Horizontal.TProgressbar',
            troughcolor=self.input_bg,
            bordercolor=self.input_bg,
            background=self.accent_green,
            lightcolor=self.accent_green,
            darkcolor=self.accent_green,
            thickness=7,
        )
        try:
            style.layout('Result.TNotebook.Tab', [])
            style.configure('Result.TNotebook', tabmargins=0, borderwidth=0, background=self.card_color)
        except Exception:
            pass
        # 결과표 (Treeview) 다크
        style.configure(
            'Treeview',
            background=self.tree_bg,
            fieldbackground=self.tree_bg,
            foreground='#dfe5f0',
            rowheight=30,
            borderwidth=0,
            relief='flat',
            font=('맑은 고딕', 10),
        )
        style.configure(
            'Treeview.Heading',
            background=self.tree_head_bg,
            foreground=self.fg_muted,
            font=('맑은 고딕', 9, 'bold'),
            borderwidth=0,
            relief='flat',
            padding=(8, 7),
        )
        style.map(
            'Treeview.Heading',
            background=[('active', '#212a38')],
            foreground=[('active', self.fg_color)],
        )
        style.map(
            'Treeview',
            background=[('selected', '#28354e')],
            foreground=[('selected', '#ffffff')],
        )
        # 스크롤바 다크 (Windows 기본 스크롤바는 색을 무시하므로 ttk 사용)
        for orient in ('Vertical', 'Horizontal'):
            style.configure(
                f'Dark.{orient}.TScrollbar',
                background=self.btn_neutral_bg,
                troughcolor=self.input_bg,
                bordercolor=self.input_bg,
                arrowcolor=self.fg_muted,
                relief='flat',
                gripcount=0,
            )
            style.map(
                f'Dark.{orient}.TScrollbar',
                background=[('active', self.btn_neutral_hover), ('pressed', self.btn_neutral_hover)],
                arrowcolor=[('active', self.fg_color)],
            )
        # 콤보박스 다크
        style.configure(
            'TCombobox',
            fieldbackground=self.input_bg,
            background=self.btn_neutral_bg,
            foreground=self.fg_color,
            arrowcolor=self.fg_muted,
            bordercolor=self.border_color,
            lightcolor=self.border_color,
            darkcolor=self.border_color,
            insertcolor=self.fg_color,
            padding=(8, 4),
        )
        style.map(
            'TCombobox',
            fieldbackground=[('readonly', self.input_bg), ('focus', self.input_bg)],
            foreground=[('disabled', self.disabled_fg)],
            bordercolor=[('focus', self.accent_color)],
            lightcolor=[('focus', self.accent_color)],
            darkcolor=[('focus', self.accent_color)],
            arrowcolor=[('active', self.fg_color)],
        )
        # 콤보박스 드롭다운 목록 다크
        try:
            self.root.option_add('*TCombobox*Listbox.background', self.input_bg)
            self.root.option_add('*TCombobox*Listbox.foreground', self.fg_color)
            self.root.option_add('*TCombobox*Listbox.selectBackground', self.accent_color)
            self.root.option_add('*TCombobox*Listbox.selectForeground', '#ffffff')
            self.root.option_add('*TCombobox*Listbox.borderWidth', 0)
        except Exception:
            pass
        self._style = style

    def _make_dark_text(self, parent, **text_kwargs):
        """tk.Text + 다크 ttk 스크롤바 묶음을 만들어 (프레임, 텍스트) 를 돌려준다."""
        bg = text_kwargs.get('bg', self.input_bg)
        frame = tk.Frame(parent, bg=bg)
        vbar = ttk.Scrollbar(frame, orient='vertical', style='Dark.Vertical.TScrollbar')
        text = tk.Text(frame, yscrollcommand=vbar.set, **text_kwargs)
        vbar.config(command=text.yview)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        return frame, text

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
        title_label = tk.Label(title_row, text='NaeilERPUpdater V5', font=('맑은 고딕', 16, 'bold'), bg=self.bg_color, fg=self.fg_color)
        title_label.pack(side=tk.LEFT)
        version_badge = tk.Label(title_row, text=f' {APP_VERSION} ', font=('맑은 고딕', 8), bg=self.card_color, fg=self.fg_muted)
        version_badge.pack(side=tk.LEFT, padx=(10, 0), pady=(5, 0))

        def _ghost_link(parent, text, command, hover_fg):
            btn = tk.Button(
                parent, text=text, font=('맑은 고딕', 8, 'bold'),
                bg=self.bg_color, fg=self.fg_muted,
                activebackground=self.bg_color, activeforeground=hover_fg,
                bd=0, relief=tk.FLAT, cursor='hand2', padx=6, pady=2,
                command=command)
            self._add_hover(btn, self.bg_color, self.bg_color, normal_fg=self.fg_muted, hover_fg=hover_fg)
            return btn

        fare_site_btn = _ghost_link(title_row, '요금조회 사이트 ↗', self.open_fare_site, self.accent_color)
        fare_site_btn.pack(side=tk.LEFT, padx=(14, 0), pady=(5, 0))

        guide_btn = _ghost_link(title_row, '가이드 매뉴얼 ↗', self.open_guide_site, self.accent_color)
        guide_btn.pack(side=tk.LEFT, padx=(4, 0), pady=(5, 0))

        template_btn = _ghost_link(title_frame, '엑셀 양식 다운받기 ↗', self.download_excel_template, self.accent_color)
        template_btn.pack(anchor=tk.W, pady=(4, 0))

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

        self._make_toolbar_btn(toolbar, '전체 지우기', self.btn_neutral_bg, self.btn_neutral_hover, self.clear_sheet, fg=self.btn_neutral_fg)
        self._make_toolbar_btn(toolbar, '실행취소', self.btn_neutral_bg, self.btn_neutral_hover, self.undo_sheet, fg=self.btn_neutral_fg)
        self._make_toolbar_btn(toolbar, '다시실행', self.btn_neutral_bg, self.btn_neutral_hover, self.redo_sheet, fg=self.btn_neutral_fg)
        self._make_toolbar_btn(toolbar, '요금구간 병합', self.accent_orange, self.accent_orange_hover, self.merge_sheet_fare_ranges)
        self._make_toolbar_btn(toolbar, '엑셀로 다운받기', self.accent_color, self.accent_hover, self.export_sheet_to_excel)
        self._make_toolbar_btn(toolbar, '작업 목록에 추가', self.btn_neutral_bg, self.btn_neutral_hover, self.add_current_sheet_to_job_queue, fg=self.btn_neutral_fg)

        self.period_mode_var = tk.BooleanVar(value=False)
        self.period_mode_cb = tk.Checkbutton(
            toolbar, text="시작일/종료일 분리 (기간 입력)",
            variable=self.period_mode_var,
            bg=self.card_color, fg=self.fg_color,
            activebackground=self.card_color, activeforeground=self.fg_color,
            selectcolor=self.input_bg,
            font=('맑은 고딕', 9),
            command=self.on_toggle_period_mode
        )
        self.period_mode_cb.pack(side=tk.RIGHT, padx=(10, 0))

        search_filter_bar = tk.Frame(sheet_card, bg=self.card_color)
        search_filter_bar.pack(fill=tk.X, pady=(0, 6))
        tk.Label(
            search_filter_bar,
            text='항공사코드',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_color,
        ).pack(side=tk.LEFT, padx=(0, 6))
        self.airline_combo = ttk.Combobox(
            search_filter_bar,
            textvariable=self.airline_var,
            values=self._airline_combo_values(),
            width=24,
            font=('맑은 고딕', 9),
        )
        self.airline_combo.pack(side=tk.LEFT)
        self.airline_combo.bind('<FocusIn>', self._select_airline_combo_text)
        self.airline_combo.bind('<ButtonRelease-1>', self._select_airline_combo_text)
        self.airline_combo.bind('<<ComboboxSelected>>', self._on_airline_combo_change)
        self.airline_combo.bind('<FocusOut>', self._on_airline_combo_change)
        tk.Label(
            search_filter_bar,
            text='요금구분',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_color,
        ).pack(side=tk.LEFT, padx=(16, 6))
        self.price_desc_entry = tk.Entry(
            search_filter_bar,
            textvariable=self.price_desc_var,
            width=18,
            bg=self.input_bg,
            fg=self.fg_color,
            insertbackground='white',
            bd=0,
            relief=tk.FLAT,
            highlightbackground=self.border_color,
            highlightcolor=self.accent_color,
            highlightthickness=1,
            font=('맑은 고딕', 9),
        )
        self.price_desc_entry.pack(side=tk.LEFT, ipady=3)
        tk.Label(
            search_filter_bar,
            text='호텔명 검색',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_color,
        ).pack(side=tk.LEFT, padx=(16, 6))
        self.hotel_name_combo = ttk.Combobox(
            search_filter_bar,
            textvariable=self.hotel_name_var,
            values=[],
            width=20,
            font=('맑은 고딕', 9),
        )
        self.hotel_name_entry = self.hotel_name_combo
        self.hotel_name_entry.pack(side=tk.LEFT, ipady=3)
        self.hotel_name_entry.bind('<<ComboboxSelected>>', self._on_hotel_combo_selected)
        self.hotel_name_entry.bind('<Return>', self._run_hotel_search_now)
        self.hotel_name_entry.bind('<KeyRelease>', self._on_hotel_combo_typed)
        self.hotel_name_entry.bind('<FocusOut>', self._on_hotel_combo_focus_out)

        self.hotel_result_frame = tk.Frame(sheet_card, bg=self.card_color)
        self.hotel_result_list = tk.Listbox(
            self.hotel_result_frame,
            height=5,
            bg=self.input_bg,
            fg=self.fg_color,
            selectbackground=self.accent_color,
            selectforeground='white',
            activestyle='none',
            bd=0,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.accent_color,
            font=('맑은 고딕', 9),
        )
        self.hotel_result_list.pack(fill=tk.X)
        self.hotel_result_list.bind('<ButtonRelease-1>', self._on_hotel_result_selected)
        self.hotel_result_list.bind('<Return>', self._on_hotel_result_selected)
        self.hotel_result_list.bind('<Double-Button-1>', self._on_hotel_result_selected)

        # 수식 입력줄(formula bar): '=' 식 편집 중 표의 칸을 클릭하면 참조가 삽입된다
        fb_frame = tk.Frame(sheet_card, bg=self.card_color)
        self.formula_frame = fb_frame
        fb_frame.pack(fill=tk.X, pady=(0, 6))
        tk.Label(fb_frame, text='fx', font=('맑은 고딕', 10, 'bold'), bg=self.card_color, fg=self.accent_orange).pack(side=tk.LEFT, padx=(0, 6))
        self.active_lbl = tk.Label(fb_frame, text='', width=13, anchor=tk.W, font=('맑은 고딕', 8), bg=self.card_color, fg=self.fg_muted)
        self.active_lbl.pack(side=tk.LEFT, padx=(0, 6))
        self.formula_entry = tk.Entry(fb_frame, textvariable=self.fb_var, bg=self.input_bg, fg=self.fg_color, insertbackground='white', bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1, font=('Consolas', 10))
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
        try:
            # 입력표 색을 V5 팔레트에 맞춘다 (실패 시 위의 dark 테마 그대로)
            self.sheet.set_options(
                table_bg=self.input_bg,
                table_fg='#dfe5f0',
                table_grid_fg='#222a37',
                table_selected_cells_border_fg=self.accent_color,
                header_bg=self.tree_head_bg,
                header_fg=self.fg_muted,
                header_grid_fg='#222a37',
                index_bg=self.tree_head_bg,
                index_fg=self.fg_muted,
                index_grid_fg='#222a37',
                top_left_bg=self.tree_head_bg,
            )
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

        self.data_source_badge = tk.Label(
            action_row,
            text='',
            bg=self.bg_color,
            fg=self.accent_orange,
            font=('맑은 고딕', 8, 'bold'),
            padx=8,
            pady=4,
        )
        self.data_source_badge.pack(side=tk.LEFT, padx=(4, 0))

        # 2-2. 다중 상품 작업 목록
        job_card = tk.Frame(self.fare_tab, bg=self.card_color, highlightbackground=self.border_color, highlightthickness=1)
        job_card.pack(fill=tk.X, padx=6, pady=(0, 8))

        job_header = tk.Frame(job_card, bg=self.card_color)
        job_header.pack(fill=tk.X)
        tk.Label(
            job_header,
            text='작업 목록',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_muted,
        ).pack(side=tk.LEFT, padx=10, pady=(7, 4))
        self.job_queue_summary_lbl = tk.Label(
            job_header,
            text='등록된 작업 없음',
            font=('맑은 고딕', 8),
            bg=self.card_color,
            fg=self.fg_muted,
        )
        self.job_queue_summary_lbl.pack(side=tk.LEFT, padx=(6, 0), pady=(7, 4))

        job_body = tk.Frame(job_card, bg=self.border_color)
        job_body.pack(fill=tk.X, padx=10)
        self.job_tree = ttk.Treeview(
            job_body,
            columns=('condition', 'status', 'rows', 'source'),
            show='headings',
            height=3,
        )
        self.job_tree.heading('condition', text='조건')
        self.job_tree.heading('status', text='진행상황')
        self.job_tree.heading('rows', text='건수')
        self.job_tree.heading('source', text='출처')
        self.job_tree.column('condition', width=390, anchor=tk.W)
        self.job_tree.column('status', width=180, anchor=tk.W)
        self.job_tree.column('rows', width=70, anchor=tk.CENTER)
        self.job_tree.column('source', width=150, anchor=tk.W)
        self.job_tree.pack(fill=tk.X, padx=1, pady=1)

        job_btn_row = tk.Frame(job_card, bg=self.card_color)
        job_btn_row.pack(fill=tk.X, padx=10, pady=(6, 8))
        self._make_toolbar_btn(job_btn_row, '입력표로 불러오기', self.btn_neutral_bg, self.btn_neutral_hover, self.load_selected_job_to_sheet, fg=self.btn_neutral_fg)
        self._make_toolbar_btn(job_btn_row, '위로', self.btn_neutral_bg, self.btn_neutral_hover, lambda: self.move_selected_job(-1), fg=self.btn_neutral_fg)
        self._make_toolbar_btn(job_btn_row, '아래로', self.btn_neutral_bg, self.btn_neutral_hover, lambda: self.move_selected_job(1), fg=self.btn_neutral_fg)
        self._make_toolbar_btn(job_btn_row, '선택 삭제', self.btn_neutral_bg, self.btn_neutral_hover, self.delete_selected_job, fg=self.btn_neutral_fg)
        self._make_toolbar_btn(job_btn_row, '목록 비우기', self.btn_neutral_bg, self.btn_neutral_hover, self.clear_job_queue, fg=self.btn_neutral_fg)

        # 3. 날짜 필터 카드
        filter_card = tk.Frame(self.fare_tab, bg=self.card_color, highlightbackground=self.border_color, highlightthickness=1)
        filter_card.pack(fill=tk.X, padx=6, pady=(0, 8))

        filter_header = tk.Frame(filter_card, bg=self.card_color)
        filter_header.pack(fill=tk.X)
        self.filter_toggle_btn = tk.Button(
            filter_header,
            text='▸ 날짜 필터 (선택)',
            bg=self.card_color,
            fg=self.fg_muted,
            activebackground=self.card_hover,
            activeforeground=self.fg_color,
            font=('맑은 고딕', 9, 'bold'),
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            anchor=tk.W,
            padx=10,
            pady=6,
            command=self.toggle_filter_panel,
        )
        self.filter_toggle_btn.pack(fill=tk.X)

        self.filter_body = tk.Frame(filter_card, bg=self.card_color)

        filter_modes_frame = tk.Frame(self.filter_body, bg=self.card_color)
        filter_modes_frame.grid(row=0, column=0, padx=6, pady=4, sticky=tk.W)
        for txt, val in [("전체 대상", "ALL"), ("특정일 이후", "FROM_DATE"), ("특정 날짜 지정", "SPECIFIC"), ("기간 범위 지정", "DATE_RANGE")]:
            rb = tk.Radiobutton(filter_modes_frame, text=txt, variable=self.filter_mode, value=val, bg=self.card_color, fg=self.fg_color, selectcolor=self.input_bg, activebackground=self.card_color, activeforeground=self.fg_color)
            rb.pack(side=tk.LEFT, padx=(0, 10))

        self.filter_input_container = tk.Frame(self.filter_body, bg=self.card_color)
        self.filter_input_container.grid(row=1, column=0, padx=6, pady=(2, 4), sticky=tk.W)
        self.filter_tip_lbl = tk.Label(self.filter_body, text='', font=('맑은 고딕', 8), bg=self.card_color, fg=self.fg_muted)
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
        log_body, self.log_txt = self._make_dark_text(
            log_wrap, height=8, bg=self.input_bg, fg=self.input_fg, insertbackground='white',
            font=('Consolas', 9), bd=0, relief=tk.FLAT, padx=10, pady=8)
        log_body.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # 6. 컨트롤 버튼
        control_frame = tk.Frame(self.fare_tab, bg=self.bg_color)
        control_frame.pack(fill=tk.X, padx=6, pady=(2, 14))

        self.start_btn = tk.Button(control_frame, text='▶  요금수정 시작', width=18, height=2, bg=self.accent_green, fg='white', font=('맑은 고딕', 10, 'bold'), activebackground=self.accent_green_hover, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', command=self.start_rpa)
        self.start_btn.pack(side=tk.LEFT)
        self._add_hover(self.start_btn, self.accent_green, self.accent_green_hover)

        self.pause_btn = tk.Button(control_frame, text='‖  일시 중지', width=13, height=2, bg=self.bg_color, fg=self.accent_orange, font=('맑은 고딕', 10, 'bold'), activebackground=self.accent_orange, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', disabledforeground=self.disabled_fg, highlightbackground='#574730', highlightthickness=1, command=self.toggle_pause)
        self.pause_btn.pack(side=tk.LEFT, padx=12)
        self.pause_btn.config(state=tk.DISABLED)
        self._add_hover(self.pause_btn, self.bg_color, self.accent_orange, normal_fg=self.accent_orange, hover_fg='white')

        self.stop_btn = tk.Button(control_frame, text='■  중지', width=13, height=2, bg=self.bg_color, fg=self.accent_red, font=('맑은 고딕', 10, 'bold'), activebackground=self.accent_red, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', disabledforeground=self.disabled_fg, highlightbackground='#573a38', highlightthickness=1, command=self.stop_rpa)
        self.stop_btn.pack(side=tk.RIGHT)
        self.stop_btn.config(state=tk.DISABLED)
        self._add_hover(self.stop_btn, self.bg_color, self.accent_red, normal_fg=self.accent_red, hover_fg='white')

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
        for key, label in [('fare', 'ERP 요금수정'), ('topas', '토파스 요금조회')]:
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

        selected_bg = self.tab_selected_bg
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

    def _make_toolbar_btn(self, parent, text, bg, hover, command, fg='white'):
        btn = tk.Button(parent, text=text, bg=bg, fg=fg, font=('맑은 고딕', 9, 'bold'), activebackground=hover, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', padx=10, pady=5, command=command)
        btn.pack(side=tk.LEFT, padx=(0, 6))
        self._add_hover(btn, bg, hover)
        self.toolbar_buttons.append(btn)
        return btn

    def _format_airline_label(self, value, text):
        value = '' if value is None else str(value)
        text = '' if text is None else str(text).strip()
        if not value:
            return AIRLINE_EMPTY_LABEL
        if text and text.lower() != 'null':
            return text if text.startswith('[') else f'[{value}] {text}'
        return f'[{value}]'

    def _airline_combo_values(self):
        return [label for _value, label in self.airline_choices]

    def _set_airline_choices(self, choices):
        current_value = self._selected_airline_code() if hasattr(self, 'airline_value_by_label') else ''
        normalized = []
        seen = set()

        def add_choice(value, text):
            value = '' if value is None else str(value)
            label = self._format_airline_label(value, text)
            key = (value, label)
            if key in seen:
                return
            seen.add(key)
            normalized.append((value, label))

        add_choice('', AIRLINE_EMPTY_LABEL)
        for value, text in choices or []:
            add_choice(value, text)

        self.airline_choices = normalized
        self.airline_value_by_label = {label: value for value, label in normalized}
        self.airline_label_by_value = {value: label for value, label in normalized}

        if hasattr(self, 'airline_combo'):
            self.airline_combo.config(values=self._airline_combo_values())
        if current_value:
            self._select_airline_code(current_value, overwrite_blank_only=False)
        elif hasattr(self, 'airline_var'):
            self.airline_var.set(AIRLINE_EMPTY_LABEL)

    def _selected_airline_code(self):
        text = self.airline_var.get() if hasattr(self, 'airline_var') else ''
        text = '' if text is None else str(text).strip()
        if not text or text == AIRLINE_EMPTY_LABEL:
            return ''
        mapped = self.airline_value_by_label.get(text) if hasattr(self, 'airline_value_by_label') else None
        if mapped is not None:
            return mapped
        bracket_match = re.match(r'^\[([^\]]+)\]', text)
        if bracket_match:
            return bracket_match.group(1)
        token = text.split()[0].strip().upper() if text.split() else ''
        if re.fullmatch(r'[A-Z0-9]{1,3}', token):
            return token
        return ''

    def _select_airline_code(self, code, overwrite_blank_only=True):
        code = '' if code is None else str(code).strip().upper()
        if overwrite_blank_only and self._selected_airline_code():
            return
        if not code:
            self.airline_var.set(AIRLINE_EMPTY_LABEL)
            return
        self.airline_var.set(self.airline_label_by_value.get(code, code))

    def _airline_code_from_route(self):
        route = self.route_var.get().strip() if hasattr(self, 'route_var') else ''
        if not route:
            return ''
        parts = [p for p in re.split(r'[^A-Za-z0-9]+', route.upper()) if p]
        if parts and re.fullmatch(r'[A-Z0-9]{2,3}', parts[-1]):
            return parts[-1]
        return ''

    def _on_airline_combo_change(self, event=None):
        code = self._selected_airline_code()
        if code and code in self.airline_label_by_value:
            self.airline_var.set(self.airline_label_by_value[code])
        elif not code:
            self.airline_var.set(AIRLINE_EMPTY_LABEL)

    def _select_airline_combo_text(self, event=None):
        if not hasattr(self, 'airline_combo'):
            return

        def select_all():
            try:
                if str(self.airline_combo.cget('state')) == tk.DISABLED:
                    return
                self.airline_combo.focus_set()
                self.airline_combo.selection_range(0, tk.END)
                self.airline_combo.icursor(tk.END)
            except Exception:
                pass

        self.root.after_idle(select_all)

    def refresh_airline_options_from_erp(self, silent=False):
        if getattr(self, 'airline_refresh_thread', None) and self.airline_refresh_thread.is_alive():
            return
        if getattr(self, 'is_running', False):
            if not silent:
                messagebox.showwarning('실행 중', 'ERP 실행 중에는 항공사 목록을 새로고침할 수 없습니다.')
            return
        if not silent:
            self.set_status('항공사코드 목록을 읽는 중입니다…', self.accent_orange)
        self.airline_refresh_thread = threading.Thread(target=self._airline_options_worker, args=(silent,), daemon=True)
        self.airline_refresh_thread.start()

    def _airline_options_worker(self, silent=False):
        driver = None
        try:
            selectors = self.config.get('selectors', {})
            driver, _browser_config = self._connect_matching_debug_browser('ERP', selectors)
            options = self._read_airline_options_from_driver(driver, selectors.get('airline_select', '#air2Cd'))
            self.root.after(0, lambda opts=options: self._apply_airline_options(opts, None, silent))
        except Exception as exc:
            self.root.after(0, lambda err=str(exc): self._apply_airline_options(None, err, silent))
        finally:
            if driver is not None:
                try:
                    driver.service.stop()
                except Exception:
                    pass

    def _apply_airline_options(self, options, error, silent=False):
        if error:
            if not silent:
                self.set_status('항공사코드 목록을 불러오지 못했습니다', self.accent_red)
                messagebox.showerror('항공사 목록 새로고침 실패', error)
            return
        self._set_airline_choices(options)
        count = max(0, len(self.airline_choices) - 1)
        if not silent:
            self.set_status(f'항공사코드 목록 {count}개를 불러왔습니다.', self.accent_green)

    def _read_airline_options_from_driver(self, driver, selector):
        if not self._switch_driver_to_frame_containing(driver, selector):
            raise RuntimeError(f'ERP 화면에서 항공사 필드({selector})를 찾지 못했습니다.')
        options = driver.execute_script(
            """
            const sel = document.querySelector(arguments[0]);
            if (!sel) return [];
            return Array.from(sel.options).map(o => ({
                value: o.value || '',
                text: (o.textContent || '').trim()
            }));
            """,
            selector,
        )
        return [(item.get('value', ''), item.get('text', '')) for item in options or []]

    def _format_hotel_label(self, record):
        name = str(record.get('name') or '').strip()
        seq = str(record.get('seq') or '').strip()
        city = str(record.get('city') or '').strip()
        nation = str(record.get('nation') or '').strip()
        parts = [name]
        area = ' / '.join(part for part in (nation, city) if part)
        if area:
            parts.append(area)
        if seq:
            parts.append(f'hotelSeq {seq}')
        return ' · '.join(part for part in parts if part)

    def _format_hotel_result_line(self, record):
        name = str(record.get('name') or '').strip()
        nation = str(record.get('nation') or '').strip()
        city = str(record.get('city') or '').strip()
        seq = str(record.get('seq') or '').strip()
        path_parts = ['호텔'] + [part for part in (nation, city) if part]
        path = ' > '.join(path_parts)
        suffix = f' · hotelSeq {seq}' if seq else ''
        return f'{name}    {path}{suffix}'

    def _normalize_hotel_record(self, raw):
        raw = raw or {}
        seq = str(raw.get('infoSeq') or raw.get('hotelSeq') or raw.get('seq') or '').strip()
        name = str(raw.get('infoTitle') or raw.get('hotelKorNm') or raw.get('name') or raw.get('korNm') or '').strip()
        if not seq or not name:
            return None
        record = {
            'seq': seq,
            'name': name,
            'eng_name': str(raw.get('infoEngTitle') or raw.get('engName') or '').strip(),
            'nation': str(raw.get('natNm') or raw.get('nation') or '').strip(),
            'city': str(raw.get('cityNm') or raw.get('city') or '').strip(),
            'use_yn': str(raw.get('useYn') or '').strip(),
            'sort_order': raw.get('sortOrder'),
        }
        record['label'] = self._format_hotel_label(record)
        return record

    def _set_hotel_choices(self, records):
        current_name, current_seq = self._selected_hotel_filter() if hasattr(self, 'hotel_record_by_label') else ('', '')
        normalized = []
        seen = set()
        for raw in records or []:
            record = self._normalize_hotel_record(raw)
            if not record or record['seq'] in seen:
                continue
            seen.add(record['seq'])
            normalized.append(record)

        self.hotel_choices = normalized
        self.hotel_record_by_label = {record['label']: record for record in normalized}
        self.hotel_record_by_seq = {record['seq']: record for record in normalized}
        if hasattr(self, 'hotel_name_combo'):
            self.hotel_name_combo.config(values=[record['label'] for record in normalized])
        if current_seq and current_seq in self.hotel_record_by_seq:
            self.current_hotel_seq = current_seq
            self.hotel_name_var.set(current_name)

    def _show_hotel_result_message(self, message):
        self.hotel_result_records = []
        if not hasattr(self, 'hotel_result_list'):
            return
        self.hotel_result_list.delete(0, tk.END)
        self.hotel_result_list.insert(tk.END, message)
        self.hotel_result_list.itemconfig(0, foreground=self.fg_muted)
        if hasattr(self, 'hotel_result_frame') and not self.hotel_result_frame.winfo_ismapped():
            pack_kwargs = {'fill': tk.X, 'pady': (0, 6)}
            if hasattr(self, 'formula_frame'):
                pack_kwargs['before'] = self.formula_frame
            self.hotel_result_frame.pack(**pack_kwargs)

    def _show_hotel_result_records(self, records):
        self.hotel_result_records = list(records or [])
        if not hasattr(self, 'hotel_result_list'):
            return
        self.hotel_result_list.delete(0, tk.END)
        if not self.hotel_result_records:
            self._show_hotel_result_message('검색 결과 없음')
            return
        for record in self.hotel_result_records:
            self.hotel_result_list.insert(tk.END, self._format_hotel_result_line(record))
        if hasattr(self, 'hotel_result_frame') and not self.hotel_result_frame.winfo_ismapped():
            pack_kwargs = {'fill': tk.X, 'pady': (0, 6)}
            if hasattr(self, 'formula_frame'):
                pack_kwargs['before'] = self.formula_frame
            self.hotel_result_frame.pack(**pack_kwargs)

    def _hide_hotel_results(self):
        self.hotel_result_records = []
        try:
            self.hotel_result_list.delete(0, tk.END)
            self.hotel_result_frame.pack_forget()
        except Exception:
            pass

    def _hotel_record_from_text(self, text):
        text = '' if text is None else str(text).strip()
        if not text:
            return None
        record = getattr(self, 'hotel_record_by_label', {}).get(text)
        if record:
            return record
        current_seq = str(getattr(self, 'current_hotel_seq', '') or '').strip()
        if current_seq:
            record = getattr(self, 'hotel_record_by_seq', {}).get(current_seq)
            if record and str(record.get('name') or '').strip() == text:
                return record
        matches = [
            record for record in getattr(self, 'hotel_choices', [])
            if str(record.get('name') or '').strip() == text
        ]
        return matches[0] if len(matches) == 1 else None

    def _select_hotel_record(self, record):
        if not record:
            self.current_hotel_seq = ''
            return
        self.current_hotel_seq = str(record.get('seq') or '').strip()
        self.hotel_name_var.set(str(record.get('name') or '').strip())

    def _selected_hotel_filter(self):
        text = self.hotel_name_var.get() if hasattr(self, 'hotel_name_var') else ''
        text = '' if text is None else str(text).strip()
        if not text:
            return '', ''
        record = self._hotel_record_from_text(text)
        if record:
            return str(record.get('name') or '').strip(), str(record.get('seq') or '').strip()
        self.current_hotel_seq = ''
        return text, ''

    def _remember_imported_hotel_condition(self, hotel_name, hotel_seq=''):
        hotel_name = '' if hotel_name is None else str(hotel_name).strip()
        hotel_seq = '' if hotel_seq is None else str(hotel_seq).strip()
        if not hotel_name:
            if hasattr(self, 'hotel_name_var'):
                self.hotel_name_var.set('')
            self.current_hotel_seq = ''
            return

        if hotel_seq:
            record = {
                'seq': hotel_seq,
                'name': hotel_name,
                'eng_name': '',
                'nation': '',
                'city': '',
                'use_yn': '',
                'sort_order': None,
            }
            record['label'] = self._format_hotel_label(record)
            existing = [
                item for item in getattr(self, 'hotel_choices', [])
                if str(item.get('seq') or '').strip() != hotel_seq
            ]
            existing.append(record)
            self.hotel_choices = existing
            self.hotel_record_by_label = {item['label']: item for item in existing}
            self.hotel_record_by_seq = {item['seq']: item for item in existing}
            if hasattr(self, 'hotel_name_combo'):
                self.hotel_name_combo.config(values=[item['label'] for item in existing])
            self._select_hotel_record(record)
            return

        self.current_hotel_seq = ''
        if hasattr(self, 'hotel_name_var'):
            self.hotel_name_var.set(hotel_name)

    def _on_hotel_combo_selected(self, event=None):
        record = self._hotel_record_from_text(self.hotel_name_var.get())
        self._select_hotel_record(record)
        self._hide_hotel_results()

    def _on_hotel_combo_typed(self, event=None):
        if event is not None and getattr(event, 'keysym', '') in {'Return', 'Tab', 'Shift_L', 'Shift_R', 'Up', 'Down', 'Left', 'Right'}:
            return
        self.current_hotel_seq = ''
        self._schedule_hotel_auto_search()

    def _on_hotel_combo_focus_out(self, event=None):
        record = self._hotel_record_from_text(self.hotel_name_var.get())
        if record:
            self._select_hotel_record(record)

    def _schedule_hotel_auto_search(self):
        if getattr(self, 'hotel_search_after_id', None):
            try:
                self.root.after_cancel(self.hotel_search_after_id)
            except Exception:
                pass
            self.hotel_search_after_id = None
        keyword = self.hotel_name_var.get().strip() if hasattr(self, 'hotel_name_var') else ''
        if not keyword:
            self._hide_hotel_results()
            return
        self._show_hotel_result_message('검색 중...')
        self.hotel_search_after_id = self.root.after(
            450,
            lambda kw=keyword: self.refresh_hotel_options_from_erp(silent=True, keyword=kw, auto=True),
        )

    def _run_hotel_search_now(self, event=None):
        if getattr(self, 'hotel_search_after_id', None):
            try:
                self.root.after_cancel(self.hotel_search_after_id)
            except Exception:
                pass
            self.hotel_search_after_id = None
        self.refresh_hotel_options_from_erp(silent=True, auto=True)
        return 'break'

    def _on_hotel_result_selected(self, event=None):
        if not hasattr(self, 'hotel_result_list'):
            return
        selection = self.hotel_result_list.curselection()
        if not selection:
            return
        index = int(selection[0])
        if not (0 <= index < len(self.hotel_result_records)):
            return
        record = self.hotel_result_records[index]
        self._select_hotel_record(record)
        self._hide_hotel_results()
        self.set_status(
            f"호텔 선택: {record.get('name')} (hotelSeq {record.get('seq')})",
            self.accent_green,
        )

    def refresh_hotel_options_from_erp(self, silent=False, keyword=None, auto=False):
        if getattr(self, 'is_running', False):
            if not silent:
                messagebox.showwarning('실행 중', 'ERP 실행 중에는 호텔 목록을 검색할 수 없습니다.')
            return
        keyword = (keyword if keyword is not None else (self.hotel_name_var.get() if hasattr(self, 'hotel_name_var') else ''))
        keyword = '' if keyword is None else str(keyword).strip()
        if not keyword:
            self._hide_hotel_results()
            if not silent and not auto:
                messagebox.showwarning('호텔검색', '검색할 호텔명 일부를 먼저 입력해 주세요.')
            return
        if not silent:
            self.set_status(f"ERP 호텔 DB에서 '{keyword}' 검색 중입니다…", self.accent_orange)
        self.hotel_search_request_id += 1
        request_id = self.hotel_search_request_id
        self.hotel_search_thread = threading.Thread(
            target=self._hotel_options_worker,
            args=(keyword, silent, request_id, auto),
            daemon=True,
        )
        self.hotel_search_thread.start()

    def _hotel_options_worker(self, keyword, silent=False, request_id=0, auto=False):
        driver = None
        try:
            selectors = self.config.get('selectors', {})
            driver, _browser_config = self._connect_matching_debug_browser('ERP', selectors)
            options = self._read_hotel_options_from_driver(driver, keyword)
            self.root.after(0, lambda opts=options, kw=keyword, rid=request_id: self._apply_hotel_options(opts, None, kw, silent, rid, auto))
        except Exception as exc:
            self.root.after(0, lambda err=str(exc), kw=keyword, rid=request_id: self._apply_hotel_options(None, err, kw, silent, rid, auto))
        finally:
            if driver is not None:
                try:
                    driver.service.stop()
                except Exception:
                    pass

    def _apply_hotel_options(self, options, error, keyword='', silent=False, request_id=0, auto=False):
        if request_id and request_id != getattr(self, 'hotel_search_request_id', 0):
            return
        if error:
            if auto:
                self._show_hotel_result_message('검색 오류')
                self.set_status(f"호텔검색 실패: {error}", self.accent_red)
            elif not silent:
                self.set_status('호텔 목록을 불러오지 못했습니다', self.accent_red)
                messagebox.showerror('호텔검색 실패', error)
            return
        self._set_hotel_choices(options)
        count = len(self.hotel_choices)
        self._show_hotel_result_records(self.hotel_choices)
        if auto:
            if count:
                self.set_status(f"'{keyword}' 호텔 후보 {count}개를 찾았습니다.", self.accent_green)
            else:
                self.set_status(f"'{keyword}' 검색 결과가 없습니다.", self.accent_orange)
        elif not silent:
            if count:
                self.set_status(f"'{keyword}' 호텔 후보 {count}개를 불러왔습니다. 목록에서 호텔을 선택해 주세요.", self.accent_green)
            else:
                self.set_status(f"'{keyword}' 호텔 후보가 없습니다.", self.accent_orange)

    def _read_hotel_options_from_driver(self, driver, keyword, limit=80):
        try:
            driver.set_script_timeout(20)
        except Exception:
            pass
        result = driver.execute_async_script(
            """
            const keyword = arguments[0] || '';
            const limit = Number(arguments[1] || 80);
            const done = arguments[arguments.length - 1];
            const params = new URLSearchParams();
            params.set('searchType', 'pop');
            params.set('infoCd', 'H');
            params.set('infoTitle', keyword);
            params.set('page', '1');
            params.set('rows', String(limit));
            params.set('pageIndex', '1');
            params.set('recordCountPerPage', String(limit));
            fetch('/erp/sy/sy02/sy02_108_list.json', {
                method: 'POST',
                credentials: 'same-origin',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: params.toString()
            }).then(async response => {
                const text = await response.text();
                if (!response.ok) {
                    done({ok: false, error: response.status + ' ' + text.slice(0, 160)});
                    return;
                }
                let data = {};
                try {
                    data = JSON.parse(text);
                } catch (e) {
                    done({ok: false, error: 'JSON parse failed: ' + String(e)});
                    return;
                }
                const root = data.responseData || data;
                const list = root.list || root.rows || data.list || data.rows || [];
                const rows = [];
                const seen = new Set();
                for (const item of list) {
                    const seq = String(item.infoSeq || item.hotelSeq || item.seq || '').trim();
                    const name = String(item.infoTitle || item.hotelKorNm || item.name || item.korNm || '').trim();
                    const infoCd = String(item.infoCd || '').trim();
                    if (!seq || !name) continue;
                    if (infoCd && infoCd !== 'H') continue;
                    if (seen.has(seq)) continue;
                    seen.add(seq);
                    rows.push(item);
                    if (rows.length >= limit) break;
                }
                done({ok: true, rows});
            }).catch(error => done({ok: false, error: String(error)}));
            """,
            keyword,
            limit,
        )
        if not result or not result.get('ok'):
            error = result.get('error') if isinstance(result, dict) else 'empty result'
            raise RuntimeError(f'ERP 호텔 DB 검색 실패: {error}')
        return result.get('rows') or []

    def _switch_driver_to_frame_containing(self, driver, selector, depth=0, max_depth=4):
        by, search_val = self._selector_locator(selector)
        try:
            if driver.find_elements(by, search_val):
                return True
        except Exception:
            pass
        if depth >= max_depth:
            return False
        try:
            frames = driver.find_elements(By.CSS_SELECTOR, 'iframe, frame')
        except Exception:
            frames = []
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                if self._switch_driver_to_frame_containing(driver, selector, depth + 1, max_depth):
                    return True
            except Exception:
                pass
            try:
                driver.switch_to.parent_frame()
            except Exception:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
        return False

    def _build_topas_tab(self):
        outer = tk.Frame(self.topas_tab, bg=self.bg_color)
        outer.pack(fill=tk.BOTH, expand=True, padx=6, pady=(10, 10))

        collect_card = tk.LabelFrame(
            outer,
            text='',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_muted,
            bd=0,
            relief=tk.FLAT,
            highlightbackground=self.border_color,
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        collect_card.pack(fill=tk.X, pady=(0, 8))

        collect_actions = tk.Frame(collect_card, bg=self.card_color)
        collect_actions.pack(fill=tk.X)

        self.topas_query_buttons = []
        self.topas_query_btn = tk.Button(
            collect_actions,
            text='AC1 자동반복',
            width=30,
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
        self.topas_query_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._add_hover(self.topas_query_btn, self.accent_green, self.accent_green_hover)
        self.topas_query_buttons.append(self.topas_query_btn)

        self.topas_stop_btn = tk.Button(
            collect_actions,
            text='■  중지',
            width=13,
            height=2,
            bg=self.card_color,
            fg=self.accent_red,
            font=('맑은 고딕', 10, 'bold'),
            activebackground=self.accent_red,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            disabledforeground=self.disabled_fg,
            highlightbackground='#573a38',
            highlightthickness=1,
            command=self.stop_topas_query,
        )
        self.topas_stop_btn.pack(side=tk.RIGHT)
        self.topas_stop_btn.config(state=tk.DISABLED)
        self._add_hover(self.topas_stop_btn, self.card_color, self.accent_red, normal_fg=self.accent_red, hover_fg='white')

        topas_progress_frame = tk.Frame(collect_card, bg=self.card_color)
        topas_progress_frame.pack(fill=tk.X, pady=(10, 4))
        self.topas_status_lbl = tk.Label(
            topas_progress_frame,
            text='TOPAS 로그인 후 첫 날짜를 직접 조회하고 수집을 시작하세요.',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_muted,
        )
        self.topas_status_lbl.pack(side=tk.LEFT)
        self.topas_progress_lbl = tk.Label(
            topas_progress_frame,
            text='0 / 0 (0%)',
            font=('맑은 고딕', 9),
            bg=self.card_color,
            fg=self.fg_muted,
        )
        self.topas_progress_lbl.pack(side=tk.RIGHT)

        self.topas_progress_bar = ttk.Progressbar(
            collect_card,
            orient='horizontal',
            mode='determinate',
            style='Accent.Horizontal.TProgressbar',
        )
        self.topas_progress_bar.pack(fill=tk.X)

        slots = tk.Frame(collect_card, bg=self.card_color)
        slots.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        slots.grid_columnconfigure(0, weight=1)
        slots.grid_columnconfigure(1, weight=1)

        self.departure_raw_txt, self.departure_parse_lbl = self._build_raw_slot(slots, '출발편', 'departure', 0)
        self.return_raw_txt, self.return_parse_lbl = self._build_raw_slot(slots, '귀국편', 'return', 1)
        self._refresh_topas_collect_button_text()

        calc_card = tk.LabelFrame(
            outer,
            text='',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_muted,
            bd=0,
            relief=tk.FLAT,
            highlightbackground=self.border_color,
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        calc_card.pack(fill=tk.X, pady=(0, 8))

        calc_row = tk.Frame(calc_card, bg=self.card_color)
        calc_row.pack(fill=tk.X)
        tk.Label(calc_row, text='항공노선', bg=self.card_color, fg=self.fg_color, font=('맑은 고딕', 9, 'bold')).pack(side=tk.LEFT)
        self.route_combo = ttk.Combobox(calc_row, textvariable=self.route_var, width=28, values=[], font=('맑은 고딕', 10))
        self.route_combo.pack(side=tk.LEFT, padx=(8, 8))
        self.route_combo.bind('<KeyRelease>', self._on_route_search_keyrelease)
        self.route_combo.bind('<Return>', self._commit_route_search)
        self.route_combo.bind('<KP_Enter>', self._commit_route_search)
        self.route_combo.bind('<<ComboboxSelected>>', self._on_route_selected)

        tk.Label(calc_row, text='박수', bg=self.card_color, fg=self.fg_color, font=('맑은 고딕', 9, 'bold')).pack(side=tk.LEFT)
        self.night_chip_frame = tk.Frame(calc_row, bg=self.card_color)
        self.night_chip_frame.pack(side=tk.LEFT)
        for night in (2, 3, 4, 5):
            var = tk.BooleanVar(value=False)
            self.night_vars[night] = var
            self._make_night_chip(night)

        custom_entry = tk.Entry(
            calc_row,
            textvariable=self.custom_night_var,
            width=5,
            bg=self.input_bg,
            fg=self.fg_color,
            insertbackground='white',
            relief=tk.FLAT,
            highlightbackground=self.border_color,
            highlightcolor=self.accent_color,
            highlightthickness=1,
            font=('Consolas', 10),
        )
        custom_entry.pack(side=tk.LEFT, padx=(10, 4), ipady=3)
        custom_entry.bind('<Return>', lambda _event: self.add_custom_night_chip())

        add_night_btn = tk.Button(
            calc_row,
            text='+추가',
            bg=self.btn_neutral_bg,
            fg=self.btn_neutral_fg,
            font=('맑은 고딕', 9, 'bold'),
            activebackground=self.btn_neutral_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            padx=8,
            pady=4,
            command=self.add_custom_night_chip,
        )
        add_night_btn.pack(side=tk.LEFT)
        self._add_hover(add_night_btn, self.btn_neutral_bg, self.btn_neutral_hover)

        calc_btn = tk.Button(
            calc_row,
            text='요금 조회',
            width=14,
            bg=self.accent_color,
            fg='white',
            font=('맑은 고딕', 10, 'bold'),
            activebackground=self.accent_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            padx=10,
            pady=5,
            command=self.calculate_topas_fares,
        )
        calc_btn.pack(side=tk.LEFT, padx=(8, 0))
        self._add_hover(calc_btn, self.accent_color, self.accent_hover)

        calc_actions = tk.Frame(calc_card, bg=self.card_color)
        calc_actions.pack(fill=tk.X, pady=(10, 0))

        fare_status_frame = tk.Frame(calc_actions, bg=self.card_color)
        fare_status_frame.pack(side=tk.RIGHT)
        self.fare_data_status_lbl = tk.Label(
            fare_status_frame,
            text='운임/시즌 미로드',
            bg=self.card_color,
            fg=self.fg_muted,
            font=('맑은 고딕', 8),
        )
        self.fare_data_status_lbl.pack(side=tk.LEFT, padx=(0, 6))
        self.refresh_fare_btn = tk.Button(
            fare_status_frame,
            text='↻',
            width=3,
            bg=self.btn_neutral_bg,
            fg=self.fg_muted,
            font=('맑은 고딕', 8, 'bold'),
            activebackground=self.btn_neutral_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            padx=2,
            pady=1,
            command=self.refresh_fare_snapshot,
        )
        self.refresh_fare_btn.pack(side=tk.LEFT)
        self._add_hover(self.refresh_fare_btn, self.btn_neutral_bg, self.btn_neutral_hover, normal_fg=self.fg_muted, hover_fg='white')

        result_card = tk.LabelFrame(
            outer,
            text='',
            font=('맑은 고딕', 9, 'bold'),
            bg=self.card_color,
            fg=self.fg_muted,
            bd=0,
            relief=tk.FLAT,
            highlightbackground=self.border_color,
            highlightthickness=1,
            padx=10,
            pady=8,
        )
        result_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.results_notebook = ttk.Notebook(result_card, style='Result.TNotebook')
        self.results_notebook.pack(fill=tk.BOTH, expand=True)

        self.root.after(700, lambda: self.refresh_fare_snapshot(force=False))

    def _build_raw_slot(self, parent, title, slot, column):
        frame = tk.Frame(parent, bg=self.card_color)
        frame.grid(row=0, column=column, sticky='nsew', padx=(0, 8) if column == 0 else (8, 0))

        header = tk.Frame(frame, bg=self.card_color)
        header.pack(fill=tk.X, pady=(0, 4))
        tk.Label(
            header,
            text=title,
            bg=self.card_color,
            fg=self.fg_color,
            font=('맑은 고딕', 9, 'bold'),
        ).pack(side=tk.LEFT)

        view_btn = tk.Button(
            header,
            text='조회내용 전체 보기',
            bg=self.btn_neutral_bg,
            fg=self.btn_neutral_fg,
            font=('맑은 고딕', 8, 'bold'),
            activebackground=self.btn_neutral_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            padx=8,
            pady=3,
            command=lambda slot_key=slot: self.show_slot_raw_popup(slot_key),
        )
        view_btn.pack(side=tk.RIGHT, padx=(4, 0))
        self._add_hover(view_btn, self.btn_neutral_bg, self.btn_neutral_hover)

        save_btn = tk.Button(
            header,
            text='텍스트 저장',
            bg=self.btn_neutral_bg,
            fg=self.btn_neutral_fg,
            font=('맑은 고딕', 8, 'bold'),
            activebackground=self.btn_neutral_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            padx=8,
            pady=3,
            command=lambda slot_key=slot: self.save_slot_raw_manually(slot_key),
        )
        save_btn.pack(side=tk.RIGHT)
        self._add_hover(save_btn, self.btn_neutral_bg, self.btn_neutral_hover)

        text_wrap = tk.Frame(frame, bg=self.border_color, bd=0)
        text_wrap.pack(fill=tk.BOTH, expand=True)
        text_body, text_widget = self._make_dark_text(
            text_wrap,
            height=8,
            bg=self.input_bg,
            fg=self.input_fg,
            insertbackground='white',
            font=('Consolas', 9),
            bd=0,
            relief=tk.FLAT,
            padx=8,
            pady=6,
            wrap=tk.NONE,
            undo=True,
        )
        text_body.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        text_widget.bind('<<Modified>>', lambda _event, slot_key=slot: self._on_raw_slot_modified(slot_key))

        summary = tk.Label(
            frame,
            text='파싱: 0일치',
            bg=self.card_color,
            fg=self.fg_muted,
            font=('맑은 고딕', 8),
        )
        summary.pack(anchor=tk.W, pady=(4, 0))
        return text_widget, summary

    def _on_raw_slot_modified(self, slot):
        widget = self._get_raw_slot_widget(slot)
        if widget is not None:
            widget.edit_modified(False)
        self._update_raw_slot_summary(slot)
        self._refresh_topas_collect_button_text()
        self._auto_select_route_from_raw()

    def _get_raw_slot_widget(self, slot):
        return self.return_raw_txt if slot == 'return' else self.departure_raw_txt

    def _get_raw_slot_label(self, slot):
        return self.return_parse_lbl if slot == 'return' else self.departure_parse_lbl

    def _get_raw_text(self, slot):
        widget = self._get_raw_slot_widget(slot)
        if widget is None:
            return ''
        return widget.get('1.0', tk.END).strip()

    def _set_raw_text(self, slot, text):
        widget = self._get_raw_slot_widget(slot)
        if widget is None:
            return
        widget.delete('1.0', tk.END)
        widget.insert('1.0', text or '')
        widget.edit_modified(False)
        self._update_raw_slot_summary(slot)
        self._refresh_topas_collect_button_text()
        self._auto_select_route_from_raw()

    def _update_raw_slot_summary(self, slot):
        text = self._get_raw_text(slot)
        label = self._get_raw_slot_label(slot)
        if label is None:
            return
        parsed = parse_topas_text(text) if text else None
        if not parsed:
            label.config(text='파싱: 0일치', fg=self.fg_muted)
            return
        color = self.accent_green if parsed.records else self.accent_orange
        suffix = ''
        if parsed.warnings:
            suffix = f' · 경고 {len(parsed.warnings)}'
        label.config(text=f'파싱: {summarize_records(parsed)}{suffix}', fg=color)

    def _slot_display_name(self, slot):
        return '귀국편' if slot == 'return' else '출발편'

    def _next_topas_target_slot(self):
        if not self._get_raw_text('departure'):
            return 'departure'
        if not self._get_raw_text('return'):
            return 'return'
        return None

    def _refresh_topas_collect_button_text(self):
        if not hasattr(self, 'topas_query_btn'):
            return
        slot = self._next_topas_target_slot()
        if slot is None:
            text = 'AC1 자동반복 (새로조회하기)'
        else:
            text = 'AC1 자동반복'
        try:
            self.topas_query_btn.config(text=text)
        except Exception:
            pass

    def _prepare_next_topas_target_slot(self):
        slot = self._next_topas_target_slot()
        if slot is not None:
            return slot

        if not messagebox.askyesno(
            '새로 조회하기',
            '출발편과 귀국편 조회내용이 모두 있습니다.\n\n'
            '새로 조회를 시작할까요?\n'
            '기존 조회내용을 모두 비우고 이번 조회를 출발편에 넣습니다.',
        ):
            return None

        self._reset_topas_query_state()
        self.route_user_modified = False
        return 'departure'

    def _suggest_topas_ac1_count(self, slot):
        if slot != 'return' or self.departure_ac1_count <= 0:
            return ''
        return str(max(1, (self.departure_ac1_count * RETURN_AC1_SUGGEST_PERCENT + 99) // 100))

    def _reset_topas_query_state(self):
        self._set_raw_text('departure', '')
        self._set_raw_text('return', '')
        self.topas_results_raw = []
        self.topas_current_ac1_count = 0
        self.departure_ac1_count = 0
        self.v5_calculation_result = None
        self.selected_result_night = None
        self.last_topas_backup_paths = []
        if hasattr(self, 'results_notebook'):
            for tab_id in self.results_notebook.tabs():
                self.results_notebook.forget(tab_id)
        if hasattr(self, 'result_summary_lbl'):
            self.result_summary_lbl.config(text='요금 조회 전', fg=self.fg_muted)
        if hasattr(self, 'route_var'):
            self.route_var.set('')
            self._refresh_route_combo_values('')
        self._append_topas_log('[새로조회] 기존 조회내용과 계산 결과를 초기화했습니다.\n')

    def _on_route_search_keyrelease(self, event=None):
        ignored_keys = {
            'Up', 'Down', 'Left', 'Right', 'Home', 'End', 'Prior', 'Next',
            'Tab', 'Escape', 'Return', 'KP_Enter',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R',
            'Caps_Lock',
        }
        if event is not None and event.keysym in ignored_keys:
            return
        self.route_user_modified = True
        self._refresh_route_combo_values(self.route_var.get())

    def _on_route_selected(self, _event=None):
        self.route_user_modified = True
        self._refresh_route_combo_values('')

    def _commit_route_search(self, _event=None):
        self._resolve_route_input(show_warning=False)
        return 'break'

    def _refresh_route_combo_values(self, query=''):
        if not hasattr(self, 'route_combo'):
            return []
        values = filter_routes(self.fare_routes, query, limit=80)
        try:
            self.route_combo['values'] = values
        except Exception:
            pass
        return values

    def _auto_select_route_from_raw(self, force=False):
        if not self.fare_routes:
            return
        if self.route_user_modified and not force:
            return

        inference = infer_route_from_topas_text(
            self._get_raw_text('departure'),
            self._get_raw_text('return'),
            self.fare_routes,
            year_hint=datetime.now().year,
        )
        if not inference.route:
            return

        if self.route_var.get().strip() != inference.route:
            self.route_var.set(inference.route)
            self._refresh_route_combo_values('')
            self._append_topas_log(f'[노선 자동 선택] {inference.route} ({inference.reason})\n')

    def _resolve_route_input(self, show_warning=True):
        query = self.route_var.get().strip()
        if query in self.fare_routes:
            self.route_user_modified = True
            self._refresh_route_combo_values('')
            return query

        matches = filter_routes(self.fare_routes, query, limit=None)
        if len(matches) == 1:
            self.route_var.set(matches[0])
            self.route_user_modified = True
            self._refresh_route_combo_values('')
            return matches[0]

        if show_warning:
            if matches:
                sample = '\n'.join(matches[:10])
                more = '' if len(matches) <= 10 else f'\n... 외 {len(matches) - 10}건'
                messagebox.showwarning(
                    '노선 확인',
                    f'노선 후보가 {len(matches)}개입니다. 더 구체적으로 입력해 주세요.\n\n{sample}{more}',
                )
            else:
                messagebox.showwarning('노선 확인', '입력한 조건과 맞는 노선을 찾지 못했습니다.')
        return None

    def show_slot_raw_popup(self, slot):
        text = self._get_raw_text(slot)
        if not text:
            messagebox.showwarning('조회내용 없음', '표시할 TOPAS 조회내용이 없습니다.')
            return
        self.show_topas_results_popup([text], 0, stopped=False)

    def save_slot_raw_manually(self, slot):
        text = self._get_raw_text(slot)
        if not text:
            messagebox.showwarning('조회내용 없음', '저장할 TOPAS 조회내용이 없습니다.')
            return
        title = '출발편' if slot == 'departure' else '귀국편'
        file_path = filedialog.asksaveasfilename(
            defaultextension='.txt',
            filetypes=[('Text Files', '*.txt'), ('All Files', '*.*')],
            title=f'{title} 텍스트 저장',
            initialfile=f'topas_{slot}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt',
        )
        if not file_path:
            return
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(text)
        self._append_topas_log(f'[저장] {title} 텍스트 저장: {file_path}\n')

    def _resolve_app_path(self, path_value):
        path_text = str(path_value or '').strip()
        if not path_text:
            return get_app_dir()
        if os.path.isabs(path_text):
            return path_text
        return os.path.join(get_app_dir(), path_text)

    def _selected_nights(self):
        selected = []
        for night, var in sorted(self.night_vars.items()):
            try:
                if var.get():
                    selected.append(int(night))
            except Exception:
                pass
        return sorted(set(n for n in selected if 1 <= n <= 30))

    def _select_only_night(self, target_night):
        keep_selected = False
        try:
            keep_selected = bool(self.night_vars[target_night].get())
        except Exception:
            pass
        if not keep_selected:
            return
        for night, var in self.night_vars.items():
            if night != target_night:
                try:
                    var.set(False)
                except Exception:
                    pass
        self._refresh_night_chips()
        self._request_recalculate_after_night_change()

    def _set_single_night(self, target_night):
        for night, var in self.night_vars.items():
            try:
                var.set(night == target_night)
            except Exception:
                pass
        self._refresh_night_chips()

    def _make_night_chip(self, night):
        """박수 선택 칩 버튼 생성 (단일 선택, 선택 시 파란 칩)."""
        btn = tk.Button(
            self.night_chip_frame,
            text=f'{night}박',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            font=('맑은 고딕', 9, 'bold'),
            padx=12,
            pady=4,
            bg=self.btn_neutral_bg,
            fg=self.fg_muted,
            activebackground=self.accent_color,
            activeforeground='white',
            command=lambda n=night: self._on_night_chip_click(n),
        )
        btn.pack(side=tk.LEFT, padx=(6, 0))
        self.night_chip_buttons[night] = btn
        self._refresh_night_chips()
        return btn

    def _on_night_chip_click(self, night):
        var = self.night_vars.get(night)
        if var is None:
            return
        try:
            var.set(not bool(var.get()))
        except Exception:
            return
        self._select_only_night(night)
        self._refresh_night_chips()

    def _refresh_night_chips(self):
        for night, btn in getattr(self, 'night_chip_buttons', {}).items():
            try:
                selected = bool(self.night_vars[night].get())
            except Exception:
                selected = False
            try:
                btn.config(
                    bg=self.accent_color if selected else self.btn_neutral_bg,
                    fg='white' if selected else self.fg_muted,
                )
            except Exception:
                pass

    def _request_recalculate_after_night_change(self):
        if not getattr(self, 'v5_calculation_result', None):
            return
        if getattr(self, '_night_auto_recalc_busy', False):
            return
        self._night_auto_recalc_busy = True
        try:
            self.root.after_idle(self._run_night_auto_recalculate)
        except Exception:
            self._run_night_auto_recalculate()

    def _run_night_auto_recalculate(self):
        try:
            self.calculate_topas_fares(auto=True)
        finally:
            self._night_auto_recalc_busy = False

    def add_custom_night_chip(self):
        raw = self.custom_night_var.get().strip()
        try:
            night = int(raw)
        except ValueError:
            messagebox.showwarning('박수 입력', '박수는 숫자로 입력해 주세요.')
            return
        if not 1 <= night <= 30:
            messagebox.showwarning('박수 입력', '박수는 1~30 사이로 입력해 주세요.')
            return
        if night in self.night_vars:
            self._set_single_night(night)
            self.custom_night_var.set('')
            self._request_recalculate_after_night_change()
            return

        var = tk.BooleanVar(value=True)
        self.night_vars[night] = var
        self._make_night_chip(night)
        self._set_single_night(night)
        self.custom_night_var.set('')
        self._request_recalculate_after_night_change()

    def refresh_fare_snapshot(self, force=True):
        if hasattr(self, 'refresh_fare_btn'):
            self.refresh_fare_btn.config(state=tk.DISABLED)
        if hasattr(self, 'fare_data_status_lbl'):
            self.fare_data_status_lbl.config(text='운임/시즌 로드 중...', fg=self.accent_orange)
        thread = threading.Thread(target=self._fare_snapshot_worker, args=(force,), daemon=True)
        thread.start()

    def _fare_snapshot_worker(self, force=True):
        try:
            cache_path = self._resolve_app_path(self.config.get('fare_cache_path', 'cache/fares_snapshot.json'))
            snapshot = load_fare_snapshot(
                self.config,
                cache_path=cache_path,
                prefer_cache_within_hours=None if force else 12,
            )
            self.root.after(0, lambda snap=snapshot: self._apply_fare_snapshot(snap, None))
        except Exception as exc:
            self.root.after(0, lambda err=str(exc): self._apply_fare_snapshot(None, err))

    def _apply_fare_snapshot(self, snapshot, error):
        if hasattr(self, 'refresh_fare_btn'):
            self.refresh_fare_btn.config(state=tk.NORMAL)
        if error:
            if hasattr(self, 'fare_data_status_lbl'):
                self.fare_data_status_lbl.config(text='운임/시즌 로드 실패', fg=self.accent_red)
            self._append_topas_log(f'[운임 로드 오류] {error}\n')
            return

        self.fare_snapshot = snapshot
        self.fare_routes = list(snapshot.routes)
        self._refresh_route_combo_values(self.route_var.get())
        self._auto_select_route_from_raw(force=not self.route_user_modified)
        status = '운임/시즌 준비됨'
        if snapshot.source == 'cache':
            status = '운임/시즌 캐시 사용'
        if hasattr(self, 'fare_data_status_lbl'):
            self.fare_data_status_lbl.config(
                text=status,
                fg=self.accent_orange if snapshot.warning else self.accent_green,
            )
        if snapshot.warning:
            self._append_topas_log(f'[운임 캐시] {snapshot.warning}\n')
        else:
            self._append_topas_log(f'[운임 로드] {status}\n')

    def calculate_topas_fares(self, auto=False):
        dep_text = self._get_raw_text('departure')
        ret_text = self._get_raw_text('return')
        nights = self._selected_nights()

        if not dep_text or not ret_text:
            messagebox.showwarning('조회내용 확인', '출발편과 귀국편 조회내용을 모두 입력해 주세요.')
            return
        if self.fare_snapshot is None:
            messagebox.showwarning('운임/시즌 데이터', '운임/시즌 데이터를 먼저 불러와야 계산할 수 있습니다.')
            return
        route = self._resolve_route_input(show_warning=True)
        if not route:
            return
        if not nights:
            messagebox.showwarning('박수 확인', '계산할 박수를 1개 이상 선택해 주세요.')
            return

        try:
            result = calculate_round_trips(
                dep_text,
                ret_text,
                self.fare_snapshot.fares,
                self.fare_snapshot.seasons,
                route,
                nights=nights,
            )
        except Exception as exc:
            messagebox.showerror('요금 조회 실패', f'계산 중 오류가 발생했습니다.\n{exc}')
            return

        self.v5_calculation_result = result
        self._render_calculation_results(result)
        if result.warnings:
            self._append_topas_log('[계산 경고]\n' + '\n'.join(f'- {w}' for w in result.warnings) + '\n')

    def _render_calculation_results(self, result):
        for tab_id in self.results_notebook.tabs():
            self.results_notebook.forget(tab_id)

        first_tab = None
        self.result_trees = {}
        for night in sorted(result.result.keys()):
            rows = result.result[night]

            tab = tk.Frame(self.results_notebook, bg=self.card_color)
            self.results_notebook.add(tab, text=f'{night}박')
            if first_tab is None:
                first_tab = tab
                self.selected_result_night = night

            header = tk.Frame(tab, bg=self.card_color)
            header.pack(fill=tk.X, pady=(4, 6))
            tk.Label(
                header,
                text=f'{night}박 결과 · {len(rows)}건',
                bg=self.card_color,
                fg=self.fg_color,
                font=('맑은 고딕', 9, 'bold'),
            ).pack(side=tk.LEFT)
            copy_btn = tk.Button(
                header,
                text='복사',
                bg=self.btn_neutral_bg,
                fg=self.btn_neutral_fg,
                font=('맑은 고딕', 8, 'bold'),
                activebackground=self.btn_neutral_hover,
                activeforeground='white',
                bd=0,
                relief=tk.FLAT,
                cursor='hand2',
                padx=8,
                pady=3,
                command=lambda n=night: self.copy_result_tsv(n),
            )
            copy_btn.pack(side=tk.RIGHT)
            self._add_hover(copy_btn, self.btn_neutral_bg, self.btn_neutral_hover)

            send_btn = tk.Button(
                header,
                text='요금수정 보내기',
                bg=self.accent_color,
                fg='white',
                font=('맑은 고딕', 8, 'bold'),
                activebackground=self.accent_hover,
                activeforeground='white',
                bd=0,
                relief=tk.FLAT,
                cursor='hand2',
                padx=8,
                pady=3,
                command=lambda n=night: self.send_calculation_to_erp_sheet(n),
            )
            send_btn.pack(side=tk.RIGHT, padx=(0, 8))
            self._add_hover(send_btn, self.accent_color, self.accent_hover)

            excel_btn = tk.Button(
                header,
                text='엑셀 다운로드',
                bg=self.btn_neutral_bg,
                fg=self.btn_neutral_fg,
                font=('맑은 고딕', 8, 'bold'),
                activebackground=self.btn_neutral_hover,
                activeforeground='white',
                bd=0,
                relief=tk.FLAT,
                cursor='hand2',
                padx=8,
                pady=3,
                command=self.download_calculation_excel,
            )
            excel_btn.pack(side=tk.RIGHT, padx=(0, 8))
            self._add_hover(excel_btn, self.btn_neutral_bg, self.btn_neutral_hover)

            process_btn = tk.Button(
                header,
                text='계산 과정 보기',
                bg=self.btn_neutral_bg,
                fg=self.btn_neutral_fg,
                font=('맑은 고딕', 8, 'bold'),
                activebackground=self.btn_neutral_hover,
                activeforeground='white',
                bd=0,
                relief=tk.FLAT,
                cursor='hand2',
                padx=8,
                pady=3,
                command=self.show_calculation_debug_popup,
            )
            process_btn.pack(side=tk.RIGHT, padx=(0, 8))
            self._add_hover(process_btn, self.btn_neutral_bg, self.btn_neutral_hover)

            tree = ttk.Treeview(tab, columns=('date', 'fare', 'status'), show='headings', height=8)
            tree.heading('date', text='출발일')
            tree.heading('fare', text='왕복요금')
            tree.heading('status', text='상태')
            tree.column('date', width=140, anchor=tk.CENTER)
            tree.column('fare', width=140, anchor=tk.E)
            tree.column('status', width=90, anchor=tk.CENTER)
            tree.pack(fill=tk.BOTH, expand=True)
            tree.tag_configure('even', background=self.tree_bg)
            tree.tag_configure('odd', background=self.tree_row_alt)
            tree.tag_configure('closed', foreground='#8a7480')
            for idx, row in enumerate(rows):
                fare_text = '' if row.is_closed else f'{js_round(row.total_fare):,}원'
                status_text = '마감' if row.is_closed else ''
                tags = ['odd' if idx % 2 else 'even']
                if row.is_closed:
                    tags.append('closed')
                tree.insert('', tk.END, values=(row.dep_date, fare_text, status_text), tags=tuple(tags))
            self.result_trees[night] = tree

        self.results_notebook.bind('<<NotebookTabChanged>>', self._on_result_tab_changed)
        if first_tab is not None:
            self.results_notebook.select(first_tab)
        summary_text = ', '.join(f'{night}박 {len(rows)}건' for night, rows in sorted(result.result.items())) or '계산 결과 없음'
        self._append_topas_log('[계산 완료] ' + summary_text + '\n')

    def _on_result_tab_changed(self, _event=None):
        if not getattr(self, 'v5_calculation_result', None):
            return
        current = self.results_notebook.select()
        for idx, tab_id in enumerate(self.results_notebook.tabs()):
            if tab_id == current:
                nights = sorted(self.v5_calculation_result.result.keys())
                if idx < len(nights):
                    self.selected_result_night = nights[idx]
                break

    def copy_result_tsv(self, night):
        if not self.v5_calculation_result:
            return
        rows = self.v5_calculation_result.result.get(int(night), [])
        text = results_to_tsv(rows)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self._append_topas_log(f'[복사] {night}박 결과 {len(rows)}건 TSV 복사\n')

    def download_calculation_excel(self):
        if not self.v5_calculation_result:
            messagebox.showwarning('결과 없음', '먼저 요금 조회를 실행해 주세요.')
            return
        route = self.route_var.get().strip() or 'route'
        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel Files', '*.xlsx')],
            title='요금 조회 결과 저장',
            initialfile=f'요금조회_{route}_{datetime.now().strftime("%Y%m%d")}.xlsx',
        )
        if not file_path:
            return
        info = {
            '항공노선': route,
            '운임 기준 시각': self.fare_snapshot.loaded_at if self.fare_snapshot else '',
            '운임 데이터 출처': self.fare_snapshot.source if self.fare_snapshot else '',
            '출발편 백업': ', '.join(self.last_topas_backup_paths),
        }
        export_results_to_excel(file_path, self.v5_calculation_result.result, info=info)
        self._append_topas_log(f'[엑셀] 요금 조회 결과 저장: {file_path}\n')

    def show_calculation_debug_popup(self):
        if not self.v5_calculation_result:
            messagebox.showwarning('결과 없음', '먼저 요금 조회를 실행해 주세요.')
            return
        popup = tk.Toplevel(self.root)
        popup.title('계산 과정')
        popup.configure(bg='#ffffff')
        popup.geometry('980x680')

        text_widget = ScrolledText(
            popup,
            bg='#ffffff',
            fg='#111827',
            insertbackground='#111827',
            font=('Consolas', 10),
            bd=0,
            relief=tk.FLAT,
            padx=12,
            pady=12,
            wrap=tk.NONE,
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert('1.0', self._format_calculation_debug())

        footer = tk.Frame(popup, bg='#ffffff', padx=12, pady=10)
        footer.pack(fill=tk.X)
        tk.Button(
            footer,
            text='닫기',
            width=10,
            bg='#f3f4f6',
            fg='#111827',
            font=('맑은 고딕', 9, 'bold'),
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=popup.destroy,
        ).pack(side=tk.RIGHT)

    def _format_calculation_debug(self):
        result = self.v5_calculation_result
        lines = []
        lines.append(f'항공노선: {self.route_var.get().strip()}')
        if result.warnings:
            lines.append('')
            lines.append('[경고]')
            lines.extend(f'- {warning}' for warning in result.warnings)
        lines.append('')
        lines.append('[출발편]')
        lines.extend(self._format_one_way_debug(result.debug.dep_debug))
        lines.append('')
        lines.append('[귀국편]')
        lines.extend(self._format_one_way_debug(result.debug.ret_debug))
        lines.append('')
        lines.append('[왕복 조합]')
        lines.extend(self._format_round_trip_debug_by_night(result.debug.combinations))
        return '\n'.join(lines)

    def _format_round_trip_debug_by_night(self, rows):
        grouped = {}
        for row in rows:
            grouped.setdefault(int(row.nights), []).append(row)

        lines = []
        if not grouped:
            return ['계산된 왕복 조합이 없습니다.']

        for index, night in enumerate(sorted(grouped.keys())):
            if index:
                lines.append('')
            rows_for_night = grouped[night]
            closed = sum(1 for row in rows_for_night if row.is_closed)
            suffix = f' · 마감 {closed}건' if closed else ''
            lines.append(f'[{night}박] {len(rows_for_night)}건{suffix}')
            lines.append('출발일\t귀국일\t출발편도\t귀국편도\t왕복\t상태\t시즌')
            for row in rows_for_night:
                status = '마감' if row.is_closed else ''
                lines.append(
                    f'{row.dep_date}\t{row.ret_date}\t'
                    f'{js_round(row.dep_fare)}\t{js_round(row.ret_fare)}\t'
                    f'{js_round(row.total_fare)}\t{status}\t{row.dep_season}/{row.ret_season}'
                )
        return lines

    def _format_one_way_debug(self, rows):
        lines = ['날짜\t시즌\t가용 클래스\t선택\t편도요금\t상태']
        for row in rows:
            status = '마감' if row.is_closed else ''
            fare = '' if row.is_closed else str(js_round(row.final_fare))
            lines.append(
                f'{row.date}\t{row.season_type}\t{",".join(row.classes)}\t'
                f'{row.selected_class}\t{fare}\t{status}'
            )
        return lines

    def send_calculation_to_erp_sheet(self, night=None):
        if not self.v5_calculation_result:
            messagebox.showwarning('결과 없음', '먼저 요금 조회를 실행해 주세요.')
            return

        nights = [night for night, rows in sorted(self.v5_calculation_result.result.items()) if rows]
        if not nights:
            messagebox.showwarning('결과 없음', '요금수정으로 보낼 계산 결과가 없습니다.')
            return
        target_night = int(night) if night is not None else self.selected_result_night
        target_night = target_night if target_night in nights else nights[0]
        rows = self.v5_calculation_result.result.get(target_night, [])
        erp_rows = to_erp_rows(rows)
        closed_rows = [row for row in rows if row.is_closed]
        if not erp_rows:
            messagebox.showwarning('전달 불가', f'{target_night}박 결과에 요금수정으로 보낼 수 있는 행이 없습니다.')
            return
        self._record_sheet_undo_state('요금수정 보내기')
        airline_code = self._airline_code_from_route()
        self._select_airline_code(airline_code, overwrite_blank_only=False)
        source_text = f'토파스 계산 ({self.route_var.get().strip()} · {target_night}박 · {len(erp_rows)}건'
        if airline_code:
            source_text += f' · 항공사코드 {airline_code}'
        source_text += ')'
        self._load_erp_rows_to_sheet(
            erp_rows,
            source_text,
            record_undo=False,
        )
        if closed_rows:
            self._show_closed_exclusion_popup(closed_rows, target_night)

    def _closed_rows_copy_text(self, closed_rows, target_night):
        lines = [f'{target_night}박 마감 제외 목록', '출발일\t귀국일\t박수\t상태']
        for row in closed_rows:
            lines.append(f'{row.dep_date}\t{row.ret_date}\t{row.nights}\t마감')
        return '\n'.join(lines)

    def _show_closed_exclusion_popup(self, closed_rows, target_night):
        result_text = self._closed_rows_copy_text(closed_rows, target_night)
        popup_bg = '#ffffff'
        popup_fg = '#111827'
        popup_border = '#d1d5db'

        popup = tk.Toplevel(self.root)
        popup.title('마감 제외')
        popup.configure(bg=popup_bg)
        popup.geometry('460x420')
        popup.minsize(390, 320)
        popup.transient(self.root)

        header = tk.Frame(popup, bg=popup_bg, padx=16, pady=0)
        header.pack(fill=tk.X, pady=(14, 8))
        tk.Label(
            header,
            text=f'마감 {len(closed_rows)}건은 ERP 입력표에서 제외했습니다.',
            font=('맑은 고딕', 10, 'bold'),
            bg=popup_bg,
            fg=popup_fg,
        ).pack(anchor=tk.W)

        body = tk.Frame(popup, bg=popup_border)
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 10))
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
            height=10,
        )
        text_widget.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        text_widget.insert('1.0', result_text)
        text_widget.config(state=tk.DISABLED)

        footer = tk.Frame(popup, bg=popup_bg, padx=16)
        footer.pack(fill=tk.X, pady=(0, 14))

        def copy_closed_rows():
            self.root.clipboard_clear()
            self.root.clipboard_append(result_text)
            self.root.update()
            self._append_topas_log(f'[복사] {target_night}박 마감 제외 목록 {len(closed_rows)}건 복사\n')

        copy_btn = tk.Button(
            footer,
            text='목록 복사',
            width=12,
            bg=self.accent_color,
            fg='white',
            font=('맑은 고딕', 9, 'bold'),
            activebackground=self.accent_hover,
            activeforeground='white',
            bd=0,
            relief=tk.FLAT,
            cursor='hand2',
            command=copy_closed_rows,
        )
        copy_btn.pack(side=tk.LEFT)
        self._add_hover(copy_btn, self.accent_color, self.accent_hover)

        close_btn = tk.Button(
            footer,
            text='확인',
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

    def _load_erp_rows_to_sheet(self, rows, source_text, record_undo=True):
        if record_undo:
            self._record_sheet_undo_state('요금수정 보내기')
        self.period_mode_var.set(False)
        self.sheet.headers(SHEET_HEADERS)
        try:
            self.sheet.set_column_widths([100, 80, 80, 80, 80, 80, 80, 80])
        except Exception:
            pass
        grid = [
            [
                row['date'],
                str(row['adult_air']),
                str(row.get('adult_hotel', '')),
                str(row.get('adult_land', '')),
                str(row.get('adult_tour', '')),
                str(row.get('adult_profit', '')),
                str(row.get('child_fare', '')),
                str(row.get('infant_fare', '')),
            ]
            for row in rows
        ]
        self.formulas.clear()
        self._results.clear()
        self._clear_merge_restore_snapshot()
        self.sheet.set_sheet_data(grid, reset_col_positions=False, reset_row_positions=True)
        self._set_source_badge(f'출처: {source_text}')
        self._select_main_tab('fare')
        if not self.panel_expanded:
            self.toggle_sheet_panel()
        self.refresh_count()
        self._sync_sheet_undo_baseline()
        self.set_status(f'{len(rows)}건을 요금수정 입력표에 불러왔습니다. 검토 후 시작하세요.', self.accent_green)

    def _set_source_badge(self, text):
        if hasattr(self, 'data_source_badge'):
            # 빈 배지가 회색 상자로 남지 않게, 내용이 있을 때만 카드 배경을 쓴다
            self.data_source_badge.config(
                text=text or '',
                bg=self.card_color if text else self.bg_color,
            )

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
        undo_snapshot = self._snapshot_sheet_state(period_mode=not is_period)

        if not is_period and self._restore_merged_fare_source_if_unchanged(current_data):
            return

        self._record_sheet_undo_state('기간 모드 전환', snapshot=undo_snapshot)

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
        self._sync_sheet_undo_baseline()

    def _normalize_sheet_rows_for_compare(self, rows, width):
        normalized = []
        for row in rows or []:
            values = list(row) + [''] * max(0, width - len(row))
            normalized.append([str(value).strip() for value in values[:width]])
        return normalized

    def _restore_merged_fare_source_if_unchanged(self, current_data):
        if self._merged_fare_restore_data is None or self._merged_fare_merged_data is None:
            return False

        self._record_sheet_undo_state('요금구간 병합 복원', snapshot=self._snapshot_sheet_state(period_mode=True))
        self.formulas.clear()
        self._results.clear()
        self.sheet.headers(["날짜", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"])
        self.sheet.set_sheet_data(
            [list(row) for row in self._merged_fare_restore_data],
            reset_col_positions=True,
            reset_row_positions=False,
        )
        try:
            self.sheet.set_column_widths([100, 80, 80, 80, 80, 80, 80, 80])
        except Exception:
            pass
        self._set_source_badge('요금구간 병합 전 데이터')
        self.refresh_count()
        self._load_active_into_fb()
        self.set_status('요금구간 병합 전 일자별 데이터로 복원했습니다.', self.accent_green)
        self._clear_merge_restore_snapshot()
        self._sync_sheet_undo_baseline()
        return True

    def _clear_merge_restore_snapshot(self):
        self._merged_fare_restore_data = None
        self._merged_fare_merged_data = None

    # ------------------------------------------------------------------
    # 그리드 조작
    # ------------------------------------------------------------------
    def _snapshot_sheet_state(self, period_mode=None):
        try:
            data = self.sheet.get_sheet_data()
        except Exception:
            data = []
        try:
            headers = list(self.sheet.headers())
        except Exception:
            headers = list(SHEET_HEADERS)
        try:
            source_text = self.data_source_badge.cget('text') if hasattr(self, 'data_source_badge') else ''
        except Exception:
            source_text = ''
        return {
            'data': copy.deepcopy(data),
            'headers': headers,
            'period_mode': bool(self.period_mode_var.get() if period_mode is None else period_mode),
            'formulas': copy.deepcopy(self.formulas),
            'results': copy.deepcopy(self._results),
            'merged_restore': copy.deepcopy(self._merged_fare_restore_data),
            'merged_merged': copy.deepcopy(self._merged_fare_merged_data),
            'source_text': source_text,
            'airline_text': self.airline_var.get() if hasattr(self, 'airline_var') else AIRLINE_EMPTY_LABEL,
            'price_desc_text': self.price_desc_var.get() if hasattr(self, 'price_desc_var') else '',
            'hotel_name_text': self.hotel_name_var.get() if hasattr(self, 'hotel_name_var') else '',
            'hotel_seq_text': getattr(self, 'current_hotel_seq', ''),
            'progress_text': self.progress_text_var.get() if hasattr(self, 'progress_text_var') else '',
            'active_cell': tuple(self._active_cell) if hasattr(self, '_active_cell') else (0, 0),
        }

    def _same_sheet_snapshot(self, left, right):
        return left == right

    def _record_sheet_undo_state(self, label='입력표 변경', snapshot=None):
        if self._applying_sheet_snapshot:
            return
        snapshot = copy.deepcopy(snapshot or self._snapshot_sheet_state())
        if self._sheet_undo_stack and self._same_sheet_snapshot(self._sheet_undo_stack[-1], snapshot):
            return
        self._sheet_undo_stack.append(snapshot)
        if len(self._sheet_undo_stack) > self._sheet_undo_limit:
            self._sheet_undo_stack = self._sheet_undo_stack[-self._sheet_undo_limit:]
        self._sheet_redo_stack.clear()

    def _sync_sheet_undo_baseline(self):
        if hasattr(self, 'sheet'):
            self._sheet_last_snapshot = self._snapshot_sheet_state()

    def _column_widths_for_headers(self, headers, period_mode):
        if period_mode or (headers and str(headers[0]).strip() == '시작일'):
            return [100, 100, 75, 75, 75, 75, 75, 70, 70]
        return [100, 80, 80, 80, 80, 80, 80, 80]

    def _restore_sheet_snapshot(self, snapshot):
        self._applying_sheet_snapshot = True
        try:
            period_mode = bool(snapshot.get('period_mode', False))
            self.period_mode_var.set(period_mode)
            headers = list(snapshot.get('headers') or (["시작일", "종료일", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"] if period_mode else SHEET_HEADERS))
            self.sheet.headers(headers)
            try:
                self.sheet.set_column_widths(self._column_widths_for_headers(headers, period_mode))
            except Exception:
                pass
            self.formulas = copy.deepcopy(snapshot.get('formulas') or {})
            self._results = copy.deepcopy(snapshot.get('results') or {})
            self._merged_fare_restore_data = copy.deepcopy(snapshot.get('merged_restore'))
            self._merged_fare_merged_data = copy.deepcopy(snapshot.get('merged_merged'))
            self.sheet.set_sheet_data(copy.deepcopy(snapshot.get('data') or []), reset_col_positions=True, reset_row_positions=True)
            self._set_source_badge(snapshot.get('source_text', ''))
            if hasattr(self, 'airline_var'):
                self.airline_var.set(snapshot.get('airline_text') or AIRLINE_EMPTY_LABEL)
            if hasattr(self, 'price_desc_var'):
                self.price_desc_var.set(snapshot.get('price_desc_text') or '')
            if hasattr(self, 'hotel_name_var'):
                self.hotel_name_var.set(snapshot.get('hotel_name_text') or '')
            self.current_hotel_seq = snapshot.get('hotel_seq_text') or ''
            if hasattr(self, 'progress_text_var'):
                self.progress_text_var.set(snapshot.get('progress_text') or '')
            active = snapshot.get('active_cell') or (0, 0)
            try:
                self._active_cell = (max(0, int(active[0])), max(0, int(active[1])))
            except Exception:
                self._active_cell = (0, 0)
            try:
                self.sheet.select_cell(*self._active_cell, run_binding_func=False)
            except Exception:
                pass
            try:
                self.sheet.redraw()
            except Exception:
                pass
        finally:
            self._applying_sheet_snapshot = False
        self.refresh_count()
        self._load_active_into_fb()
        self._sheet_last_snapshot = self._snapshot_sheet_state()

    def undo_sheet(self):
        while self._sheet_undo_stack:
            current = self._snapshot_sheet_state()
            snapshot = self._sheet_undo_stack.pop()
            if not self._same_sheet_snapshot(current, snapshot):
                self._sheet_redo_stack.append(current)
                self._restore_sheet_snapshot(snapshot)
                self.set_status('입력표 실행취소 완료', self.accent_green)
                return
        try:
            self.sheet.undo()
        except Exception:
            pass
        self._on_sheet_modified()

    def redo_sheet(self):
        while self._sheet_redo_stack:
            current = self._snapshot_sheet_state()
            snapshot = self._sheet_redo_stack.pop()
            if not self._same_sheet_snapshot(current, snapshot):
                self._sheet_undo_stack.append(current)
                self._restore_sheet_snapshot(snapshot)
                self.set_status('입력표 다시실행 완료', self.accent_green)
                return
        try:
            self.sheet.redo()
        except Exception:
            pass
        self._on_sheet_modified()

    # ------------------------------------------------------------------
    # 수식 엔진: 셀에는 결과 표시, 식은 self.formulas 에 보관 + 자동 재계산
    # ------------------------------------------------------------------
    def _on_sheet_modified(self, event=None):
        if self._recalc_busy or self._applying_sheet_snapshot:
            return
        before = copy.deepcopy(self._sheet_last_snapshot)
        current_before_recalc = self._snapshot_sheet_state()
        if before is not None and not self._same_sheet_snapshot(before, current_before_recalc):
            self._record_sheet_undo_state('셀 편집', snapshot=before)
        self._recalc_busy = True
        try:
            self._sync_formulas_from_cells()
            self._recalc_formulas()
        finally:
            self._recalc_busy = False
        self.refresh_count()
        self._sync_sheet_undo_baseline()

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
        self._record_sheet_undo_state('수식 전체 적용')

        for i in range(last + 1):
            self.formulas[(i, col)] = formula
            self._results.pop((i, col), None)
        self._recalc_busy = True
        try:
            self._recalc_formulas()
        finally:
            self._recalc_busy = False

        self.refresh_count()
        self._load_active_into_fb()
        self._sync_sheet_undo_baseline()
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
        self._record_sheet_undo_state('수식줄 입력')
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
        self._sync_sheet_undo_baseline()
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
        self._record_sheet_undo_state('행 추가')
        data = self.sheet.get_sheet_data()
        data.append(["", "", "", "", "", "", "", ""])
        self.sheet.set_sheet_data(data, reset_col_positions=False, reset_row_positions=True)
        self.refresh_count()
        self._load_active_into_fb()
        self._sync_sheet_undo_baseline()

    def delete_selected_rows(self):
        try:
            rows = self.sheet.get_selected_rows(get_cells_as_rows=True)
        except Exception:
            rows = set()
        rows = sorted(rows, reverse=True)
        if not rows:
            messagebox.showinfo('행 삭제', '삭제할 행을 먼저 선택해 주세요. (왼쪽 행 번호를 클릭하면 행 전체가 선택됩니다.)')
            return
        self._record_sheet_undo_state('행 삭제')
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
        self._sync_sheet_undo_baseline()

    def clear_sheet(self):
        if not messagebox.askyesno('전체 지우기', '입력한 모든 행을 지울까요?'):
            return
        self._record_sheet_undo_state('전체 지우기')
        self.formulas.clear()
        self._results.clear()
        self._clear_merge_restore_snapshot()
        self.sheet.set_sheet_data([["", "", "", "", "", "", "", ""] for _ in range(INITIAL_BLANK_ROWS)], reset_col_positions=False, reset_row_positions=True)
        self._set_source_badge('')
        if hasattr(self, 'airline_var'):
            self.airline_var.set(AIRLINE_EMPTY_LABEL)
        if hasattr(self, 'price_desc_var'):
            self.price_desc_var.set('')
        if hasattr(self, 'hotel_name_var'):
            self.hotel_name_var.set('')
        self.current_hotel_seq = ''
        if hasattr(self, 'progress_text_var'):
            self.progress_text_var.set('')
        self.selected_airline_code = ''
        self.selected_price_desc = ''
        self.selected_hotel_name = ''
        self.selected_hotel_seq = ''
        self.selected_progress_text = ''
        self.fares_data = []
        self.editing_job_index = None
        self.refresh_count()
        self._load_active_into_fb()
        self._sync_sheet_undo_baseline()
        self.set_status('입력표를 지우고 항공사코드 목록을 새로고침합니다.', self.fg_muted)
        self.refresh_airline_options_from_erp(silent=True)

    def _normalize_job_airline_code(self, value):
        raw = '' if value is None else str(value).strip()
        if not raw or raw == AIRLINE_EMPTY_LABEL:
            return ''
        if raw in getattr(self, 'airline_value_by_label', {}):
            return self.airline_value_by_label[raw]
        bracket_match = re.match(r'^\[([^\]]+)\]', raw)
        if bracket_match:
            return bracket_match.group(1).strip().upper()
        upper = raw.upper()
        if re.fullmatch(r'[A-Z0-9]{1,3}', upper):
            return upper

        def compact(text):
            return re.sub(r'\s+', '', str(text or '')).lower()

        needle = compact(raw)
        for code, label in getattr(self, 'airline_choices', []):
            if code and needle and needle in compact(label):
                return str(code).strip().upper()
        return upper

    def _job_condition_text(self, job):
        price_desc = str(job.get('price_desc') or '').strip() or '전체 요금구분'
        airline = str(job.get('airline_code') or '').strip() or '전체 항공사'
        hotel_name = str(job.get('hotel_name') or '').strip()
        hotel_seq = str(job.get('hotel_seq') or '').strip()
        parts = [f'요금구분 : {price_desc}', f'항공사 : {airline}']
        if hotel_name:
            hotel_label = f'{hotel_name} (hotelSeq {hotel_seq})' if hotel_seq else hotel_name
            parts.append(f'호텔명 : {hotel_label}')
        return ' / '.join(parts)

    def _normalize_job(self, job):
        rows = [dict(row) for row in (job.get('rows') or [])]
        price_desc = str(job.get('price_desc') or '').strip()
        airline_code = self._normalize_job_airline_code(job.get('airline_code'))
        hotel_name = str(job.get('hotel_name') or '').strip()
        hotel_seq = str(job.get('hotel_seq') or '').strip()
        progress_text = str(job.get('progress_text') or '').strip()
        for row in rows:
            row['price_desc'] = str(row.get('price_desc') or price_desc).strip()
            row['airline_code'] = self._normalize_job_airline_code(row.get('airline_code') or airline_code)
            row['hotel_name'] = str(row.get('hotel_name') or hotel_name).strip()
            row['hotel_seq'] = str(row.get('hotel_seq') or hotel_seq).strip()
            row['progress_status'] = str(row.get('progress_status') or progress_text).strip()
        progress = dict(job.get('_progress') or {})
        progress.setdefault('status', '대기')
        progress.setdefault('done', 0)
        progress.setdefault('success', 0)
        progress.setdefault('fail', 0)
        progress.setdefault('skip', 0)
        progress['total'] = len(rows)
        normalized_job = {
            'price_desc': price_desc,
            'airline_code': airline_code,
            'hotel_name': hotel_name,
            'hotel_seq': hotel_seq,
            'progress_text': progress_text,
            'rows': rows,
            'source': str(job.get('source') or '').strip(),
            '_progress': progress,
        }
        label = self._job_condition_text(normalized_job)
        for row in rows:
            row['_job_label'] = label
            row['_job_source'] = normalized_job['source']
        return normalized_job

    def _assign_job_queue_metadata(self):
        for idx, job in enumerate(self.job_queue):
            job['_queue_index'] = idx
            progress = dict(job.get('_progress') or {})
            progress.setdefault('status', '대기')
            progress.setdefault('done', 0)
            progress.setdefault('success', 0)
            progress.setdefault('fail', 0)
            progress.setdefault('skip', 0)
            progress['total'] = len(job.get('rows') or [])
            job['_progress'] = progress
            label = self._job_condition_text(job)
            for row in job.get('rows') or []:
                row['_job_index'] = idx
                row['_job_label'] = label
                row['_job_source'] = job.get('source') or ''

    def _job_status_text(self, job):
        progress = dict(job.get('_progress') or {})
        status = str(progress.get('status') or '대기')
        total = int(progress.get('total') or len(job.get('rows') or []) or 0)
        done = int(progress.get('done') or 0)
        success = int(progress.get('success') or 0)
        fail = int(progress.get('fail') or 0)
        skip = int(progress.get('skip') or 0)
        if status == '대기':
            return f'대기 0/{total}'
        return f'{status} {done}/{total} · 성공 {success} / 실패 {fail} / 건너뜀 {skip}'

    def _set_job_progress_ui(self, job_index, status=None, result_status=None):
        if job_index is None:
            return
        try:
            idx = int(job_index)
        except (TypeError, ValueError):
            return

        def apply_update():
            if not (0 <= idx < len(self.job_queue)):
                return
            job = self.job_queue[idx]
            progress = dict(job.get('_progress') or {})
            progress.setdefault('total', len(job.get('rows') or []))
            progress.setdefault('done', 0)
            progress.setdefault('success', 0)
            progress.setdefault('fail', 0)
            progress.setdefault('skip', 0)
            if status:
                progress['status'] = status
            if result_status:
                progress['done'] = int(progress.get('done') or 0) + 1
                if result_status == 'SUCCESS':
                    progress['success'] = int(progress.get('success') or 0) + 1
                elif result_status == 'SKIP':
                    progress['skip'] = int(progress.get('skip') or 0) + 1
                else:
                    progress['fail'] = int(progress.get('fail') or 0) + 1

                total = int(progress.get('total') or 0)
                done = int(progress.get('done') or 0)
                if total and done >= total:
                    if int(progress.get('fail') or 0):
                        progress['status'] = '완료(실패 있음)'
                    elif int(progress.get('success') or 0):
                        progress['status'] = '완료'
                    else:
                        progress['status'] = '전체 건너뜀'
                else:
                    progress['status'] = '진행 중'
            job['_progress'] = progress
            self._refresh_job_queue_view()

        try:
            self.root.after(0, apply_update)
        except Exception:
            apply_update()

    def _reset_job_progress_for_run(self, jobs):
        for idx, job in enumerate(self.job_queue):
            total = len(job.get('rows') or [])
            job['_progress'] = {
                'status': '대기',
                'done': 0,
                'success': 0,
                'fail': 0,
                'skip': 0,
                'total': total,
            }
        for job in jobs:
            idx = job.get('_queue_index')
            if isinstance(idx, int) and 0 <= idx < len(self.job_queue):
                self.job_queue[idx]['_progress']['total'] = len(job.get('rows') or [])
        self._refresh_job_queue_view()

    def _refresh_job_queue_view(self):
        if not hasattr(self, 'job_tree'):
            return
        for item in self.job_tree.get_children():
            self.job_tree.delete(item)
        total_rows = 0
        for idx, raw_job in enumerate(self.job_queue, start=1):
            job = raw_job
            row_count = len(job.get('rows') or [])
            total_rows += row_count
            condition = f"{idx}. {self._job_condition_text(job)}"
            self.job_tree.insert(
                '',
                tk.END,
                iid=str(idx - 1),
                values=(condition, self._job_status_text(job), f'{row_count}건', job.get('source') or '수동'),
            )
        if hasattr(self, 'job_queue_summary_lbl'):
            if self.job_queue:
                self.job_queue_summary_lbl.config(text=f'{len(self.job_queue)}개 작업 · {total_rows}건')
            else:
                self.job_queue_summary_lbl.config(text='등록된 작업 없음')

    def _update_start_button_label(self):
        if not hasattr(self, 'start_btn'):
            return
        label = '▶  작업목록 수정 시작' if self.job_queue else '▶  요금수정 시작'
        self.start_btn.config(text=label)

    def _selected_job_indices(self):
        if not hasattr(self, 'job_tree'):
            return []
        indices = []
        for iid in self.job_tree.selection():
            try:
                idx = int(iid)
            except (TypeError, ValueError):
                continue
            if 0 <= idx < len(self.job_queue):
                indices.append(idx)
        return sorted(set(indices))

    def _select_job_index(self, index):
        if not hasattr(self, 'job_tree') or index is None:
            return
        if not (0 <= index < len(self.job_queue)):
            return
        iid = str(index)
        try:
            self.job_tree.selection_set(iid)
            self.job_tree.focus(iid)
            self.job_tree.see(iid)
        except Exception:
            pass

    def _set_job_queue(self, jobs, select_index=None):
        self.job_queue = [self._normalize_job(job) for job in jobs if job.get('rows')]
        self._assign_job_queue_metadata()
        if self.editing_job_index is not None and not (0 <= self.editing_job_index < len(self.job_queue)):
            self.editing_job_index = None
        self._refresh_job_queue_view()
        self._update_start_button_label()
        total_rows = sum(len(job.get('rows') or []) for job in self.job_queue)
        self.progress_bar['maximum'] = max(total_rows, 1)
        self.progress_bar['value'] = 0
        self.progress_lbl.config(text=f'0 / {total_rows} (0%)')
        if select_index is not None:
            self._select_job_index(max(0, min(select_index, len(self.job_queue) - 1)))

    def _sheet_has_nonempty_data(self):
        try:
            rows = self.sheet.get_sheet_data()
        except Exception:
            return False
        for row in rows:
            if any(str(cell).strip() for cell in row if cell is not None):
                return True
        return False

    def _clear_sheet_after_job_registration(self):
        self.formulas.clear()
        self._results.clear()
        self._clear_merge_restore_snapshot()
        self.period_mode_var.set(False)
        self.sheet.headers(SHEET_HEADERS)
        try:
            self.sheet.set_column_widths([100, 80, 80, 80, 80, 80, 80, 80])
        except Exception:
            pass
        self.sheet.set_sheet_data(
            [["", "", "", "", "", "", "", ""] for _ in range(INITIAL_BLANK_ROWS)],
            reset_col_positions=False,
            reset_row_positions=True,
        )
        self._set_source_badge('')
        if hasattr(self, 'airline_var'):
            self.airline_var.set(AIRLINE_EMPTY_LABEL)
        if hasattr(self, 'price_desc_var'):
            self.price_desc_var.set('')
        if hasattr(self, 'hotel_name_var'):
            self.hotel_name_var.set('')
        self.current_hotel_seq = ''
        if hasattr(self, 'progress_text_var'):
            self.progress_text_var.set('')
        self.selected_airline_code = ''
        self.selected_price_desc = ''
        self.selected_hotel_name = ''
        self.selected_hotel_seq = ''
        self.selected_progress_text = ''
        self.fares_data = []
        self.editing_job_index = None
        self._load_active_into_fb()
        self._sync_sheet_undo_baseline()

    def add_current_sheet_to_job_queue(self):
        if self.is_running:
            messagebox.showwarning('실행 중', 'ERP 실행 중에는 작업을 추가할 수 없습니다.')
            return
        rows, errors = self.read_sheet_data()
        if errors:
            preview = "\n".join(errors[:10])
            if len(errors) > 10:
                preview += f"\n... 외 {len(errors) - 10}건"
            msg = ("표에서 다음 행에 문제가 있습니다:\n\n"
                   f"{preview}\n\n"
                   "문제 행은 건너뛰고 정상 행만 작업 목록에 추가할까요?")
            if not messagebox.askyesno('입력 확인', msg):
                return
        filtered = self._apply_date_filter(rows)
        if not filtered:
            messagebox.showwarning('데이터 없음', '작업 목록에 추가할 유효한 요금 행이 없습니다.')
            return
        airline_code = self._selected_airline_code()
        price_desc = self.price_desc_var.get().strip() if hasattr(self, 'price_desc_var') else ''
        hotel_name, hotel_seq = self._selected_hotel_filter()
        progress_text = ''
        if not any([price_desc, airline_code, hotel_name]):
            messagebox.showwarning('조건 필요', '작업 목록에 추가하려면 요금구분, 항공사코드, 호텔명 중 하나 이상 입력해 주세요.')
            return
        rows_for_job = []
        for row in filtered:
            item = dict(row)
            item['price_desc'] = price_desc
            item['airline_code'] = airline_code
            item['hotel_name'] = hotel_name
            item['hotel_seq'] = hotel_seq
            item['progress_status'] = str(item.get('progress_status') or progress_text).strip()
            rows_for_job.append(item)
        new_job = self._normalize_job({
            'price_desc': price_desc,
            'airline_code': airline_code,
            'hotel_name': hotel_name,
            'hotel_seq': hotel_seq,
            'progress_text': progress_text,
            'rows': rows_for_job,
            'source': '입력표',
        })
        replace_index = self.editing_job_index
        confirm_action = '수정 반영' if replace_index is not None and 0 <= replace_index < len(self.job_queue) else '추가'
        hotel_confirm_text = hotel_name or '전체 호텔'
        if hotel_seq:
            hotel_confirm_text = f'{hotel_confirm_text} (hotelSeq {hotel_seq})'
        elif hotel_name:
            hotel_confirm_text = f'{hotel_confirm_text} (ERP DB 미선택)'
        confirm_msg = (
            f"아래 조건으로 작업 목록에 {confirm_action}할까요?\n\n"
            f"요금구분: {price_desc or '전체 요금구분'}\n"
            f"항공사: {airline_code or '전체 항공사'}\n"
            f"호텔명: {hotel_confirm_text}\n"
            f"대상 행: {len(rows_for_job)}건"
        )
        if not messagebox.askyesno('작업 목록 확인', confirm_msg):
            return
        if replace_index is not None and 0 <= replace_index < len(self.job_queue):
            self.job_queue[replace_index] = new_job
            select_index = replace_index
            action_text = '수정 반영'
        else:
            self.job_queue.append(new_job)
            select_index = len(self.job_queue) - 1
            action_text = '추가'
        self._set_job_queue(self.job_queue, select_index=select_index)
        self._clear_sheet_after_job_registration()
        self._select_job_index(select_index)
        self.set_status(f'작업 목록에 {self._job_condition_text(new_job)} · {len(rows_for_job)}건을 {action_text}했습니다.', self.accent_green)

    def load_selected_job_to_sheet(self):
        if self.is_running:
            return
        selected = self._selected_job_indices()
        if not selected:
            messagebox.showinfo('작업 불러오기', '요금 입력표로 불러올 작업을 먼저 선택해 주세요.')
            return
        if len(selected) > 1:
            messagebox.showinfo('작업 불러오기', '한 번에 하나의 작업만 입력표로 불러올 수 있습니다.')
            return
        if self._sheet_has_nonempty_data():
            if not messagebox.askyesno('입력표 덮어쓰기', '현재 요금 입력표 내용을 지우고 선택한 작업을 불러올까요?'):
                return
        index = selected[0]
        job = self._normalize_job(self.job_queue[index])
        self._load_job_to_sheet(job)
        self.editing_job_index = index
        self._select_job_index(index)
        self.set_status(f'작업을 입력표로 불러왔습니다. 수정 후 `작업 목록에 추가`를 누르면 목록에 반영됩니다.', self.accent_orange)

    def _load_job_to_sheet(self, job):
        rows = [dict(row) for row in (job.get('rows') or [])]
        is_period = any(str(row.get('date_end') or row.get('date') or '') != str(row.get('date') or '') for row in rows)
        self._record_sheet_undo_state('작업 입력표 불러오기')
        self.period_mode_var.set(is_period)
        headers = ["시작일", "종료일", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"] if is_period else SHEET_HEADERS
        col_widths = [100, 100, 75, 75, 75, 75, 75, 70, 70] if is_period else [100, 80, 80, 80, 80, 80, 80, 80]
        self.sheet.headers(headers)
        try:
            self.sheet.set_column_widths(col_widths)
        except Exception:
            pass
        def display_value(row, field):
            if str(row.get('progress_status_field') or '') == field and str(row.get('progress_status') or '').strip():
                return str(row.get('progress_status') or '').strip()
            return str(row.get(field, ''))

        if is_period:
            grid = [[
                row.get('date', ''),
                row.get('date_end') or row.get('date', ''),
                display_value(row, 'adult_air'),
                display_value(row, 'adult_hotel'),
                display_value(row, 'adult_land'),
                display_value(row, 'adult_tour'),
                display_value(row, 'adult_profit'),
                display_value(row, 'child_fare'),
                display_value(row, 'infant_fare'),
            ] for row in rows]
        else:
            grid = [[
                row.get('date', ''),
                display_value(row, 'adult_air'),
                display_value(row, 'adult_hotel'),
                display_value(row, 'adult_land'),
                display_value(row, 'adult_tour'),
                display_value(row, 'adult_profit'),
                display_value(row, 'child_fare'),
                display_value(row, 'infant_fare'),
            ] for row in rows]
        self.formulas.clear()
        self._results.clear()
        self._clear_merge_restore_snapshot()
        self.sheet.set_sheet_data(grid, reset_col_positions=False, reset_row_positions=True)
        self.price_desc_var.set(str(job.get('price_desc') or ''))
        self._select_airline_code(job.get('airline_code') or '', overwrite_blank_only=False)
        self.hotel_name_var.set(str(job.get('hotel_name') or ''))
        self.current_hotel_seq = str(job.get('hotel_seq') or '').strip()
        self.progress_text_var.set(str(job.get('progress_text') or ''))
        self._set_source_badge(f"편집 중: {self._job_condition_text(job)}")
        if not self.panel_expanded:
            self.toggle_sheet_panel()
        self.refresh_count()
        self._load_active_into_fb()
        self._sync_sheet_undo_baseline()

    def move_selected_job(self, direction):
        if self.is_running:
            return
        selected = self._selected_job_indices()
        if not selected:
            messagebox.showinfo('순서 변경', '순서를 바꿀 작업을 먼저 선택해 주세요.')
            return
        if len(selected) > 1:
            messagebox.showinfo('순서 변경', '한 번에 하나의 작업만 이동할 수 있습니다.')
            return
        index = selected[0]
        new_index = index + int(direction)
        if not (0 <= new_index < len(self.job_queue)):
            return
        self.job_queue[index], self.job_queue[new_index] = self.job_queue[new_index], self.job_queue[index]
        if self.editing_job_index == index:
            self.editing_job_index = new_index
        elif self.editing_job_index == new_index:
            self.editing_job_index = index
        self._set_job_queue(self.job_queue, select_index=new_index)
        self.set_status('작업 실행 순서를 변경했습니다.', self.fg_muted)

    def delete_selected_job(self):
        if self.is_running:
            return
        selected = sorted(self._selected_job_indices(), reverse=True)
        if not selected:
            messagebox.showinfo('작업 삭제', '삭제할 작업을 먼저 선택해 주세요.')
            return
        old_editing_index = self.editing_job_index
        for idx in selected:
            if 0 <= idx < len(self.job_queue):
                del self.job_queue[idx]
        if old_editing_index is not None:
            if old_editing_index in selected:
                self.editing_job_index = None
            else:
                self.editing_job_index = old_editing_index - sum(1 for idx in selected if idx < old_editing_index)
        self._set_job_queue(self.job_queue)
        self.set_status('선택한 작업을 목록에서 삭제했습니다.', self.fg_muted)

    def clear_job_queue(self):
        if self.is_running:
            return
        if not self.job_queue:
            return
        if not messagebox.askyesno('작업 목록 비우기', '등록된 작업 목록을 모두 비울까요?'):
            return
        self.editing_job_index = None
        self._set_job_queue([])
        self.set_status('작업 목록을 비웠습니다.', self.fg_muted)

    @staticmethod
    def _merge_fare_records_preserving_gaps(records):
        sorted_records = sorted(records, key=lambda item: (item['start'], item['end']))
        merged = []
        gaps = []
        coverage_end = None

        for record in sorted_records:
            if coverage_end is not None and record['start'] > coverage_end + timedelta(days=1):
                gaps.append((coverage_end + timedelta(days=1), record['start'] - timedelta(days=1)))

            if (
                merged
                and merged[-1]['key'] == record['key']
                and record['start'] == merged[-1]['end'] + timedelta(days=1)
            ):
                merged[-1]['end'] = record['end']
            else:
                merged.append(record.copy())

            coverage_end = record['end'] if coverage_end is None else max(coverage_end, record['end'])

        return merged, gaps

    @staticmethod
    def _format_gap_summary(gaps, limit=5):
        labels = []
        for start, end in gaps[:limit]:
            if start == end:
                labels.append(start.strftime('%Y-%m-%d'))
            else:
                labels.append(f"{start.strftime('%Y-%m-%d')}~{end.strftime('%Y-%m-%d')}")
        if len(gaps) > limit:
            labels.append(f"외 {len(gaps) - limit}구간")
        return ', '.join(labels)

    def merge_sheet_fare_ranges(self):
        try:
            source_rows = self.sheet.get_sheet_data()
        except Exception as e:
            messagebox.showerror('요금구간 병합 실패', f'입력표를 읽는 중 오류가 발생했습니다.\n{e}')
            return

        is_period = bool(self.period_mode_var.get())
        fare_start_col = 2 if is_period else 1
        restore_source_rows = None if is_period else [
            (list(row) + [''] * max(0, 8 - len(row)))[:8]
            for row in source_rows
            if any(str(value).strip() for value in row)
        ]

        def normalize_fare_cell(value):
            text = '' if value is None else str(value).strip()
            compact = text.replace(',', '')
            if compact:
                try:
                    return str(int(float(compact)))
                except ValueError:
                    pass
            return text

        records = []
        for row_idx, row in enumerate(source_rows, start=1):
            row_values = list(row)
            width = 9 if is_period else 8
            row_values += [''] * max(0, width - len(row_values))
            if not any(str(value).strip() for value in row_values):
                continue

            start_text = normalize_date(row_values[0])
            end_text = normalize_date(row_values[1]) if is_period and str(row_values[1]).strip() else start_text
            if not start_text or not end_text:
                messagebox.showwarning('요금구간 병합', f'{row_idx}행 날짜를 확인해 주세요.')
                return

            try:
                start_date = datetime.strptime(start_text, '%Y-%m-%d').date()
                end_date = datetime.strptime(end_text, '%Y-%m-%d').date()
            except ValueError:
                messagebox.showwarning('요금구간 병합', f'{row_idx}행 날짜 형식을 확인해 주세요.')
                return
            if end_date < start_date:
                messagebox.showwarning('요금구간 병합', f'{row_idx}행 종료일이 시작일보다 빠릅니다.')
                return

            fares = row_values[fare_start_col:fare_start_col + 7]
            fare_key = tuple(normalize_fare_cell(value) for value in fares)
            if not any(fare_key):
                continue
            records.append({
                'start': start_date,
                'end': end_date,
                'fares': [str(value).strip() for value in fares],
                'key': fare_key,
            })

        if not records:
            messagebox.showwarning('요금구간 병합', '병합할 요금 데이터가 없습니다.')
            return

        merged, gaps = self._merge_fare_records_preserving_gaps(records)

        output_rows = [
            [item['start'].strftime('%Y-%m-%d'), item['end'].strftime('%Y-%m-%d')] + item['fares']
            for item in merged
        ]

        self._record_sheet_undo_state('요금구간 병합')
        self._merged_fare_restore_data = restore_source_rows
        self._merged_fare_merged_data = [list(row) for row in output_rows]
        self.period_mode_var.set(True)
        self.formulas.clear()
        self._results.clear()
        self.sheet.headers(["시작일", "종료일", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아", "유아"])
        self.sheet.set_sheet_data(output_rows, reset_col_positions=True, reset_row_positions=True)
        try:
            self.sheet.set_column_widths([100, 100, 75, 75, 75, 75, 75, 70, 70])
        except Exception:
            pass
        self._set_source_badge('요금구간 병합')
        self.refresh_count()
        self._load_active_into_fb()
        self._sync_sheet_undo_baseline()
        gap_note = ''
        if gaps:
            gap_note = f' · 빈 날짜 {len(gaps)}구간 제외'
        status_color = self.accent_orange if gaps else self.accent_green
        self.set_status(f'요금구간 병합 완료: {len(records)}건 -> {len(merged)}건{gap_note}', status_color)
        if gaps:
            gap_summary = self._format_gap_summary(gaps)
            messagebox.showwarning(
                '요금구간 병합',
                '중간에 비는 날짜가 있어 해당 날짜는 기간에 포함하지 않고 구간을 나눴습니다.\n\n'
                f'빈 날짜: {gap_summary}\n\n'
                '비운항/마감일이면 그대로 진행하면 됩니다.',
            )

    def import_excel_to_sheet(self):
        path = excel_loader.select_excel_file()
        if not path:
            return
        try:
            job_res = excel_loader.load_fare_jobs_from_excel(path)
        except Exception as e:
            messagebox.showerror('요금불러오기 실패', f'엑셀 작업 목록을 읽는 중 오류가 발생했습니다.\n{e}')
            return
        if job_res and job_res.get('detected'):
            jobs = job_res.get('jobs') or []
            errors = job_res.get('errors') or []
            if not jobs:
                preview = "\n".join(errors[:10]) if errors else '조건열은 감지됐지만 유효한 작업 행이 없습니다.'
                messagebox.showwarning('작업 목록 없음', preview)
                return
            if self.job_queue:
                append = messagebox.askyesno(
                    '작업 목록 처리',
                    '이미 등록된 작업 목록이 있습니다.\n\n'
                    '예: 기존 목록 뒤에 추가\n'
                    '아니오: 기존 목록을 지우고 새로 불러오기',
                )
                next_jobs = self.job_queue + jobs if append else jobs
            else:
                next_jobs = jobs
            self._set_job_queue(next_jobs)
            self._clear_sheet_after_job_registration()
            self._set_source_badge(f'출처: 다중 작업 엑셀 ({os.path.basename(path)})')
            total_rows = sum(len(job.get('rows') or []) for job in jobs)
            self.set_status(f'엑셀에서 {len(jobs)}개 작업 · {total_rows}건을 작업 목록에 불러왔습니다.', self.accent_green)
            if errors:
                preview = "\n".join(errors[:10])
                if len(errors) > 10:
                    preview += f"\n... 외 {len(errors) - 10}건"
                messagebox.showwarning('일부 행 확인', f'일부 행은 보정 또는 제외했습니다.\n\n{preview}')
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
        if self.job_queue and messagebox.askyesno(
            '작업 목록 확인',
            '단일 엑셀을 입력표로 불러옵니다.\n\n'
            '기존 작업 목록이 있으면 실행 시 작업 목록이 우선됩니다.\n'
            '기존 작업 목록을 비울까요?',
        ):
            self._set_job_queue([])
        self._record_sheet_undo_state('엑셀 불러오기')
        self._clear_merge_restore_snapshot()
        
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
        self._set_source_badge('')
        # 불러온 결과를 바로 볼 수 있도록 패널을 펼친다
        if not self.panel_expanded:
            self.toggle_sheet_panel()
        self.refresh_count()
        self._load_active_into_fb()
        self._sync_sheet_undo_baseline()
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
            fare_text_cells = [
                ("adult_air", adult_air_cell),
                ("adult_hotel", adult_hotel_cell),
                ("adult_land", adult_land_cell),
                ("adult_tour", adult_tour_cell),
                ("adult_profit", adult_profit_cell),
                ("child_fare", child_cell),
                ("infant_fare", infant_cell),
            ]
            progress_status_field, progress_status_text = next(
                ((field, cell) for field, cell in fare_text_cells if self._progress_status_from_text(cell)[0]),
                ("", ""),
            )

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
                raw_cell = row_cells[col_idx] if col_idx < len(row_cells) else None
                raw_text = str(raw_cell).strip() if raw_cell is not None else ""
                if not raw_text:
                    return ""
                if self._progress_status_from_text(raw_text)[0]:
                    return ""
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
                "infant_fare": infant_fare,
                "progress_status": progress_status_text,
                "progress_status_field": progress_status_field,
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
    def toggle_filter_panel(self):
        self.filter_panel_expanded = not bool(getattr(self, 'filter_panel_expanded', False))
        if hasattr(self, 'filter_body'):
            if self.filter_panel_expanded:
                self.filter_body.pack(fill=tk.X, padx=4, pady=(0, 6))
            else:
                self.filter_body.pack_forget()
        if hasattr(self, 'filter_toggle_btn'):
            marker = '▾' if self.filter_panel_expanded else '▸'
            self.filter_toggle_btn.config(text=f'{marker} 날짜 필터 (선택)')

    def _on_filter_mode_change(self, *args):
        try:
            if not hasattr(self, 'filter_input_container') or not hasattr(self, 'filter_tip_lbl'):
                return
            for widget in self.filter_input_container.winfo_children():
                widget.destroy()

            mode = self.filter_mode.get()
            entry_kwargs = dict(bg=self.input_bg, fg=self.fg_color, insertbackground='white', bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)

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
            if hasattr(self, 'filter_toggle_btn'):
                self.filter_toggle_btn.config(state=state)
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
            self.airline_combo.config(state=tk.DISABLED if locked else tk.NORMAL)
        except Exception:
            pass
        try:
            self.price_desc_entry.config(state=tk.DISABLED if locked else tk.NORMAL)
        except Exception:
            pass
        try:
            self.hotel_name_entry.config(state=tk.DISABLED if locked else tk.NORMAL)
        except Exception:
            pass
        try:
            self.hotel_search_btn.config(state=state)
        except Exception:
            pass
        try:
            if locked:
                self.sheet.disable_bindings("edit_cell", "paste", "cut", "delete", "undo")
            else:
                self.sheet.enable_bindings("edit_cell", "paste", "cut", "delete", "undo")
        except Exception:
            pass

    def _current_direct_run_conditions(self):
        hotel_name, hotel_seq = self._selected_hotel_filter()
        return {
            'price_desc': self.price_desc_var.get().strip() if hasattr(self, 'price_desc_var') else '',
            'airline_code': self._selected_airline_code(),
            'hotel_name': hotel_name,
            'hotel_seq': hotel_seq,
        }

    @staticmethod
    def _pending_erp_condition_imports(current_conditions, erp_conditions):
        current_conditions = current_conditions or {}
        erp_conditions = erp_conditions or {}
        pending = {}
        if not str(current_conditions.get('price_desc') or '').strip():
            price_desc = str(erp_conditions.get('price_desc') or '').strip()
            if price_desc:
                pending['price_desc'] = price_desc
        if not str(current_conditions.get('airline_code') or '').strip():
            airline_code = str(erp_conditions.get('airline_code') or '').strip().upper()
            if airline_code:
                pending['airline_code'] = airline_code
                pending['airline_text'] = str(erp_conditions.get('airline_text') or '').strip()
        if not str(current_conditions.get('hotel_name') or '').strip():
            hotel_name = str(erp_conditions.get('hotel_name') or '').strip()
            if hotel_name:
                pending['hotel_name'] = hotel_name
                pending['hotel_seq'] = str(erp_conditions.get('hotel_seq') or '').strip()
        return pending

    def _format_erp_condition_import_message(self, pending):
        lines = [
            'ERP 화면에 입력된 검색 조건을 발견했습니다.',
            '',
            '프로그램에서 비어 있는 조건만 이번 작업 조건으로 가져옵니다.',
            '',
        ]
        if pending.get('price_desc'):
            lines.append(f"요금구분: {pending['price_desc']}")
        if pending.get('airline_code'):
            airline_text = pending.get('airline_text') or pending['airline_code']
            lines.append(f"항공사: {airline_text}")
        if pending.get('hotel_name'):
            hotel_text = pending['hotel_name']
            if pending.get('hotel_seq'):
                hotel_text = f"{hotel_text} (hotelSeq {pending['hotel_seq']})"
            else:
                hotel_text = f"{hotel_text} (hotelSeq 없음)"
            lines.append(f"호텔명: {hotel_text}")
        lines.extend([
            '',
            '이 조건을 가져와서 실행할까요?',
            '',
            '아니오를 누르면 프로그램에 입력된 조건 그대로 실행합니다.',
        ])
        return '\n'.join(lines)

    def _apply_imported_erp_conditions(self, pending):
        pending = pending or {}
        if pending.get('price_desc') and hasattr(self, 'price_desc_var'):
            self.price_desc_var.set(str(pending.get('price_desc') or '').strip())
        if pending.get('airline_code'):
            self._select_airline_code(pending.get('airline_code'), overwrite_blank_only=False)
        if pending.get('hotel_name'):
            self._remember_imported_hotel_condition(
                pending.get('hotel_name'),
                pending.get('hotel_seq'),
            )

    def _read_erp_screen_conditions(self, selectors):
        selectors = selectors or {}
        price_selector = selectors.get('price_desc_input', '#priceDesc')
        airline_selector = selectors.get('airline_select', '#air2Cd')
        hotel_name_selector = selectors.get('hotel_name_input', '#hotelKorNm')
        hotel_seq_selector = selectors.get('hotel_seq_input', '#hotelSeq')
        result = self.driver.execute_script(
            """
            const priceSelector = arguments[0];
            const airlineSelector = arguments[1];
            const hotelNameSelector = arguments[2];
            const hotelSeqSelector = arguments[3];
            const valueOf = (selector) => {
                const el = document.querySelector(selector);
                return el ? String(el.value || '').trim() : '';
            };
            const priceDesc = valueOf(priceSelector);
            const hotelName = valueOf(hotelNameSelector);
            const hotelSeq = valueOf(hotelSeqSelector);
            const airline = document.querySelector(airlineSelector);
            let airlineCode = '';
            let airlineText = '';
            if (airline) {
                airlineCode = String(airline.value || '').trim().toUpperCase();
                const selected = airline.options && airline.selectedIndex >= 0 ? airline.options[airline.selectedIndex] : null;
                airlineText = selected ? String(selected.textContent || '').trim() : '';
            }
            if (!airlineCode || airlineText === '_선택_') {
                airlineCode = '';
                airlineText = '';
            }
            return {
                price_desc: priceDesc,
                airline_code: airlineCode,
                airline_text: airlineText,
                hotel_name: hotelName,
                hotel_seq: hotelSeq
            };
            """,
            price_selector,
            airline_selector,
            hotel_name_selector,
            hotel_seq_selector,
        )
        return result or {}

    def _try_import_erp_conditions_for_direct_run(self):
        current = self._current_direct_run_conditions()
        if all(str(current.get(key) or '').strip() for key in ('price_desc', 'airline_code', 'hotel_name')):
            return current

        selectors = self.config.get('selectors', {})
        driver = None
        previous_driver = self.driver
        try:
            driver, _browser_config = self._connect_matching_debug_browser('ERP', selectors)
            self.driver = driver
            if not self.find_and_switch_frame(selectors.get('search_date_input', '#searchStDate')):
                return current
            erp_conditions = self._read_erp_screen_conditions(selectors)
        except Exception:
            return current
        finally:
            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass
            self.driver = previous_driver

        pending = self._pending_erp_condition_imports(current, erp_conditions)
        if not pending:
            return current
        if not messagebox.askyesno('ERP 조건 가져오기', self._format_erp_condition_import_message(pending)):
            return current

        self._apply_imported_erp_conditions(pending)
        return self._current_direct_run_conditions()

    def _rpa_row_conditions(self, row):
        row = row or {}
        is_job_queue_row = row.get('_job_index') is not None

        def value_for(field, selected_value=''):
            value = str(row.get(field) or '').strip()
            if value or is_job_queue_row:
                return value
            return str(selected_value or '').strip()

        return {
            'airline_code': value_for('airline_code', self.selected_airline_code).upper(),
            'price_desc': value_for('price_desc', self.selected_price_desc),
            'hotel_name': value_for('hotel_name', self.selected_hotel_name),
            'hotel_seq': value_for('hotel_seq', self.selected_hotel_seq),
            'progress_text': value_for('progress_status', self.selected_progress_text),
        }

    # ------------------------------------------------------------------
    # RPA 실행 제어
    # ------------------------------------------------------------------
    def start_rpa(self):
        if self.is_running:
            messagebox.showwarning('실행 중', '이미 실행 중인 작업이 있습니다.')
            return

        if self.job_queue:
            jobs = [self._normalize_job(job) for job in self.job_queue if job.get('rows')]
            for queue_index, job in enumerate(jobs):
                job['_queue_index'] = queue_index
                for row in job.get('rows') or []:
                    row['_job_index'] = queue_index
                    row['_job_label'] = self._job_condition_text(job)
                    row['_job_source'] = job.get('source') or ''
            if not jobs:
                messagebox.showwarning('작업 없음', '실행할 작업 목록이 없습니다.')
                return
            total_rows = sum(len(job.get('rows') or []) for job in jobs)
            preview_lines = [
                f"{idx}. {self._job_condition_text(job)} · {len(job.get('rows') or [])}건"
                for idx, job in enumerate(jobs[:10], start=1)
            ]
            if len(jobs) > 10:
                preview_lines.append(f"... 외 {len(jobs) - 10}개 작업")
            msg = (
                f"총 {len(jobs)}개 작업, {total_rows}건을 순서대로 실행합니다.\n\n"
                + "\n".join(preview_lines)
                + "\n\n계속 실행할까요?"
            )
            if not messagebox.askyesno('작업 목록 실행 확인', msg):
                return
            self.rpa_jobs_to_run = jobs
            self.fares_data = [row for job in jobs for row in (job.get('rows') or [])]
            first_job = jobs[0]
            self.selected_airline_code = first_job.get('airline_code') or ''
            self.selected_price_desc = first_job.get('price_desc') or ''
            self.selected_hotel_name = first_job.get('hotel_name') or ''
            self.selected_hotel_seq = first_job.get('hotel_seq') or ''
            self.selected_progress_text = first_job.get('progress_text') or ''
            self._reset_job_progress_for_run(jobs)
        else:
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
            run_conditions = self._try_import_erp_conditions_for_direct_run()
            self.selected_airline_code = str(run_conditions.get('airline_code') or '').strip().upper()
            self.selected_price_desc = str(run_conditions.get('price_desc') or '').strip()
            self.selected_hotel_name = str(run_conditions.get('hotel_name') or '').strip()
            self.selected_hotel_seq = str(run_conditions.get('hotel_seq') or '').strip()
            self.selected_progress_text = ''
            for row in self.fares_data:
                row['price_desc'] = self.selected_price_desc
                row['airline_code'] = self.selected_airline_code
                row['hotel_name'] = self.selected_hotel_name
                row['hotel_seq'] = self.selected_hotel_seq
            self.rpa_jobs_to_run = [{
                'price_desc': self.selected_price_desc,
                'airline_code': self.selected_airline_code,
                'hotel_name': self.selected_hotel_name,
                'hotel_seq': self.selected_hotel_seq,
                'progress_text': self.selected_progress_text,
                'rows': filtered,
                'source': '입력표',
            }]

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
        self._update_start_button_label()
        self.pause_btn.config(state=tk.DISABLED, text='‖  일시 중지', bg=self.accent_orange)
        self.stop_btn.config(state=tk.DISABLED, text='■  중지')
        self.is_running = False
        self.is_paused = False
        self.rpa_jobs_to_run = []
        self.current_rpa_job = None
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

    def _legacy_debugger_address(self):
        return self.config.get("debugger_address", self.config.get("debuggerAddress", "127.0.0.1:9222"))

    def _debug_browser_config(self, target_name):
        normalized = str(target_name or '').upper()
        if normalized == 'TOPAS':
            address = self.config.get('topas_debugger_address') or self._legacy_debugger_address()
            profile_dir = self.config.get('topas_chrome_profile_dir', 'ChromeProfile')
            default_port = 9222
        else:
            address = self.config.get('erp_debugger_address') or '127.0.0.1:9223'
            profile_dir = self.config.get('erp_chrome_profile_dir', 'ChromeProfile_ERP')
            default_port = 9223

        host, port = self._parse_debugger_address(address, default_port)
        return {
            'address': f'{host}:{port}',
            'host': host,
            'port': port,
            'profile_dir': str(profile_dir or ('ChromeProfile' if normalized == 'TOPAS' else 'ChromeProfile_ERP')),
        }

    def _debug_browser_configs(self, target_name):
        normalized = str(target_name or '').upper()
        legacy = self._legacy_debugger_address()
        configs = [self._debug_browser_config(normalized)]

        if normalized == 'TOPAS':
            fallbacks = [
                (legacy, self.config.get('topas_chrome_profile_dir', 'ChromeProfile'), 9222),
                (self.config.get('erp_debugger_address'), self.config.get('erp_chrome_profile_dir', 'ChromeProfile_ERP'), 9223),
            ]
        else:
            fallbacks = [
                (legacy, self.config.get('topas_chrome_profile_dir', 'ChromeProfile'), 9222),
                (self.config.get('topas_debugger_address'), self.config.get('topas_chrome_profile_dir', 'ChromeProfile'), 9222),
            ]

        for address, profile_dir, default_port in fallbacks:
            if not address:
                continue
            host, port = self._parse_debugger_address(address, default_port)
            profile_default = 'ChromeProfile' if port == 9222 else 'ChromeProfile_ERP'
            configs.append({
                'address': f'{host}:{port}',
                'host': host,
                'port': port,
                'profile_dir': str(profile_dir or profile_default),
            })

        unique = []
        seen = set()
        for config in configs:
            if config['address'] in seen:
                continue
            seen.add(config['address'])
            unique.append(config)
        return unique

    def _parse_debugger_address(self, address, default_port):
        text = str(address or '').strip()
        if not text:
            return '127.0.0.1', int(default_port)
        if ':' not in text:
            try:
                return '127.0.0.1', int(text)
            except ValueError:
                return text, int(default_port)
        host, port_text = text.rsplit(':', 1)
        try:
            port = int(port_text)
        except ValueError:
            port = int(default_port)
        return host or '127.0.0.1', port

    def _connect_debug_browser_config(self, config):
        options = webdriver.ChromeOptions()
        options.add_experimental_option("debuggerAddress", config['address'])
        return webdriver.Chrome(options=options), config

    def _connect_debug_browser(self, target_name):
        return self._connect_debug_browser_config(self._debug_browser_config(target_name))

    def _selector_locator(self, selector):
        is_xpath = str(selector or '').startswith('//') or str(selector or '').startswith('xpath=')
        return (
            By.XPATH if is_xpath else By.CSS_SELECTOR,
            str(selector or '').replace('xpath=', '') if str(selector or '').startswith('xpath=') else str(selector or ''),
        )

    def _driver_has_selector_in_frames(self, driver, selector, depth=0, max_depth=4):
        by, search_val = self._selector_locator(selector)
        try:
            if driver.find_elements(by, search_val):
                return True
        except Exception:
            pass
        if depth >= max_depth:
            return False

        try:
            frames = driver.find_elements(By.CSS_SELECTOR, 'iframe, frame')
        except Exception:
            frames = []
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                if self._driver_has_selector_in_frames(driver, selector, depth + 1, max_depth):
                    return True
            except Exception:
                pass
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass
        return False

    def _browser_locations(self, driver):
        locations = []
        try:
            original_handle = driver.current_window_handle
        except Exception:
            original_handle = None
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
        if original_handle:
            try:
                driver.switch_to.window(original_handle)
                driver.switch_to.default_content()
            except Exception:
                pass
        return ' / '.join(locations) if locations else '(확인 불가)'

    def _switch_to_task_window(self, driver, task_name, url_keywords=(), title_keywords=(), selector=None):
        try:
            original_handle = driver.current_window_handle
        except Exception:
            original_handle = None

        try:
            handles = list(driver.window_handles)
        except Exception:
            handles = []

        url_keywords = tuple(str(keyword).lower() for keyword in url_keywords)
        title_keywords = tuple(str(keyword).lower() for keyword in title_keywords)
        matches = []

        for index, handle in enumerate(handles):
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                title = driver.title or ''
                url = driver.current_url or ''
                title_l = title.lower()
                url_l = url.lower()
                score = 0
                if selector and self._driver_has_selector_in_frames(driver, selector):
                    score += 1000
                if url_keywords and any(keyword in url_l for keyword in url_keywords):
                    score += 100
                if title_keywords and any(keyword in title_l for keyword in title_keywords):
                    score += 50
                if score:
                    matches.append((score, -index, handle, title, url))
            except Exception:
                continue

        if matches:
            matches.sort(reverse=True)
            score, _index, handle, title, url = matches[0]
            driver.switch_to.window(handle)
            driver.switch_to.default_content()
            print(f"[브라우저 선택] {task_name}: {title or '(제목 없음)'} - {url or '(주소 없음)'}")
            return handle

        if original_handle:
            try:
                driver.switch_to.window(original_handle)
                driver.switch_to.default_content()
            except Exception:
                pass

        raise RuntimeError(
            f'{task_name} 작업에 사용할 브라우저 탭을 찾지 못했습니다.\n'
            f'현재 연결된 브라우저: {self._browser_locations(driver)}'
        )

    def _switch_to_topas_window(self, driver):
        handle = self._switch_to_task_window(
            driver,
            'TOPAS',
            url_keywords=('topassellconnect.com', 'topas'),
            title_keywords=('topas', 'sell connect'),
            selector=self.TOPAS_SHELL_ROOT,
        )
        self._topas_window_handle = handle
        driver.switch_to.default_content()
        return handle

    def _switch_to_erp_window(self, driver, selectors):
        return self._switch_to_task_window(
            driver,
            'ERP 요금수정',
            url_keywords=('erp.naeiltour.co.kr',),
            title_keywords=('erp', 'mclick', 'naeiltour'),
            selector=selectors.get('search_date_input') if selectors else None,
        )

    def _connect_matching_debug_browser(self, target_name, selectors=None):
        normalized = str(target_name or '').upper()
        errors = []
        for config in self._debug_browser_configs(normalized):
            driver = None
            try:
                driver, browser_config = self._connect_debug_browser_config(config)
                if normalized == 'TOPAS':
                    self._switch_to_topas_window(driver)
                else:
                    self._switch_to_erp_window(driver, selectors or {})
                return driver, browser_config
            except Exception as exc:
                errors.append(f"{config['address']}: {str(exc).splitlines()[0]}")
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:
                        pass

        label = 'TOPAS' if normalized == 'TOPAS' else 'ERP 요금수정'
        checked = ', '.join(config['address'] for config in self._debug_browser_configs(normalized))
        detail = '\n'.join(f'- {error}' for error in errors) if errors else '- 연결 가능한 디버그 브라우저 없음'
        raise RuntimeError(
            f'{label} 작업에 사용할 브라우저 탭을 URL 기준으로 찾지 못했습니다.\n'
            f'확인한 디버그 주소: {checked}\n'
            f'{detail}'
        )

    # ------------------------------------------------------------------
    # TOPAS 조회 제어
    # ------------------------------------------------------------------
    def prompt_topas_query_count(self):
        if self.is_running:
            messagebox.showwarning('실행 중', '이미 실행 중인 작업이 있습니다.')
            return

        slot = self._prepare_next_topas_target_slot()
        if slot is None:
            return

        dialog = tk.Toplevel(self.root)
        title = self._slot_display_name(slot)
        dialog.title(f'{title} 토파스 조회 횟수')
        dialog.configure(bg=self.bg_color)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        box = tk.Frame(dialog, bg=self.bg_color, padx=18, pady=16)
        box.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            box,
            text=f'{title} AC1 실행 횟수',
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

        suggested_count = self._suggest_topas_ac1_count(slot)
        count_var = tk.StringVar(value=suggested_count)
        count_entry = tk.Entry(
            box,
            textvariable=count_var,
            width=18,
            bg=self.input_bg,
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
            text=(
                f'출발편 {self.departure_ac1_count}회 기준 10% 여유: {suggested_count}회'
                if suggested_count
                else '예: 200 입력 시 AC1을 200번 실행합니다.'
            ),
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
            self.start_topas_query(count, batch_size, slot=slot)

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

    def start_topas_query(self, count, batch_size=1, slot='departure'):
        if self.is_running:
            messagebox.showwarning('실행 중', '이미 실행 중인 작업이 있습니다.')
            return

        self.is_running = True
        self.topas_stop_requested = False
        self.topas_results_raw = []
        self.topas_target_slot = slot
        self.topas_current_ac1_count = max(1, int(count))
        self._topas_reset_element_cache()

        for btn in getattr(self, 'topas_query_buttons', [self.topas_query_btn]):
            btn.config(state=tk.DISABLED)
        self.topas_stop_btn.config(state=tk.NORMAL)
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        self._set_inputs_locked(True)

        self.topas_progress_bar['maximum'] = count
        self.topas_progress_bar['value'] = 0
        self.topas_progress_lbl.config(text=f'0 / {count} (0%)')
        target = '출발편' if slot == 'departure' else '귀국편'
        self._set_topas_status(f'{target} TOPAS 연결 준비 중', self.accent_orange)
        if hasattr(self, 'topas_log_txt'):
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
            self._topas_reset_element_cache()
            driver, browser_config = self._connect_matching_debug_browser('TOPAS')
            self.driver = driver
            self.root.after(
                0,
                lambda addr=browser_config['address']: self._append_topas_log(
                    f'[TOPAS] Entry 브라우저 탭을 확인했습니다. ({addr})\n'
                ),
            )

            first_block = self._topas_get_latest_block(driver)
            if first_block is None:
                if not self._topas_has_entry_shell(driver):
                    location = self._topas_debug_location(driver)
                    raise RuntimeError(
                        f"디버그 브라우저({browser_config['address']})에서 TOPAS Entry 화면을 찾지 못했습니다.\n"
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
        count = max(1, int(count))
        keys = self._topas_build_ac1_keys(count)
        last_error = None

        for attempt in range(3):
            prompt = self._topas_wait_prompt_input(driver, timeout=5 if attempt == 0 else 2)
            if prompt is None:
                last_error = RuntimeError('TOPAS 명령 입력창을 찾지 못했습니다.')
                continue

            try:
                self._topas_prepare_prompt_input(driver, prompt)
                prompt.send_keys(*keys)
                return
            except Exception as exc:
                last_error = exc
                self._topas_prompt_el = None

            try:
                prompt = self._topas_wait_prompt_input(driver, timeout=2)
                if prompt is None:
                    continue
                self._topas_prepare_prompt_input(driver, prompt)
                actions = ActionChains(driver)
                actions.move_to_element(prompt).click()
                for _ in range(count):
                    actions.send_keys('AC1').send_keys(Keys.ENTER)
                actions.perform()
                return
            except Exception as exc:
                last_error = exc
                self._topas_prompt_el = None
                time.sleep(0.2)

        detail = str(last_error).splitlines()[0] if last_error else '입력창 준비 실패'
        raise RuntimeError(
            'TOPAS 명령 입력창이 현재 입력 가능한 상태가 아닙니다. '
            'Entry 화면 로딩이 끝났는지 확인하거나, TOPAS 화면을 한 번 클릭한 뒤 다시 실행해 주세요. '
            f'원인: {detail}'
        )

    def _topas_build_ac1_keys(self, count):
        keys = []
        for _ in range(max(1, int(count))):
            keys.extend(('AC1', Keys.ENTER))
        return keys

    def _topas_prepare_prompt_input(self, driver, prompt):
        driver.execute_script(
            """
            const el = arguments[0];
            el.scrollIntoView({block: 'center', inline: 'center'});
            el.focus();
            el.value = '';
            el.dispatchEvent(new Event('input', {bubbles: true}));
            """,
            prompt,
        )
        try:
            prompt.click()
        except Exception:
            pass
        time.sleep(0.05)

    def _topas_wait_prompt_input(self, driver, timeout=8):
        deadline = time.perf_counter() + timeout
        while time.perf_counter() < deadline:
            prompt = self._topas_find_prompt_input(driver)
            if prompt is not None:
                return prompt
            self._topas_prompt_el = None
            time.sleep(0.1)
        return None

    def _topas_wait_next_block(self, driver, previous_block, timeout=None):
        return self._topas_wait_next_blocks(driver, previous_block, 1, timeout=timeout)[0]

    def _topas_wait_next_blocks(self, driver, previous_block, expected_count, timeout=None):
        if timeout is None:
            timeout = self._float_config('topas_batch_timeout', 80, 10, 300)
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
        cached = getattr(self, '_topas_prompt_el', None)
        if self._topas_prompt_is_interactable(driver, cached):
            return cached

        self._topas_prompt_el = None
        found = self._topas_find_interactable_element(driver, self.TOPAS_PROMPT_INPUT)
        self._topas_prompt_el = found
        return found

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

        handles = list(driver.window_handles)
        preferred_handle = getattr(self, '_topas_window_handle', None)
        if preferred_handle in handles:
            handles = [preferred_handle]

        for handle in handles:
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                found = self._topas_find_element_in_frames(driver, selector)
                if found is not None:
                    self._topas_window_handle = handle
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

        frame_count = len(driver.find_elements(By.CSS_SELECTOR, 'iframe, frame'))
        for i in range(frame_count):
            try:
                driver.switch_to.frame(i)
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

    def _topas_find_interactable_element(self, driver, selector):
        try:
            original_handle = driver.current_window_handle
        except Exception:
            original_handle = None

        handles = list(driver.window_handles)
        preferred_handle = getattr(self, '_topas_window_handle', None)
        if preferred_handle in handles:
            handles = [preferred_handle]

        for handle in handles:
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                found = self._topas_find_interactable_element_in_frames(driver, selector)
                if found is not None:
                    self._topas_window_handle = handle
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

    def _topas_find_interactable_element_in_frames(self, driver, selector, depth=0, max_depth=4):
        for element in driver.find_elements(By.CSS_SELECTOR, selector):
            if self._topas_prompt_is_interactable(driver, element):
                return element
        if depth >= max_depth:
            return None

        frame_count = len(driver.find_elements(By.CSS_SELECTOR, 'iframe, frame'))
        for i in range(frame_count):
            try:
                driver.switch_to.frame(i)
                found = self._topas_find_interactable_element_in_frames(driver, selector, depth + 1, max_depth)
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

    def _topas_prompt_is_interactable(self, driver, prompt):
        if prompt is None:
            return False
        try:
            if not prompt.is_displayed() or not prompt.is_enabled():
                return False
            return bool(
                driver.execute_script(
                    """
                    const el = arguments[0];
                    if (!el || !el.isConnected) return false;
                    const rect = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    return rect.width > 0 &&
                        rect.height > 0 &&
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        style.pointerEvents !== 'none' &&
                        !el.disabled &&
                        !el.readOnly;
                    """,
                    prompt,
                )
            )
        except Exception:
            return False

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
        return self._browser_locations(driver)

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
                self._apply_topas_results_to_slot(results, stopped=True)
            return

        if stopped:
            self._set_topas_status('사용자 중지 · 부분 결과 보관', self.accent_orange)
            self._append_topas_log(f'[중지 완료] {len(results)}개 원문을 보관했습니다. 총 {elapsed:.1f}s\n')
        else:
            self._set_topas_status('조회 완료', self.accent_green)
            self._append_topas_log(f'[조회 완료] {len(results)}개 원문을 보관했습니다. 총 {elapsed:.1f}s\n')

        if results:
            self._apply_topas_results_to_slot(results, stopped=stopped)

    def clean_up_ui_after_topas(self):
        self.is_running = False
        self.topas_stop_requested = False
        for btn in getattr(self, 'topas_query_buttons', [self.topas_query_btn]):
            btn.config(state=tk.NORMAL)
        self.topas_stop_btn.config(state=tk.DISABLED, text='■  중지')
        self.start_btn.config(state=tk.NORMAL)
        self._update_start_button_label()
        self.pause_btn.config(state=tk.DISABLED, text='‖  일시 중지', bg=self.accent_orange)
        self.stop_btn.config(state=tk.DISABLED, text='■  중지')
        self._set_inputs_locked(False)

    def _apply_topas_results_to_slot(self, results, stopped=False):
        slot = getattr(self, 'topas_target_slot', 'departure')
        raw_text = join_raw_blocks(results)
        self._set_raw_text(slot, raw_text)

        direction = 'departure' if slot == 'departure' else 'return'
        route = self.route_var.get().strip() or 'UNKNOWN'
        raw_dir = self._resolve_app_path(self.config.get('topas_raw_dir', 'logs/topas_raw'))
        try:
            backup = save_raw_backup(results, raw_dir, route=route, direction=direction)
            self.last_topas_backup_paths.append(str(backup.path))
            self._append_topas_log(f'[백업] {backup.path}\n')
        except Exception as exc:
            self._append_topas_log(f'[백업 오류] {exc}\n')

        label = '출발편' if slot == 'departure' else '귀국편'
        suffix = '부분 ' if stopped else ''
        self._append_topas_log(f'[슬롯 반영] {label}에 {suffix}조회내용 {len(results)}개를 넣었습니다.\n')
        if slot == 'departure':
            completed_ac1 = max(0, len(results) - 1)
            if not stopped and self.topas_current_ac1_count:
                completed_ac1 = self.topas_current_ac1_count
            self.departure_ac1_count = completed_ac1
            suggested = self._suggest_topas_ac1_count('return')
            if suggested:
                self._append_topas_log(
                    f'[귀국편 추천] 출발편 AC1 {completed_ac1}회 기준으로 귀국편 기본 횟수를 {suggested}회로 제안합니다.\n'
                )
        if slot == 'return':
            self.topas_current_ac1_count = 0

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
            text=f'{status} · 조회내용 {len(results)}개 · {elapsed:.1f}s',
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
            self._append_topas_log('[복사] TOPAS 조회내용 전체를 클립보드에 복사했습니다.\n')

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
    def _set_erp_airline_filter(self, selectors, airline_code):
        airline_code = '' if airline_code is None else str(airline_code).strip().upper()
        selector = selectors.get('airline_select', '#air2Cd')
        result = self.driver.execute_script(
            """
            const selector = arguments[0];
            const value = arguments[1];
            const sel = document.querySelector(selector);
            if (!sel) {
                return {ok: false, reason: 'not_found'};
            }
            sel.value = value;
            const selected = sel.options && sel.selectedIndex >= 0 ? sel.options[sel.selectedIndex] : null;
            if (value && sel.value !== value) {
                return {ok: false, reason: 'missing_option', value: sel.value || ''};
            }
            sel.dispatchEvent(new Event('change', {bubbles: true}));
            try {
                if (window.jQuery) {
                    window.jQuery(sel).trigger('change').trigger('chosen:updated');
                }
            } catch (e) {}
            return {
                ok: true,
                value: sel.value || '',
                text: selected ? (selected.textContent || '').trim() : ''
            };
            """,
            selector,
            airline_code,
        )
        if not result or not result.get('ok'):
            reason = result.get('reason') if isinstance(result, dict) else 'unknown'
            if reason == 'missing_option':
                raise RuntimeError(f"항공사코드 필드에 '{airline_code}' 코드 옵션이 없습니다.")
            raise RuntimeError(f"항공사코드 필드({selector})를 설정하지 못했습니다.")
        return result

    def _set_erp_text_filter(self, selectors, selector_key, default_selector, value, label):
        value = '' if value is None else str(value).strip()
        selector = selectors.get(selector_key, default_selector)
        result = self.driver.execute_script(
            """
            const selector = arguments[0];
            const value = arguments[1];
            const el = document.querySelector(selector);
            if (!el) {
                return {ok: false, reason: 'not_found'};
            }
            el.value = value;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            el.dispatchEvent(new Event('blur', {bubbles: true}));
            return {ok: true, value: el.value || ''};
            """,
            selector,
            value,
        )
        if not result or not result.get('ok'):
            raise RuntimeError(f"{label} 필드({selector})를 설정하지 못했습니다.")
        actual = str(result.get('value') or '').strip()
        if actual != value:
            raise RuntimeError(f"{label} 필드 값이 기대값과 다릅니다. 기대={value!r}, 실제={actual!r}")
        return result

    def _set_erp_hotel_filter(self, selectors, hotel_name, expected_seq='', timeout=3.0, poll=0.15):
        hotel_name = '' if hotel_name is None else str(hotel_name).strip()
        expected_seq = '' if expected_seq is None else str(expected_seq).strip()
        name_selector = selectors.get('hotel_name_input', '#hotelKorNm')
        seq_selector = selectors.get('hotel_seq_input', '#hotelSeq')
        result = self.driver.execute_script(
            """
            const nameSelector = arguments[0];
            const seqSelector = arguments[1];
            const value = arguments[2];
            const el = document.querySelector(nameSelector);
            const seq = document.querySelector(seqSelector);
            if (!el) {
                return {ok: false, reason: 'name_not_found'};
            }
            el.value = value;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            if (!value) {
                if (seq) {
                    seq.value = '';
                    seq.dispatchEvent(new Event('change', {bubbles: true}));
                }
                el.dispatchEvent(new Event('blur', {bubbles: true}));
                return {ok: true, value: el.value || '', seq: seq ? (seq.value || '') : '', cleared: true};
            }
            try {
                if (typeof autoHotel === 'function') {
                    autoHotel({type: 'blur'}, el);
                }
            } catch (e) {
                return {ok: false, reason: 'autoHotel_error', error: String(e)};
            }
            el.dispatchEvent(new Event('blur', {bubbles: true}));
            return {ok: true, value: el.value || '', seq: seq ? (seq.value || '') : '', cleared: false};
            """,
            name_selector,
            seq_selector,
            hotel_name,
        )
        if not result or not result.get('ok'):
            reason = result.get('reason') if isinstance(result, dict) else 'unknown'
            if reason == 'name_not_found':
                raise RuntimeError(f"호텔명 필드({name_selector})를 찾지 못했습니다.")
            raise RuntimeError(f"호텔명 필드 설정 실패: {reason}")

        if not hotel_name:
            return result

        deadline = time.time() + timeout
        seq_value = str(result.get('seq') or '').strip()
        actual_name = str(result.get('value') or '').strip()
        while time.time() < deadline and not seq_value:
            time.sleep(poll)
            seq_value = self.driver.execute_script(
                """
                const seq = document.querySelector(arguments[0]);
                return seq ? (seq.value || '') : '';
                """,
                seq_selector,
            )
            seq_value = str(seq_value or '').strip()
            actual_name = self.driver.execute_script(
                """
                const el = document.querySelector(arguments[0]);
                return el ? (el.value || '') : '';
                """,
                name_selector,
            )
            actual_name = str(actual_name or '').strip()

        if not seq_value:
            raise RuntimeError(
                f"호텔명 '{hotel_name}'이 ERP 호텔키({seq_selector})로 매칭되지 않았습니다. "
                "호텔명을 ERP 자동완성에서 단일 매칭되는 이름으로 입력해 주세요."
            )
        if expected_seq and seq_value != expected_seq:
            raise RuntimeError(
                f"호텔명 '{hotel_name}'의 ERP 매칭값이 선택한 hotelSeq와 다릅니다. "
                f"선택={expected_seq}, ERP자동완성={seq_value}"
            )
        return {'ok': True, 'value': actual_name or hotel_name, 'seq': seq_value}

    def _set_erp_date_input(self, selector, value, label, attempts=4, pause=0.2):
        value = '' if value is None else str(value).strip()

        def read_value():
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, selector)
                return (el.get_attribute('value') or '').strip().replace('.', '-').replace('/', '-')
            except Exception:
                return ''

        last_value = ''
        for _ in range(max(1, attempts)):
            el = self.driver.find_element(By.CSS_SELECTOR, selector)
            result = self.driver.execute_script(
                """
                const el = arguments[0];
                const value = arguments[1];
                if (!el) {
                    return {ok:false, value:''};
                }
                el.value = '';
                el.dispatchEvent(new Event('input', {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                el.value = value;
                el.dispatchEvent(new Event('input', {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                el.dispatchEvent(new Event('blur', {bubbles:true}));
                try {
                    if (window.jQuery) {
                        window.jQuery(el).val(value).trigger('change').trigger('blur');
                    }
                } catch (e) {}
                return {ok:true, value: el.value || ''};
                """,
                el,
                value,
            )
            if not result or not result.get('ok'):
                raise RuntimeError(f"{label} 입력 필드({selector})를 설정하지 못했습니다.")
            time.sleep(pause)
            last_value = read_value()
            if last_value == value:
                return last_value
        raise RuntimeError(f"{label}이 ERP에 '{last_value}'(으)로 남아 기대값({value})과 다릅니다.")

    @staticmethod
    def _progress_status_from_text(text):
        normalized = re.sub(r"\s+", "", str(text or ""))
        if "예약마감" in normalized:
            return "05", "예약마감"
        return "", ""

    def _current_page_progress_status_matches(self, progress_code, progress_label):
        """현재 페이지의 진행구분이 이미 목표값이면 중복 정보일괄수정을 피한다."""
        grid_id = self.config.get('grid_id', '#gridMain')
        try:
            rows = self.driver.execute_script(
                """
                try {
                    return (typeof AUIGrid !== 'undefined') ? AUIGrid.getGridData(arguments[0]) : [];
                } catch(e) {
                    return [];
                }
                """,
                grid_id,
            ) or []
        except Exception:
            return False
        if not rows:
            return False
        target_label = re.sub(r"\s+", "", str(progress_label or ""))
        for row in rows:
            code = str(row.get('procCd') or '').strip()
            label = re.sub(
                r"\s+",
                "",
                str(row.get('procNm') or row.get('procDesc') or row.get('procCdNm') or ''),
            )
            if code == progress_code:
                continue
            if target_label and target_label in label:
                continue
            return False
        return True

    def _current_page_progress_status_counts(self):
        grid_id = self.config.get('grid_id', '#gridMain')
        try:
            rows = self.driver.execute_script(
                """
                try {
                    return (typeof AUIGrid !== 'undefined') ? AUIGrid.getGridData(arguments[0]) : [];
                } catch(e) {
                    return [];
                }
                """,
                grid_id,
            ) or []
        except Exception:
            rows = []

        counts = {}
        reservation_closed = 0
        for row in rows:
            code = str(row.get('procCd') or '').strip()
            label = re.sub(
                r"\s+",
                "",
                str(row.get('procNm') or row.get('procDesc') or row.get('procCdNm') or ''),
            )
            key = f"{code}|{label or '-'}"
            counts[key] = counts.get(key, 0) + 1
            if code == "05" or "예약마감" in label:
                reservation_closed += 1
        return {
            "total": len(rows),
            "reservation_closed": reservation_closed,
            "counts": counts,
        }

    @staticmethod
    def _should_skip_price_update_for_all_closed(progress_counts):
        total_count = int((progress_counts or {}).get("total") or 0)
        reservation_closed_count = int((progress_counts or {}).get("reservation_closed") or 0)
        return total_count > 0 and reservation_closed_count >= total_count

    def _apply_current_page_progress_status(self, selectors, progress_code, progress_label,
                                            driver_timeout, erp_short_pause, erp_poll_interval):
        """현재 조회/페이지의 선택 행에 정보일괄수정 진행구분을 적용한다."""
        header_chk = self.driver.find_element(By.CSS_SELECTOR, selectors["header_all_checkbox"])
        if not header_chk.is_selected():
            self.driver.execute_script("arguments[0].click();", header_chk)
            try:
                WebDriverWait(self.driver, 2, poll_frequency=erp_poll_interval).until(
                    lambda d: d.find_element(By.CSS_SELECTOR, selectors["header_all_checkbox"]).is_selected()
                )
            except Exception:
                time.sleep(erp_short_pause)

        event_modify_selector = selectors.get("event_modify_button", "#eventModify")
        event_modify_btn = self.driver.find_element(By.CSS_SELECTOR, event_modify_selector)
        self.driver.execute_script("arguments[0].click();", event_modify_btn)

        proc_select_selector = selectors.get("progress_status_select", "#procCd")
        proc_chk_selector = selectors.get("progress_status_checkbox", "#procCdChk")
        save_selector = selectors.get("bulk_update_save_button", "#appSave")
        wait = WebDriverWait(self.driver, driver_timeout, poll_frequency=erp_poll_interval)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, proc_select_selector)))
        wait.until(
            lambda d: d.execute_script(
                """
                const sel = document.querySelector(arguments[0]);
                return !!(sel && Array.from(sel.options || []).some(opt => opt.value === arguments[1]));
                """,
                proc_select_selector,
                progress_code,
            )
        )

        result = self.driver.execute_script(
            """
            const chk = document.querySelector(arguments[0]);
            const sel = document.querySelector(arguments[1]);
            const code = arguments[2];
            if (!chk || !sel) {
                return {ok:false, reason:'missing_field'};
            }
            chk.checked = true;
            if (window.jQuery) {
                window.jQuery(chk).prop('checked', true).trigger('change');
            } else {
                chk.dispatchEvent(new Event('change', {bubbles:true}));
            }
            sel.value = code;
            sel.dispatchEvent(new Event('change', {bubbles:true}));
            const selected = sel.options && sel.selectedIndex >= 0 ? sel.options[sel.selectedIndex] : null;
            return {ok:true, checked: chk.checked, value: sel.value || '', text: selected ? (selected.textContent || '').trim() : ''};
            """,
            proc_chk_selector,
            proc_select_selector,
            progress_code,
        )
        if not result or not result.get("ok") or not result.get("checked") or result.get("value") != progress_code:
            raise RuntimeError(f"진행구분 필드를 {progress_label}({progress_code})로 설정하지 못했습니다.")

        save_btn = self.driver.find_element(By.CSS_SELECTOR, save_selector)
        self.driver.execute_script("arguments[0].click();", save_btn)

        error_keywords = ('408', '지연', '연결이 원활', '관리자에게 문의', '오류가', '실패')
        save_alert_error = None
        for _ in range(4):
            try:
                alert = self.driver.switch_to.alert
                atext = alert.text or ''
                print(f" -> [진행구분 얼럿 감지]: {atext}")
                if any(k in atext for k in error_keywords):
                    save_alert_error = atext
                    alert.accept()
                    break
                alert.accept()
                time.sleep(erp_short_pause)
            except Exception:
                break
        if save_alert_error:
            raise RuntimeError(f"진행구분 저장 오류 가능: {save_alert_error[:120]}")

        wait.until(lambda d: not d.find_elements(By.CSS_SELECTOR, proc_select_selector))
        self.wait_until_grid_ready_after_save(selectors, driver_timeout)

    def find_and_switch_frame(self, selector):
        by, search_val = self._selector_locator(selector)

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
        browser_config = self._debug_browser_config(target_name)
        self.chrome_launch_btn.config(state=tk.DISABLED, text='여는 중…')
        if target_name == 'TOPAS':
            self._set_topas_status(f"TOPAS 브라우저를 여는 중입니다… ({browser_config['address']})", self.accent_orange)
        else:
            self.set_status(f"ERP 브라우저를 여는 중입니다… ({browser_config['address']})", self.accent_orange)
        self.root.update_idletasks()

        def _restore():
            self.chrome_launch_btn.config(state=tk.NORMAL, text='브라우저 켜기')
            if target_name == 'TOPAS':
                self._set_topas_status(
                    f"TOPAS 브라우저 준비 완료 ({browser_config['address']}) · 로그인 후 첫 조회를 실행하세요",
                    self.accent_green,
                )
            else:
                self.set_status(
                    f"ERP 브라우저 준비 완료 ({browser_config['address']}) · 로그인 후 시작하세요",
                    self.accent_green,
                )

        try:
            subprocess.Popen(
                [bat_path, target_url, str(browser_config['port']), browser_config['profile_dir']],
                creationflags=subprocess.CREATE_NEW_CONSOLE,
            )
            print(
                f"[브라우저 기동] 디버깅용 크롬 브라우저 기동 명령을 전달했습니다. "
                f"({target_name}, {browser_config['address']}, {browser_config['profile_dir']})"
            )
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
            erp_candidates = ', '.join(config['address'] for config in self._debug_browser_configs('ERP'))
            print(f"[RPA 시작] ERP 디버그 브라우저 후보: {erp_candidates}")

            driver_timeout = int(self.config.get("timeout", 25))
            erp_poll_interval = self._float_config('erp_poll_interval', 0.25, 0.1, 1.0)
            erp_short_pause = self._float_config('erp_short_pause', 0.2, 0.05, 0.5)

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
                self.root.after(0, self.hide_loading)
                self.root.after(0, lambda msg=err_text: messagebox.showerror("설정 오류", msg))
                return

            try:
                # Selenium Manager가 로컬 크롬 버전에 맞춰 드라이버를 자동 제어하도록 서비스 오버헤드 제거
                self.driver, browser_config = self._connect_matching_debug_browser('ERP', selectors)
                print(f"[RPA 시작] 선택된 ERP 디버그 브라우저: {browser_config['address']}")
            except Exception as connect_ex:
                err_text = str(connect_ex)
                print(f"[오류] {err_text}")
                self.root.after(0, self.hide_loading)
                self.root.after(
                    0,
                    lambda msg=err_text: messagebox.showerror(
                        "브라우저 선택 실패",
                        f"{msg}\n\n"
                        "요금수정 화면이 디버그 Chrome 안에 열려 있는지 확인해 주세요.",
                    ),
                )
                return

            self.root.after(0, self.hide_loading)

            total_items = len(self.fares_data)

            print(f"[RPA 정보] 총 {total_items}개의 날짜 데이터 처리를 시작합니다.")
            last_job_label = None
            if not self.find_and_switch_frame(selectors["search_date_input"]):
                err_text = "출발일자 입력 필드를 찾을 수 없어 ERP 조회 조건을 설정하지 못했습니다."
                print(f"[오류] {err_text}")
                self.root.after(0, lambda msg=err_text: messagebox.showerror("ERP 화면 오류", msg))
                return

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
                job_index = row.get("_job_index")
                job_label = str(row.get("_job_label") or "").strip()
                history_date_str = f"{job_label} / {date_log_str}" if job_label else date_log_str
                if job_label and job_label != last_job_label:
                    print(f"\n################ 작업 시작: {job_label} ################")
                    source = str(row.get("_job_source") or "").strip()
                    if source:
                        print(f" -> 출처: {source}")
                    last_job_label = job_label
                self._set_job_progress_ui(job_index, status='진행 중')

                adult_air = str(row.get("adult_air", "")).strip()
                adult_hotel = str(row.get("adult_hotel", "")).strip()
                adult_land = str(row.get("adult_land", "")).strip()
                adult_tour = str(row.get("adult_tour", "")).strip()
                adult_profit = str(row.get("adult_profit", "")).strip()
                child_val = str(row.get("child_fare", "")).strip()
                infant_val = str(row.get("infant_fare", "")).strip()
                row_conditions = self._rpa_row_conditions(row)
                airline_code = row_conditions['airline_code']
                price_desc = row_conditions['price_desc']
                hotel_name = row_conditions['hotel_name']
                hotel_seq = row_conditions['hotel_seq']
                progress_text = row_conditions['progress_text']
                progress_code, progress_label = self._progress_status_from_text(progress_text)

                print(f"\n========================================================")
                print(f"[{index+1}/{total_items}] 대상 날짜: {date_log_str}")
                print(f" -> 입력 데이터: 성인(항공={adult_air}, 호텔={adult_hotel}, 지상={adult_land}, 경비={adult_tour}, 수익={adult_profit}), 소아={child_val}, 유아={infant_val}")
                if progress_code:
                    print(f" -> 진행구분 변경 예약: {progress_label}({progress_code})")

                if not self.find_and_switch_frame(selectors["search_date_input"]):
                    err_msg = "출발일자 입력 필드를 찾을 수 없습니다."
                    print(f" -> [오류] {err_msg}")
                    rpa_history.append({"date": history_date_str, "status": "FAIL", "error": err_msg})
                    self._set_job_progress_ui(job_index, result_status='FAIL')
                    self.update_progress_ui(index + 1, total_items)
                    continue

                try:
                    start_value = self._set_erp_date_input(
                        selectors["search_date_input"],
                        date_val,
                        "시작일",
                        pause=erp_short_pause,
                    )
                    end_value = self._set_erp_date_input(
                        selectors["search_date_end_input"],
                        date_end_val,
                        "종료일",
                        pause=erp_short_pause,
                    )
                    print(f" -> 날짜 필터 설정: {start_value} ~ {end_value}")

                    airline_result = self._set_erp_airline_filter(selectors, airline_code)
                    if airline_code:
                        airline_text = airline_result.get('text') or airline_code
                        print(f" -> 항공사 필터 설정: {airline_text}")
                    elif index == 0:
                        print(" -> 항공사 필터 초기화: 전체 항공사 대상")

                    price_desc_result = self._set_erp_text_filter(
                        selectors,
                        'price_desc_input',
                        '#priceDesc',
                        price_desc,
                        '요금구분',
                    )
                    if price_desc:
                        print(f" -> 요금구분 필터 설정: {price_desc_result.get('value') or price_desc}")
                    elif index == 0:
                        print(" -> 요금구분 필터 초기화: 전체 요금구분 대상")

                    hotel_result = self._set_erp_hotel_filter(
                        selectors,
                        hotel_name,
                        expected_seq=hotel_seq,
                        timeout=min(driver_timeout, 4),
                        poll=erp_poll_interval,
                    )
                    if hotel_name:
                        print(f" -> 호텔명 필터 설정: {hotel_result.get('value') or hotel_name} (hotelSeq={hotel_result.get('seq')})")
                    elif index == 0:
                        print(" -> 호텔명 필터 초기화: 전체 호텔 대상")

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
                        rpa_history.append({"date": history_date_str, "status": "SKIP", "error": "조회결과 없음"})
                        self._set_job_progress_ui(job_index, result_status='SKIP')
                        self.update_progress_ui(index + 1, total_items)
                        continue

                    if hotel_name:
                        expected_hotel_seq = str(hotel_result.get('seq') or '').strip()
                        hotel_summary = self.driver.execute_script(
                            """
                            try {
                                const rows = (typeof AUIGrid !== 'undefined') ? (AUIGrid.getGridData(arguments[0]) || []) : [];
                                const expected = String(arguments[1] || '').trim();
                                let mismatches = [];
                                for (let i = 0; i < rows.length; i++) {
                                    const seq = String(rows[i].hotelSeq == null ? '' : rows[i].hotelSeq).trim();
                                    if (expected && seq !== expected) {
                                        mismatches.push({row: i + 1, hotelSeq: seq, hotelKorNm: rows[i].hotelKorNm || ''});
                                    }
                                }
                                return {total: rows.length, expected, mismatchCount: mismatches.length, mismatches: mismatches.slice(0, 3)};
                            } catch(e) {
                                return {total: 0, expected: String(arguments[1] || '').trim(), mismatchCount: -1, error: String(e)};
                            }
                            """,
                            grid_id,
                            expected_hotel_seq,
                        )
                        hotel_summary = hotel_summary or {'mismatchCount': -1, 'error': 'empty verification result'}
                        mismatch_count = int(hotel_summary.get('mismatchCount') or 0)
                        if mismatch_count < 0:
                            raise RuntimeError(f"호텔명 필터 검증 중 오류: {hotel_summary.get('error')}")
                        if mismatch_count > 0:
                            sample = hotel_summary.get('mismatches') or []
                            raise RuntimeError(
                                f"호텔명 필터 검증 실패: 기대 hotelSeq={expected_hotel_seq}, "
                                f"다른 호텔 행 {mismatch_count}건 감지 {sample}"
                            )
                        print(f" -> 호텔명 필터 검증: {hotel_summary.get('total', 0)}건 hotelSeq={expected_hotel_seq}")

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
                    inputs_mapping = {
                        key: value for key, value in inputs_mapping.items()
                        if str(value).strip() != ""
                    }
                    if not inputs_mapping and not progress_code:
                        print(f" -> [건너뜀] {date_log_str}에 입력할 요금값 또는 예약마감 변경값이 없습니다.")
                        rpa_history.append({"date": history_date_str, "status": "SKIP", "error": "입력/변경값 없음"})
                        self._set_job_progress_ui(job_index, result_status='SKIP')
                        self.update_progress_ui(index + 1, total_items)
                        continue

                    # 페이지별 재시도/페이싱 설정 (서버 일시 지연 408 등 대응)
                    page_max_retries = max(1, int(self.config.get('page_max_retries', 3)))
                    page_pause = self._float_config('erp_page_pause', 1.0, 0.0, 10.0)
                    done_actions = []
                    if inputs_mapping:
                        done_actions.append("요금 업데이트")
                    if progress_code:
                        done_actions.append("진행구분 변경")
                    done_summary = " + ".join(done_actions) if done_actions else "작업"

                    target_page = 1
                    pages_done = 0
                    pages_failed = []
                    pages_price_skipped_all_closed = 0

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
                                    actions = []
                                    if inputs_mapping:
                                        actions.append("요금 입력")
                                    if progress_code:
                                        actions.append(f"진행구분 {progress_label}")
                                    print(f" -> [페이지 {target_page}/{total_pages}] 전체선택 → {' + '.join(actions)} → 저장{suffix}")

                                if inputs_mapping:
                                    progress_counts = self._current_page_progress_status_counts()
                                    reservation_closed_count = int(progress_counts.get("reservation_closed") or 0)
                                    total_count = int(progress_counts.get("total") or 0)
                                    if self._should_skip_price_update_for_all_closed(progress_counts):
                                        pages_price_skipped_all_closed += 1
                                        print(
                                            " -> [요금 업데이트 스킵] 현재 페이지의 모든 행이 예약마감입니다. "
                                            "ERP가 수정 가능한 요금 선택을 허용하지 않아 이 페이지의 요금 업데이트를 건너뜁니다."
                                        )
                                    else:
                                        if reservation_closed_count:
                                            print(
                                                f" -> [요금 업데이트 안내] 현재 페이지 {total_count}건 중 "
                                                f"예약마감 {reservation_closed_count}건은 ERP 정책상 제외되고, "
                                                "수정 가능한 행만 업데이트됩니다."
                                            )
                                        self._save_current_page(
                                            selectors, wait, inputs_mapping, driver_timeout,
                                            erp_short_pause, erp_poll_interval
                                        )
                                if progress_code:
                                    if inputs_mapping:
                                        self.navigate_to_grid_page(selectors, target_page, driver_timeout)
                                    if self._current_page_progress_status_matches(progress_code, progress_label):
                                        print(f" -> 진행구분 확인: 현재 페이지가 이미 {progress_label} 상태입니다.")
                                    else:
                                        print(f" -> 진행구분 정보일괄수정: {progress_label}")
                                        self._apply_current_page_progress_status(
                                            selectors, progress_code, progress_label,
                                            driver_timeout, erp_short_pause, erp_poll_interval
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
                        rpa_history.append({"date": history_date_str, "status": "FAIL", "error": msg})
                        self._set_job_progress_ui(job_index, result_status='FAIL')
                    else:
                        skip_note = ""
                        if pages_price_skipped_all_closed:
                            skip_note = f" (전체 예약마감 페이지 {pages_price_skipped_all_closed}개 요금 업데이트 스킵)"
                        if total_pages > 1:
                            print(f" -> [성공] {date_log_str} {done_summary} 완료 (전체 {pages_done}/{total_pages}페이지){skip_note}")
                        else:
                            print(f" -> [성공] {date_log_str} {done_summary} 완료{skip_note}")
                        if inputs_mapping and not progress_code and pages_price_skipped_all_closed == total_pages:
                            rpa_history.append({"date": history_date_str, "status": "SKIP", "error": "전체 예약마감으로 요금 업데이트 대상 없음"})
                            self._set_job_progress_ui(job_index, result_status='SKIP')
                        else:
                            rpa_history.append({"date": history_date_str, "status": "SUCCESS", "error": ""})
                            self._set_job_progress_ui(job_index, result_status='SUCCESS')

                except Exception as row_ex:
                    err_msg = str(row_ex).replace("\n", " ")
                    print(f" -> [실패] 오류 발생: {err_msg}")
                    rpa_history.append({"date": history_date_str, "status": "FAIL", "error": err_msg})
                    self._set_job_progress_ui(job_index, result_status='FAIL')

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
        mutex_name = "Global\\NaeilERPUpdaterV5_SingleInstance_Mutex"
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
            messagebox.showinfo("이미 실행 중", "NaeilERPUpdater V5가 이미 실행 중입니다.\n열려 있는 창을 확인해 주세요.")
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
