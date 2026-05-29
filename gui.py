# -*- coding: utf-8 -*-
import os
import sys
import json
import time
import csv
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
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import selenium.webdriver.chrome.options
import selenium.webdriver.chrome.service
import selenium.webdriver.chrome.webdriver
from webdriver_manager.chrome import ChromeDriverManager
import excel_loader
import update_client

APP_VERSION = "v1.1.7"
UPDATER_EXE_NAME = "NaeilERPUpdaterUpdater.exe"

def get_app_dir():
    """실행 파일 또는 소스 파일 옆의 쓰기 가능한 폴더 경로를 반환합니다."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(rel_path):
    """번들된 리소스(이미지/GIF 등)의 실제 경로를 반환한다.
    PyInstaller로 묶인 경우 _MEIPASS를 우선 확인하고, 없으면 앱 폴더에서 찾는다."""
    base = getattr(sys, '_MEIPASS', None)
    if base:
        cand = os.path.join(base, rel_path)
        if os.path.exists(cand):
            return cand
    return os.path.join(get_app_dir(), rel_path)

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
    def __init__(self, root):
        self.root = root
        self.root.title(f'Naeil Tour ERP 요금 업데이트 RPA ({APP_VERSION})')
        self.root.geometry('820x660')
        self.root.resizable(False, False)

        # --- 색상 팔레트 (다크 테마 정돈) ---
        self.bg_color = '#15161c'        # 전체 배경
        self.card_color = '#21232c'      # 카드/패널 배경
        self.card_hover = '#2a2d38'
        self.fg_color = '#f3f4f6'        # 기본 글자
        self.fg_muted = '#969aa6'        # 보조 글자
        self.accent_color = '#3b82f6'    # 메인 블루
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
        self.excel_path = ''
        self.fares_data = []
        self.rpa_thread = None
        self.is_running = False
        self.is_paused = False
        self.is_user_stopped = False
        self.driver = None
        self.console_redirector = None
        
        # 날짜 필터 관리 변수
        self.filter_mode = tk.StringVar(value="ALL") # ALL, FROM_DATE, SPECIFIC
        self.filter_value = tk.StringVar(value="")

        # 로딩 애니메이션 상태
        self._loading_frames = []
        self._loading_idx = 0
        self._loading_after_id = None
        self._loading_overlay = None

        self.build_ui()
        self.sync_config_to_ui()
        self._load_loading_frames()
        self.root.after(800, self.check_for_updates_on_startup)

    def load_config(self):
        base_dir = get_app_dir()
        config_path = os.path.join(base_dir, 'config.json')
        default_config = {
            'headless': False,
            'dry_run': False, # 항상 LIVE
            'timeout': 25,
            'login_timeout': 180,
            'use_existing_browser': True,
            'debugger_address': '127.0.0.1:9222',
            'erp_url': 'http://erp.naeiltour.co.kr',
            'history_log_path': 'logs/update_history.csv',
            'screenshot_dir': 'logs/screenshots',
            'update_enabled': True,
            'update_latest_url': '',
            'update_check_timeout': 8,
            'update_download_timeout': 30,
            'selectors': {
                'login_id': "input[name='userId']",
                'login_pw': "input[name='userPw']",
                'login_btn': '#btnLogin',
                'search_date_input': '#searchStDate',
                'search_date_end_input': '#searchEnDate',
                'search_button': '#gridMain_r',
                'header_all_checkbox': 'td.aui-grid-row-check-header input',
                'update_button': '#priceUpdate',
                'adult_fare_input': '#addAir01',
                'child_fare_input': '#addProfit41',
                'save_button': '#priceSave',
                'cancel_button': '#popCloseBtn',
                'date_cell_in_row': '.aui-grid-default-column'
            }
        }
        
        if not os.path.exists(config_path):
            try:
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(default_config, f, indent=2, ensure_ascii=False)
            except Exception as e:
                print(f"[경고] 설정 파일 생성 실패: {str(e)}")
            return default_config
        else:
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_conf = json.load(f)
                    # 기본값과 머지
                    for k, v in default_config.items():
                        if k not in user_conf:
                            user_conf[k] = v
                    return user_conf
            except Exception as e:
                print(f"[경고] 설정 로드 실패, 기본값 사용: {str(e)}")
                return default_config

    def save_config(self):
        base_dir = get_app_dir()
        config_path = os.path.join(base_dir, 'config.json')
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"[경고] 설정 저장 실패: {str(e)}")

    def sync_config_to_ui(self):
        self.config['dry_run'] = False
        self.save_config()

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

    def _setup_styles(self):
        """ttk 위젯(진행바 등) 다크 테마 스타일 구성."""
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

    def _add_hover(self, widget, normal_bg, hover_bg):
        """버튼에 마우스 오버 색상 전환 효과를 부여한다."""
        def on_enter(_e):
            if str(widget['state']) != 'disabled':
                widget.config(bg=hover_bg)
        def on_leave(_e):
            if str(widget['state']) != 'disabled':
                widget.config(bg=normal_bg)
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)

    def set_status(self, text, color=None):
        """상태 표시줄(점 + 텍스트)을 한 번에 갱신한다."""
        color = color or self.fg_color
        if hasattr(self, 'status_dot'):
            self.status_dot.config(fg=color)
        self.status_lbl.config(text=text, fg=color)

    def _load_loading_frames(self, size=140):
        """캐릭터 로딩 GIF의 프레임을 미리 PhotoImage로 변환해 둔다.
        Pillow/리소스가 없으면 조용히 비활성화한다."""
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
            # 애니메이션은 부가 기능이므로 실패해도 앱은 정상 동작
            print(f"[안내] 로딩 애니메이션 비활성화: {e}")

    def show_loading(self, text='잠시만 기다려 주세요…'):
        """화면 중앙에 캐릭터 로딩 오버레이를 띄운다."""
        if not self._loading_frames:
            # GIF가 없으면 상태 텍스트만 갱신
            self.set_status(text, self.accent_orange)
            return
        if self._loading_overlay is not None:
            # 이미 떠 있으면 문구만 갱신
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
        """로딩 오버레이를 닫는다."""
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

    def build_ui(self):
        # 1. 헤더 영역
        header_frame = tk.Frame(self.root, bg=self.bg_color, pady=14)
        header_frame.pack(fill=tk.X, padx=24)

        # 우측: ERP 켜기 버튼 (제목 줄 높이에 맞춰 수직 중앙 정렬)
        self.chrome_launch_btn = tk.Button(
            header_frame, text='ERP 켜기', width=14, height=2,
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
        version_badge = tk.Label(title_row, text=f' {APP_VERSION} ', font=('맑은 고딕', 8, 'bold'), bg=self.card_color, fg=self.accent_color)
        version_badge.pack(side=tk.LEFT, padx=(8, 0), pady=(4, 0))

        desc_label = tk.Label(title_frame, text='엑셀 데이터를 기반으로 ERP 항공요금을 자동으로 반영합니다.', font=('맑은 고딕', 9), bg=self.bg_color, fg=self.fg_muted)
        desc_label.pack(anchor=tk.W, pady=(4, 0))
        
        # 2. 파일 및 필터 카드
        settings_frame = tk.LabelFrame(self.root, text=' 작업 설정 ', font=('맑은 고딕', 9, 'bold'), bg=self.card_color, fg=self.fg_muted, bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightthickness=1, padx=10, pady=8)
        settings_frame.pack(fill=tk.X, padx=24, pady=(6, 12))

        # 엑셀 경로 선택 행
        excel_lbl = tk.Label(settings_frame, text='엑셀 파일 경로', font=('맑은 고딕', 9), bg=self.card_color, fg=self.fg_color)
        excel_lbl.grid(row=0, column=0, padx=10, pady=7, sticky=tk.W)

        self.excel_entry = tk.Entry(settings_frame, width=58, bg=self.bg_color, fg=self.fg_color, insertbackground='white', bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        self.excel_entry.grid(row=0, column=1, padx=5, pady=7, ipady=4, sticky=tk.W)

        excel_btn = tk.Button(settings_frame, text='찾기...', width=8, bg=self.accent_color, fg='white', font=('맑은 고딕', 9, 'bold'), activebackground=self.accent_hover, activeforeground='white', bd=0, relief=tk.FLAT, cursor='hand2', command=self.browse_excel)
        excel_btn.grid(row=0, column=2, padx=10, pady=7, ipady=2, sticky=tk.W)
        self._add_hover(excel_btn, self.accent_color, self.accent_hover)
        
        # 날짜 필터 선택 행
        filter_lbl = tk.Label(settings_frame, text='날짜 필터 옵션', font=('맑은 고딕', 9), bg=self.card_color, fg=self.fg_color)
        filter_lbl.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        
        filter_modes_frame = tk.Frame(settings_frame, bg=self.card_color)
        filter_modes_frame.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        rad_all = tk.Radiobutton(filter_modes_frame, text="전체 대상", variable=self.filter_mode, value="ALL", bg=self.card_color, fg=self.fg_color, selectcolor=self.card_color, activebackground=self.card_color, activeforeground=self.fg_color)
        rad_all.pack(side=tk.LEFT, padx=(0, 10))
        
        rad_from = tk.Radiobutton(filter_modes_frame, text="특정일 이후", variable=self.filter_mode, value="FROM_DATE", bg=self.card_color, fg=self.fg_color, selectcolor=self.card_color, activebackground=self.card_color, activeforeground=self.fg_color)
        rad_from.pack(side=tk.LEFT, padx=10)
        
        rad_specific = tk.Radiobutton(filter_modes_frame, text="특정 날짜 지정", variable=self.filter_mode, value="SPECIFIC", bg=self.card_color, fg=self.fg_color, selectcolor=self.card_color, activebackground=self.card_color, activeforeground=self.fg_color)
        rad_specific.pack(side=tk.LEFT, padx=10)
        
        # 필터 값 입력 행
        filter_val_lbl = tk.Label(settings_frame, text='필터 값 입력', font=('맑은 고딕', 9), bg=self.card_color, fg=self.fg_color)
        filter_val_lbl.grid(row=2, column=0, padx=10, pady=(5, 10), sticky=tk.W)

        self.filter_val_entry = tk.Entry(settings_frame, textvariable=self.filter_value, width=40, bg=self.bg_color, fg=self.fg_color, insertbackground='white', bd=0, relief=tk.FLAT, highlightbackground=self.border_color, highlightcolor=self.accent_color, highlightthickness=1)
        self.filter_val_entry.grid(row=2, column=1, padx=5, pady=(5, 10), ipady=4, sticky=tk.W)
        
        filter_tip_lbl = tk.Label(settings_frame, text='(예: 2026-06-07 또는 20260607, 20260608)', font=('맑은 고딕', 8), bg=self.card_color, fg=self.fg_muted)
        filter_tip_lbl.grid(row=2, column=1, columnspan=2, padx=(300, 0), pady=(5, 10), sticky=tk.W)

        # 3. 진행 상태 바 영역
        progress_frame = tk.Frame(self.root, bg=self.bg_color)
        progress_frame.pack(fill=tk.X, padx=24, pady=(2, 4))

        self.status_dot = tk.Label(progress_frame, text='●', font=('맑은 고딕', 11), bg=self.bg_color, fg=self.fg_muted)
        self.status_dot.pack(side=tk.LEFT, padx=(0, 6))
        self.status_lbl = tk.Label(progress_frame, text='대기 중', font=('맑은 고딕', 9, 'bold'), bg=self.bg_color, fg=self.fg_color)
        self.status_lbl.pack(side=tk.LEFT)

        self.progress_lbl = tk.Label(progress_frame, text='0 / 0 (0%)', font=('맑은 고딕', 9), bg=self.bg_color, fg=self.fg_muted)
        self.progress_lbl.pack(side=tk.RIGHT)

        self.progress_bar = ttk.Progressbar(self.root, orient='horizontal', mode='determinate', style='Accent.Horizontal.TProgressbar')
        self.progress_bar.pack(fill=tk.X, padx=24, pady=(0, 12))

        # 4. 로그 영역
        log_title = tk.Label(self.root, text='실시간 작업 내용', font=('맑은 고딕', 9, 'bold'), bg=self.bg_color, fg=self.fg_muted)
        log_title.pack(anchor=tk.W, padx=24, pady=(0, 3))

        log_wrap = tk.Frame(self.root, bg=self.border_color, bd=0)
        log_wrap.pack(fill=tk.BOTH, padx=24, pady=(0, 12))
        self.log_txt = ScrolledText(log_wrap, height=13, bg='#0c0d11', fg='#c9cdd6', insertbackground='white', font=('Consolas', 9), bd=0, relief=tk.FLAT, padx=10, pady=8)
        self.log_txt.pack(fill=tk.BOTH, padx=1, pady=1)

        # 5. 컨트롤 버튼 영역
        control_frame = tk.Frame(self.root, bg=self.bg_color)
        control_frame.pack(fill=tk.X, padx=24, pady=(2, 12))

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

    def browse_excel(self):
        file_path = excel_loader.select_excel_file()
        if file_path:
            self.excel_path = file_path
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)
            self.load_excel_data()

    def load_excel_data(self):
        self.excel_path = self.excel_entry.get().strip()
        if not self.excel_path:
            return
        
        try:
            # 엑셀 데이터 원본 로드
            raw_data = excel_loader.load_and_validate_fares(self.excel_path)
            if not raw_data:
                self.set_status('처리할 유효한 요금 행이 없습니다.', self.accent_red)
                self.fares_data = []
                return

            # 필터 옵션 적용
            mode = self.filter_mode.get()
            val = self.filter_value.get().strip()
            self.fares_data = excel_loader.filter_fares_by_date(raw_data, mode, val)

            if self.fares_data:
                total_cnt = len(self.fares_data)
                min_date = self.fares_data[0]['date']
                max_date = self.fares_data[-1]['date']
                self.set_status(f'엑셀 로드 성공 · 총 {total_cnt}건 ({min_date} ~ {max_date})', self.accent_green)
                self.progress_lbl.config(text=f'0 / {total_cnt} (0%)')
                self.progress_bar['maximum'] = total_cnt
                self.progress_bar['value'] = 0
            else:
                self.set_status('필터 결과 대상 데이터가 존재하지 않습니다.', self.accent_red)
        except Exception as e:
            self.set_status(f'엑셀 파싱 실패 - {str(e)}', self.accent_red)
            self.fares_data = []

    def start_rpa(self):
        self.excel_path = self.excel_entry.get().strip()
        if not self.excel_path:
            messagebox.showwarning('파일 누락', '요금 업데이트용 엑셀 파일을 먼저 선택해 주세요.')
            return
            
        self.load_excel_data()
        if not self.fares_data:
            messagebox.showwarning('데이터 누락', '파싱 및 로드된 대상 요금 데이터가 없습니다. 필터를 확인해 주세요.')
            return
            
        self.is_running = True
        self.is_paused = False
        self.is_user_stopped = False
        
        # 버튼 UI 잠금
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL, text='‖  일시 중지', bg=self.accent_orange)
        self.stop_btn.config(state=tk.NORMAL)
        self.excel_entry.config(state=tk.DISABLED)
        self.filter_val_entry.config(state=tk.DISABLED)
        
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
        self.excel_entry.config(state=tk.NORMAL)
        self.filter_val_entry.config(state=tk.NORMAL)
        self.progress_bar['value'] = 0
        self.progress_lbl.config(text=f'0 / {len(self.fares_data)} (0%)')
        self.is_running = False
        self.is_paused = False

    def finish_rpa_on_ui(self, history):
        try:
            self.generate_and_show_report(history)
        finally:
            self.clean_up_ui_after_rpa()

    def find_and_switch_frame(self, selector):
        by = By.XPATH if selector.startswith('//') or selector.startswith('xpath=') else By.CSS_SELECTOR
        search_val = selector.replace('xpath=', '') if selector.startswith('xpath=') else selector
        
        # 1. Default Content 확인
        try:
            self.driver.switch_to.default_content()
            if len(self.driver.find_elements(by, search_val)) > 0:
                return True
        except Exception:
            pass

        # 2. iframe들 1계층 순회 탐색
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

    def save_screenshot(self, driver, folder, prefix, date_str):
        # 사용자 룰에 의해 screenshot 저장기능 비활성화(no-op)
        pass

    def write_log(self, path, date_str, adult, child, status, err=''):
        # 사용자 룰에 의해 update_history.csv 이력 누적 비활성화(no-op)
        pass

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
        timeout = max(1, int(timeout))
        warmup_seconds = min(2, timeout)
        unlocked_ticks = 0
        saw_lock = False

        self.driver.switch_to.default_content()
        self.find_and_switch_frame(selectors["search_date_input"])

        print(f" -> 저장 후 그리드 활성화 대기 중(스마트 폴링, 최대 {timeout}초)...")

        for elapsed in range(1, timeout + 1):
            if not self.is_running:
                return False

            time.sleep(1.0)

            try:
                is_locked = self.is_grid_locked()
            except Exception as poll_ex:
                print(f" -> [대기 확인 경고] 그리드 차단막 확인 중 예외 발생: {str(poll_ex)}")
                is_locked = False

            if is_locked:
                saw_lock = True
                unlocked_ticks = 0
                print(f" -> [{elapsed}/{timeout}초] ERP 처리 중: 그리드 차단막 감지")
                continue

            unlocked_ticks += 1

            if saw_lock:
                print(f" -> [안정화 완료] {elapsed}초 만에 그리드 차단막이 해제되었습니다.")
                return True

            if unlocked_ticks >= warmup_seconds:
                print(f" -> [안정화 완료] {elapsed}초 동안 차단막 없이 그리드가 활성 상태입니다.")
                return True

            print(f" -> [{elapsed}/{timeout}초] 차단막 없음: 안정 상태 재확인 중")

        print(f" -> [주의] 그리드 차단막 대기 {timeout}초 초과. 현재 상태로 다음 날짜를 진행합니다.")
        return False

    def launch_debug_chrome(self):
        base_dir = get_app_dir()
        bat_path = os.path.join(base_dir, 'chrome_debug.bat')
        if not os.path.exists(bat_path):
            messagebox.showerror('파일 없음', 'chrome_debug.bat 파일을 찾을 수 없습니다.')
            return

        # 즉시 시각 피드백: 버튼/상태로 "여는 중"을 바로 표시 (중복 클릭 방지)
        self.chrome_launch_btn.config(state=tk.DISABLED, text='여는 중…')
        self.set_status('ERP 브라우저를 여는 중입니다…', self.accent_orange)
        self.root.update_idletasks()

        def _restore():
            self.chrome_launch_btn.config(state=tk.NORMAL, text='ERP 켜기')
            self.set_status('ERP 브라우저 준비 완료 · 로그인 후 시작하세요', self.accent_green)

        try:
            import subprocess
            subprocess.Popen([bat_path], creationflags=subprocess.CREATE_NEW_CONSOLE)
            print('[브라우저 기동] 디버깅용 크롬 브라우저 기동 명령을 전달했습니다.')
            # 크롬이 뜰 시간을 준 뒤 버튼 복구
            self.root.after(3000, _restore)
        except Exception as e:
            print(f'[오류] 크롬 기동 실패: {str(e)}')
            self.chrome_launch_btn.config(state=tk.NORMAL, text='ERP 켜기')
            self.set_status('크롬 기동 실패', self.accent_red)
            messagebox.showerror('기동 오류', f'크롬 기동에 실패했습니다:\n{str(e)}')

    def rpa_worker_loop(self):
        # 결과 보고서 작성을 위해 상세 이력 리스트 생성
        rpa_history = []
        
        try:
            # 연결 단계 로딩 애니메이션 표시 (캐릭터)
            self.root.after(0, lambda: self.show_loading('ERP에 연결하고 있어요…'))
            print("[RPA 시작] Selenium 웹드라이버 연결 중...")
            options = webdriver.ChromeOptions()
            debug_addr = self.config.get("debugger_address", self.config.get("debuggerAddress", "127.0.0.1:9222"))
            options.add_experimental_option("debuggerAddress", debug_addr)
            
            # 셀레늄 타임아웃
            driver_timeout = int(self.config.get("timeout", 25))
            
            try:
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=options)
            except Exception as connect_ex:
                print(f"[오류] 브라우저 연결 실패: {str(connect_ex)}")
                print(" -> 크롬 브라우저가 디버깅 모드로 켜져 있는지 확인하고 모든 일반 크롬 창을 닫아주세요.")
                self.root.after(0, lambda: messagebox.showerror("연결 실패", f"디버그 브라우저 연결에 실패했습니다.\n{str(connect_ex)}"))
                return
            
            # 연결 완료 → 로딩 애니메이션 닫기
            self.root.after(0, self.hide_loading)

            selectors = self.config["selectors"]
            total_items = len(self.fares_data)

            print(f"[RPA 정보] 총 {total_items}개의 날짜 데이터 처리를 시작합니다.")
            
            for index, row in enumerate(self.fares_data):
                # 중단 요청 처리
                if not self.is_running:
                    print("[RPA 알림] 사용자가 작업을 강제 중단했습니다.")
                    break
                    
                # 일시 중지 대기 루프
                while self.is_paused and self.is_running:
                    time.sleep(0.5)
                    
                if not self.is_running:
                    break
                    
                date_val = str(row["date"]).strip()
                adult_val = str(row.get("adult_fare", row.get("adult", ""))).strip()
                child_val = str(row.get("child_fare", row.get("child", ""))).strip()
                
                print(f"\n========================================================")
                print(f"[{index+1}/{total_items}] 대상 날짜: {date_val}")
                print(f" -> 입력 데이터: 항공비={adult_val}, 알선수익={child_val}")
                
                # 1. 출발일자 입력란 찾기
                if not self.find_and_switch_frame(selectors["search_date_input"]):
                    err_msg = "출발일자 입력 필드를 찾을 수 없습니다."
                    print(f" -> [오류] {err_msg}")
                    rpa_history.append({"date": date_val, "status": "FAIL", "error": err_msg})
                    self.update_progress_ui(index + 1, total_items)
                    continue
                
                try:
                    # 2. 날짜 필드 주입
                    st_date_input = self.driver.find_element(By.CSS_SELECTOR, selectors["search_date_input"])
                    en_date_input = self.driver.find_element(By.CSS_SELECTOR, selectors["search_date_end_input"])
                    
                    self.driver.execute_script("arguments[0].value = '';", st_date_input)
                    st_date_input.send_keys(date_val)
                    
                    self.driver.execute_script("arguments[0].value = '';", en_date_input)
                    en_date_input.send_keys(date_val)
                    
                    # 3. 조회 버튼 클릭
                    search_btn = self.driver.find_element(By.CSS_SELECTOR, selectors["search_button"])
                    self.driver.execute_script("arguments[0].click();", search_btn)
                    print(" -> 조회 버튼을 클릭했습니다. 조회 데이터 로딩 대기 중...")
                    
                    # 4~5. 조회 결과 로딩 대기 + 대상 날짜 안정화 매칭 (스마트 폴링)
                    # AUIGrid 조회는 지연(~10초)이 있고, 전환 중에는 이전 날짜 행과
                    # 새 날짜 행이 잠시 공존한다. 이때 전체선택을 누르면 이전 날짜까지
                    # 선택될 수 있으므로, 그리드의 날짜 셀이 '대상 날짜 단일'로
                    # 안정될 때까지 폴링한 뒤에만 매칭 성공으로 본다.
                    wait = WebDriverWait(self.driver, driver_timeout)
                    norm_date = date_val.replace('-', '').replace('.', '').replace('/', '')
                    matched = False
                    deadline = time.time() + driver_timeout

                    while time.time() < deadline:
                        if not self.is_running:
                            break

                        dates_seen = set()
                        for cell in self.driver.find_elements(By.CSS_SELECTOR, selectors["date_cell_in_row"]):
                            try:
                                cell_txt = (cell.text or '').strip()
                            except Exception:
                                continue
                            # YYYY-MM-DD 형태의 날짜 셀만 수집
                            if len(cell_txt) == 10 and cell_txt[4] == '-' and cell_txt[7] == '-':
                                dates_seen.add(cell_txt.replace('-', '').replace('.', '').replace('/', ''))

                        # 그리드가 '대상 날짜'로만 채워졌을 때만 매칭 (전환 중 이전 날짜 잔상 배제)
                        if dates_seen == {norm_date}:
                            matched = True
                            break
                        time.sleep(0.5)

                    if not matched:
                        print(f" -> [조회결과 없음] {date_val} 일자 데이터를 반영하지 못했습니다. ERP에서 직접 날짜 조회를 확인해 주세요.")
                        rpa_history.append({"date": date_val, "status": "SKIP", "error": "조회결과 없음"})
                        self.update_progress_ui(index + 1, total_items)
                        continue
                        
                    # 6. 헤더 전체 체크박스 선택
                    header_chk = self.driver.find_element(By.CSS_SELECTOR, selectors["header_all_checkbox"])
                    if not header_chk.is_selected():
                        self.driver.execute_script("arguments[0].click();", header_chk)
                        time.sleep(0.5)
                        
                    # 7. 요금업데이트 버튼 클릭
                    update_btn = self.driver.find_element(By.CSS_SELECTOR, selectors["update_button"])
                    self.driver.execute_script("arguments[0].click();", update_btn)
                    time.sleep(1.0)
                    
                    # 8. 요금 입력창(모달/팝업 영역) 포커스 이동 및 데이터 기입
                    # 요금 필드가 나타나는지 감지
                    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, selectors["adult_fare_input"])))
                    
                    adult_input = self.driver.find_element(By.CSS_SELECTOR, selectors["adult_fare_input"])
                    child_input = self.driver.find_element(By.CSS_SELECTOR, selectors["child_fare_input"])
                    
                    self.driver.execute_script("arguments[0].value = '';", adult_input)
                    adult_input.send_keys(adult_val)
                    
                    self.driver.execute_script("arguments[0].value = '';", child_input)
                    child_input.send_keys(child_val)
                    time.sleep(0.5)
                    
                    # 9. 저장 클릭 및 얼럿 처리
                    save_btn = self.driver.find_element(By.CSS_SELECTOR, selectors["save_button"])
                    self.driver.execute_script("arguments[0].click();", save_btn)
                    
                    # 10. 얼럿 다중 수락
                    handled = False
                    for _ in range(3):
                        try:
                            alert = self.driver.switch_to.alert
                            print(f" -> [얼럿 감지]: {alert.text}")
                            alert.accept()
                            handled = True
                            time.sleep(0.5)
                        except Exception:
                            break
                            
                    # 저장 모달창 닫기 (만약 닫히지 않고 남아있다면 예외 방지용)
                    try:
                        self.driver.switch_to.default_content()
                        # 프레임 재진입 필요
                        self.find_and_switch_frame(selectors["cancel_button"])
                        cancel_btn = self.driver.find_elements(By.CSS_SELECTOR, selectors["cancel_button"])
                        if cancel_btn and cancel_btn[0].is_displayed():
                            self.driver.execute_script("arguments[0].click();", cancel_btn[0])
                    except Exception:
                        pass
                    
                    # 11. 저장 후 스마트 폴링: 그리드 차단막이 해제될 때까지 최대 설정값만큼 대기
                    self.wait_until_grid_ready_after_save(selectors, driver_timeout)
                        
                    print(f" -> [성공] {date_val} 요금 업데이트 완료")
                    rpa_history.append({"date": date_val, "status": "SUCCESS", "error": ""})
                    
                except Exception as row_ex:
                    err_msg = str(row_ex).replace("\n", " ")
                    print(f" -> [실패] 오류 발생: {err_msg}")
                    rpa_history.append({"date": date_val, "status": "FAIL", "error": err_msg})
                    
                    # 에러 시 혹시 모를 팝업을 닫고 디폴트 프레임 복원
                    try:
                        self.driver.switch_to.alert.accept()
                    except Exception:
                        pass
                    
                # UI 진척도 업데이트
                self.update_progress_ui(index + 1, total_items)
                
            print("\n========================================================")
            print("[RPA 종료] 전체 날짜에 대한 작업이 마무리되었습니다.")
            print("========================================================\n")
            
        except Exception as loop_ex:
            print(f"[치명적 오류] 작업 루프 처리 실패: {str(loop_ex)}")
            
        finally:
            # 12. 결과 보고서 출력 및 UI 복구는 Tkinter 메인 스레드에서만 처리
            history_snapshot = list(rpa_history)
            try:
                self.root.after(0, lambda: self.finish_rpa_on_ui(history_snapshot))
            except Exception:
                pass

    def update_progress_ui(self, current, total):
        pct = int((current / total) * 100)
        self.root.after(0, lambda: self.set_status(f'진행 중 ({pct}%)', self.accent_orange))
        self.root.after(0, lambda: self.progress_lbl.config(text=f'{current} / {total} ({pct}%)'))
        self.root.after(0, lambda: self.progress_bar.config(value=current))

    def generate_and_show_report(self, history):
        if not history:
            return
            
        total_cnt = len(history)
        success_cnt = sum(1 for x in history if x["status"] == "SUCCESS")
        fail_cnt = total_cnt - success_cnt
        
        # 실패/스킵 목록 파싱
        # 날짜순 정렬
        sorted_history = sorted(history, key=lambda x: x["date"])
        
        # 단발성 실패 및 연속 실패 분석
        single_failures = []
        consecutive_failures = [] # list of lists of dicts
        
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
                
        # 리포트 문자열 생성
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
                # 대표 오류 메시지 추출 (가장 많이 발생한 에러명 선택)
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
        
        report_str = "\n".join(report)
        
        # 콘솔 화면(ScrolledText) 출력
        print(report_str)
        
        # 사용자가 수동으로 중단한 경우 결과 팝업창을 띄우지 않는다.
        if getattr(self, 'is_user_stopped', False):
            return
            
        # 알림 팝업 창 노출
        def show_popup():
            msg_box = tk.Toplevel(self.root)
            msg_box.title("요금 수정 작업 결과 요약")
            msg_box.geometry("600x500")
            msg_box.configure(bg=self.bg_color)
            msg_box.transient(self.root)
            msg_box.grab_set()
            
            lbl = tk.Label(msg_box, text="[작업 완료] 결과 요약 보고서", font=('맑은 고딕', 12, 'bold'), bg=self.bg_color, fg=self.fg_color, pady=10)
            lbl.pack()
            
            txt = ScrolledText(msg_box, bg='#0d0d0f', fg='#c5c5cb', font=('Consolas', 9.5), bd=1, relief=tk.SOLID)
            txt.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
            txt.insert(tk.END, report_str)
            txt.config(state=tk.DISABLED)
            
            btn = tk.Button(msg_box, text="확인", width=12, bg=self.accent_color, fg='white', font=('맑은 고딕', 9, 'bold'), bd=0, command=msg_box.destroy)
            btn.pack(pady=(0, 15))
            
        self.root.after(0, show_popup)

def _acquire_single_instance_lock():
    """중복 실행 방지용 Windows 네임드 뮤텍스를 잡는다.
    이미 실행 중이면 False를 반환한다."""
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32
        mutex_name = "Global\\NaeilERPUpdater_SingleInstance_Mutex"
        handle = kernel32.CreateMutexW(None, wintypes.BOOL(True), mutex_name)
        ERROR_ALREADY_EXISTS = 183
        if kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
            return False, handle
        return True, handle
    except Exception:
        # 뮤텍스 사용 불가 환경(비윈도우 등)에서는 잠금 없이 진행
        return True, None


if __name__ == "__main__":
    # PyInstaller 스플래시(실행 직후 "로딩 중" 창)가 있으면 닫는다
    try:
        import pyi_splash  # type: ignore
        pyi_splash.update_text('프로그램을 준비하고 있습니다…')
    except Exception:
        pyi_splash = None

    # 콘솔이 없는 환경(pythonw/스플래시)에서도 안전하게
    try:
        if sys.stdout is not None and getattr(sys.stdout, 'encoding', None) != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
    if sys.stdout is None:
        sys.stdout = NullStream()

    # 중복 실행 방지: 여러 번 클릭해도 창이 하나만 뜨도록
    acquired, _mutex_handle = _acquire_single_instance_lock()
    if not acquired:
        try:
            if 'pyi_splash' in dir() and pyi_splash:
                pyi_splash.close()
        except Exception:
            pass
        try:
            _alert = tk.Tk()
            _alert.withdraw()
            messagebox.showinfo("이미 실행 중", "프로그램이 이미 실행 중입니다.\n열려 있는 창을 확인해 주세요.")
            _alert.destroy()
        except Exception:
            pass
        sys.exit(0)

    root_win = tk.Tk()
    app = RpaGuiApp(root_win)

    # UI 준비가 끝났으니 스플래시 닫기
    try:
        if pyi_splash:
            pyi_splash.close()
    except Exception:
        pass

    def on_window_close():
        if app.is_running:
            if messagebox.askyesno("앱 종료", "현재 RPA가 작동 중입니다. 중지하고 프로그램을 종료하시겠습니까?"):
                app.is_running = False
                root_win.destroy()
        else:
            root_win.destroy()

    root_win.protocol("WM_DELETE_WINDOW", on_window_close)
    root_win.mainloop()
