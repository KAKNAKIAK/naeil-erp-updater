import argparse
import contextlib
import json
import os
import sys
import time
from datetime import datetime

import pandas as pd

import excel_loader
import gui

FARE_FIELDS = (
    ("adult_air", ["항공비", "성인항공비", "운임", "최저가", "adult_air", "adultair", "air"], "성인 항공비"),
    ("adult_hotel", ["호텔비", "성인호텔비", "adult_hotel", "hotel"], "성인 호텔비"),
    ("adult_land", ["지상비", "성인지상비", "adult_land", "land"], "성인 지상비"),
    ("adult_tour", ["여행경비", "경비", "adult_tour", "tour", "expense"], "성인 여행경비"),
    ("adult_profit", ["알선수익", "수익", "adult_profit", "profit"], "성인 알선수익"),
    ("child_fare", ["소아", "소아요금", "child", "child_fare", "childfare"], "소아 요금"),
    ("infant_fare", ["유아", "유아요금", "infant", "infant_fare", "infantfare"], "유아 요금"),
)
FARE_KEYS = tuple(field for field, _aliases, _label in FARE_FIELDS)
SKIP_TOKENS = {"조회불가", "조회오류"}


def load_cli_config(config_path=None):
    """Load erp_updater config.json merged over RpaGuiApp default config/selectors."""
    runner = HeadlessErpRunner.__new__(HeadlessErpRunner)
    cfg = gui.RpaGuiApp.load_config(runner)
    if config_path:
        with open(config_path, "r", encoding="utf-8") as f:
            user_cfg = json.load(f)
        user_selectors = user_cfg.get("selectors")
        merged = dict(cfg)
        merged.update(user_cfg)
        if isinstance(user_selectors, dict):
            selectors = dict(cfg.get("selectors") or {})
            selectors.update(user_selectors)
            merged["selectors"] = selectors
        cfg = merged
    return cfg


def _parse_optional_fare(val, label, row_no, warnings):
    """Parse fare cell preserving blank/missing cells as '' so ERP leaves untouched fields alone."""
    if val is None or pd.isna(val):
        return ""
    text = str(val).strip()
    if not text or excel_loader._has_progress_status_text(text):
        return ""
    try:
        parsed = int(float(text.replace(",", "")))
        if parsed < 0:
            warnings.append(f"{row_no}행: {label} 요금이 음수({parsed})라 제외했습니다.")
            return ""
        return str(parsed)
    except (ValueError, TypeError):
        warnings.append(f"{row_no}행: {label} 요금 오류값({text})을 제외했습니다.")
        return ""


def _read_excel_jobs(
    file_path,
    *,
    price_desc="",
    airline_code="",
    departure_flight="",
    hotel_name="",
    progress_text="",
    allow_unfiltered=False,
):
    """Read multi-job (condition columns) or single-sheet Excel without turning blanks into 0."""
    warnings = []
    if not file_path or not os.path.exists(file_path):
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {file_path}")

    df = pd.read_excel(file_path)
    if df.empty:
        return {"detected_multi_job": False, "is_period": False, "raw_jobs": [], "warnings": ["엑셀 파일에 데이터가 없습니다."]}

    headers = [str(c).strip() for c in df.columns]
    price_col = excel_loader._find_column(headers, ["요금구분", "가격구분", "price_desc", "pricedesc", "price description"])
    airline_col = excel_loader._find_column(headers, ["항공사코드", "항공사", "airline", "airline_code", "airlinecode", "air2cd"])
    flight_col = excel_loader._find_column(headers, ["출발편", "출발편명", "출발항공편", "편명", "departure_flight", "departureflight", "trans_flight", "transflight", "flight"])
    hotel_col = excel_loader._find_column(headers, ["호텔명", "호텔", "hotel_name", "hotelname", "hotel_kor_nm", "hotelkornm"])
    progress_col = excel_loader._find_column(headers, ["진행구분", "진행상태", "progress", "progress_status", "proccd"])

    detected_multi = price_col is not None and airline_col is not None
    if not detected_multi:
        has_condition = any(
            str(v or "").strip() for v in (price_desc, airline_code, departure_flight, hotel_name)
        )
        if not has_condition and not allow_unfiltered:
            raise ValueError(
                "단일 엑셀은 최소 1개 조건(--price-desc/--airline/--flight/--hotel)이 필요합니다. "
                "전체 행사 대상 실행이 의도라면 --allow-unfiltered를 지정하세요."
            )

    condition_cols = {
        idx for idx in (price_col, airline_col, flight_col, hotel_col, progress_col) if idx is not None
    }
    data_cols = [idx for idx in range(len(headers)) if idx not in condition_cols]
    start_col = excel_loader._find_column(
        headers, ["시작일", "날짜", "출발일", "출발일자", "start", "start_date", "startdate", "date"]
    )
    end_col = excel_loader._find_column(
        headers, ["종료일", "종료날짜", "종료", "end", "end_date", "enddate"]
    )
    if start_col is None and data_cols:
        start_col = data_cols[0]

    if end_col is None and start_col is not None:
        candidates = [idx for idx in data_cols if idx != start_col]
        if candidates:
            total = date_like = 0
            for val in df.iloc[:, candidates[0]]:
                marker = excel_loader._is_date_like(val)
                if marker is None:
                    continue
                total += 1
                if marker:
                    date_like += 1
            if total > 0 and date_like >= total * 0.6:
                end_col = candidates[0]

    is_period = end_col is not None
    ordered_fare_cols = [idx for idx in data_cols if idx not in {start_col, end_col}]
    has_named_fares = any(excel_loader._find_column(headers, aliases) is not None for _, aliases, _ in FARE_FIELDS)
    col_map = []
    for offset, (field, aliases, label) in enumerate(FARE_FIELDS):
        found = excel_loader._find_column(headers, aliases)
        if found is None and not has_named_fares:
            found = ordered_fare_cols[offset] if offset < len(ordered_fare_cols) else None
        col_map.append((field, found, label))

    default_p_code, default_p_label = excel_loader.progress_status_from_text(progress_text)
    default_progress = default_p_label if default_p_code else str(progress_text or "").strip()
    source_name = os.path.basename(file_path)

    jobs_by_key = {}
    job_order = []
    for idx, row in df.iterrows():
        row_no = idx + 2
        row_price_desc = excel_loader._cell_text(row, price_col) if detected_multi else str(price_desc or "").strip()
        row_airline = (excel_loader._cell_text(row, airline_col) if detected_multi else str(airline_code or "").strip()).upper()
        row_flight = excel_loader._cell_text(row, flight_col) if (detected_multi and flight_col is not None) else str(departure_flight or "").strip()
        row_hotel = excel_loader._cell_text(row, hotel_col) if (detected_multi and hotel_col is not None) else str(hotel_name or "").strip()
        if detected_multi and (not row_price_desc or not row_airline):
            warnings.append(f"{row_no}행: 요금구분과 항공사코드는 필수입니다.")
            continue

        date_str = excel_loader._normalize_dt(row.iloc[start_col]) if start_col is not None else None
        if not date_str:
            warnings.append(f"{row_no}행: 시작 날짜 형식 오류")
            continue
        end_raw = row.iloc[end_col] if end_col is not None else row.iloc[start_col]
        date_end_str = excel_loader._normalize_dt(end_raw) or date_str
        if date_str > date_end_str:
            warnings.append(f"{row_no}행: 시작 날짜({date_str})가 종료 날짜({date_end_str})보다 늦습니다.")
            continue

        row_progress = excel_loader._cell_text(row, progress_col) if progress_col is not None else ""
        raw_vals = [(field, row.iloc[col] if col is not None else None) for field, col, _label in col_map]

        if any(str(v).strip() in SKIP_TOKENS for _f, v in raw_vals if v is not None and not pd.isna(v)):
            warnings.append(f"{row_no}행({date_str}): 조회불가/조회오류 행 건너뜀")
            continue

        progress_field = next(
            (field for field, val in raw_vals if val is not None and not pd.isna(val) and excel_loader._has_progress_status_text(val)),
            "",
        )
        if not row_progress:
            row_progress = next(
                (str(val).strip() for _field, val in raw_vals if val is not None and not pd.isna(val) and excel_loader._has_progress_status_text(val)),
                default_progress,
            )
        p_code, p_label = excel_loader.progress_status_from_text(row_progress)
        if p_code:
            row_progress = p_label

        parsed = {
            field: _parse_optional_fare(row.iloc[col] if col is not None else None, label, row_no, warnings)
            for field, col, label in col_map
        }
        if not any(parsed[k] != "" for k in FARE_KEYS) and not row_progress:
            continue

        row_data = {
            "row_index": row_no,
            "date": date_str,
            "date_end": date_end_str,
            **parsed,
            "price_desc": row_price_desc,
            "airline_code": row_airline,
            "departure_flight": row_flight,
            "hotel_name": row_hotel,
            "progress_status": row_progress,
            "progress_status_field": progress_field,
        }
        key = (row_price_desc, row_airline, row_flight, row_hotel)
        if key not in jobs_by_key:
            jobs_by_key[key] = {
                "price_desc": row_price_desc,
                "airline_code": row_airline,
                "departure_flight": row_flight,
                "hotel_name": row_hotel,
                "progress_text": "",
                "rows": [],
                "source": source_name,
            }
            job_order.append(key)
        jobs_by_key[key]["rows"].append(row_data)

    return {
        "detected_multi_job": detected_multi,
        "is_period": is_period,
        "raw_jobs": [jobs_by_key[k] for k in job_order if jobs_by_key[k]["rows"]],
        "warnings": warnings,
    }


def _normalize_cli_date(value, label):
    if value is None or str(value).strip() == "":
        return None
    normalized = excel_loader._normalize_dt(value)
    if not normalized:
        raise ValueError(f"{label} 날짜 형식이 올바르지 않습니다: {value}")
    return normalized


def filter_rows_by_date_range(rows, start=None, end=None):
    """Overlap filter matching gui._apply_date_filter (date_end >= start and date <= end)."""
    start_norm = _normalize_cli_date(start, "--start")
    end_norm = _normalize_cli_date(end, "--end")
    if start_norm and end_norm and start_norm > end_norm:
        raise ValueError(f"--start({start_norm})가 --end({end_norm})보다 늦습니다.")
    if not start_norm and not end_norm:
        return list(rows)
    out = []
    for row in rows:
        d_start = str(row["date"]).strip()
        d_end = str(row.get("date_end") or d_start).strip()
        if start_norm and d_end < start_norm:
            continue
        if end_norm and d_start > end_norm:
            continue
        out.append(row)
    return out


def merge_job_rows(rows):
    """Merge consecutive identical fare/progress dates while preserving gaps."""
    if not rows:
        return [], []
    records = []
    for row in rows:
        start_dt = datetime.strptime(str(row["date"]).strip(), "%Y-%m-%d").date()
        end_dt = datetime.strptime(str(row.get("date_end") or row["date"]).strip(), "%Y-%m-%d").date()
        key = tuple(str(row.get(k, "")).strip() for k in FARE_KEYS) + (
            str(row.get("progress_status") or "").strip(),
            str(row.get("progress_status_field") or "").strip(),
        )
        records.append({"start": start_dt, "end": end_dt, "key": key, "row": dict(row)})

    merged_records, gaps = gui.RpaGuiApp._merge_fare_records_preserving_gaps(records)
    merged_rows = []
    for rec in merged_records:
        item = dict(rec["row"])
        item["date"] = rec["start"].strftime("%Y-%m-%d")
        item["date_end"] = rec["end"].strftime("%Y-%m-%d")
        merged_rows.append(item)
    return merged_rows, gaps


def load_jobs_for_cli(
    excel_path,
    *,
    price_desc="",
    airline_code="",
    departure_flight="",
    hotel_name="",
    progress_text="",
    start=None,
    end=None,
    merge=False,
    allow_unfiltered=False,
):
    """Load Excel into normalized job list + flattened fares_data for HeadlessErpRunner."""
    parsed = _read_excel_jobs(
        excel_path,
        price_desc=price_desc,
        airline_code=airline_code,
        departure_flight=departure_flight,
        hotel_name=hotel_name,
        progress_text=progress_text,
        allow_unfiltered=allow_unfiltered,
    )

    helper = HeadlessErpRunner(config={}, dry_run=True)
    jobs = []
    for raw_job in parsed["raw_jobs"]:
        norm = helper._normalize_job(raw_job)
        rows = filter_rows_by_date_range(norm.get("rows") or [], start=start, end=end)
        if merge and rows:
            rows, _gaps = merge_job_rows(rows)
        if not rows:
            continue
        norm["rows"] = rows
        jobs.append(norm)

    fares_data = []
    for q_idx, job in enumerate(jobs):
        job["_queue_index"] = q_idx
        label = helper._job_condition_text(job)
        for row in job["rows"]:
            row["_job_index"] = q_idx
            row["_job_label"] = label
            row["_job_source"] = job.get("source") or os.path.basename(excel_path)
            fares_data.append(row)

    return {
        "detected_multi_job": parsed["detected_multi_job"],
        "is_period": parsed["is_period"],
        "jobs": jobs,
        "fares_data": fares_data,
        "warnings": parsed["warnings"],
    }


class _HeadlessRoot:
    def after(self, _delay_ms, callback=None, *args):
        if callable(callback):
            callback(*args)

    def update_idletasks(self):
        pass


class HeadlessErpRunner(gui.RpaGuiApp):
    """Runs RpaGuiApp's proven ERP RPA loop without Tkinter windows."""

    def __init__(self, config, *, dry_run=True, explicit_erp_address=None):
        self.config = dict(config or {})
        self.dry_run = bool(dry_run)
        # Never inherit dry_run=False from config.json unless explicitly asked via --apply.
        self.config["dry_run"] = self.dry_run
        self.explicit_erp_address = str(explicit_erp_address or "").strip() or None
        if self.explicit_erp_address:
            self.config["erp_debugger_address"] = self.explicit_erp_address

        self.root = _HeadlessRoot()
        self.driver = None
        self.is_running = True
        self.is_paused = False
        self.is_user_stopped = False
        self.fares_data = []
        self.selected_airline_code = ""
        self.selected_price_desc = ""
        self.selected_departure_flight = ""
        self.selected_hotel_name = ""
        self.selected_hotel_seq = ""
        self.selected_hotel_seq_strict = False
        self.selected_progress_text = ""
        self.airline_value_by_label = {}
        self.airline_choices = []

        self.final_history = []
        self.planned_total = 0
        self.fatal_error = ""
        self.dry_run_actions = []

    def _debug_browser_configs(self, target_name):
        if self.explicit_erp_address and str(target_name or "").upper() != "TOPAS":
            host, port = self._parse_debugger_address(self.explicit_erp_address, 9223)
            return [{
                "address": f"{host}:{port}",
                "host": host,
                "port": port,
                "profile_dir": str(self.config.get("erp_chrome_profile_dir", "ChromeProfile_ERP")),
            }]
        return super()._debug_browser_configs(target_name)

    def show_loading(self, message="처리 중..."):
        print(f"[CLI] {message}")

    def hide_loading(self):
        pass

    def update_progress_ui(self, current, total):
        pct = int((current / total) * 100) if total else 0
        print(f"[진행] {current}/{total} ({pct}%)")

    def _set_job_progress_ui(self, job_index, status=None, result_status=None):
        pass

    def finish_rpa_on_ui(self, history, planned_cnt=None):
        self.final_history = list(history or [])
        self.planned_total = int(planned_cnt) if planned_cnt is not None else len(self.final_history)

    def _should_set_reservation_requested_for_hotel_update(self, inputs_mapping):
        if self.dry_run:
            # shortcut: dry-run skips live 06->04 pre-transition verification; logged in _save_current_page instead
            return False
        return super()._should_set_reservation_requested_for_hotel_update(inputs_mapping)

    def _save_current_page(self, selectors, wait, inputs_mapping, driver_timeout,
                           erp_short_pause, erp_poll_interval, target_page=1):
        if not self.dry_run:
            return super()._save_current_page(
                selectors, wait, inputs_mapping, driver_timeout,
                erp_short_pause, erp_poll_interval, target_page=target_page,
            )
        if str(inputs_mapping.get("adult_hotel_input", "")).strip():
            rows = self._current_page_progress_rows()
            pending = self._progress_status_row_count(rows, "06")
            if pending:
                print(f" -> [DRY-RUN] 호텔비 수정 전 대기예약(06) {pending}건 예약신청 선변경 예정")
        self.dry_run_actions.append({
            "action": "price_update",
            "inputs": dict(inputs_mapping),
        })
        print(f" -> [DRY-RUN] 저장 생략 (입력 예정: {dict(inputs_mapping)})")

    def _apply_current_page_progress_status(self, selectors, progress_code, progress_label,
                                            driver_timeout, erp_short_pause, erp_poll_interval,
                                            source_progress_code=None, target_page=1):
        if not self.dry_run:
            return super()._apply_current_page_progress_status(
                selectors, progress_code, progress_label,
                driver_timeout, erp_short_pause, erp_poll_interval,
                source_progress_code=source_progress_code, target_page=target_page,
            )
        self.dry_run_actions.append({
            "action": "progress_status",
            "progress_code": progress_code,
            "progress_label": progress_label,
            "source_progress_code": source_progress_code,
            "target_page": target_page,
        })
        print(f" -> [DRY-RUN] 진행구분 변경 생략 ({progress_label}({progress_code}), page={target_page})")
        return 0

    def _verify_progress_rows(self, rows, expected_ids, progress_code):
        if self.dry_run:
            return
        return super()._verify_progress_rows(rows, expected_ids, progress_code)

    def check_connection(self):
        selectors = self.config.get("selectors") or {}
        driver, browser_cfg = self._connect_matching_debug_browser("ERP", selectors)
        self.driver = driver
        has_screen = bool(self.find_and_switch_frame(selectors.get("search_date_input", "#searchStDate")))
        return {
            "connected": True,
            "address": browser_cfg["address"],
            "fareScreenReady": has_screen,
        }

    def run_fares(self, fares_data):
        self.fares_data = list(fares_data)
        self.final_history = []
        self.planned_total = len(self.fares_data)
        self.fatal_error = ""
        self.dry_run_actions = []

        orig_showerror = gui.messagebox.showerror
        orig_ensure = gui.ensure_window_visible

        def _capture_error(title, msg, **_kwargs):
            self.fatal_error = f"{title}: {msg}"
            print(f"[오류] {self.fatal_error}")

        gui.messagebox.showerror = _capture_error
        gui.ensure_window_visible = lambda _win: None
        try:
            self.rpa_worker_loop()
            if not self.fatal_error and getattr(self, "_last_loop_error", ""):
                self.fatal_error = f"작업 루프 오류: {self._last_loop_error}"
        finally:
            gui.messagebox.showerror = orig_showerror
            gui.ensure_window_visible = orig_ensure
        return self.final_history


def _emit_json(orig_stdout, payload):
    orig_stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    orig_stdout.flush()


def build_arg_parser():
    ap = argparse.ArgumentParser(description="Naeil ERP 행사요금일괄수정 CLI (기본: dry-run)")
    ap.add_argument("--check", action="store_true", help="ERP 디버그 Chrome 연결 및 행사요금일괄수정 화면 확인")
    ap.add_argument("--excel", help="요금 엑셀 파일 경로 (.xlsx/.xls)")
    ap.add_argument("--price-desc", default="", help="단일 엑셀 요금구분 필터 (예: 다낭)")
    ap.add_argument("--airline", default="", help="단일 엑셀 항공사코드 필터 (예: LJ)")
    ap.add_argument("--flight", default="", help="단일 엑셀 출발편 필터 (예: LJ081)")
    ap.add_argument("--hotel", default="", help="단일 엑셀 호텔명 필터")
    ap.add_argument("--progress", default="", help="단일 엑셀 기본 진행구분 (예: 예약마감)")
    ap.add_argument("--start", help="날짜 필터 시작일 YYYY-MM-DD (겹침 기준)")
    ap.add_argument("--end", help="날짜 필터 종료일 YYYY-MM-DD (겹침 기준)")
    ap.add_argument("--merge", action="store_true", help="연속 동일 요금·진행구간 병합 (중간 빈 날짜 분리)")
    ap.add_argument("--plan", action="store_true", help="브라우저 연결 없이 파싱·병합 계획만 JSON 출력")
    ap.add_argument("--apply", action="store_true", help="실제 ERP 저장 수행 (미지정 시 dry-run)")
    ap.add_argument("--allow-unfiltered", action="store_true", help="단일 엑셀 조건 없이 전체 행사 실행 허용")
    ap.add_argument("--erp-address", help="ERP 디버그 Chrome 주소 지정 (예: 127.0.0.1:9224)")
    ap.add_argument("--config", help="config.json 경로 지정")
    return ap


def main(argv=None):
    ap = build_arg_parser()
    args = ap.parse_args(argv)
    orig_stdout = sys.stdout
    started = time.time()

    if not args.check and not args.excel:
        ap.error("--check 또는 --excel 중 하나는 지정해야 합니다.")

    mode = "check" if args.check else ("plan" if args.plan else ("apply" if args.apply else "dry_run"))
    payload = {
        "ok": False,
        "mode": mode,
        "excel": os.path.abspath(args.excel) if args.excel else "",
        "erpAddress": args.erp_address or "",
        "jobsCount": 0,
        "planned": 0,
        "executed": 0,
        "unrun": 0,
        "success": 0,
        "skip": 0,
        "fail": 0,
        "warnings": [],
        "error": "",
        "elapsedSec": 0.0,
        "items": [],
    }

    try:
        with contextlib.redirect_stdout(sys.stderr):
            config = load_cli_config(args.config)
            erp_addr = args.erp_address or config.get("erp_debugger_address", "127.0.0.1:9223")
            payload["erpAddress"] = erp_addr

            if args.check:
                runner = HeadlessErpRunner(config, dry_run=True, explicit_erp_address=args.erp_address)
                info = runner.check_connection()
                payload.update(info)
                payload["erpAddress"] = info["address"]
                payload["ok"] = bool(info["connected"] and info["fareScreenReady"])
                if not info["fareScreenReady"]:
                    payload["error"] = "행사요금일괄수정 화면(출발일자 입력칸)을 찾지 못했습니다."
                payload["elapsedSec"] = round(time.time() - started, 2)
                _emit_json(orig_stdout, payload)
                return 0 if payload["ok"] else 1

            loaded = load_jobs_for_cli(
                args.excel,
                price_desc=args.price_desc,
                airline_code=args.airline,
                departure_flight=args.flight,
                hotel_name=args.hotel,
                progress_text=args.progress,
                start=args.start,
                end=args.end,
                merge=args.merge,
                allow_unfiltered=args.allow_unfiltered,
            )
            payload["warnings"] = loaded["warnings"]
            payload["jobsCount"] = len(loaded["jobs"])
            payload["planned"] = len(loaded["fares_data"])

            if not loaded["fares_data"]:
                payload["error"] = "실행할 유효한 요금 행이 없습니다."
                payload["elapsedSec"] = round(time.time() - started, 2)
                _emit_json(orig_stdout, payload)
                return 1

            if args.plan:
                payload["ok"] = True
                payload["items"] = [
                    {
                        "date": r["date"],
                        "dateEnd": r.get("date_end", r["date"]),
                        "priceDesc": r.get("price_desc", ""),
                        "airline": r.get("airline_code", ""),
                        "flight": r.get("departure_flight", ""),
                        "hotel": r.get("hotel_name", ""),
                        "progress": r.get("progress_status", ""),
                        "fares": {k: r.get(k, "") for k in FARE_KEYS if str(r.get(k, "")).strip() != ""},
                    }
                    for r in loaded["fares_data"]
                ]
                payload["elapsedSec"] = round(time.time() - started, 2)
                _emit_json(orig_stdout, payload)
                return 0

            runner = HeadlessErpRunner(
                config,
                dry_run=(not args.apply),
                explicit_erp_address=args.erp_address,
            )
            history = runner.run_fares(loaded["fares_data"])
            success = sum(1 for h in history if h.get("status") == "SUCCESS")
            skip = sum(1 for h in history if h.get("status") == "SKIP")
            fail = sum(1 for h in history if h.get("status") == "FAIL")
            unrun = max(0, payload["planned"] - len(history))

            payload["executed"] = len(history)
            payload["unrun"] = unrun
            payload["success"] = success
            payload["skip"] = skip
            payload["fail"] = fail
            payload["items"] = history
            if runner.dry_run:
                payload["dryRunActions"] = len(runner.dry_run_actions)
            if runner.fatal_error:
                payload["error"] = runner.fatal_error
            elif unrun > 0:
                payload["error"] = f"전체 {payload['planned']}건 중 {unrun}건 미실행"

            payload["ok"] = (fail == 0 and unrun == 0 and not runner.fatal_error)
            payload["elapsedSec"] = round(time.time() - started, 2)
    except Exception as exc:
        payload["error"] = str(exc)
        payload["elapsedSec"] = round(time.time() - started, 2)
        _emit_json(orig_stdout, payload)
        return 1

    _emit_json(orig_stdout, payload)
    return 0 if payload["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
