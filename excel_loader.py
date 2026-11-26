# -*- coding: utf-8 -*-
import os
import sys
import re
from datetime import date
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def select_excel_file(parent=None):
    """
    윈도우 파일 대화상자를 띄워 사용자가 엑셀 파일을 선택하게 합니다.
    parent가 주어지면 해당 부모 창을 사용하고, None이면 독립 임시 root를 생성하여 안전하게 정리합니다.
    """
    if parent is not None:
        return filedialog.askopenfilename(
            parent=parent,
            title="요금 업데이트 엑셀 파일 선택",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    try:
        file_path = filedialog.askopenfilename(
            parent=root,
            title="요금 업데이트 엑셀 파일 선택",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
    finally:
        try:
            root.destroy()
        except Exception:
            pass
    return file_path

def _normalize_dt(raw_dt):
    if pd.isna(raw_dt):
        return None
    if isinstance(raw_dt, pd.Timestamp):
        return raw_dt.strftime('%Y-%m-%d')
    dt_str = re.split(r'[ T]', str(raw_dt).strip(), maxsplit=1)[0]
    dt_str = dt_str.replace('/', '-').replace('.', '-').strip('-')
    if len(dt_str) == 8 and dt_str.isdigit():
        try:
            return date(int(dt_str[:4]), int(dt_str[4:6]), int(dt_str[6:])).isoformat()
        except ValueError:
            return None
    parts = dt_str.split('-')
    if len(parts) == 3:
        try:
            y, m, d = parts
            if len(y) == 4:
                return date(int(y), int(m), int(d)).isoformat()
        except (ValueError, TypeError):
            return None
    return None


def _is_date_like(val):
    """값이 '실제 날짜'로 보이는지 판정한다.
    반환: True(날짜) / False(날짜 아님) / None(빈 값 — 판단 보류)
    8자리 요금(예: 41000000)이 날짜로 오인되지 않도록 연/월/일 범위까지 확인한다."""
    if pd.isna(val):
        return None
    if isinstance(val, pd.Timestamp):
        return True
    s = str(val).strip()
    if not s:
        return None
    norm = _normalize_dt(val)
    if not norm:
        return False
    try:
        y, m, d = (int(x) for x in norm.split('-'))
        return 2000 <= y <= 2100 and 1 <= m <= 12 and 1 <= d <= 31
    except (ValueError, TypeError):
        return False


def _normalize_header_name(value):
    text = "" if value is None else str(value).strip().lower()
    return re.sub(r"[\s_\-./()]+", "", text)


def _find_column(headers, aliases):
    normalized_aliases = {_normalize_header_name(alias) for alias in aliases}
    for idx, header in enumerate(headers):
        if _normalize_header_name(header) in normalized_aliases:
            return idx
    return None


def _cell_text(row, col_idx):
    if col_idx is None:
        return ""
    try:
        value = row.iloc[col_idx]
    except Exception:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


PROGRESS_STATUS_CODES = {
    "대기예약": "06",
    "예약신청": "04",
    "예약마감": "05",
}

PROGRESS_STATUS_ALIASES = {
    "대기예약": ("대기예약", "예약대기"),
    "예약신청": ("예약신청",),
    "예약마감": ("예약마감",),
}

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


def progress_status_from_text(value):
    """입력표/엑셀 진행구분 문구를 ERP 진행구분 코드로 변환한다."""
    normalized = re.sub(r"\s+", "", str(value or ""))
    for label, code in PROGRESS_STATUS_CODES.items():
        aliases = PROGRESS_STATUS_ALIASES.get(label, (label,))
        if any(alias in normalized for alias in aliases):
            return code, label
    return "", ""


def _has_progress_status_text(value):
    return bool(progress_status_from_text(value)[0])


def _resolve_fare_columns(headers, data_cols, start_col, end_col):
    ordered_fare_cols = [idx for idx in data_cols if idx not in {start_col, end_col}]
    has_named_fares = any(_find_column(headers, aliases) is not None for _, aliases, _ in FARE_FIELDS)
    col_map = []
    for offset, (field, aliases, label) in enumerate(FARE_FIELDS):
        found = _find_column(headers, aliases)
        if found is None and not has_named_fares and start_col is not None:
            found = ordered_fare_cols[offset] if offset < len(ordered_fare_cols) else None
        col_map.append((field, found, label))
    return col_map


def _parse_optional_fare(val, label, row_no, warnings, *, as_str=False):
    """Parse fare cell preserving blank/missing/invalid cells as '' so ERP leaves untouched fields alone."""
    if val is None or pd.isna(val):
        return ""
    text = str(val).strip()
    if not text or _has_progress_status_text(text):
        return ""
    try:
        parsed = int(float(text.replace(",", "")))
        if parsed < 0:
            warnings.append(f"{row_no}행: {label} 요금이 음수({parsed})라 제외했습니다.")
            return ""
        return str(parsed) if as_str else parsed
    except (ValueError, TypeError):
        warnings.append(f"{row_no}행: {label} 요금 오류값({text})을 제외했습니다.")
        return ""


def load_fare_jobs_from_excel(file_path, history_log_path=None):
    """조건열이 있는 엑셀을 작업 큐용 그룹으로 읽는다.

    조건열 감지 기준: 요금구분 + 항공사코드 열이 모두 있어야 한다. 출발편/호텔명/진행구분은 선택 조건열이다.
    반환값: {
        detected: bool,
        jobs: [{price_desc, airline_code, departure_flight, hotel_name, rows, source}],
        is_period: bool,
        errors: [str]
    }
    """
    result = {"detected": False, "jobs": [], "is_period": False, "errors": []}
    if not file_path or not os.path.exists(file_path):
        return result

    try:
        df = pd.read_excel(file_path)
    except Exception as exc:
        result["errors"].append(f"엑셀 읽기 실패: {exc}")
        return result

    if df.empty:
        result["errors"].append("엑셀 파일에 데이터가 없습니다.")
        return result

    headers = [str(c).strip() for c in df.columns]
    price_col = _find_column(headers, ["요금구분", "가격구분", "price_desc", "pricedesc", "price description"])
    airline_col = _find_column(headers, ["항공사코드", "항공사", "airline", "airline_code", "airlinecode", "air2cd"])
    departure_flight_col = _find_column(headers, ["출발편", "출발편명", "출발항공편", "편명", "departure_flight", "departureflight", "trans_flight", "transflight", "transFlight", "flight"])
    hotel_col = _find_column(headers, ["호텔명", "호텔", "hotel_name", "hotelname", "hotel_kor_nm", "hotelkornm", "hotelKorNm"])
    progress_col = _find_column(headers, ["진행구분", "진행상태", "progress", "progress_status", "proccd"])
    if price_col is None or airline_col is None:
        return result

    result["detected"] = True
    condition_cols = {idx for idx in (price_col, airline_col, departure_flight_col, hotel_col, progress_col) if idx is not None}

    start_col = _find_column(headers, ["시작일", "날짜", "출발일", "출발일자", "start", "start_date", "startdate", "date"])
    end_col = _find_column(headers, ["종료일", "종료날짜", "종료", "end", "end_date", "enddate"])
    data_cols = [idx for idx in range(len(headers)) if idx not in condition_cols]
    if start_col is None and data_cols:
        start_col = data_cols[0]

    if end_col is None and start_col is not None:
        candidates = [idx for idx in data_cols if idx != start_col]
        if candidates:
            total = date_like = 0
            for value in df.iloc[:, candidates[0]]:
                marker = _is_date_like(value)
                if marker is None:
                    continue
                total += 1
                if marker:
                    date_like += 1
            if total > 0 and date_like >= total * 0.6:
                end_col = candidates[0]

    is_period_excel = end_col is not None
    result["is_period"] = is_period_excel

    fare_cols = _resolve_fare_columns(headers, data_cols, start_col, end_col)

    completed_dates = set()
    if history_log_path and os.path.exists(history_log_path):
        try:
            history_df = pd.read_csv(history_log_path, encoding="utf-8")
            if not history_df.empty and "status" in history_df.columns and "date" in history_df.columns:
                success_rows = history_df[history_df["status"].str.strip() == "SUCCESS"]
                completed_dates = set(success_rows["date"].astype(str).str.strip())
        except Exception as exc:
            result["errors"].append(f"이력 로그 읽기 실패: {exc}")

    jobs_by_key = {}
    job_order = []
    source_name = os.path.basename(file_path)
    for idx, row in df.iterrows():
        row_no = idx + 2
        price_desc = _cell_text(row, price_col)
        airline_code = _cell_text(row, airline_col).upper()
        departure_flight = _cell_text(row, departure_flight_col)
        hotel_name = _cell_text(row, hotel_col)
        progress_text = _cell_text(row, progress_col)
        if not price_desc or not airline_code:
            result["errors"].append(f"{row_no}행: 요금구분과 항공사코드는 필수입니다.")
            continue

        date_str = _normalize_dt(row.iloc[start_col]) if start_col is not None else None
        if not date_str:
            result["errors"].append(f"{row_no}행: 시작 날짜 형식 오류")
            continue
        end_raw = row.iloc[end_col] if end_col is not None else row.iloc[start_col]
        date_end_str = _normalize_dt(end_raw) or date_str
        if date_str > date_end_str:
            result["errors"].append(f"{row_no}행: 시작 날짜({date_str})가 종료 날짜({date_end_str})보다 늦습니다.")
            continue
        if date_str in completed_dates:
            continue

        raw_fare_values = [
            (field, row.iloc[col] if col is not None else "")
            for field, col, _label in fare_cols
        ]
        if any(str(v).strip() in SKIP_TOKENS for _f, v in raw_fare_values if v is not None and not pd.isna(v)):
            result["errors"].append(f"{row_no}행({date_str}): 조회불가/조회오류 행 건너뜀")
            continue
        progress_status_field = next(
            (field for field, value in raw_fare_values if _has_progress_status_text(value)),
            "",
        )
        if not progress_text:
            progress_text = next((str(v).strip() for _field, v in raw_fare_values if _has_progress_status_text(v)), "")
        _progress_code, canonical_progress_label = progress_status_from_text(progress_text)
        if canonical_progress_label:
            progress_text = canonical_progress_label

        parsed = {
            field: _parse_optional_fare(row.iloc[col] if col is not None else None, label, row_no, result["errors"])
            for field, col, label in fare_cols
        }
        if not any(parsed[k] != "" for k in FARE_KEYS) and not progress_text:
            continue
        row_data = {
            "row_index": row_no,
            "date": date_str,
            "date_end": date_end_str,
            **parsed,
            "price_desc": price_desc,
            "airline_code": airline_code,
            "departure_flight": departure_flight,
            "hotel_name": hotel_name,
            "progress_status": progress_text,
            "progress_status_field": progress_status_field,
        }
        key = (price_desc, airline_code, departure_flight, hotel_name)
        if key not in jobs_by_key:
            jobs_by_key[key] = {
                "price_desc": price_desc,
                "airline_code": airline_code,
                "departure_flight": departure_flight,
                "hotel_name": hotel_name,
                "progress_text": "",
                "rows": [],
                "source": source_name,
            }
            job_order.append(key)
        jobs_by_key[key]["rows"].append(row_data)

    result["jobs"] = [jobs_by_key[key] for key in job_order if jobs_by_key[key]["rows"]]
    return result

def load_and_validate_fares(file_path, history_log_path=None):
    """
    선택된 엑셀 파일을 읽고 요금 데이터를 파싱 및 유효성 검사합니다.
    history_log_path가 주어지면 이전 SUCCESS 이력 날짜를 자동 스킵합니다.
    반환값: (valid_rows, is_period_excel)
    """
    if not file_path:
        print("[오류] 파일이 선택되지 않았습니다.")
        return None, False

    if not os.path.exists(file_path):
        print(f"[오류] 파일이 존재하지 않습니다: {file_path}")
        return None, False

    # 1. 성공 이력 로드 (비활성 상태 — history_log_path를 넘기지 않으면 스킵 없음)
    completed_dates = set()
    if history_log_path and os.path.exists(history_log_path):
        try:
            history_df = pd.read_csv(history_log_path, encoding='utf-8')
            if not history_df.empty and 'status' in history_df.columns and 'date' in history_df.columns:
                success_rows = history_df[history_df['status'].str.strip() == 'SUCCESS']
                completed_dates = set(success_rows['date'].astype(str).str.strip())
                print(f"[정보] 성공 이력 로그에서 {len(completed_dates)}개의 완료된 날짜를 로드했습니다.")
        except Exception as e:
            print(f"[경고] 이력 로그 읽기 실패 (처음부터 새로 진행합니다): {str(e)}")

    try:
        # 2. 엑셀 파일 읽기
        df = pd.read_excel(file_path)

        if df.empty:
            print("[오류] 엑셀 파일에 데이터가 없습니다.")
            return None, False

        print(f"[정보] 엑셀 로딩 완료. 총 {len(df)}개의 행을 감지했습니다.")

        # 3. 기간 모드 여부 자동 감지
        headers = [str(c).strip() for c in df.columns]
        is_period_excel = False

        # (a) 헤더에 '종료' 계열 키워드가 있으면 기간 모드
        period_headers = ['종료일', '종료날짜', '종료', 'end date', 'enddate', 'end_date']
        for h in headers:
            if h.lower() in period_headers:
                is_period_excel = True
                break

        # (b) 헤더 키워드가 없어도(헤더가 비어 Unnamed 로 잡히는 파일 포함)
        #     '2번째 열이 날짜로 채워져 있으면' 기간 모드로 본다.
        #     단일 모드는 2번째 열이 요금(숫자)이라 날짜 비율이 낮다 → 열 개수와 무관하게 검사.
        if not is_period_excel and df.shape[1] >= 2:
            col0_dates = col1_dates = total = 0
            for v0, v1 in zip(df.iloc[:, 0], df.iloc[:, 1]):
                r0, r1 = _is_date_like(v0), _is_date_like(v1)
                if r0 is None and r1 is None:
                    continue  # 양쪽 모두 빈 행은 판정에서 제외
                total += 1
                if r0:
                    col0_dates += 1
                if r1:
                    col1_dates += 1
            # 시작일·종료일이 둘 다 날짜인 행이 다수(60% 이상)면 기간 양식으로 확정
            if total > 0 and col0_dates >= total * 0.6 and col1_dates >= total * 0.6:
                is_period_excel = True
                print(f"[정보] 2번째 열 날짜 비율로 기간 양식을 감지했습니다 "
                      f"(시작일 {col0_dates}/{total}, 종료일 {col1_dates}/{total}).")

        valid_rows = []
        for idx, row in df.iterrows():
            try:
                row_len = len(row)
                raw_date = row.iloc[0] if row_len > 0 else None
                
                if is_period_excel:
                    raw_date_end = row.iloc[1] if row_len > 1 else None
                    raw_adult_air = row.iloc[2] if row_len > 2 else None
                    raw_adult_hotel = row.iloc[3] if row_len > 3 else None
                    raw_adult_land = row.iloc[4] if row_len > 4 else None
                    raw_adult_tour = row.iloc[5] if row_len > 5 else None
                    raw_adult_profit = row.iloc[6] if row_len > 6 else None
                    raw_child_fare = row.iloc[7] if row_len > 7 else None
                    raw_infant_fare = row.iloc[8] if row_len > 8 else None
                else:
                    raw_date_end = raw_date
                    raw_adult_air = row.iloc[1] if row_len > 1 else None
                    # 호환성 지원: 엑셀 열이 3개뿐인 경우(기존 v1 3열 양식: 날짜, 항공비, 알선수익)
                    if row_len == 3:
                        raw_adult_hotel = None
                        raw_adult_land = None
                        raw_adult_tour = None
                        raw_adult_profit = row.iloc[2]
                        raw_child_fare = None
                        raw_infant_fare = None
                    else:
                        raw_adult_hotel = row.iloc[2] if row_len > 2 else None
                        raw_adult_land = row.iloc[3] if row_len > 3 else None
                        raw_adult_tour = row.iloc[4] if row_len > 4 else None
                        raw_adult_profit = row.iloc[5] if row_len > 5 else None
                        raw_child_fare = row.iloc[6] if row_len > 6 else None
                        raw_infant_fare = row.iloc[7] if row_len > 7 else None

                # 날짜 유효성 검증 및 정규화
                date_str = _normalize_dt(raw_date)
                if not date_str:
                    continue

                date_end_str = _normalize_dt(raw_date_end)
                if not date_end_str:
                    date_end_str = date_str

                # 이미 성공한 날짜인지 검사하여 자동 스킵
                if date_str in completed_dates:
                    print(f"[이어서 진행] {idx+2}행: 시작 날짜 {date_str}는 이미 이전 실행에서 성공하여 스킵합니다.")
                    continue

                raw_cells = (
                    raw_adult_air, raw_adult_hotel, raw_adult_land,
                    raw_adult_tour, raw_adult_profit, raw_child_fare, raw_infant_fare,
                )
                if any(str(v).strip() in SKIP_TOKENS for v in raw_cells if v is not None and not pd.isna(v)):
                    print(f"[정보] {idx+2}행({date_str}): 조회불가/조회오류 행 건너뜀")
                    continue

                # 각 요금 필드 안전하게 파싱 (빈 셀은 '' 유지, 진행구분 문구는 보존)
                def parse_fare(val, label):
                    if val is None or pd.isna(val):
                        return ""
                    text = str(val).strip()
                    if not text:
                        return ""
                    p_code, p_label = progress_status_from_text(text)
                    if p_code:
                        return p_label
                    try:
                        f_val = int(float(text.replace(',', '')))
                        if f_val < 0:
                            print(f"[정보] {idx+2}행: {label} 요금이 음수({f_val})이므로 제외합니다.")
                            return ""
                        return f_val
                    except (ValueError, TypeError):
                        print(f"[정보] {idx+2}행: {label} 요금 오류값({val})이 입력되어 제외합니다.")
                        return ""

                adult_air = parse_fare(raw_adult_air, "성인 항공비")
                adult_hotel = parse_fare(raw_adult_hotel, "성인 호텔비")
                adult_land = parse_fare(raw_adult_land, "성인 지상비")
                adult_tour = parse_fare(raw_adult_tour, "성인 여행경비")
                adult_profit = parse_fare(raw_adult_profit, "성인 알선수익")
                child_fare = parse_fare(raw_child_fare, "소아 요금")
                infant_fare = parse_fare(raw_infant_fare, "유아 요금")

                valid_rows.append({
                    "row_index": idx + 2,
                    "date": date_str,
                    "date_end": date_end_str,
                    "adult_air": adult_air,
                    "adult_hotel": adult_hotel,
                    "adult_land": adult_land,
                    "adult_tour": adult_tour,
                    "adult_profit": adult_profit,
                    "child_fare": child_fare,
                    "infant_fare": infant_fare
                })
            except Exception as e:
                print(f"[경고] {idx+2}행 파싱 오류: {str(e)} (건너뜀)")

        print(f"[정보] 유효성 검사 완료. 실행할 대상 행은 총 {len(valid_rows)}개이며, 기간 모드 여부는 {is_period_excel}입니다.")
        return valid_rows, is_period_excel

    except Exception as e:
        print(f"[오류] 엑셀 데이터 처리 중 에러 발생: {str(e)}")
        return None, False


if __name__ == "__main__":
    print("엑셀 선택기 작동 테스트...")
    path = select_excel_file()
    if path:
        print(f"선택된 파일: {path}")
        data, _ = load_and_validate_fares(path)
        if data:
            print("상위 5개 데이터 미리보기:")
            for item in data[:5]:
                print(item)
    else:
        print("선택 취소됨.")
