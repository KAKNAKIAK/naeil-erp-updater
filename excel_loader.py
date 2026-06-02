# -*- coding: utf-8 -*-
import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def select_excel_file():
    """
    윈도우 파일 대화상자를 띄워 사용자가 엑셀 파일을 선택하게 합니다.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    file_path = filedialog.askopenfilename(
        title="요금 업데이트 엑셀 파일 선택",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    root.destroy()
    return file_path

def _normalize_dt(raw_dt):
    if pd.isna(raw_dt):
        return None
    if isinstance(raw_dt, pd.Timestamp):
        return raw_dt.strftime('%Y-%m-%d')
    dt_str = str(raw_dt).strip().split(' ')[0]
    dt_str = dt_str.replace('/', '-').replace('.', '-').strip('-')
    if len(dt_str) == 8 and dt_str.isdigit():
        return f"{dt_str[:4]}-{dt_str[4:6]}-{dt_str[6:]}"
    parts = dt_str.split('-')
    if len(parts) == 3:
        try:
            y, m, d = parts
            if len(y) == 4:
                return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
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
                    raw_adult_air = row.iloc[2] if row_len > 2 else 0
                    raw_adult_hotel = row.iloc[3] if row_len > 3 else 0
                    raw_adult_land = row.iloc[4] if row_len > 4 else 0
                    raw_adult_tour = row.iloc[5] if row_len > 5 else 0
                    raw_adult_profit = row.iloc[6] if row_len > 6 else 0
                    raw_child_fare = row.iloc[7] if row_len > 7 else 0
                    raw_infant_fare = row.iloc[8] if row_len > 8 else 0
                else:
                    raw_date_end = raw_date
                    raw_adult_air = row.iloc[1] if row_len > 1 else 0
                    # 호환성 지원: 엑셀 열이 3개뿐인 경우(기존 v1 3열 양식: 날짜, 항공비, 알선수익)
                    if row_len == 3:
                        raw_adult_hotel = 0
                        raw_adult_land = 0
                        raw_adult_tour = 0
                        raw_adult_profit = row.iloc[2]
                        raw_child_fare = 0
                        raw_infant_fare = 0
                    else:
                        raw_adult_hotel = row.iloc[2] if row_len > 2 else 0
                        raw_adult_land = row.iloc[3] if row_len > 3 else 0
                        raw_adult_tour = row.iloc[4] if row_len > 4 else 0
                        raw_adult_profit = row.iloc[5] if row_len > 5 else 0
                        raw_child_fare = row.iloc[6] if row_len > 6 else 0
                        raw_infant_fare = row.iloc[7] if row_len > 7 else 0

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

                # 각 요금 필드 안전하게 파싱 (기본값 0)
                def parse_fare(val, label):
                    if pd.isna(val):
                        return 0
                    try:
                        f_val = int(float(str(val).replace(',', '').strip()))
                        if f_val < 0:
                            print(f"[정보] {idx+2}행: {label} 요금이 음수({f_val})이므로 0원으로 대체합니다.")
                            return 0
                        return f_val
                    except (ValueError, TypeError):
                        print(f"[정보] {idx+2}행: {label} 요금 오류값({val})이 입력되어 0원으로 대체합니다.")
                        return 0

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


def filter_fares_by_date(fares_data, mode="ALL", value=""):
    """
    날짜 필터를 적용하여 처리 대상 행만 반환합니다.

    mode:
        "ALL"        — 전체 행 (필터 없음)
        "FROM_DATE"  — value 이후 날짜만 (value 포함)
        "SPECIFIC"   — 콤마 구분 날짜 목록에 해당하는 행만
        "DATE_RANGE" — 물결표(~) 구분 시작일~종료일 범위 내 행만

    value:
        FROM_DATE 일 때: "2025-06-01"
        SPECIFIC 일 때: "2025-01-01, 2025-01-05, 2025-03-10"
        DATE_RANGE 일 때: "2026-06-01~2026-06-10" 또는 "20260601~20260610"
    """
    if not fares_data:
        return fares_data

    if mode == "ALL" or not value.strip():
        return fares_data

    if mode == "FROM_DATE":
        from_date = value.strip().replace('/', '-').replace('.', '-')
        # YYYYMMDD 변환
        if len(from_date) == 8 and from_date.isdigit():
            from_date = f"{from_date[:4]}-{from_date[4:6]}-{from_date[6:]}"
        filtered = [r for r in fares_data if r["date"] >= from_date]
        print(f"[날짜 필터] 시작일 {from_date} 이후 → {len(filtered)}건 대상")
        return filtered

    if mode == "SPECIFIC":
        raw_dates = [d.strip().replace('/', '-').replace('.', '-') for d in value.split(",") if d.strip()]
        target_dates = set()
        for d in raw_dates:
            if len(d) == 8 and d.isdigit():
                target_dates.add(f"{d[:4]}-{d[4:6]}-{d[6:]}")
            else:
                target_dates.add(d)
        filtered = [r for r in fares_data if r["date"] in target_dates]
        print(f"[날짜 필터] 지정 날짜 {len(target_dates)}개 중 → {len(filtered)}건 대상")
        return filtered

    if mode == "DATE_RANGE":
        parts = []
        if "~" in value:
            parts = value.split("~")
        elif " - " in value: # YYYY-MM-DD - YYYY-MM-DD
            parts = value.split(" - ")
        else:
            # 공백 구분 시도 (두 단어인 경우)
            sp = value.split()
            if len(sp) == 2:
                parts = sp

        if len(parts) >= 2:
            start_raw = parts[0].strip().replace('/', '-').replace('.', '-')
            end_raw = parts[1].strip().replace('/', '-').replace('.', '-')
            # YYYYMMDD 변환
            if len(start_raw) == 8 and start_raw.isdigit():
                start_raw = f"{start_raw[:4]}-{start_raw[4:6]}-{start_raw[6:]}"
            if len(end_raw) == 8 and end_raw.isdigit():
                end_raw = f"{end_raw[:4]}-{end_raw[4:6]}-{end_raw[6:]}"
            
            filtered = [r for r in fares_data if start_raw <= r["date"] <= end_raw]
            print(f"[날짜 필터] 기간 지정 {start_raw} ~ {end_raw} → {len(filtered)}건 대상")
            return filtered
        else:
            print(f"[경고] 기간 범위 형식이 올바르지 않습니다: {value} (구분자 ~ 필요)")
            return []

    return fares_data


def print_summary_report(history_log_path="logs/update_history.csv"):
    """
    작업 이력 로그(CSV)를 분석하여 성공 및 실패(단발성/연속) 통계를 출력합니다.
    """
    if not os.path.exists(history_log_path):
        print(f"[보고서] 실행 이력 파일이 없습니다: {history_log_path}")
        return

    try:
        df = pd.read_csv(history_log_path, encoding='utf-8')
        if df.empty or 'status' not in df.columns or 'date' not in df.columns:
            print("[보고서] 분석할 실행 이력이 충분하지 않습니다.")
            return

        df['date'] = df['date'].astype(str).str.strip()
        df['status'] = df['status'].astype(str).str.strip()
        
        # 정렬 (날짜 순)
        df = df.sort_values(by='date').reset_index(drop=True)

        total_cnt = len(df['date'].unique())
        success_dates = set(df[df['status'] == 'SUCCESS']['date'].unique())
        failed_df = df[df['status'] != 'SUCCESS'].copy()
        
        # 실패/스킵 정보 추출
        # 날짜별 최종 상태 판단 (마지막 상태 기준)
        date_status = {}
        for idx, row in df.iterrows():
            date_status[row['date']] = (row['status'], row.get('error_message', ''))

        all_dates = sorted(list(date_status.keys()))
        success_cnt = sum(1 for d in all_dates if date_status[d][0] == 'SUCCESS')
        fail_cnt = len(all_dates) - success_cnt

        # 연속 실패 및 단발성 실패 분석
        single_failures = []
        consecutive_failures = [] # list of dict: {'start', 'end', 'reason'}

        current_streak = []
        
        for i, d in enumerate(all_dates):
            status, err = date_status[d]
            if status != 'SUCCESS':
                current_streak.append((d, err))
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

        print("\n========================================================")
        print("                 요금 업데이트 결과 보고서")
        print("========================================================")
        
        print("\n[단발성 실패]")
        if single_failures:
            for d, err in single_failures:
                err_msg = str(err).strip() if pd.notna(err) and str(err).strip() else "알 수 없는 오류"
                print(f"- {d} (오류: {err_msg})")
        else:
            print("- 없음")

        print("\n[연속 실패]")
        if consecutive_failures:
            for streak in consecutive_failures:
                start_date = streak[0][0]
                end_date = streak[-1][0]
                days = len(streak)
                # 에러 메시지 빈도 수 집계로 대표 에러 메시지 결정
                err_msgs = [str(x[1]).strip() for x in streak if pd.notna(x[1]) and str(x[1]).strip()]
                rep_err = max(set(err_msgs), key=err_msgs.count) if err_msgs else "알 수 없는 오류"
                print(f"- {start_date} ~ {end_date} ({days}일간 연속 실패 / 대표오류: {rep_err})")
        else:
            print("- 없음")

        print("\n[요약 통계]")
        print(f"- 전체 대상: {len(all_dates)}일")
        print(f"- 성공: {success_cnt}일")
        print(f"- 실패/스킵: {fail_cnt}일")
        print("========================================================\n")

    except Exception as e:
        print(f"[경고] 결과 보고서 생성 중 오류 발생: {str(e)}")


if __name__ == "__main__":
    print("엑셀 선택기 작동 테스트...")
    path = select_excel_file()
    if path:
        print(f"선택된 파일: {path}")
        data = load_and_validate_fares(path)
        if data:
            print("상위 5개 데이터 미리보기:")
            for item in data[:5]:
                print(item)
    else:
        print("선택 취소됨.")

