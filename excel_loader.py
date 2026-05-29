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

def load_and_validate_fares(file_path, history_log_path=None):
    """
    선택된 엑셀 파일을 읽고 요금 데이터를 파싱 및 유효성 검사합니다.
    history_log_path가 주어지면 이전 SUCCESS 이력 날짜를 자동 스킵합니다.
    """
    if not file_path:
        print("[오류] 파일이 선택되지 않았습니다.")
        return None

    if not os.path.exists(file_path):
        print(f"[오류] 파일이 존재하지 않습니다: {file_path}")
        return None

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
            return None

        print(f"[정보] 엑셀 로딩 완료. 총 {len(df)}개의 행을 감지했습니다.")

        valid_rows = []
        for idx, row in df.iterrows():
            try:
                raw_date = row.iloc[0]
                raw_adult = row.iloc[1]
                raw_child = row.iloc[2]

                # 날짜 유효성 검사
                if pd.isna(raw_date):
                    continue

                if isinstance(raw_date, pd.Timestamp):
                    date_str = raw_date.strftime('%Y-%m-%d')
                else:
                    date_str = str(raw_date).strip().split(' ')[0]
                    if len(date_str) == 8 and date_str.isdigit():
                        date_str = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"

                # 이미 성공한 날짜인지 검사하여 자동 스킵
                if date_str in completed_dates:
                    print(f"[이어서 진행] {idx+2}행: 날짜 {date_str}는 이미 이전 실행에서 성공하여 스킵합니다.")
                    continue

                # 요금 유효성 검사
                if pd.isna(raw_adult) or pd.isna(raw_child):
                    print(f"[경고] {idx+2}행: 요금이 누락되었습니다. (건너뜀)")
                    continue

                adult_fare = int(float(str(raw_adult).replace(',', '').strip()))
                child_fare = int(float(str(raw_child).replace(',', '').strip()))

                if adult_fare < 0 or child_fare < 0:
                    print(f"[경고] {idx+2}행: 요금은 음수일 수 없습니다. (건너뜀)")
                    continue

                valid_rows.append({
                    "row_index": idx + 2,
                    "date": date_str,
                    "adult_fare": adult_fare,
                    "child_fare": child_fare
                })
            except Exception as e:
                print(f"[경고] {idx+2}행 파싱 오류: {str(e)} (건너뜀)")

        print(f"[정보] 유효성 검사 완료. 실행할 대상 행은 총 {len(valid_rows)}개입니다.")
        return valid_rows

    except Exception as e:
        print(f"[오류] 엑셀 데이터 처리 중 에러 발생: {str(e)}")
        return None


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

