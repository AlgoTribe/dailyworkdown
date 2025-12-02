import pyautogui
import time
import datetime
import calendar
import os
import glob
import keyboard  # ESC 감지용 (pip install keyboard 필요)

# ========== 사용자 설정 부분 ==========

# ▶ 시작 / 끝 날짜 (원하는 대로 수정)
START_DATE = datetime.date(2025, 9, 1)   # 예: 2025년 4월 14일부터
END_DATE   = datetime.date(2025, 10, 30)  # 예: 2025년 12월 31일까지

# ▶ 다운로드 폴더 경로 (네 PC에 맞게 수정)
DOWNLOAD_DIR = r"C:\Users\USER\Downloads"

# ▶ 달력 그리드 설정
#  - 실제 캘린더: "일월화수목금토" (일요일 시작)
#  - 2025-03-01(토)이 1행 7열 위치, 좌표가 (2595, 388)
#  - 칸 간격 33px → 1행 1열 x = 2595 - 6*33 = 2397
STEP = 33          # 가로/세로 칸 간격(px)
BASE_X = 2397      # 1행 1열 x좌표
BASE_Y = 388       # 1행 1열 y좌표

# ▶ 버튼 좌표들
CALENDAR_BTN    = (2480, 300)  # 달력 여는 버튼
NEXT_MONTH_BTN  = (2595, 322)  # 다음달 이동 버튼
QUERY_BTN       = (3770, 300)  # 조회 버튼
EXCEL_BTN       = (3770, 530)  # 엑셀 내려받기
CSV_BTN         = (3770, 560)  # CSV 버튼
CONFIRM_BTN     = (3680, 590)  # 확인 버튼

# ▶ 각 클릭 후 대기시간 (로딩 고려)
CLICK_SLEEP = 1

pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = True

# ⚠️ 달력: "일요일 시작" 기준
calendar.setfirstweekday(calendar.SUNDAY)  # 1열=일요일, 7열=토요일

# ========== 공통 유틸 ==========

def esc_check():
    """ESC 눌리면 True 리턴"""
    return keyboard.is_pressed('esc')


def click(x, y, sleep=CLICK_SLEEP):
    """좌표 이동 후 클릭 + 대기 (중간에 ESC 체크)"""
    if esc_check():
        raise KeyboardInterrupt("ESC 감지 - 중단")
    pyautogui.moveTo(x, y, duration=0.1)
    pyautogui.click()
    time.sleep(sleep)


# ========== 달력 좌표 관련 함수 ==========

def get_cell_xy(row, col):
    """
    달력 그리드 (row, col) → 실제 화면 좌표 (x, y)
    row : 1~6 (행, 위→아래)
    col : 1~7 (열, 1=일요일, 7=토요일)
    """
    x = BASE_X + (col - 1) * STEP
    y = BASE_Y + (row - 1) * STEP
    return x, y


def get_row_col_for_date(d: datetime.date):
    """
    날짜 d(연/월/일)가 달력 그리드에서 몇 행 몇 열인지 계산.
    calendar.monthcalendar는 setfirstweekday 설정(일요일 시작)을 반영함.
    """
    mc = calendar.monthcalendar(d.year, d.month)  # 최대 6행 × 7열
    for r_idx, week in enumerate(mc):
        if d.day in week:
            c_idx = week.index(d.day)
            # r_idx, c_idx는 0-base → 1-base로 변환
            row = r_idx + 1
            col = c_idx + 1
            return row, col
    raise ValueError(f"달력에서 날짜를 찾을 수 없음: {d}")


# ========== 파일 이름 변경 유틸 ==========

def get_latest_csv_path(folder):
    files = glob.glob(os.path.join(folder, "*.csv"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def rename_latest_csv(current_date: datetime.date):
    """
    다운로드 폴더에서 가장 최근 CSV를 찾아서
    'YYYY.MM.DD.csv' 형식으로 이름 변경
    """
    latest = get_latest_csv_path(DOWNLOAD_DIR)
    if latest is None:
        print(f"[경고] {DOWNLOAD_DIR} 에 csv 파일이 없습니다.")
        return

    new_name = current_date.strftime("%Y.%m.%d") + ".csv"
    new_path = os.path.join(DOWNLOAD_DIR, new_name)

    # 동일 이름 있을 경우 뒤에 _번호 붙이기
    if os.path.abspath(latest) != os.path.abspath(new_path):
        base, ext = os.path.splitext(new_path)
        idx = 1
        while os.path.exists(new_path):
            new_path = f"{base}_{idx}{ext}"
            idx += 1
        os.rename(latest, new_path)
        print(f"[파일이름 변경] {os.path.basename(latest)} → {os.path.basename(new_path)}")
    else:
        print(f"[알림] 이미 이름이 맞는 파일 존재: {os.path.basename(new_path)}")


# ========== 메인 루프 ==========

def run():
    current = START_DATE
    prev_date = None  # 이전 루프에서 처리한 날짜

    try:
        while current <= END_DATE:
            if esc_check():
                print("ESC 감지 → 스크립트 중지합니다.")
                break

            is_weekend = current.weekday() >= 5  # 토(5), 일(6) → 주말

            print(f"\n=== 처리 날짜: {current} (주말: {is_weekend}) ===")

            # 1. 달력 버튼 클릭
            click(*CALENDAR_BTN)

            # 1-1. 월이 바뀌었으면 다음달 버튼 클릭
            if prev_date is not None and (current.year, current.month) != (prev_date.year, prev_date.month):
                print("→ 새로운 달 진입, 다음달 버튼 클릭")
                click(*NEXT_MONTH_BTN)

            # 2. 날짜 칸 클릭 (주말이든 평일이든 달력은 맞춰줌)
            row, col = get_row_col_for_date(current)
            x, y = get_cell_xy(row, col)
            print(f"날짜 클릭 → row={row}, col={col}, xy=({x}, {y})")
            click(x, y)

            if not is_weekend:
                # ===== 평일만 조회/다운로드 수행 =====
                # 3. 조회 버튼 클릭
                click(*QUERY_BTN)

                # 4. 엑셀 내려받기 버튼 클릭
                click(*EXCEL_BTN)

                # 5. CSV 버튼 클릭
                click(*CSV_BTN)

                # 6. 확인 버튼 클릭
                click(*CONFIRM_BTN)

                # 7. CSV 파일 이름 변경
                time.sleep(2)  # 다운로드 여유시간
                if esc_check():
                    print("ESC 감지 → 스크립트 중지합니다.")
                    break
                rename_latest_csv(current)
            else:
                print("→ 주말이므로 조회/다운로드는 건너뜀")

            # 8. 다음 날짜로
            prev_date = current
            current += datetime.timedelta(days=1)

    except KeyboardInterrupt as e:
        print(str(e))

    print("스크립트 종료 완료.")


if __name__ == "__main__":
    # 처음에는 click() 안의 pyautogui.click()을 잠깐 주석 처리하고
    # row/col/xy 출력만 보면서 날짜랑 좌표가 맞는지 먼저 확인하는 걸 추천.
    run()
