# creon_launcher.py — 크레온 자동 접속 론처 (강의용)
# CreonPlus를 자동 실행하고, 로그인 → 연결까지 확인한다.
# ※ 반드시 관리자 권한으로 실행되는 파이썬에서 돌려야 한다.

import time
import ctypes
from datetime import datetime

from pywinauto import application
import win32com.client

# ────────────────────────────────────────────────
# 본인 정보를 입력하세요 (따옴표 안에)
CREON_ID      = "본인 아이디"
CREON_PW      = "본인 비밀번호"
CREON_PW_CERT = "공인인증서 비밀번호"
# ────────────────────────────────────────────────

# coStarter 경로 (크레온 기본 설치 경로)
COSTARTER = r"C:\CREON\STARTER\coStarter.exe"

# 연결 대기 최대 시간(초)
CONNECT_TIMEOUT = 180


def is_admin() -> bool:
    "관리자 권한으로 실행 중인지 확인"
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def is_connected() -> bool:
    "CreonPlus가 연결되어 있는지 확인"
    try:
        cybos = win32com.client.Dispatch("CpUtil.CpCybos")
        return bool(cybos.IsConnect)
    except Exception:
        return False


def main():
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] 크레온 론처 시작...")

    # 1) 관리자 권한 확인
    if not is_admin():
        print("관리자 권한으로 실행해야 합니다.")
        return

    # 2) 이미 연결돼 있으면 그대로 종료
    if is_connected():
        print("이미 연결되어 있습니다.")
        return

    # 3) coStarter로 CreonPlus 자동 실행 + 로그인
    cmd = (
        f"{COSTARTER} /prj:cp "
        f"/id:{CREON_ID} /pwd:{CREON_PW} /pwdcert:{CREON_PW_CERT} /autostart"
    )
    application.Application().start(cmd)
    print("CreonPlus 실행 중... 연결을 기다립니다.")

    # 4) 연결될 때까지 1초마다 확인 (최대 CONNECT_TIMEOUT초)
    end = time.time() + CONNECT_TIMEOUT
    while time.time() < end:
        if is_connected():
            print("크레온 PLUS 연결 성공!")
            return
        time.sleep(1)

    print("연결 실패 (시간 초과).")


if __name__ == "__main__":
    main()
