# account_info.py — 크레온 계좌 정보 조회 (강의용)
# 접속된 상태에서 내 계좌번호와 보유종목(잔고)을 조회해 콘솔에 출력한다.
# ※ 반드시 관리자 권한 파이썬에서, CreonPlus 로그인(연결) 후 실행할 것.

import win32com.client


def main():
    # 1) 연결 확인
    cybos = win32com.client.Dispatch("CpUtil.CpCybos")
    if cybos.IsConnect != 1:
        print("크레온 PLUS에 연결되어 있지 않습니다. 먼저 론처로 접속하세요.")
        return

    # 2) 거래 객체 초기화
    trade = win32com.client.Dispatch("CpTrade.CpTdUtil")
    if trade.TradeInit(0) != 0:
        print("주문 초기화(TradeInit) 실패")
        return

    # 3) 계좌번호 · 상품구분 가져오기
    acc = trade.AccountNumber[0]            # 계좌번호
    acc_flag = trade.GoodsList(acc, 1)      # 주식 상품 구분
    print("계좌번호:", acc)                  # (강의 화면에서는 마스킹)

    # 4) 잔고(보유종목) 조회 객체 — CpTd6033
    obj = win32com.client.Dispatch("CpTrade.CpTd6033")
    obj.SetInputValue(0, acc)               # 계좌번호
    obj.SetInputValue(1, acc_flag[0])       # 상품구분
    obj.SetInputValue(2, 50)                # 요청 건수(최대 50)
    obj.BlockRequest()

    # 5) 통신 상태 확인
    if obj.GetDibStatus() != 0:
        print("조회 실패:", obj.GetDibMsg1())
        return

    # 6) 결과 출력
    cnt = obj.GetHeaderValue(7)             # 보유종목 수
    print(f"보유종목 수: {cnt}\n")

    print("종목명 / 장부단가 / 평가손익")
    for i in range(cnt):
        name      = obj.GetDataValue(0, i)   # 종목명
        buy_price = obj.GetDataValue(17, i)  # 체결장부단가
        eval_pl   = obj.GetDataValue(11, i)  # 평가손익
        print(f"{name}  /  {buy_price:,}원  /  {eval_pl:,}원")


if __name__ == "__main__":
    main()
