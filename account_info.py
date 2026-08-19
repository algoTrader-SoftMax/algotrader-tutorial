# account_info.py — 크레온 계좌 정보 조회 (강의용)
# 접속된 상태에서 평가금액·평가손익·예수금·보유종목을 조회해 콘솔에 출력한다.
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

    acc = trade.AccountNumber[0]            # 계좌번호
    acc_flag = trade.GoodsList(acc, 1)      # 주식 상품 구분

    # 3) 잔고 조회 — CpTd6033 (평가금액·평가손익·보유종목)
    bal = win32com.client.Dispatch("CpTrade.CpTd6033")
    bal.SetInputValue(0, acc)
    bal.SetInputValue(1, acc_flag[0])
    bal.SetInputValue(2, 50)
    bal.BlockRequest()

    eval_amount = bal.GetHeaderValue(3)    # 평가금액
    eval_profit = bal.GetHeaderValue(4)    # 평가손익
    stock_count = bal.GetHeaderValue(7)    # 보유종목 수

    # 4) 예수금 조회 — CpTdNew5331A (D+2 예수금)
    cash = win32com.client.Dispatch("CpTrade.CpTdNew5331A")
    cash.SetInputValue(0, acc)
    cash.SetInputValue(1, acc_flag[0])
    cash.BlockRequest()
    cash_d2 = cash.GetHeaderValue(9)       # 예수금 D+2

    # 5) 계좌 요약 출력 (평가금액 → 평가손익 → 예수금)
    print("=" * 44)
    print("[ 계좌 요약 ]")
    print(f"평가금액 : {eval_amount:,.0f}원")
    print(f"평가손익 : {eval_profit:,.0f}원")
    print(f"예수금(D+2) : {cash_d2:,.0f}원")
    print("=" * 44)

    # 6) 보유종목 출력 (종목명 / 장부단가 / 수익률)
    print(f"\n[ 보유종목 {stock_count}개 ]")
    print("종목명 / 장부단가 / 수익률")
    for i in range(stock_count):
        name      = bal.GetDataValue(0, i)    # 종목명
        buy_price = bal.GetDataValue(17, i)   # 체결장부단가
        ret_rate  = bal.GetDataValue(11, i)   # 수익률(%)
        print(f"{name}  /  {buy_price:,.2f}원  /  {ret_rate:.2f}%")


if __name__ == "__main__":
    main()