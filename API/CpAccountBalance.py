import win32com.client

class CpAccountBalance:
    """
    CpTrade.CpTd6033 기능을 포함하는 클래스
    설명: 계좌별 잔고 및 주문체결 평가현황(수익률, 평가손익 등)을 조회합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpTrade.CpTd6033")

    def get_balance_data(self, acc_no, acc_flag, yield_type="2"):
        """
        잔고 및 평가 현황 데이터를 요청하고 (요약 정보, 종목 리스트)를 반환합니다.
        acc_no: 계좌번호
        acc_flag: 상품관리구분코드
        yield_type: "1" (100% 기준), "2" (0% 기준 - 일반적)
        """
        # 1. 입력 데이터 설정
        self.obj.SetInputValue(0, acc_no)       # 계좌번호
        self.obj.SetInputValue(1, acc_flag)     # 상품관리구분코드
        self.obj.SetInputValue(2, 50)           # 요청건수 (최대 50개)
        self.obj.SetInputValue(3, yield_type)   # 수익률 구분 ("2": 0% 기준)
        self.obj.SetInputValue(4, "1")          # 시장 구분 ("1": KRX)

        # 2. 데이터 요청
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None, []

        # 3. 헤더 정보 (계좌 요약) 추출
        summary = {
            'account_name': self.obj.GetHeaderValue(0),      # 계좌명
            'total_eval_amt': self.obj.GetHeaderValue(3),    # 총 평가금액 (예수금 포함)
            'total_profit_loss': self.obj.GetHeaderValue(4), # 총 평가손익
            'total_yield': self.obj.GetHeaderValue(8),       # 총 수익률
            'd2_deposit': self.obj.GetHeaderValue(9),        # D+2 예상 예수금
            'stock_eval_amt': self.obj.GetHeaderValue(11),   # 순수 주식 잔고 평가금액
        }

        # 4. 개별 종목 리스트 추출
        count = self.obj.GetHeaderValue(7) # 수신 종목 개수
        stocks = []
        for i in range(count):
            item = {
                'name': self.obj.GetDataValue(0, i),          # 종목명
                'code': self.obj.GetDataValue(12, i),         # 종목코드
                'total_qty': self.obj.GetDataValue(7, i),     # 체결잔고수량
                'sellable_qty': self.obj.GetDataValue(15, i), # 매도가능수량
                'buy_price': self.obj.GetDataValue(17, i),    # 체결장부단가 (매수평균가)
                'eval_amt': self.obj.GetDataValue(9, i),      # 평가금액
                'profit_loss': self.obj.GetDataValue(10, i),  # 평가손익
                'yield': round(self.obj.GetDataValue(11, i), 2), # 수익률
            }
            stocks.append(item)
            
        return summary, stocks

# --- 사용 예시 ---
if __name__ == "__main__":
    # 주문 초기화(TradeInit)가 완료된 상태여야 합니다.
    # trade_mgr = CpTradeUtil()
    # trade_mgr.trade_init()

    balance_mgr = CpAccountBalance()
    
    # 계좌번호와 상품코드는 본인의 정보를 입력하세요.
    summary, stocks = balance_mgr.get_balance_data("YOUR_ACC_NO", "01")

    if summary:
        print(f"\n===== [{summary['account_name']}] 자산 현황 =====")
        print(f"총 평가금액: {summary['total_eval_amt']:,}원")
        print(f"총 평가손익: {summary['total_profit_loss']:,}원")
        print(f"총 수익률: {summary['total_yield']}%")
        print(f"D+2 예상 예수금: {summary['d2_deposit']:,}원")
        print("-" * 40)

        for s in stocks:
            print(f"{s['name']}({s['code']}) | {s['yield']}% | {s['profit_loss']:,}원 | {s['total_qty']}주 보유")