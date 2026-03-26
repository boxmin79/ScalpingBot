import win32com.client

class CpDayTradeProfitLoss:
    """
    CpTrade.CpTd6032 기능을 포함하는 클래스
    설명: 체결 기준으로 주식 당일 매매 손익 데이터를 요청하고 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpTrade.CpTd6032")

    def get_profit_loss_data(self, acc_no, acc_flag):
        """
        당일 매매 손익 데이터를 조회하여 (요약 정보, 종목별 상세)를 반환합니다.
        acc_no: 계좌번호
        acc_flag: 상품관리구분코드
        """
        # 1. 입력 데이터 설정
        self.obj.SetInputValue(0, acc_no)       # 계좌번호
        self.obj.SetInputValue(1, acc_flag)     # 상품관리구분코드
        self.obj.SetInputValue(2, "1")          # 거래소구분: "1" KRX

        # 2. 데이터 요청
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None, []

        # 3. 헤더 정보 (당일 전체 요약) 추출
        # 단위 주의: 헤더의 손익 금액은 '천원' 단위입니다.
        summary = {
            'req_count': self.obj.GetHeaderValue(0),         # 조회 요청건수
            'total_eval_pl': self.obj.GetHeaderValue(1),     # 잔량평가손익금액 (단위: 천원)
            'total_realized_pl': self.obj.GetHeaderValue(2), # 매도실현손익금액 (단위: 천원)
            'total_yield': self.obj.GetHeaderValue(3),       # 총 수익률 (float)
        }

        # 4. 종목별 상세 내역 추출
        count = self.obj.GetHeaderValue(0)
        results = []
        for i in range(count):
            item = {
                'name': self.obj.GetDataValue(0, i),          # 종목명
                'prev_balance': self.obj.GetDataValue(2, i),  # 전일잔고
                'buy_qty': self.obj.GetDataValue(3, i),       # 금일매수수량
                'sell_qty': self.obj.GetDataValue(4, i),      # 금일매도수량
                'current_balance': self.obj.GetDataValue(5, i),# 금일잔고
                'avg_buy_price': self.obj.GetDataValue(6, i), # 평균매입단가
                'avg_sell_price': self.obj.GetDataValue(7, i),# 평균매도단가
                'current_price': self.obj.GetDataValue(8, i), # 현재가
                'eval_pl': self.obj.GetDataValue(9, i),       # 잔량평가손익
                'realized_pl': self.obj.GetDataValue(10, i),  # 매도실현손익
                'yield': self.obj.GetDataValue(11, i),        # 수익률(%) (float)
                'code': self.obj.GetDataValue(12, i),         # 종목코드
            }
            results.append(item)
            
        return summary, results

# --- 사용 예시 ---
if __name__ == "__main__":
    # 이 클래스는 주문 초기화(TradeInit)가 선행되어야 합니다.
    # trade_mgr = CpTradeUtil()
    # trade_mgr.trade_init()

    pl_mgr = CpDayTradeProfitLoss()
    
    # 본인의 계좌 정보를 입력하세요.
    summary, details = pl_mgr.get_profit_loss_data("YOUR_ACC_NO", "01")

    if summary:
        print(f"\n===== 오늘 매매 성과 요약 =====")
        print(f"매도 실현 손익: {summary['total_realized_pl']} (천원)")
        print(f"보유 잔량 평가: {summary['total_eval_pl']} (천원)")
        print(f"오늘의 총 수익률: {summary['total_yield']}%")
        print("-" * 35)

        for d in details:
            print(f"[{d['name']}]")
            print(f"  실현손익: {d['realized_pl']}원 / 평가손익: {d['eval_pl']}원 / 수익률: {d['yield']}%")
            print(f"  매수: {d['buy_qty']}주 / 매도: {d['sell_qty']}주 / 현재고: {d['current_balance']}주")