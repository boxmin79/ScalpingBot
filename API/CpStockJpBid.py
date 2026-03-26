import win32com.client
import pythoncom

# [공통] 이벤트 핸들러 클래스
class CpEvent:
    def set_params(self, client, name):
        self.client = client
        self.name = name

    def OnReceived(self):
        # 데이터 수신 시 클라이언트의 process_received 호출
        if hasattr(self.client, 'process_received'):
            self.client.process_received()

class CpStockJpBid:
    """
    Dscbo1.StockJpBid 기능을 포함하는 클래스
    설명: 주식/ETF/ELW의 1차~10차 호가 및 LP 호가 잔량을 실시간 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockJpBid")

    def subscribe(self, code):
        """실시간 호가 수신 신청"""
        self.obj.SetInputValue(0, code)
        
        # 이벤트 핸들러 연결
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self, code)
        
        self.obj.Subscribe()
        print(f"[{code}] 실시간 호가(LP포함) 구독 시작")

    def unsubscribe(self):
        """실시간 호가 수신 해제"""
        self.obj.Unsubscribe()
        print("실시간 호가 구독 해지 완료")

    def process_received(self):
        """데이터 수신 시 호출되는 콜백 메서드"""
        code = self.obj.GetHeaderValue(0)
        time = self.obj.GetHeaderValue(1)
        
        # 1. 일반 호가 데이터 (1차~10차)
        # 인덱스 규칙: 1~5차(3~22), 6~10차(27~46)
        quotes = []
        for i in range(1, 11):
            if i <= 5:
                base = (i - 1) * 4 + 3
            else:
                base = (i - 6) * 4 + 27
            
            quotes.append({
                'level': i,
                'ask_p': self.obj.GetHeaderValue(base),     # 매도호가
                'bid_p': self.obj.GetHeaderValue(base + 1), # 매수호가
                'ask_r': self.obj.GetHeaderValue(base + 2), # 매도잔량
                'bid_r': self.obj.GetHeaderValue(base + 3)  # 매수잔량
            })

        # 2. 총 잔량 및 기타
        total_ask_rem = self.obj.GetHeaderValue(23)
        total_bid_rem = self.obj.GetHeaderValue(24)
        mid_price = self.obj.GetHeaderValue(69) # 중간가격

        # 3. LP 잔량 (ELW 거래 시 중요)
        # 인덱스 47번부터 10차까지 순차적
        lp_quotes = []
        for i in range(1, 11):
            base = (i - 1) * 2 + 47
            lp_quotes.append({
                'level': i,
                'lp_ask_r': self.obj.GetHeaderValue(base),
                'lp_bid_r': self.obj.GetHeaderValue(base + 1)
            })

        # 결과 출력 (최우선 호가와 LP 합계 위주)
        print(f"\n[실시간 호가] {code} | 시각: {time}")
        print(f"   1차 매도: {quotes[0]['ask_p']:,} ({quotes[0]['ask_r']:,}) <LP: {lp_quotes[0]['lp_ask_r']:,}>")
        print(f"   1차 매수: {quotes[0]['bid_p']:,} ({quotes[0]['bid_r']:,}) <LP: {lp_quotes[0]['lp_bid_r']:,}>")
        print(f"   총 잔량: 매도 {total_ask_rem:,} / 매수 {total_bid_rem:,} | 중간가: {mid_price:,}")
        print("-" * 50)

# --- 사용 예시 ---
if __name__ == "__main__":
    jp_bid = CpStockJpBid()
    
    # ELW나 거래량이 많은 주식 종목코드 입력
    target_code = "A005930" 
    jp_bid.subscribe(target_code)
    
    try:
        while True:
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        jp_bid.unsubscribe()