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

class CpStockJpBidCnld:
    """
    Dscbo1.StockJpBidCnld 기능을 포함하는 클래스
    설명: 통합(KRX+NXT) 주식/ETF/ELW의 10차 호가 및 LP 잔량을 실시간 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockJpBidCnld")

    def subscribe(self, code):
        """통합 실시간 호가 수신 신청"""
        self.obj.SetInputValue(0, code)
        
        # 이벤트 핸들러 연결
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self, code)
        
        self.obj.Subscribe()
        print(f"[{code}] 통합 실시간 호가(KRX+NXT) 구독 시작")

    def unsubscribe(self):
        """실시간 수신 해제"""
        self.obj.Unsubscribe()
        print("통합 실시간 호가 구독 해지 완료")

    def process_received(self):
        """데이터 수신 시 호출되는 콜백 메서드"""
        # 1. 통합(Total) 10차 호가 데이터 (인덱스 0~46)
        total_quotes = []
        for i in range(1, 11):
            base = (i - 1) * 4 + 3 if i <= 5 else (i - 6) * 4 + 27
            total_quotes.append({
                'level': i,
                'ask_p': self.obj.GetHeaderValue(base),     # 매도호가
                'bid_p': self.obj.GetHeaderValue(base + 1), # 매수호가
                'ask_r': self.obj.GetHeaderValue(base + 2), # 매도잔량
                'bid_r': self.obj.GetHeaderValue(base + 3)  # 매수잔량
            })

        # 2. KRX 전용 호가 잔량 (인덱스 69~88)
        krx_remains = []
        for i in range(1, 11):
            base = (i - 1) * 2 + 69
            krx_remains.append({
                'level': i,
                'ask_r': self.obj.GetHeaderValue(base),
                'bid_r': self.obj.GetHeaderValue(base + 1)
            })

        # 3. NXT 전용 호가 잔량 (인덱스 94~113)
        nxt_remains = []
        for i in range(1, 11):
            base = (i - 1) * 2 + 94
            nxt_remains.append({
                'level': i,
                'ask_r': self.obj.GetHeaderValue(base),
                'bid_r': self.obj.GetHeaderValue(base + 1)
            })

        # 4. 요약 정보
        code = self.obj.GetHeaderValue(0)
        time = self.obj.GetHeaderValue(1)
        total_ask = self.obj.GetHeaderValue(23)
        total_bid = self.obj.GetHeaderValue(24)
        
        # 결과 출력 (통합 1차 호가와 거래소별 비중 출력 예시)
        print(f"\n[통합 실시간 호가] {code} | 시각: {time}")
        print(f"   통합 1차: 매도 {total_quotes[0]['ask_p']:,} / 매수 {total_quotes[0]['bid_p']:,}")
        print(f"   거래소별 1차 잔량 비교:")
        print(f"      KRX - 매도: {krx_remains[0]['ask_r']:,} / 매수: {krx_remains[0]['bid_r']:,}")
        print(f"      NXT - 매도: {nxt_remains[0]['ask_r']:,} / 매수: {nxt_remains[0]['bid_r']:,}")
        print(f"   통합 총잔량: 매도 {total_ask:,} / 매수 {total_bid:,}")
        print("-" * 50)

# --- 사용 예시 ---
if __name__ == "__main__":
    cnld_bid = CpStockJpBidCnld()
    
    # 통합 거래가 지원되는 종목 입력
    target_code = "A005930" 
    cnld_bid.subscribe(target_code)
    
    try:
        while True:
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        cnld_bid.unsubscribe()