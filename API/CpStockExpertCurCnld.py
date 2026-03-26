import win32com.client
import pythoncom

# [공통] 이벤트 핸들러 클래스
class CpEvent:
    def set_params(self, client, name):
        self.client = client
        self.name = name

    def OnReceived(self):
        if hasattr(self.client, 'process_received'):
            self.client.process_received()

class CpStockExpertCurCnld:
    """
    Dscbo1.StockExpertCurCnld 기능을 포함하는 클래스
    설명: 통합(KRX+NXT) 주식의 실시간 예상체결 시세를 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockExpertCurCnld")

    def subscribe(self, code):
        """통합 실시간 예상체결가 수신 신청"""
        self.obj.SetInputValue(0, code)
        
        # 이벤트 핸들러 연결
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self, code)
        
        self.obj.Subscribe()
        print(f"[{code}] 통합 실시간 예상체결 시세 구독 시작")

    def unsubscribe(self):
        """실시간 수신 해제"""
        self.obj.Unsubscribe()
        print("통합 예상체결 시세 구독 해지 완료")

    def process_received(self):
        """데이터 수신 시 호출되는 콜백 메서드"""
        code = self.obj.GetHeaderValue(0)        # 주식코드
        name = self.obj.GetHeaderValue(7)        # 종목명
        time = self.obj.GetHeaderValue(1)        # 시간
        exp_price = self.obj.GetHeaderValue(2)   # 예상체결가
        exp_vol = self.obj.GetHeaderValue(4)     # 예상체결수량
        
        # 집행거래소 구분
        exch_code = self.obj.GetHeaderValue(9)   # 'K': KRX, 'N': NXT
        exch_name = "KRX(한국거래소)" if exch_code == 'K' else "NXT(대체거래소)"
        
        # 세션 코드 (장 시작 전, 장 종료 전 등 구분)
        session = self.obj.GetHeaderValue(8)

        # 결과 출력
        print(f"\n[통합 예상체결] {name}({code}) | {time}")
        print(f"   집행소: {exch_name} | 세션: {session}")
        print(f"   예상가: {exp_price:,}원 | 예상수량: {exp_vol:,}주")
        print(f"   매도호가: {self.obj.GetHeaderValue(5):,} | 매수호가: {self.obj.GetHeaderValue(6):,}")
        print("-" * 50)

# --- 사용 예시 ---
if __name__ == "__main__":
    expert_cur = CpStockExpertCurCnld()
    
    # 동시호가 시간(08:30~09:00, 15:20~15:30)에 테스트하는 것이 가장 정확합니다.
    target_code = "A005930" 
    expert_cur.subscribe(target_code)
    
    print("통합 예상체결 데이터를 감시 중입니다. (종료: Ctrl+C)")
    
    try:
        while True:
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        expert_cur.unsubscribe()