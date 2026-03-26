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

class CpStockExpectCur:
    """
    DsCbo1.StockExpectCur 기능을 포함하는 클래스
    설명: 주식의 예상체결 시세(단일가 매매 상태)를 실시간으로 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("DsCbo1.StockExpectCur")
        # 세션 코드 매핑
        self.map_session = {
            '1': '시가 단일가', 
            '2': '장중 단일가(VI 등)', 
            '3': '종가 단일가'
        }

    def subscribe(self, code):
        """실시간 예상체결가 수신 신청"""
        self.obj.SetInputValue(0, code)
        
        # 이벤트 핸들러 연결
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self, code)
        
        self.obj.Subscribe()
        print(f"[{code}] 실시간 예상체결 데이터 구독 시작")

    def unsubscribe(self):
        """실시간 수신 해제"""
        self.obj.Unsubscribe()
        print("예상체결 데이터 구독 해지 완료")

    def process_received(self):
        """데이터 수신 시 호출되는 콜백 메서드"""
        code = self.obj.GetHeaderValue(0)        # 종목코드
        name = self.obj.GetHeaderValue(7)        # 종목명
        time = self.obj.GetHeaderValue(1)        # 시간 (HHMM)
        exp_price = self.obj.GetHeaderValue(2)   # 예상체결가
        exp_vol = self.obj.GetHeaderValue(4)     # 예상체결수량
        
        # 세션 구분 (현재 어떤 상황인지 파악)
        session_code = self.obj.GetHeaderValue(8)
        session_name = self.map_session.get(session_code, "기타")

        # 결과 출력
        print(f"\n[실시간 예상체결] {name}({code}) | {time}")
        print(f"   현재 세션: {session_name}")
        print(f"   예상가: {exp_price:,}원 | 예상수량: {exp_vol:,}주")
        print(f"   매도호가: {self.obj.GetHeaderValue(5):,} | 매수호가: {self.obj.GetHeaderValue(6):,}")
        print("-" * 45)

# --- 사용 예시 ---
if __name__ == "__main__":
    expect_cur = CpStockExpectCur()
    
    # 삼성전자 구독 (장 시작 전 08:30~09:00 사이에 실행하면 데이터가 들어옵니다)
    target_code = "A005930"
    expect_cur.subscribe(target_code)
    
    print("실시간 예상체결 데이터를 수신 중입니다. 종료하려면 Ctrl+C...")
    
    try:
        while True:
            # 이벤트를 수신하기 위해 펌핑 루프 필요
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        expect_cur.unsubscribe()