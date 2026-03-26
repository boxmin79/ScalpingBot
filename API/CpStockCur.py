import win32com.client
import pythoncom

# [공통] 이벤트 핸들러 클래스 (기존에 만든 것과 동일)
class CpEvent:
    def set_params(self, client, name):
        self.client = client
        self.name = name

    def OnReceived(self):
        # 데이터 수신 시 클라이언트의 process_received 호출
        if hasattr(self.client, 'process_received'):
            self.client.process_received()

class CpStockCur:
    """
    Dscbo1.StockCur 기능을 포함하는 클래스
    설명: 주식/ELW의 체결 데이터를 실시간으로 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockCur")
        
        # 대비부호 매핑
        self.map_diff_status = {
            '1': '상한', '2': '상승', '3': '보합', '4': '하한', '5': '하락',
            '6': '기세상한', '7': '기세상승', '8': '기세하한', '9': '기세하락'
        }

    def subscribe(self, code):
        """특정 종목의 실시간 시세 수신 신청"""
        self.obj.SetInputValue(0, code)
        
        # 이벤트 핸들러 연결 (win32com.client.WithEvents 사용)
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self, code)
        
        self.obj.Subscribe()
        print(f"[{code}] 실시간 시세 구독 시작")

    def unsubscribe(self):
        """실시간 시세 수신 해제"""
        self.obj.Unsubscribe()
        print("실시간 시세 구독 해지 완료")

    def process_received(self):
        """데이터 수신 시 호출되는 콜백 메서드"""
        # 헤더 데이터 추출
        code = self.obj.GetHeaderValue(0)        # 종목코드
        name = self.obj.GetHeaderValue(1)        # 종목명
        price = self.obj.GetHeaderValue(13)      # 현재가 또는 예상체결가
        vol = self.obj.GetHeaderValue(9)         # 누적거래량
        instant_vol = self.obj.GetHeaderValue(17) # 순간체결수량
        
        # 시간 처리 (HHMM + SS)
        hhmm = self.obj.GetHeaderValue(3)
        ss = self.obj.GetHeaderValue(18)
        
        # 상태 플래그
        price_flag = self.obj.GetHeaderValue(19)  # '1':예상체결가, '2':장중체결
        market_flag = self.obj.GetHeaderValue(20) # '1':장전, '2':장중, '5':장후 등
        diff_flag = self.obj.GetHeaderValue(22)   # 대비부호
        
        status_text = "장중" if market_flag == '2' else "예상"
        diff_text = self.map_diff_status.get(diff_flag, diff_flag)

        # 결과 출력
        print(f"[{status_text}] {name}({code}) {hhmm:04d}{ss:02d}")
        print(f"   현재가: {price:,} | 대비: {diff_text} | 순간체결: {instant_vol} | 누적거래: {vol:,}")
        print("-" * 40)

# --- 사용 예시 ---
if __name__ == "__main__":
    # 1. 클래스 생성
    cur = CpStockCur()
    
    # 2. 삼성전자(A005930) 구독
    cur.subscribe("A005930")
    
    print("실시간 데이터를 수신 중입니다. 종료하려면 Ctrl+C를 누르세요.")
    
    try:
        while True:
            # 3. 윈도우 메시지 펌핑 (이벤트를 수신하기 위해 필수)
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        cur.unsubscribe()