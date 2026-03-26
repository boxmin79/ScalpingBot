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

class CpStockBsccnsCnld:
    """
    Dscbo1.StockBsccnsCnld 기능을 포함하는 클래스
    설명: 통합(KRX+NXT) 주식/업종/ELW의 체결 시세를 실시간으로 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockBsccnsCnld")
        
        # 대비부호(22번) 매핑
        self.map_diff_status = {
            '1': '상한', '2': '상승', '3': '보합', '4': '하한', '5': '하락',
            '6': '기세상한', '7': '기세상승', '8': '기세하한', '9': '기세하락'
        }

    def subscribe(self, code):
        """통합 실시간 시세 수신 신청"""
        self.obj.SetInputValue(0, code)
        
        # 이벤트 핸들러 연결
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self, code)
        
        self.obj.Subscribe()
        print(f"[{code}] 통합 실시간 시세(KRX+NXT) 구독 시작")

    def unsubscribe(self):
        """실시간 시세 수신 해제"""
        self.obj.Unsubscribe()
        print("통합 실시간 시세 구독 해지 완료")

    def process_received(self):
        """데이터 수신 시 호출되는 콜백 메서드"""
        code = self.obj.GetHeaderValue(0)        # 종목코드
        name = self.obj.GetHeaderValue(1)        # 종목명
        price = self.obj.GetHeaderValue(13)      # 현재가
        instant_vol = self.obj.GetHeaderValue(17) # 순간체결수량
        
        # 집행거래소 구분 (중요!)
        exch_code = self.obj.GetHeaderValue(29)  # 'K': KRX, 'N': NXT
        exch_name = "한국거래소(KRX)" if exch_code == 'K' else "대체거래소(NXT)"
        
        # 시간 정보
        hhmm = self.obj.GetHeaderValue(3)
        ss = self.obj.GetHeaderValue(18)
        
        # 대비 정보
        diff_flag = self.obj.GetHeaderValue(22)
        diff_text = self.map_diff_status.get(diff_flag, diff_flag)

        # 결과 출력
        print(f"\n[통합 체결] {name}({code}) | {hhmm:04d}{ss:02d}")
        print(f"   체결소: {exch_name}")
        print(f"   현재가: {price:,} | 대비: {diff_text} | 순간체결: {instant_vol:,}")
        print("-" * 45)

# --- 사용 예시 ---
if __name__ == "__main__":
    # 통합 시세 모듈 생성
    cnld_cur = CpStockBsccnsCnld()
    
    # 통합 거래 대상 종목(예: 삼성전자) 구독
    target_code = "A005930"
    cnld_cur.subscribe(target_code)
    
    print("통합 실시간 데이터를 수신 중입니다. 종료하려면 Ctrl+C...")
    
    try:
        while True:
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        cnld_cur.unsubscribe()