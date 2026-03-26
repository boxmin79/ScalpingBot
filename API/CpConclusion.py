import win32com.client
import pythoncom

# 1. 공통 이벤트 핸들러 (이미 만드셨다면 재사용 가능)
class CpEvent:
    def set_params(self, client, name):
        self.client = client
        self.name = name

    def OnReceived(self):
        # 데이터 수신 시 클라이언트의 process_received 호출
        if hasattr(self.client, 'process_received'):
            self.client.process_received()

# 2. 실시간 체결 수신 클래스
class CpConclusion:
    """
    Dscbo1.CpConclusion 기능을 포함하는 클래스
    설명: 내 계좌에서 발생한 주문의 접수/체결/거부 내역을 실시간으로 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.CpConclusion")
        
        # 상태 코드 매핑 (가독성을 위해)
        self.map_side = {'1': '매도', '2': '매수'}
        self.map_type = {'1': '체결', '2': '확인', '3': '거부', '4': '접수', '5': '접수대기'}
        self.map_order_kind = {'1': '정상', '2': '정정', '3': '취소'}

    def subscribe(self):
        """실시간 수신 신청"""
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self, "CpConclusion")
        self.obj.Subscribe()
        print("--- 주식 체결 실시간 모니터링 시작 ---")

    def unsubscribe(self):
        """실시간 수신 해제"""
        self.obj.Unsubscribe()
        print("--- 주식 체결 실시간 모니터링 종료 ---")

    def process_received(self):
        """데이터 수신 시 실행되는 로직"""
        # 헤더에서 데이터 추출
        acc_name = self.obj.GetHeaderValue(1)    # 계좌명
        stock_name = self.obj.GetHeaderValue(2)  # 종목명
        amount = self.obj.GetHeaderValue(3)      # 체결수량
        price = self.obj.GetHeaderValue(4)       # 체결가격
        order_no = self.obj.GetHeaderValue(5)    # 주문번호
        orignal_no = self.obj.GetHeaderValue(6)  # 원주문번호
        
        stock_code = self.obj.GetHeaderValue(9)  # 종목코드
        
        # 코드값 변환
        side_code = self.obj.GetHeaderValue(12)  # 매매구분
        side = self.map_side.get(side_code, side_code)
        
        exec_code = self.obj.GetHeaderValue(14)  # 체결구분
        status = self.map_type.get(exec_code, exec_code)
        
        kind_code = self.obj.GetHeaderValue(16)  # 정정취소구분
        kind = self.map_order_kind.get(kind_code, kind_code)

        # 결과 출력 (또는 DB 저장/알림 로직)
        print(f"\n[알림] 주문/체결 이벤트 발생!")
        print(f"상태: {status} ({kind}) | {side} | {stock_name}({stock_code})")
        print(f"수량: {amount} / 가격: {price} / 주문번호: {order_no}")
        print(f"계좌: {acc_name}")
        print("-" * 30)

# --- 실행 예시 ---
if __name__ == "__main__":
    # 이 클래스를 사용하기 위해선 로그인이 되어 있어야 하며, 
    # 실제 주문을 넣었을 때 이벤트가 발생합니다.
    
    watcher = CpConclusion()
    watcher.subscribe()

    print("체결 대기 중... (종료하려면 Ctrl+C)")
    
    try:
        while True:
            # 실시간 이벤트를 수신하기 위해 메시지 루프 필요
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        watcher.unsubscribe()