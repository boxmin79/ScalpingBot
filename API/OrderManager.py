import win32com.client
import pythoncom

# 1. 실시간 이벤트를 수신할 핸들러 클래스
class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        # 이벤트 발생 시 OrderManager의 로직 호출
        self.client.process_conclusion()

class OrderManager:
    """
    주식 주문(신규, 정정, 취소) 및 실시간 체결 확인을 담당하는 클래스
    """
    def __init__(self):
        # 주문 관련 오브젝트
        self.obj_new_order = win32com.client.Dispatch("CpTrade.CpTd0311")
        self.obj_modify_order = win32com.client.Dispatch("CpTrade.CpTd0313")
        self.obj_cancel_order = win32com.client.Dispatch("CpTrade.CpTd0314")
        
        # 실시간 체결 오브젝트
        self.obj_conclusion = win32com.client.Dispatch("Dscbo1.CpConclusion")
        
        # 상태값 매핑 (영문)
        self.map_side = {'1': 'SELL', '2': 'BUY'}
        self.map_status = {'1': 'CONCLUDED', '2': 'CONFIRMED', '3': 'REJECTED', '4': 'ACCEPTED', '5': 'PENDING'}
        self.map_order_kind = {'1': 'NORMAL', '2': 'MODIFY', '3': 'CANCEL'}

        self.callback = None # 체결 시 실행할 외부 함수 저장용
    
    def set_callback(self, func):
        """체결 알림을 받을 함수를 등록"""
        self.callback = func
            
    # --- 실시간 수신 설정 ---
    def subscribe_conclusion(self):
        """실시간 체결 수신 시작"""
        handler = win32com.client.WithEvents(self.obj_conclusion, CpEvent)
        handler.set_params(self)
        self.obj_conclusion.Subscribe()
        print("Subscribed to real-time conclusion.")

    def unsubscribe_conclusion(self):
        """실시간 체결 수신 종료"""
        self.obj_conclusion.Unsubscribe()
        print("Unsubscribed from real-time conclusion.")

    def process_conclusion(self):
        """CpConclusion 이벤트 발생 시 호출"""
        code = self.obj_conclusion.GetHeaderValue(9)
        name = self.obj_conclusion.GetHeaderValue(1)
        actual_price = self.obj_conclusion.GetHeaderValue(4) # 체결가
        concluded_qty = self.obj_conclusion.GetHeaderValue(3) # 🎯 체결 수량 추출
        status = self.obj_conclusion.GetHeaderValue(14)      # 체결 상태
        # 🎯 [추가] 매수/매도 구분값 추출 (인덱스 12)
        # 대신증권 기준: '1'은 매도, '2'는 매수
        side = self.obj_conclusion.GetHeaderValue(12)
        
        # '1'(체결) 또는 '2'(확인) 상태일 때 콜백 전달
        if status in ['1', '2'] and self.callback:
            concl_data = {
                'code': code,
                'name': name,
                'actual_price': actual_price,
                'concluded_qty': concluded_qty, # 🎯 수량 데이터 추가
                'side': str(side) # 🎯 RealtimeManager에서 비교하기 쉽게 문자열로 변환 
            }
            self.callback(concl_data)
                
    # --- 기존 주문 메서드 ---
    def request_new_order(self, acc_no, acc_flag, code, qty, price, order_type="2", hoga_flag="01"):
        self.obj_new_order.SetInputValue(0, order_type)
        self.obj_new_order.SetInputValue(1, acc_no)
        self.obj_new_order.SetInputValue(2, acc_flag)
        self.obj_new_order.SetInputValue(3, code)
        self.obj_new_order.SetInputValue(4, qty)
        self.obj_new_order.SetInputValue(5, price)
        self.obj_new_order.SetInputValue(8, hoga_flag) # "01": 보통가, "03": 시장가
        
        ret = self.obj_new_order.BlockRequest()
        if ret != 0:
            print(f"New Order Request Failed (Code: {ret})")
            return None
        #################################################
        # 🎯 [추가] 증권사 서버의 응답 메시지 확인
        rq_status = self.obj_new_order.GetDibStatus()
        rq_msg = self.obj_new_order.GetDibMsg1()
        
        if rq_status != 0:
            print(f"❌ [주문 거부] 사유: {rq_msg} (코드: {rq_status})")
        else:
            # GetHeaderValue(8)은 주문번호를 가져오는 코드입니다. (인덱스는 API 버전에 따라 다를 수 있음)
            order_no = self.obj_new_order.GetHeaderValue(8) 
            if order_no == 0 or order_no == "" or order_no is None:
                print(f"⚠️ [주문 이상] 전송은 되었으나 주문번호를 받지 못했습니다. 사유: {rq_msg}")
            else:
                print(f"✅ [주문 접수] 주문번호: {order_no} | 상태: {rq_msg}")
        ##################################################                
        order_no = self.obj_new_order.GetHeaderValue(8)
        print(f"New Order Success - Order No: {order_no}")
        return order_no

    def request_modify_order(self, org_order_no, acc_no, acc_flag, code, qty, price):
        self.obj_modify_order.SetInputValue(1, org_order_no)
        self.obj_modify_order.SetInputValue(2, acc_no)
        self.obj_modify_order.SetInputValue(3, acc_flag)
        self.obj_modify_order.SetInputValue(4, code)
        self.obj_modify_order.SetInputValue(5, qty)
        self.obj_modify_order.SetInputValue(6, price)
        
        ret = self.obj_modify_order.BlockRequest()
        if ret != 0:
            print(f"Modify Order Request Failed (Code: {ret})")
            return None

        new_order_no = self.obj_modify_order.GetHeaderValue(8)
        print(f"Modify Order Success - New Order No: {new_order_no}")
        return new_order_no

    def request_cancel_order(self, org_order_no, acc_no, acc_flag, code, qty=0):
        self.obj_cancel_order.SetInputValue(1, org_order_no)
        self.obj_cancel_order.SetInputValue(2, acc_no)
        self.obj_cancel_order.SetInputValue(3, acc_flag)
        self.obj_cancel_order.SetInputValue(4, code)
        self.obj_cancel_order.SetInputValue(5, qty)
        
        ret = self.obj_cancel_order.BlockRequest()
        if ret != 0:
            print(f"Cancel Order Request Failed (Code: {ret})")
            return False

        print(f"Cancel Order Requested (Org No: {org_order_no})")
        return True

# --- 메인 실행 예시 ---
if __name__ == "__main__":
    manager = OrderManager()
    
    # 실시간 감시 시작
    manager.subscribe_conclusion()

    print("Listening for conclusions... Press Ctrl+C to stop.")
    
    try:
        while True:
            # 이 함수가 호출되어야 COM 이벤트(OnReceived)가 처리됩니다.
            pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        manager.unsubscribe_conclusion()