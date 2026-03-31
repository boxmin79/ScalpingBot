import win32com.client

class OrderManager:
    """
    주식 신규 주문, 정정, 취소를 담당하는 클래스
    """
    def __init__(self):
        # 1. 신규 주문 오브젝트
        self.obj_new_order = win32com.client.Dispatch("CpTrade.CpTd0311")
        # 2. 정정 주문 오브젝트
        self.obj_modify_order = win32com.client.Dispatch("CpTrade.CpTd0313")
        # 3. 취소 주문 오브젝트
        self.obj_cancel_order = win32com.client.Dispatch("CpTrade.CpTd0314")

    def request_new_order(self, acc_no, acc_flag, code, qty, price, order_type="2"):
        """
        [CpTd0311] 신규 현금 주문 (매수/매도)
        order_type: "1"-매도, "2"-매수
        """
        self.obj_new_order.SetInputValue(0, order_type)   # 1:매도, 2:매수
        self.obj_new_order.SetInputValue(1, acc_no)       # 계좌번호
        self.obj_new_order.SetInputValue(2, acc_flag)     # 상품관리구분코드
        self.obj_new_order.SetInputValue(3, code)         # 종목코드
        self.obj_new_order.SetInputValue(4, qty)          # 주문수량
        self.obj_new_order.SetInputValue(5, price)        # 주문단가
        self.obj_new_order.SetInputValue(8, "01")         # 호가구분: 01 보통
        
        ret = self.obj_new_order.BlockRequest()
        if ret != 0:
            print(f"신규 주문 요청 실패 (에러코드: {ret})")
            return None

        # 주문 성공 시 서버에서 부여한 주문번호 반환
        order_no = self.obj_new_order.GetHeaderValue(8)
        print(f"신규 주문 성공 - 주문번호: {order_no}")
        return order_no

    def request_modify_order(self, org_order_no, acc_no, acc_flag, code, qty, price):
        """
        [CpTd0313] 가격/수량 정정 주문
        org_order_no: 정정하고자 하는 원주문 번호
        qty: 0으로 설정 시 잔량 전체 정정
        """
        self.obj_modify_order.SetInputValue(1, org_order_no) # 원주문 번호
        self.obj_modify_order.SetInputValue(2, acc_no)       # 계좌번호
        self.obj_modify_order.SetInputValue(3, acc_flag)     # 상품관리구분코드
        self.obj_modify_order.SetInputValue(4, code)         # 종목코드
        self.obj_modify_order.SetInputValue(5, qty)          # 정정 수량
        self.obj_modify_order.SetInputValue(6, price)        # 정정 단가
        
        ret = self.obj_modify_order.BlockRequest()
        if ret != 0:
            print(f"정정 주문 요청 실패 (에러코드: {ret})")
            return None

        new_order_no = self.obj_modify_order.GetHeaderValue(8)
        print(f"정정 주문 성공 - 새로운 주문번호: {new_order_no}")
        return new_order_no

    def request_cancel_order(self, org_order_no, acc_no, acc_flag, code, qty=0):
        """
        [CpTd0314] 취소 주문
        qty: 0으로 설정 시 잔량 전체 취소
        """
        self.obj_cancel_order.SetInputValue(1, org_order_no) # 원주문 번호
        self.obj_cancel_order.SetInputValue(2, acc_no)       # 계좌번호
        self.obj_cancel_order.SetInputValue(3, acc_flag)     # 상품관리구분코드
        self.obj_cancel_order.SetInputValue(4, code)         # 종목코드
        self.obj_cancel_order.SetInputValue(5, qty)          # 취소 수량 (0: 잔량전부)
        
        ret = self.obj_cancel_order.BlockRequest()
        if ret != 0:
            print(f"취소 주문 요청 실패 (에러코드: {ret})")
            return False

        print(f"취소 주문 완료 (원주문번호: {org_order_no})")
        return True