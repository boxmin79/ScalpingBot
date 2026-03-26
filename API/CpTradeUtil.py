import win32com.client

class CpTradeUtil:
    """
    CpTrade.CpTdUtil 기능을 포함하는 클래스
    설명: 주문 오브젝트 사용을 위한 초기화 및 계좌 정보를 관리합니다.
    [필독] 모든 주문 관련 작업 전 반드시 trade_init()을 먼저 호출해야 합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpTrade.CpTdUtil")

    def trade_init(self):
        """
        주문을 하기 위한 예비 과정(초기화)을 수행합니다.
        호출 시 비밀번호 입력 창이 뜰 수 있습니다.
        반환값: 0-정상, -1-오류, 1-OTP/보안카드 에러, 3-취소
        """
        # VC++의 경우 0을 설정하므로 파이썬에서도 0을 인자로 전달합니다.
        ret = self.obj.TradeInit(0)
        
        if ret == 0:
            print("주문 초기화 성공 (정상)")
        elif ret == -1:
            print("주문 초기화 실패 (오류/비밀번호 틀림)")
        elif ret == 1:
            print("주문 초기화 실패 (OTP/보안카드 오입력)")
        elif ret == 3:
            print("주문 초기화 취소")
        
        return ret

    def get_account_numbers(self):
        """
        사용자의 계좌 목록을 배열(tuple) 형태로 반환합니다.
        [주의] trade_init()이 성공한 이후에만 정상적으로 값을 가져옵니다.
        """
        return self.obj.AccountNumber

    def get_goods_list(self, acc_no, filter_type=-1):
        """
        해당 계좌의 상품별 계좌 목록을 반환합니다.
        acc_no: 계좌번호 (string)
        filter_type: 
            -1 : 전체
            1 : 주식
            2 : 선물/옵션
            16 : EUREX
            64 : 해외선물
            (조합 가능: 주식+선물 = 3)
        """
        return self.obj.GoodsList(acc_no, filter_type)

# --- 활용 예시 ---
if __name__ == "__main__":
    trade_util = CpTradeUtil()
    
    # 1. 주문 초기화 (가장 먼저 실행)
    if trade_util.trade_init() == 0:
        
        # 2. 계좌 목록 가져오기
        accounts = trade_util.get_account_numbers()
        print(f"보유 계좌 목록: {accounts}")
        
        if accounts:
            # 3. 첫 번째 계좌의 주식(1) 상품 코드 확인
            acc = accounts[0]
            stock_goods = trade_util.get_goods_list(acc, 1)
            print(f"계좌 {acc}의 주식 상품 코드: {stock_goods}")