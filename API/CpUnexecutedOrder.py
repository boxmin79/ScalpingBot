import win32com.client

class CpUnexecutedOrder:
    """
    CpTrade.CpTd5339 기능을 포함하는 클래스
    설명: 계좌별 미체결 잔량 데이터를 요청하고 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpTrade.CpTd5339")

    def get_unexecuted_list(self, acc_no, acc_flag, stock_code="", count=20):
        """
        미체결 내역을 조회하여 리스트로 반환합니다.
        acc_no: 계좌번호
        acc_flag: 상품관리구분코드
        stock_code: 종목코드 (생략 시 전종목)
        count: 요청 개수 (최대 20개)
        """
        # 1. 입력 데이터 설정
        self.obj.SetInputValue(0, acc_no)       # 계좌번호
        self.obj.SetInputValue(1, acc_flag)     # 상품관리구분코드
        self.obj.SetInputValue(3, stock_code)   # 종목코드
        self.obj.SetInputValue(4, "0")          # 주문구분: "0" 전체
        self.obj.SetInputValue(5, "0")          # 정렬구분: "0" 순차
        self.obj.SetInputValue(6, "0")          # 주문종가구분: "0" 전체
        self.obj.SetInputValue(7, count)        # 요청개수
        self.obj.SetInputValue(8, "0")          # 거래소유형: "0" 전체

        results = []
        
        while True:
            # 2. 데이터 요청
            ret = self.obj.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = self.obj.GetHeaderValue(5) # 실제 수신 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    'order_no': self.obj.GetDataValue(1, i),      # 주문번호
                    'org_order_no': self.obj.GetDataValue(2, i),  # 원주문번호
                    'code': self.obj.GetDataValue(3, i),          # 종목코드
                    'name': self.obj.GetDataValue(4, i),          # 종목명
                    'content': self.obj.GetDataValue(5, i),       # 주문내용
                    'qty': self.obj.GetDataValue(6, i),           # 주문수량
                    'price': self.obj.GetDataValue(7, i),         # 주문단가
                    'exec_qty': self.obj.GetDataValue(8, i),      # 체결수량
                    'cancelable_qty': self.obj.GetDataValue(11, i), # 정정취소가능수량 (핵심)
                    'side_code': self.obj.GetDataValue(13, i),    # 매매구분 (1:매도, 2:매수)
                    'order_type': self.obj.GetDataValue(21, i),   # 주문호가구분코드
                    'result_status': self.obj.GetDataValue(30, i) # 주문접수결과 (0:대기, 1:정상, 2:접수)
                }
                results.append(item)

            # 5. 연속 데이터 유무 확인 (Paging)
            if self.obj.Continue == False:
                break
                
        return results

# --- 사용 예시 ---
if __name__ == "__main__":
    # 주문 초기화(TradeInit)가 선행되어야 합니다.
    # trade_mgr = CpTradeUtil()
    # trade_mgr.trade_init()

    unexecuted_mgr = CpUnexecutedOrder()
    
    # 내 계좌의 모든 미체결 내역 조회
    # acc_no, acc_flag는 자신의 계좌 정보를 사용하세요.
    history = unexecuted_mgr.get_unexecuted_list("YOUR_ACC_NO", "01")

    print(f"\n현재 미체결 내역 총 {len(history)}건")
    for row in history:
        side = "매수" if row['side_code'] == '2' else "매도"
        print(f"종목: {row['name']} | {side} | 주문가: {row['price']} | "
              f"미체결량: {row['cancelable_qty']} | 주문번호: {row['order_no']}")