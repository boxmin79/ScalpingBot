import win32com.client

class CpOrderHistory:
    """
    CpTrade.CpTd5341 기능을 포함하는 클래스
    설명: 금일 계좌별 주문/체결 내역 조회 데이터를 요청하고 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpTrade.CpTd5341")

    def get_history_list(self, acc_no, acc_flag, stock_code="", count=20):
        """
        주문/체결 내역을 조회하여 리스트로 반환합니다.
        acc_no: 계좌번호
        acc_flag: 상품관리구분코드
        stock_code: 종목코드 (생략 시 전종목)
        count: 요청 개수 (최대 20개)
        """
        # 1. 입력 데이터 설정
        self.obj.SetInputValue(0, acc_no)       # 계좌번호
        self.obj.SetInputValue(1, acc_flag)     # 상품관리구분코드
        self.obj.SetInputValue(2, stock_code)   # 종목코드
        self.obj.SetInputValue(3, 0)            # 시작주문번호 (0: 처음부터)
        self.obj.SetInputValue(4, ord('1'))     # 정렬구분: '1' 역순(최근순)
        self.obj.SetInputValue(5, count)        # 요청개수
        self.obj.SetInputValue(6, ord('2'))     # 조회구분: '2' 건별
        self.obj.SetInputValue(7, ord('0'))     # 거래소유형: '0' 전체

        results = []
        
        while True:
            # 2. 데이터 요청
            ret = self.obj.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = self.obj.GetHeaderValue(6) # 실제 수신 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    'order_no': self.obj.GetDataValue(1, i),      # 주문번호
                    'org_order_no': self.obj.GetDataValue(2, i),  # 원주문번호
                    'code': self.obj.GetDataValue(3, i),          # 종목코드
                    'name': self.obj.GetDataValue(4, i),          # 종목이름
                    'content': self.obj.GetDataValue(5, i),       # 주문내용
                    'qty': self.obj.GetDataValue(7, i),           # 주문수량
                    'price': self.obj.GetDataValue(8, i),         # 주문단가
                    'exec_total': self.obj.GetDataValue(9, i),    # 총체결수량
                    'exec_qty': self.obj.GetDataValue(10, i),     # 이번 체결수량
                    'exec_price': self.obj.GetDataValue(11, i),   # 체결단가
                    'side_code': self.obj.GetDataValue(35, i),    # 매매구분코드 (1:매도, 2:매수)
                    'type_code': self.obj.GetDataValue(36, i),    # 정정취소구분코드 (1:정상, 2:정정, 3:취소)
                    'time': self.obj.GetDataValue(42, i),         # 체결상세 시분초
                }
                results.append(item)

            # 5. 연속 데이터 유무 확인 (Paging)
            if self.obj.Continue == False:
                break
            
            # 다음 조회를 위해 잠시 대기 (TR 과부하 방지)
            import time
            time.sleep(0.2)
            
        return results

# --- 사용 예시 ---
if __name__ == "__main__":
    # 1. 초기화 (주문 초기화 클래스가 먼저 실행되어야 함)
    # td_util = CpTradeUtil() ...
    
    # 2. 조회 실행
    history_mgr = CpOrderHistory()
    
    # 계좌번호와 상품구분코드는 CpTradeUtil에서 가져온 값을 사용하세요.
    # 예: "12345678", "01"
    history = history_mgr.get_history_list("YOUR_ACC_NO", "01")

    print(f"\n조회된 내역 총 {len(history)}건")
    for row in history[:10]: # 최근 10건만 출력
        side = "매수" if row['side_code'] == '2' else "매도"
        status = "정상"
        if row['type_code'] == '2': status = "정정"
        elif row['type_code'] == '3': status = "취소"
        
        print(f"[{row['time']}] {row['name']} | {side}({status}) | 수량:{row['qty']} | 체결:{row['exec_total']} | 번호:{row['order_no']}")