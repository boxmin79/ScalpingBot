import path_finder
import win32com.client
import time

class AccountManager:
    """계좌관리클래스
    통신종류: Requests/Reply
    모듈위치 CpTrade.dll
    
    결제예정예수금 가계산: CpTrade.CpTd0732
    계좌별 매수 가능금액/수량: CpTrade.CpTdNew5331A
    계좌별 매도 가능수량: CpTrade.CpTdNew5331B
    
    계좌별 미체결잔량: CpTrade.CpTd5339
    국내 주식 주문 주문체결 내역: CpTrade.CpTd5341
    금일/전일 쳬결기준 내역: CpTrade.CpTd5342
    
    계좌별 잔고 평가현황: CpTrade.CpTd6033
    체결기준 주식 당일매매손익:  CpTrade.CpTd6032
    """
    def __init__(self, acc_no:str=None, acc_flag:str=None):
        # 기본 오프젝트 로드
        self.cybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.util = win32com.client.Dispatch("CpTrade.CpTdUtil")
        
        if self.cybos.IsConnect:
            # print("CreonPlus에 연결되었습니다.")
            # print(f"주문 초기화: type: {type(ti)}, value: {ti}")
            if not self.util.TradeInit(0):
                # print("주문 초기화 완료")
                if acc_no is not None:
                    self.acc_no = acc_no  # 계좌번호 설정 <class 'str'>
                else: 
                    self.acc_no = self.util.AccountNumber[0]
                    print(f"계좌번호: {self.acc_no}")
                    
                if acc_flag is not None:
                    self.acc_flag = acc_flag  
                else:
                    self.acc_flag = self.util.GoodsList(self.acc_no, 1)[0] # 상품구분코드 설정 <class 'str'>
                    print(f"상품관리구분코드: {self.acc_flag}")   
            else:
                print("주문 초기화 실패")
        else:
            print("CreonPlus에 연결되지 않았습니다.")

    
    def get_expected_deposit(self):
        """
        [CpTd0732] 주식 결제예정 예수금 가계산 조회
        D+2 결제 기준의 실제 가용 자금을 확인합니다.

        Args:
            acc_no (str, optional): 계좌번호. Defaults to None.
            acc_flag (str, optional): 상품관리구분코드. Defaults to None.

        Returns:
            Dict:   
        """
        # [CpTd0732] 주식 결제예정 예수금 가계산 조회
        # D+2 결제 기준의 실제 가용 자금을 확인합니다.
        
        obj = win32com.client.Dispatch("CpTrade.CpTd0732")     # 결제예정 예수금 가계산
        
        # 1. 입력값 설정
        obj.SetInputValue(0, self.acc_no)
        obj.SetInputValue(1, self.acc_flag)

        # 2. 데이터 요청
        res = obj.BlockRequest()
        
        if res != 0:
            print(f"❌ CpTd0732 요청 실패 (에러코드: {res})")
            return None
        result = {
        "계좌번호":obj.GetHeaderValue(0), # (string)
        "상품관리구분코드":obj.GetHeaderValue(1), # (string)
        "계좌명":obj.GetHeaderValue(2), # (string)
        "예수금":obj.GetHeaderValue(3), # (longlong)
        "미수금":obj.GetHeaderValue(4), # (long)
        "전일장내현금매도":obj.GetHeaderValue(5), # (long)
        "전일장내현금매수":obj.GetHeaderValue(6), # (long)
        "전일신용융자매도":obj.GetHeaderValue(7), # (long)
        "전일신용융자매수":obj.GetHeaderValue(8), # (long)
        "전일신용대주매도":obj.GetHeaderValue(9), # (long)
        "전일신용대주매수":obj.GetHeaderValue(10), # (long)
        "전일현금수수료":obj.GetHeaderValue(11), # (long)
        "전일현금제세금":obj.GetHeaderValue(12), # (long)
        "전일현금정산금":obj.GetHeaderValue(13), # (long)
        "전일장외단주매도":obj.GetHeaderValue(14), # (long)
        "전일장외단주매수":obj.GetHeaderValue(15), # (long)
        "전일장외수수료":obj.GetHeaderValue(16), # (long)
        "전일장외제세금":obj.GetHeaderValue(17), # (long)
        "전일장외정산금":obj.GetHeaderValue(18), # (long)
        "전일합계매도금":obj.GetHeaderValue(19), # (long)   
        "전일합계매수금":obj.GetHeaderValue(20), # (long)
        "전일합계수수료":obj.GetHeaderValue(21), # (long)   
        "전일합계제세금":obj.GetHeaderValue(22), # (long)
        "전일합계정산금":obj.GetHeaderValue(23), # (long)
        "전일장내현금신규융자":obj.GetHeaderValue(24), # (long)
        "전일신용융자융자상환":obj.GetHeaderValue(25), # (long)
        "전일장내현금신규대주":obj.GetHeaderValue(26), # (long)
        "전일신용융자대주상환":obj.GetHeaderValue(27), # (long)
        "전일장내현금신용상환":obj.GetHeaderValue(28), # (long)
        "전일상환융자이자":obj.GetHeaderValue(29), # (long)
        "전일상황이용료":obj.GetHeaderValue(30), # (long)
        "전일현금거래세":obj.GetHeaderValue(31), # (long)
        "전일대주소득":obj.GetHeaderValue(32), # (long)
        "전일대주주민세":obj.GetHeaderValue(33), # (long)
        "금일장내현금매도":obj.GetHeaderValue(34), # (long)
        "금일장내현금매수":obj.GetHeaderValue(35), # (long)
        "금일신용융자매도":obj.GetHeaderValue(36), # (long)
        "금일신용융자매수":obj.GetHeaderValue(37), # (long)
        "금일신용대주매도":obj.GetHeaderValue(38), # (long)
        "금일신용대주매수":obj.GetHeaderValue(39), # (long)
        "금일현금수수료":obj.GetHeaderValue(40), # (long)
        "금일현금제세금":obj.GetHeaderValue(41), # (long)
        "금일현금정산금":obj.GetHeaderValue(42), # (long)
        "금일장외단주매도":obj.GetHeaderValue(43), # (long)
        "금일장외단주매수":obj.GetHeaderValue(44), # (long)
        "금일장외수수료":obj.GetHeaderValue(45), # (long)
        "금일장외제세금":obj.GetHeaderValue(46), # (long)
        "금일장외정산금":obj.GetHeaderValue(47), # (long)
        "금일합계매도금":obj.GetHeaderValue(48), # (long)
        "금일합계매수금":obj.GetHeaderValue(49), # (long)
        "금일합계수수료":obj.GetHeaderValue(50), # (long)
        "금일합계제세금":obj.GetHeaderValue(51), # (long)
        "금일합계정산금":obj.GetHeaderValue(52), # (long)
        "금일장내현금신규융자":obj.GetHeaderValue(53), # (long)
        "금일신용융자융자상환":obj.GetHeaderValue(54), # (long)
        "금일장내현금신규대주":obj.GetHeaderValue(55), # (long)
        "금일신용융자대주상환":obj.GetHeaderValue(56), # (long)
        "금일장내현금신용상환":obj.GetHeaderValue(57), # (long)
        "금일상환융자이자":obj.GetHeaderValue(58), # (long)
        "금일상황이용료":obj.GetHeaderValue(59), # (long)
        "금일현금거래세":obj.GetHeaderValue(60), # (long)
        "금일대주소득세":obj.GetHeaderValue(61), # (long)
        "금일대주주민세":obj.GetHeaderValue(62), # (long)
        "익일영업일":obj.GetHeaderValue(63), # (long)
        "익영업일예수금":obj.GetHeaderValue(64), # (longlong)
        "결제일":obj.GetHeaderValue(65), # (long)
        "결제일예수금":obj.GetHeaderValue(66), # (longlong)
        }

        return result
    
    def get_buyable_data(self, 
                         stk_code: str, 
                         price: int = 0, 
                         quote_type: str = '01', 
                         query_type: str = '2'):
        """
        계좌별 매수 가능 금액/수량을 조회합니다.
        
        Args:
            stock_code (str): 종목코드
            price (int): 주문 단가 (수량 조회 시 필수)
            quote_type (str): 호가구분코드 ('01': 보통, '03': 시장가 등)
            query_type (str): '1': 금액조회, '2': 수량조회
        """
         
        obj = win32com.client.Dispatch("CpTrade.CpTdNew5331A")          # 매수 가능 금액/수량 조회 v
        
        # 1. 입력 데이터 설정
        obj.SetInputValue(0, self.acc_no)       # 계좌번호
        obj.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        obj.SetInputValue(2, stk_code)         # 종목코드
        obj.SetInputValue(3, quote_type)         # 호가구분 (보통/시장가 등)
        obj.SetInputValue(4, price)              # 주문단가 (시장가면 0)
        obj.SetInputValue(5, 'Y')                # 증거금 100% 징수여부 (N: 계좌별 설정)
        obj.SetInputValue(6, query_type)         # 조회구분 ('2': 수량조회 중심)

        # 2. 데이터 요청
        ret = obj.BlockRequest()
        if ret != 0:
            print(f"매수 가능 조회 실패 (에러코드: {ret})")
            return None

        # 3. 결과 파싱 (스캘핑에 필요한 핵심 필드 위주)
        # 이 오브젝트는 HeaderValue에서 데이터를 제공합니다.
        data = {
            "종목코드": obj.GetHeaderValue(0),                # (string)
            "종목명": obj.GetHeaderValue(1),                  # (string)
            "증거금율구분코드": obj.GetHeaderValue(2),          # (char)
            "증거금20%주문가능금액": obj.GetHeaderValue(3),    # (longlong)
            "증거금30%주문가능금액": obj.GetHeaderValue(4),    # (longlong)
            "증거금40%주문가능금액": obj.GetHeaderValue(5),    # (longlong)
            "증거금50%주문가능금액": obj.GetHeaderValue(6),    # (longlong)
            "증거금60%주문가능금액": obj.GetHeaderValue(7),    # (longlong)
            "증거금70%주문가능금액": obj.GetHeaderValue(8),    # (longlong)
            "증거금100%주문가능금액": obj.GetHeaderValue(9),   # (longlong)
            "현금주문가능금액": obj.GetHeaderValue(10),         # (longlong)
            "증거금20%주문가능수량": obj.GetHeaderValue(11),    # (long)
            "증거금30%주문가능수량": obj.GetHeaderValue(12),    # (long)
            "증거금40%주문가능수량": obj.GetHeaderValue(13),    # (long)
            "증거금50%주문가능수량": obj.GetHeaderValue(14),    # (long)
            "증거금60%주문가능수량": obj.GetHeaderValue(15),    # (long)
            "증거금70%주문가능수량": obj.GetHeaderValue(16),    # (long)
            "증거금100%주문가능수량": obj.GetHeaderValue(17),   # (long)
            "현금주문가능수량": obj.GetHeaderValue(18),         # (long)
            "증거금20%융자주문가능금액": obj.GetHeaderValue(19), # (longlong)
            "증거금30%융자주문가능금액": obj.GetHeaderValue(20), # (longlong)
            "증거금40%융자주문가능금액": obj.GetHeaderValue(21), # (longlong)
            "증거금50%융자주문가능금액": obj.GetHeaderValue(22), # (longlong)
            "증거금60%융자주문가능금액": obj.GetHeaderValue(23), # (longlong)
            "증거금70%융자주문가능금액": obj.GetHeaderValue(24), # (longlong)
            "융자주문가능금액": obj.GetHeaderValue(25),         # (longlong)
            "증거금20%융자주문가능수량": obj.GetHeaderValue(26), # (long)
            "증거금30%융자주문가능수량": obj.GetHeaderValue(27), # (long)
            "증거금40%융자주문가능수량": obj.GetHeaderValue(28), # (long)
            "증거금50%융자주문가능수량": obj.GetHeaderValue(29), # (long)
            "증거금60%융자주문가능수량": obj.GetHeaderValue(30), # (long)
            "증거금70%융자주문가능수량": obj.GetHeaderValue(31), # (long)
            "융자주문가능수량": obj.GetHeaderValue(32),         # (long)
            "대주가능수량": obj.GetHeaderValue(33),            # (long)
            "매도가능수량": obj.GetHeaderValue(34),            # (long)
            "매입80%가능금액": obj.GetHeaderValue(35),         # (longlong)
            "매입80%가능수량": obj.GetHeaderValue(36),         # (long)
            "매입100%가능금액": obj.GetHeaderValue(37),        # (longlong)
            "매입100%가능수량": obj.GetHeaderValue(38),        # (long)
            "매입110%가능금액": obj.GetHeaderValue(39),        # (longlong)
            "매입110%가능수량": obj.GetHeaderValue(40),        # (long)
            "매입120%가능금액": obj.GetHeaderValue(41),        # (longlong)
            "매입120%가능수량": obj.GetHeaderValue(42),        # (long)
            "매입140%가능금액": obj.GetHeaderValue(43),        # (longlong)
            "매입140%가능수량": obj.GetHeaderValue(44),        # (long)
            "예수금": obj.GetHeaderValue(45),                 # (longlong)
            "대용금": obj.GetHeaderValue(46),                 # (longlong)
            "가능예수금": obj.GetHeaderValue(47),               # (longlong)
            "가능대용금": obj.GetHeaderValue(48),               # (longlong)
            "신용상환차익금액": obj.GetHeaderValue(49),         # (longlong)
            "신용담보현금금액": obj.GetHeaderValue(50),         # (longlong)
            "미결제환원금": obj.GetHeaderValue(51),            # (longlong)
            "결제환원금": obj.GetHeaderValue(52),              # (longlong)
            "대주가능금액": obj.GetHeaderValue(53),            # (longlong)
            "주식적용증거금구분코드": obj.GetHeaderValue(54),     # (char)
            "주식적용증거금구분내용": obj.GetHeaderValue(55)      # (string)
            }

        return data
    
    def get_sellable_qty(self, 
                         stk_code:str='',
                         stk_bond_flag:str='1',
                         cash_credit_flag:str='1',
                         rqst_count=20):
        """
        계좌별매도주문가능수량데이터를요청하고수신한다

        Args:
            stock_code (str, optional): _description_. Defaults to ''.
            stk_bond_flag (str, optional): _description_. Defaults to '1'.
            cash_credit_flag (str, optional): _description_. Defaults to '1'.
            rqst_count (int, optional): _description_. Defaults to 20.

        Returns:
            _type_: _description_
        """
        obj = win32com.client.Dispatch("CpTrade.CpTdNew5331B")         # 매도 가능 수량 조회 v
        # 1. 입력 데이터 설정
        obj.SetInputValue(0, self.acc_no)       # 계좌번호
        obj.SetInputValue(1, self.acc_flag)     # 상품구분
        obj.SetInputValue(2, stk_code)         # 특정 종목 조회
        obj.SetInputValue(3, stk_bond_flag)          # 1: 주식
        obj.SetInputValue(4, cash_credit_flag)  # 1: 현금(현금주문 가능 수량)
        obj.SetInputValue(10, rqst_count)              # 요청개수


        # 2. 데이터 요청
        obj.BlockRequest()

       # 3. 결과 파싱
        recv_ecount = self.obj.GetHeaderValue(0)
        if recv_ecount == 0:
            return 0
        data = {
            "종목코드": obj.GetDataValue(0, 0),          # (string)
            "종목명": obj.GetDataValue(1, 0),            # (string)
            "증거금율구분코드": obj.GetDataValue(2, 0),    # (char)
            "신용대출구분내용": obj.GetDataValue(3, 0),    # (string)
            "주식매매수량단위": obj.GetDataValue(4, 0),    # (longlong)
            "대출가능여부코드": obj.GetDataValue(5, 0),    # (char)
            "잔고수량": obj.GetDataValue(6, 0),          # (long)
            "전일매수체결수량": obj.GetDataValue(7, 0),    # (string)
            "전일매도체결수량": obj.GetDataValue(8, 0),    # (long)
            "지정수량": obj.GetDataValue(9, 0),          # (long)
            "금일매수체결수량": obj.GetDataValue(10, 0),   # (long)
            "금일매도체결수량": obj.GetDataValue(11, 0),   # (long)
            "매도가능수량": obj.GetDataValue(12, 0),       # (long)
            "신용대출금액": obj.GetDataValue(13, 0),       # (long)
            "주문구분코드": obj.GetDataValue(14, 0),       # (char)
            "신용대출구분코드": obj.GetDataValue(15, 0),   # (string)
            "신용대출일자": obj.GetDataValue(16, 0),       # (long)
            "만기일자": obj.GetDataValue(17, 0),          # (long)
            "가잔고수량": obj.GetDataValue(18, 0)         # (long)
}
        return data   
        
    def get_trade_history(self, 
                          rqst_count:int=20,
                          date_flag:str="1", 
                          target_code:str=""):
        """
        종목별금일/전일체결기준내역조회

        Args:
            count (int, optional):  요청개수[default:7] - 최대 20개
            date_flag (str, optional): 요청일구분코드 - '1' 금일, '2' 전일[default]
            target_code (str, optional): 요청종목코드[default:"" - 모든종목]

        Returns:
            _type_: _description_
        """
        obj = win32com.client.Dispatch("CpTrade.CpTd5342")      # 금일/전일 체결 기준 내역 v
        # 1. 입력 데이터 설정
        obj.SetInputValue(0, self.acc_no)      # 계좌번호
        obj.SetInputValue(1, self.acc_flag) # 상품구분
        obj.SetInputValue(2, rqst_count)               # 한 번에 최대 20개 조회
        obj.SetInputValue(3, date_flag)        # 금일/전일 구분
        obj.SetInputValue(4, target_code)      # 종목코드 (기본 전체)

        trades = []
        
        while True:
            # 2. 데이터 요청
            obj.BlockRequest()

            # 헤더 정보 (요청 결과 요약)
            recv_count = obj.GetHeaderValue(8)      # 수신 개수
            trade_date = obj.GetHeaderValue(3) # 매매일
            
            # 3. 데이터 파싱 (아이템별 상세 정보)
            for i in range(recv_count):
                item = {
                    "종목코드": obj.GetDataValue(0, i),          # (string)
                    "종목명": obj.GetDataValue(1, i),            # (string)
                    "체결수량": obj.GetDataValue(3, i),          # (long)
                    "수수료": obj.GetDataValue(4, i),            # (long)
                    "농특세": obj.GetDataValue(5, i),            # (longlong)
                    "매매구분내용": obj.GetDataValue(10, i),      # (string)
                    "거래세": obj.GetDataValue(12, i),           # (long)
                    "매수일": obj.GetDataValue(13, i),           # (string)
                    "과세구분내용": obj.GetDataValue(14, i),      # (string)
                    "신용차금": obj.GetDataValue(17, i),         # (long)
                    "매매거래유형코드": obj.GetDataValue(19, i),   # (string)
                    "매매거래유형내용": obj.GetDataValue(20, i),   # (string)
                    "매매구분코드": obj.GetDataValue(21, i),      # (string)
                    "약정금액": obj.GetDataValue(22, i),         # (ulonglong)
                    "결제금액": obj.GetDataValue(23, i),         # (ulonglong)
                    "정산금액": obj.GetDataValue(24, i),         # (ulonglong)
                    "체결단가": obj.GetDataValue(28, i),         # (double)
                    "거래소주문유형": obj.GetDataValue(29, i)     # (string)
                    }
                trades.append(item)

            # 4. 연속 데이터 처리 (20개가 넘을 경우)
            if not obj.Continue:
                break
        
        return trades, trade_date
    
    def get_unexecuted_list(self,
                            stock_code:str='',
                            ord_flag:str='0',
                            sort_flag:str='0',
                            close_type:str='1',
                            rqst_count:int=20,
                            mkt_type:str='1'):
        """
        계좌별미체결잔량데이터를요청하고수신한다

        Args:
            stock_code (str, optional): 종목코드(입력값생략가능, 생략시전체종목이조회됨)
            ord_type (str, optional):  주문구분코드[default:'0'] - '0' 전체, '1' 거래소주식, '2' 장내채권, '3' 코스닥주식', '4' 장외단주', '5' 프리보드
            sort_type (str, optional): 정렬구분코드[default:'0'] - '0' 순차, '1' 역순
            close_type (str, optional): 주문종가구분코드[default:'0'] - '0'전체, '1' 일반, '2' 시간외종가,  '3' 시간외단일가
            count (int, optional): 요청개수 (최대 20개)
            mkt_type (str, optional): 거래소주문유형[default:1] - '0' 전체, '1' KRX, '2' NXT

        Returns:
            _type_: _description_
        """
        obj = win32com.client.Dispatch("CpTrade.CpTd5339")      # 미체결 잔량 조회 v
        # 1. 입력 데이터 설정
        obj.SetInputValue(0, self.acc_no)       # 계좌번호
        obj.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        obj.SetInputValue(3, stock_code)   # 종목코드
        obj.SetInputValue(4, ord_flag)          # 주문구분: "0" 전체
        obj.SetInputValue(5, sort_flag)          # 정렬구분: "0" 순차
        obj.SetInputValue(6, close_type)          # 주문종가구분: "0" 전체
        obj.SetInputValue(7, rqst_count)        # 요청개수
        obj.SetInputValue(8, mkt_type)          # 거래소유형: "0" 전체

        results = []
        
        while True:
            # 2. 데이터 요청
            ret = obj.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = obj.GetHeaderValue(5) # 실제 수신 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    "상품관리구분코드": obj.GetHeaderValue(0, 0),        # (string)
                    "주문번호": obj.GetHeaderValue(1, 0),                # (long)
                    "원주문번호": obj.GetHeaderValue(2, 0),              # (long)
                    "종목코드": obj.GetHeaderValue(3, 0),                # (string)
                    "종목명": obj.GetHeaderValue(4, 0),                  # (string)
                    "주문구분내용": obj.GetHeaderValue(5, 0),            # (string)
                    "주문수량": obj.GetHeaderValue(6, 0),                # (long)
                    "주문단가": obj.GetHeaderValue(7, 0),                # (long)
                    "체결수량": obj.GetHeaderValue(8, 0),                # (long)
                    "신용구분": obj.GetHeaderValue(9, 0),                # (string)
                    "계좌번호": obj.GetHeaderValue(10, 0),               # (string)
                    "정정취소가능수량": obj.GetHeaderValue(11, 0),        # (long)
                    "매매구분코드": obj.GetHeaderValue(13, 0),           # (string)
                    "대출일": obj.GetHeaderValue(17, 0),                 # (string)
                    "주문입력매체코드": obj.GetHeaderValue(18, 0),        # (string)
                    "주문호가구분코드내용": obj.GetHeaderValue(19, 0),    # (string)
                    "주문호가구분코드": obj.GetHeaderValue(21, 0),        # (string)
                    "주문구분코드": obj.GetHeaderValue(22, 0),           # (string)
                    "주문구분내용_상세": obj.GetHeaderValue(23, 0),       # (string)
                    "현금신용대용구분코드": obj.GetHeaderValue(24, 0),     # (string)
                    "주문종가구분코드": obj.GetHeaderValue(25, 0),        # (string)
                    "주문입력매체코드내용": obj.GetHeaderValue(26, 0),    # (string)
                    "정정주문수량": obj.GetHeaderValue(27, 0),           # (long)
                    "취소주문수량": obj.GetHeaderValue(28, 0),           # (long)
                    "조건단가": obj.GetHeaderValue(29, 0),               # (double)
                    "주문접수결과": obj.GetHeaderValue(30, 0),           # (string)
                    "거래소주문유형": obj.GetHeaderValue(31, 0)          # (string)
                }
                results.append(item)

            # 5. 연속 데이터 유무 확인 (Paging)
            if obj.Continue == False:
                break
                
        return results
    
    def get_today_history_list(self, 
                         stock_code:str="", 
                         ord_no:int=0, 
                         sort_flag:str='1',
                         rqst_count=20,
                         type_code:str='2',
                         mkt_type:str='1'):
        """금일계좌별주문/체결내역조회데이터를요청하고수신한다

        Args:
            stock_code (str, optional): 종목코드[default:""]- 생략시전종목에대해서조회가됨
            ord_no (int, optional): 시작주문번호[default:0] - 생략시전주문에대해서조회가됨
            sort (str, optional): 정렬구분코드[default:'1'] - '0' 순차, '1' 역순(최근순)
            count (int, optional): 요청개수[default:7] - 최대 20개
            type_code (str, optional): 조회구분코드[default:'2'] - '1' 단가별, '2' 건별, '3' 합산
            mkt_type (str, optional): 거래소주문유형[default:'1'] - '0' 전체, '1' KRX, '2' NXT

        Returns:
            _type_: _description_
        """
        obj = win32com.client.Dispatch("CpTrade.CpTd5341")          # 주문 및 체결 내역 v
        # 1. 입력 데이터 설정
        obj.SetInputValue(0, self.acc_no)       # 계좌번호
        obj.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        obj.SetInputValue(2, stock_code)   # 종목코드
        obj.SetInputValue(3, ord_no)            # 시작주문번호 (0: 처음부터)
        obj.SetInputValue(4, sort_flag)     # 정렬구분: '1' 역순(최근순)
        obj.SetInputValue(5, rqst_count)        # 요청개수
        obj.SetInputValue(6, type_code)     # 조회구분: '2' 건별
        obj.SetInputValue(7, mkt_type)     # 거래소유형: '0' 전체

        results = []
        
        while True:
            # 2. 데이터 요청
            ret = obj.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = obj.GetHeaderValue(6) # 실제 수신 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    "상품관리구분코드": obj.GetHeaderValue(0, 0),        # (string)
                    "주문번호": obj.GetHeaderValue(1, 0),                # (long)
                    "원주문번호": obj.GetHeaderValue(2, 0),              # (long)
                    "종목코드": obj.GetHeaderValue(3, 0),                # (string)
                    "종목이름": obj.GetHeaderValue(4, 0),                # (string)
                    "주문내용": obj.GetHeaderValue(5, 0),                # (string)
                    "주문호가구분코드내용": obj.GetHeaderValue(6, 0),    # (string)
                    "주문수량": obj.GetHeaderValue(7, 0),                # (long)
                    "주문단가": obj.GetHeaderValue(8, 0),                # (long)
                    "총체결수량": obj.GetHeaderValue(9, 0),              # (long)
                    "체결수량": obj.GetHeaderValue(10, 0),               # (long)
                    "체결단가": obj.GetHeaderValue(11, 0),               # (long)
                    "확인수량": obj.GetHeaderValue(12, 0),               # (long)
                    "정정취소구분내용": obj.GetHeaderValue(13, 0),       # (string)
                    "거부사유내용": obj.GetHeaderValue(14, 0),           # (string)
                    "채권매수일": obj.GetHeaderValue(16, 0),             # (string)
                    "거래세과세구분내용": obj.GetHeaderValue(17, 0),     # (string)
                    "현금신용대용구분내용": obj.GetHeaderValue(18, 0),   # (string)
                    "주문입력매체코드내용": obj.GetHeaderValue(19, 0),   # (string)
                    "종합계좌": obj.GetHeaderValue(21, 0),               # (string)
                    "정정취소가능수량": obj.GetHeaderValue(22, 0),       # (long)
                    "매매구분내용": obj.GetHeaderValue(24, 0),           # (string)
                    "대출일": obj.GetHeaderValue(27, 0),                 # (string)
                    "거래소거부사유내용": obj.GetHeaderValue(28, 0),     # (string)
                    "주문구분코드": obj.GetHeaderValue(29, 0),           # (string)
                    "주문구분내용": obj.GetHeaderValue(30, 0),           # (string)
                    "현금신용대용구분코드": obj.GetHeaderValue(31, 0),   # (string)
                    "주문입력매체코드": obj.GetHeaderValue(32, 0),       # (string)
                    "거래세과세구분코드": obj.GetHeaderValue(33, 0),     # (string)
                    "주문호가구분코드": obj.GetHeaderValue(34, 0),       # (string)
                    "매매구분코드": obj.GetHeaderValue(35, 0),           # (string)
                    "정정취소구분코드": obj.GetHeaderValue(36, 0),       # (string)
                    "거래세": obj.GetHeaderValue(38, 0),                 # (double)
                    "농어촌특별세": obj.GetHeaderValue(39, 0),           # (double)
                    "조건단가": obj.GetHeaderValue(40, 0),               # (double)
                    "거래소주문유형": obj.GetHeaderValue(41, 0),         # (string)
                    "체결상세시분초": obj.GetHeaderValue(42, 0)          # (string)
                }



                results.append(item)

            # 5. 연속 데이터 유무 확인 (Paging)
            if obj.Continue == False:
                break
            
            # 다음 조회를 위해 잠시 대기 (TR 과부하 방지)
            time.sleep(0.2)
            
        return results
    
    def get_profit_loss_data(self):
        """
        당일 매매 손익 데이터를 조회하여 (요약 정보, 종목별 상세)를 반환합니다.
        """
        obj = win32com.client.Dispatch("CpTrade.CpTd6032")      # 당일 매매 손익(체결 기준) v
        # 1. 입력 데이터 설정
        obj.SetInputValue(0, self.acc_no)       # 계좌번호
        obj.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        obj.SetInputValue(2, "1")          # 거래소구분: "1" KRX

        # 2. 데이터 요청
        ret = obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None, []

        # 3. 헤더 정보 (당일 전체 요약) 추출
        # 단위 주의: 헤더의 손익 금액은 '천원' 단위입니다.
        summary = {
            '조회 요청건수': obj.GetHeaderValue(0),         # 조회 요청건수
            '잔량평가손익금액': obj.GetHeaderValue(1),     # 잔량평가손익금액 (단위: 천원)
            '매도실현손익금액': obj.GetHeaderValue(2), # 매도실현손익금액 (단위: 천원)
            '수익률': obj.GetHeaderValue(3),       # 총 수익률 (float)
        }

        # 4. 종목별 상세 내역 추출
        results = []
        for i in range(summary['조회 요청건수']):
            item = {
                '종목명': obj.GetDataValue(0, i),               # 종목명 (string) 
                '신용일자': obj.GetDataValue(1, i),             # 종목명 (string) 
                '전일잔고': obj.GetDataValue(2, i),             # 전일잔고 (string) 
                '금일매수수량': obj.GetDataValue(3, i),         # 금일매수수량 (string) 
                '금일매도수량': obj.GetDataValue(4, i),         # 금일매도수량 (string) 
                '금일잔고': obj.GetDataValue(5, i),             # 금일잔고 (string) 
                '평균매입단가': obj.GetDataValue(6, i),         # 평균매입단가 (string) 
                '평균매도단가': obj.GetDataValue(7, i),         # 평균매도단가 (string) 
                '현재가': obj.GetDataValue(8, i),               # 현재가 (string) 
                '잔량평가손익': obj.GetDataValue(9, i),         # 잔량평가손익 (string) 
                '매도실현손익': obj.GetDataValue(10, i),        # 매도실현손익 (string) 
                '수익률': obj.GetDataValue(11, i),              # 수익률(%) (float)
                '종목코드': obj.GetDataValue(12, i),            # 종목코드 (string) 
                'NXT거래가능여부': obj.GetDataValue(13, i),     # NXT 거래가능여부 (Y/N) 
            }
            results.append(item)
            
        return summary, results
        
    def get_balance_data(self, yield_type="2"):
        """
        잔고 및 평가 현황 데이터를 요청하고 (요약 정보, 종목 리스트)를 반환합니다.
        acc_no: 계좌번호
        acc_flag: 상품관리구분코드
        yield_type: "1" (100% 기준), "2" (0% 기준 - 일반적)
        """
        obj = win32com.client.Dispatch("CpTrade.CpTd6033")  # 계좌 잔고 및 평가 현황 v
        
        # 1. 입력 데이터 설정
        obj.SetInputValue(0, self.acc_no)       # 계좌번호
        obj.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        obj.SetInputValue(2, 50)           # 요청건수 (최대 50개)
        obj.SetInputValue(3, yield_type)   # 수익률 구분 ("2": 0% 기준)
        obj.SetInputValue(4, "1")          # 시장 구분 ("1": KRX)

        # 2. 데이터 요청
        ret = obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None, []

        # 3. 헤더 정보 (계좌 요약) 추출
        summary = {
                "계좌명": obj.GetHeaderValue(0),                # (string)
                "결제잔고수량": obj.GetHeaderValue(1),           # (long)
                "체결잔고수량": obj.GetHeaderValue(2),           # (long)
                "총평가금액": obj.GetHeaderValue(3),             # (longlong) - 예수금, 대주, 잔고 합산
                "평가손익": obj.GetHeaderValue(4),               # (longlong)
                "대출금액": obj.GetHeaderValue(6),               # (longlong)
                "수신개수": obj.GetHeaderValue(7),               # (long) - 종목별 잔고 개수
                "수익율": obj.GetHeaderValue(8),                 # (double)
                "D2예상예수금": obj.GetHeaderValue(9),           # (longlong)
                "대주평가금액": obj.GetHeaderValue(10),          # (longlong) - 총평가 내 금액
                "잔고평가금액": obj.GetHeaderValue(11),          # (longlong) - 총평가 내 금액
                "대주금액": obj.GetHeaderValue(12)               # (longlong)
            }

        # 4. 개별 종목 리스트 추출
        rqst_count = summary['수신개수'] # 수신 종목 개수
        stocks = []
        for i in range(rqst_count):
            item = {
                "종목명": obj.GetDataValue(0, i),            # (string)
                "신용구분": obj.GetDataValue(1, i),          # (char)
                "대출일": obj.GetDataValue(2, i),            # (string)
                "결제잔고수량": obj.GetDataValue(3, i),       # (long)
                "결제장부단가": obj.GetDataValue(4, i),       # (long)
                "전일체결수량": obj.GetDataValue(5, i),       # (long)
                "금일체결수량": obj.GetDataValue(6, i),       # (long)
                "체결잔고수량": obj.GetDataValue(7, i),       # (long)
                "평가금액": obj.GetDataValue(9, i),          # (longlong) - 천원 미만 내림
                "평가손익": obj.GetDataValue(10, i),         # (longlong) - 천원 미만 내림
                "수익률": obj.GetDataValue(11, i),           # (double)
                "종목코드": obj.GetDataValue(12, i),         # (string)
                "주문구분": obj.GetDataValue(13, i),         # (char)
                "매도가능수량": obj.GetDataValue(15, i),      # (long)
                "만기일": obj.GetDataValue(16, i),           # (string)
                "체결장부단가": obj.GetDataValue(17, i),      # (double)
                "손익단가": obj.GetDataValue(18, i),         # (longlong)
                "NXT거래가능여부": obj.GetDataValue(19, i)    # (string) - Y/N
            }
            stocks.append(item)
            
        return summary, stocks
        
        
if __name__ == "__main__":
    am = AccountManager()
    summary, data = am.get_balance_data()
    print("summay")
    for key, value in summary.items():
        print(f"{key}: {value}")
    
    for item in data:
        for key, value in item.items():
            print(f"{key}: {value}")
        print("\n")
    
    
    
        
        