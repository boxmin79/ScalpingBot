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
            
        # 1. 자금 및 매매 가능 확인 관련
        self.obj_deposit_settlement = win32com.client.Dispatch("CpTrade.CpTd0732")     # 결제예정 예수금 가계산
        self.obj_buyable = win32com.client.Dispatch("CpTrade.CpTdNew5331A")          # 매수 가능 금액/수량 조회 v
        self.obj_sellalble = win32com.client.Dispatch("CpTrade.CpTdNew5331B")         # 매도 가능 수량 조회 v

        # 2. 주문 및 체결 내역 관련
        self.obj_unexecuted_orders = win32com.client.Dispatch("CpTrade.CpTd5339")      # 미체결 잔량 조회 v
        self.obj_order_history = win32com.client.Dispatch("CpTrade.CpTd5341")          # 주문 및 체결 내역 v
        self.obj_execution_details = win32com.client.Dispatch("CpTrade.CpTd5342")      # 금일/전일 체결 기준 내역 v

        # 3. 계좌 잔고 및 손익 평가 관련
        self.obj_portfolio_status = win32com.client.Dispatch("CpTrade.CpTd6033")       # 계좌 잔고 및 평가 현황 v
        self.obj_daily_profit_loss = win32com.client.Dispatch("CpTrade.CpTd6032")      # 당일 매매 손익(체결 기준) v

        self.acc_no = acc_no
        self.acc_flag = acc_flag
    
    def get_expected_deposit(self):
        """
        [CpTd0732] 주식 결제예정 예수금 가계산 조회
        D+2 결제 기준의 실제 가용 자금을 확인합니다.
        """
        # 1. 입력값 설정
        self.obj_deposit_settlement.SetInputValue(0,  self.acc_no)
        self.obj_deposit_settlement.SetInputValue(1, self.acc_flag)

        # 2. 데이터 요청
        res = self.obj_deposit_settlement.BlockRequest()
        
        if res != 0:
            print(f"❌ CpTd0732 요청 실패 (에러코드: {res})")
            return None

        # 3. 주요 데이터 추출 (총 66개 필드 중 핵심 필드 위주)
        result = {
            'account_no': self.obj_deposit_settlement.GetHeaderValue(0),
            'account_name': self.obj_deposit_settlement.GetHeaderValue(2),
            'current_deposit': self.obj_deposit_settlement.GetHeaderValue(3),   # 현재 예수금
            'receivables': self.obj_deposit_settlement.GetHeaderValue(4),
            
            # 오늘 매수/매도 합계 (정산용)
            'today_sell': self.obj_deposit_settlement.GetHeaderValue(48),       # 금일 합계 매도금
            'today_buy': self.obj_deposit_settlement.GetHeaderValue(49),        # 금일 합계 매수금
            'today_fee': self.obj_deposit_settlement.GetHeaderValue(50),        # 금일 합계 수수료
            'today_tax': self.obj_deposit_settlement.GetHeaderValue(51),        # 금일 합계 제세금
            'today_total': self.obj_deposit_settlement.GetHeaderValue(52),      # 금일 합계 정산금
            
            # 결제일 기준 가용 자금 (스캘핑 시 가장 중요)
            'd1_deposit': self.obj_deposit_settlement.GetHeaderValue(64),       # 익영업일(D+1) 예상 예수금
            'd2_deposit': self.obj_deposit_settlement.GetHeaderValue(66),       # 결제일(D+2) 예상 예수금
        }
        
        return result
    
    def get_buyable_data(self, 
                         stock_code: str, 
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
        # 1. 입력 데이터 설정
        self.obj_buyable.SetInputValue(0, self.acc_no)       # 계좌번호
        self.obj_buyable.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        self.obj_buyable.SetInputValue(2, stock_code)         # 종목코드
        self.obj_buyable.SetInputValue(3, quote_type)         # 호가구분 (보통/시장가 등)
        self.obj_buyable.SetInputValue(4, price)              # 주문단가 (시장가면 0)
        self.obj_buyable.SetInputValue(5, 'Y')                # 증거금 100% 징수여부 (N: 계좌별 설정)
        self.obj_buyable.SetInputValue(6, query_type)         # 조회구분 ('2': 수량조회 중심)

        # 2. 데이터 요청
        ret = self.obj_buyable.BlockRequest()
        if ret != 0:
            print(f"매수 가능 조회 실패 (에러코드: {ret})")
            return None

        # 3. 결과 파싱 (스캘핑에 필요한 핵심 필드 위주)
        # 이 오브젝트는 HeaderValue에서 데이터를 제공합니다.
        result = {
            'code': self.obj_buyable.GetHeaderValue(0),       # 종목코드
            'name': self.obj_buyable.GetHeaderValue(1),       # 종목명
            'margin_type': self.obj_buyable.GetHeaderValue(2), # 증거금률구분
            
            # --- 금액 관련 ---
            'cash_buyable_amt': self.obj_buyable.GetHeaderValue(10), # 현금주문 가능금액
            'total_deposit': self.obj_buyable.GetHeaderValue(45),    # 예수금
            'available_deposit': self.obj_buyable.GetHeaderValue(47), # 가능예수금
            
            # --- 수량 관련 (query_type '2'일 때 정확) ---
            'cash_buyable_qty': self.obj_buyable.GetHeaderValue(18),  # 🎯 현금주문 가능수량 (가장 중요)
            
        }

        return result
    
    def get_sellable_qty(self, 
                         stock_code:str='',
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
        # 1. 입력 데이터 설정
        self.obj_sellalble.SetInputValue(0, self.acc_no)       # 계좌번호
        self.obj_sellalble.SetInputValue(1, self.acc_flag)     # 상품구분
        self.obj_sellalble.SetInputValue(2, stock_code)         # 특정 종목 조회
        self.obj_sellalble.SetInputValue(3, stk_bond_flag)          # 1: 주식
        self.obj_sellalble.SetInputValue(4, cash_credit_flag)  # 1: 현금(현금주문 가능 수량)
        self.obj_sellalble.SetInputValue(10, rqst_count)              # 요청개수


        # 2. 데이터 요청
        self.obj_sellalble.BlockRequest()

       # 3. 결과 파싱
        count = self.obj.GetHeaderValue(0)
        if count == 0:
            return 0
        data = {
            'code': self.obj_sellalble.GetDataValue(0, 0),
            'name': self.obj_sellalble.GetDataValue(1, 0),
            'sellable_qty': self.obj_sellalble.GetDataValue(12, 0)
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
        # 1. 입력 데이터 설정
        self.obj_execution_details.SetInputValue(0, self.acc_no)      # 계좌번호
        self.obj_execution_details.SetInputValue(1, self.acc_flag) # 상품구분
        self.obj_execution_details.SetInputValue(2, rqst_count)               # 한 번에 최대 20개 조회
        self.obj_execution_details.SetInputValue(3, date_flag)        # 금일/전일 구분
        self.obj_execution_details.SetInputValue(4, target_code)      # 종목코드 (기본 전체)

        trades = []
        
        while True:
            # 2. 데이터 요청
            self.obj_execution_details.BlockRequest()

            # 헤더 정보 (요청 결과 요약)
            recv_count = self.obj_execution_details.GetHeaderValue(8)      # 수신 개수
            trade_date = self.obj_execution_details.GetHeaderValue(3) # 매매일
            
            # 3. 데이터 파싱 (아이템별 상세 정보)
            for i in range(recv_count):
                item = {
                    'code': self.obj_execution_details.GetDataValue(0, i),      # 종목코드
                    'name': self.obj_execution_details.GetDataValue(1, i).strip(), # 종목명
                    'qty': self.obj_execution_details.GetDataValue(3, i),        # 체결수량
                    'price': self.obj_execution_details.GetDataValue(28, i),     # 🎯 체결단가
                    'side': '매도' if self.obj_execution_details.GetDataValue(10, i) == "1" else '매수', # 매매구분
                    'fee': self.obj_execution_details.GetDataValue(4, i),        # 수수료
                    'tax': self.obj_execution_details.GetDataValue(12, i),       # 거래세
                    'settle_amount': self.obj_execution_details.GetDataValue(24, i) # 🎯 정산금액 (세금/수수료 반영)
                }
                trades.append(item)

            # 4. 연속 데이터 처리 (20개가 넘을 경우)
            if not self.obj_execution_details.Continue:
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
        # 1. 입력 데이터 설정
        self.obj_unexecuted_orders.SetInputValue(0, self.acc_no)       # 계좌번호
        self.obj_unexecuted_orders.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        self.obj_unexecuted_orders.SetInputValue(3, stock_code)   # 종목코드
        self.obj_unexecuted_orders.SetInputValue(4, ord_flag)          # 주문구분: "0" 전체
        self.obj_unexecuted_orders.SetInputValue(5, sort_flag)          # 정렬구분: "0" 순차
        self.obj_unexecuted_orders.SetInputValue(6, close_type)          # 주문종가구분: "0" 전체
        self.obj_unexecuted_orders.SetInputValue(7, rqst_count)        # 요청개수
        self.obj_unexecuted_orders.SetInputValue(8, mkt_type)          # 거래소유형: "0" 전체

        results = []
        
        while True:
            # 2. 데이터 요청
            ret = self.obj_unexecuted_orders.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = self.obj_unexecuted_orders.GetHeaderValue(5) # 실제 수신 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    'order_no': self.obj_unexecuted_orders.GetDataValue(1, i),      # 주문번호
                    'org_order_no': self.obj_unexecuted_orders.GetDataValue(2, i),  # 원주문번호
                    'code': self.obj_unexecuted_orders.GetDataValue(3, i),          # 종목코드
                    'name': self.obj_unexecuted_orders.GetDataValue(4, i),          # 종목명
                    'content': self.obj_unexecuted_orders.GetDataValue(5, i),       # 주문내용
                    'qty': self.obj_unexecuted_orders.GetDataValue(6, i),           # 주문수량
                    'price': self.obj_unexecuted_orders.GetDataValue(7, i),         # 주문단가
                    'exec_qty': self.obj_unexecuted_orders.GetDataValue(8, i),      # 체결수량
                    'cancelable_qty': self.obj_unexecuted_orders.GetDataValue(11, i), # 정정취소가능수량 (핵심)
                    'side_code': self.obj_unexecuted_orders.GetDataValue(13, i),    # 매매구분 (1:매도, 2:매수)
                    'order_type': self.obj_unexecuted_orders.GetDataValue(21, i),   # 주문호가구분코드
                    'result_status': self.obj_unexecuted_orders.GetDataValue(30, i) # 주문접수결과 (0:대기, 1:정상, 2:접수)
                }
                results.append(item)

            # 5. 연속 데이터 유무 확인 (Paging)
            if self.obj_unexecuted_orders.Continue == False:
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
        # 1. 입력 데이터 설정
        self.obj_order_history.SetInputValue(0, self.acc_no)       # 계좌번호
        self.obj_order_history.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        self.obj_order_history.SetInputValue(2, stock_code)   # 종목코드
        self.obj_order_history.SetInputValue(3, ord_no)            # 시작주문번호 (0: 처음부터)
        self.obj_order_history.SetInputValue(4, sort_flag)     # 정렬구분: '1' 역순(최근순)
        self.obj_order_history.SetInputValue(5, rqst_count)        # 요청개수
        self.obj_order_history.SetInputValue(6, type_code)     # 조회구분: '2' 건별
        self.obj_order_history.SetInputValue(7, mkt_type)     # 거래소유형: '0' 전체

        results = []
        
        while True:
            # 2. 데이터 요청
            ret = self.obj_order_history.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = self.obj_order_history.GetHeaderValue(6) # 실제 수신 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    'order_no': self.obj_order_history.GetDataValue(1, i),      # 주문번호
                    'org_order_no': self.obj_order_history.GetDataValue(2, i),  # 원주문번호
                    'code': self.obj_order_history.GetDataValue(3, i),          # 종목코드
                    'name': self.obj_order_history.GetDataValue(4, i),          # 종목이름
                    'content': self.obj_order_history.GetDataValue(5, i),       # 주문내용
                    'qty': self.obj_order_history.GetDataValue(7, i),           # 주문수량
                    'price': self.obj_order_history.GetDataValue(8, i),         # 주문단가
                    'exec_total': self.obj_order_history.GetDataValue(9, i),    # 총체결수량
                    'exec_qty': self.obj_order_history.GetDataValue(10, i),     # 이번 체결수량
                    'exec_price': self.obj_order_history.GetDataValue(11, i),   # 체결단가
                    'confirm_qty': self.obj_order_history.GetDataValue(12, i),  # 확인수량
                    'side_code': self.obj_order_history.GetDataValue(35, i),    # 매매구분코드 (1:매도, 2:매수)
                    'type_code': self.obj_order_history.GetDataValue(36, i),    # 정정취소구분코드 (1:정상, 2:정정, 3:취소)
                    'time': self.obj_order_history.GetDataValue(42, i),         # 체결상세 시분초
                }
                results.append(item)

            # 5. 연속 데이터 유무 확인 (Paging)
            if self.obj_order_history.Continue == False:
                break
            
            # 다음 조회를 위해 잠시 대기 (TR 과부하 방지)
            time.sleep(0.2)
            
        return results
    
    def get_profit_loss_data(self):
        """
        당일 매매 손익 데이터를 조회하여 (요약 정보, 종목별 상세)를 반환합니다.
        """
        
        # 1. 입력 데이터 설정
        self.obj_daily_profit_loss.SetInputValue(0, self.acc_no)       # 계좌번호
        self.obj_daily_profit_loss.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        self.obj_daily_profit_loss.SetInputValue(2, "1")          # 거래소구분: "1" KRX

        # 2. 데이터 요청
        ret = self.obj_daily_profit_loss.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None, []

        # 3. 헤더 정보 (당일 전체 요약) 추출
        # 단위 주의: 헤더의 손익 금액은 '천원' 단위입니다.
        summary = {
            'req_count': self.obj_daily_profit_loss.GetHeaderValue(0),         # 조회 요청건수
            'total_eval_pl': self.obj_daily_profit_loss.GetHeaderValue(1),     # 잔량평가손익금액 (단위: 천원)
            'total_realized_pl': self.obj_daily_profit_loss.GetHeaderValue(2), # 매도실현손익금액 (단위: 천원)
            'total_yield': self.obj_daily_profit_loss.GetHeaderValue(3),       # 총 수익률 (float)
        }

        # 4. 종목별 상세 내역 추출
        results = []
        for i in range(summary['req_count']):
            item = {
                'name': self.obj_daily_profit_loss.GetDataValue(0, i),          # 종목명
                'prev_balance': self.obj_daily_profit_loss.GetDataValue(2, i),  # 전일잔고
                'buy_qty': self.obj_daily_profit_loss.GetDataValue(3, i),       # 금일매수수량
                'sell_qty': self.obj_daily_profit_loss.GetDataValue(4, i),      # 금일매도수량
                'current_balance': self.obj_daily_profit_loss.GetDataValue(5, i),# 금일잔고
                'avg_buy_price': self.obj_daily_profit_loss.GetDataValue(6, i), # 평균매입단가
                'avg_sell_price': self.obj_daily_profit_loss.GetDataValue(7, i),# 평균매도단가
                'current_price': self.obj_daily_profit_loss.GetDataValue(8, i), # 현재가
                'eval_pl': self.obj_daily_profit_loss.GetDataValue(9, i),       # 잔량평가손익
                'realized_pl': self.obj_daily_profit_loss.GetDataValue(10, i),  # 매도실현손익
                'yield': self.obj_daily_profit_loss.GetDataValue(11, i),        # 수익률(%) (float)
                'code': self.obj_daily_profit_loss.GetDataValue(12, i),         # 종목코드
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
        # 1. 입력 데이터 설정
        self.obj_portfolio_status.SetInputValue(0, self.acc_no)       # 계좌번호
        self.obj_portfolio_status.SetInputValue(1, self.acc_flag)     # 상품관리구분코드
        self.obj_portfolio_status.SetInputValue(2, 50)           # 요청건수 (최대 50개)
        self.obj_portfolio_status.SetInputValue(3, yield_type)   # 수익률 구분 ("2": 0% 기준)
        self.obj_portfolio_status.SetInputValue(4, "1")          # 시장 구분 ("1": KRX)

        # 2. 데이터 요청
        ret = self.obj_portfolio_status.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None, []

        # 3. 헤더 정보 (계좌 요약) 추출
        summary = {
            'Account_Name': self.obj_portfolio_status.GetHeaderValue(0), # 0 - (string) 계좌명
            # 1 - (long) 결제잔고수량
            'Execution_Qty': self.obj_portfolio_status.GetHeaderValue(2), # 2 - (long) 체결잔고수량
            'Total_Eval_Amt' : self.obj_portfolio_status.GetHeaderValue(3), # 3 - (longlong) 총 평가금액(단위:원) - 예수금, 대주평가,잔고평가금액등을 감안한 총평가 금액
            'Total_Profit_Loss': self.obj_portfolio_status.GetHeaderValue(4), # 4 - (longlong) 평가손익(단위:원)
            # 5 - 사용하지않음
            'Total_Yield': self.obj_portfolio_status.GetHeaderValue(8), # 8 - (double) 수익율
            'D2_Deposit': self.obj_portfolio_status.GetHeaderValue(9), # 9 - (longlong) D+2 예상예수금
            # 10 - (longlong) 총평가 내 대주평가금액
            'Stock_Eval_Amt': self.obj_portfolio_status.GetHeaderValue(11), # 11 - (longlong) 총평가 내 잔고평가금액 - 잔고에 대한 평가된 금액
            # 12 - (longlong) 대주금액
        }

        # 4. 개별 종목 리스트 추출
        rqst_count = self.obj_portfolio_status.GetHeaderValue(7) # 수신 종목 개수
        stocks = []
        for i in range(rqst_count):
            item = {
                'name': self.obj_portfolio_status.GetDataValue(0, i), # 0 - (string) 종목명
                # 1 - (char)신용구분
                # 2 - (string) 대출일
                # 3 - (long)결제잔고수량
                # 4 - (long)결제장부단가
                # 5 - (long)전일체결수량
                # 6 - (long)금일체결수량
                'total_qty': self.obj_portfolio_status.GetDataValue(7, i), # 7 - (long)체결잔고수량
                'eval_amt': self.obj_portfolio_status.GetDataValue(9, i), # 9 - (longlong)평가금액(단위:원) - 천원미만은내림
                'profit_loss': self.obj_portfolio_status.GetDataValue(10, i), # 10 - (longlong)평가손익(단위:원) - 천원미만은내림
                'yield': round(self.obj_portfolio_status.GetDataValue(11, i), 2), # 11 - (double)수익률
                'code': self.obj_portfolio_status.GetDataValue(12, i), # 12 - (string) 종목코드
                # 13 - (char)주문구분
                'sellable_qty': self.obj_portfolio_status.GetDataValue(15, i), # 15 - (long)매도가능수량
                # 16 - (string) 만기일
                'buy_price': self.obj_portfolio_status.GetDataValue(17, i), # 17 - (double) 체결장부단가
                # 18 - (longlong) 손익단가
                # 19 - (string)  NXT 거래가능여부(Y/N) 
                            }
            stocks.append(item)
            
        return summary, stocks
        
        
if __name__ == "__main__":
    obj_td_util = win32com.client.Dispatch("CpTrade.CpTdUtil")
    init_status = obj_td_util.TradeInit(0)
    if init_status == 0:
        acc_no = obj_td_util.AccountNumber[0]
        acc_flag = obj_td_util.GoodsList(acc_no, 1)[0]
        print(f"acc_no: {acc_no}, acc_flag: {acc_flag}")
        
    am = AccountManager(acc_no, acc_flag)
    # am.get_balance_data()
    data = am.get_expected_deposit()
    print(data)
    
        
        