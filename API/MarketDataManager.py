import win32com.client
import time

# 1. 코스닥 실시간 이벤트를 처리할 핸들러 클래스
class KOSDAQStatusHandler:
    def set_params(self, obj):
        self.obj = obj

    def OnReceived(self):
        # 코스닥 실시간 데이터 수신 시 호출
        data = {
            'market': 'KOSDAQ',
            'rising': self.obj.GetHeaderValue(0),        # 상승 종목 수
            'upper_limit': self.obj.GetHeaderValue(1),   # 상한 종목 수
            'unchanged': self.obj.GetHeaderValue(2),     # 보합 종목 수
            'falling': self.obj.GetHeaderValue(3),       # 하락 종목 수
            'lower_limit': self.obj.GetHeaderValue(4),   # 하한 종목 수
            
            # 단위 보정 (천주 -> 주, 백만 -> 원)
            'total_vol': self.obj.GetHeaderValue(5) * 1000,
            'total_amt': self.obj.GetHeaderValue(6) * 1000000,
        }
        
        # 실시간 ADR(등락비율) 계산
        print(f"📈 코스닥 시장동향: 상승 {data['rising']} / 하락 {data['falling']}")
        
        # 봇의 전역 상태 업데이트 또는 전략 로직 호출
        
# 1. 실시간 이벤트를 처리할 핸들러 클래스
class KOSPIStatusHandler:
    def set_params(self, obj):
        self.obj = obj

    def OnReceived(self):
        # 실시간 데이터 수신 시 호출됨
        data = {
            'rising': self.obj.GetHeaderValue(0),       # 상승 종목 수
            'upper_limit': self.obj.GetHeaderValue(1),   # 상한 종목 수
            'unchanged': self.obj.GetHeaderValue(2),     # 보합 종목 수
            'falling': self.obj.GetHeaderValue(3),       # 하락 종목 수
            'lower_limit': self.obj.GetHeaderValue(4),   # 하한 종목 수
            'total_vol': self.obj.GetHeaderValue(5) * 1000,    # 총 거래량 (단위 보정)
            'total_amt': self.obj.GetHeaderValue(6) * 1000000, # 총 거래대금 (단위 보정)
        }
        # 여기서 봇의 전역 상태를 업데이트하거나 로그를 남깁니다.
        print(f"📈 코스피시장동향: 상승 {data['rising']} / 하락 {data['falling']}")
        
class MarketDataManager:
    """
    [DsCbo1.StockMst / CpSysDib.StockMstM / DsCbo1.StockMst2] 통합 시세 관리자
    단일 종목, 다중 종목, 상세 호가 정보를 효율적으로 관리합니다.
    """
    def __init__(self):
        # 1. 단일 종목 시세 (StockMst)
        self.obj_mst = win32com.client.Dispatch("DsCbo1.StockMst")
        # 2. 다중 종목 시세 (StockMstM - 최대 200개)
        self.obj_mst_m = win32com.client.Dispatch("CpSysDib.StockMstM")
        # 3. 상세 호가 및 기술적 지표 (StockMst2)
        self.obj_mst_2 = win32com.client.Dispatch("DsCbo1.StockMst2")
        self.obj_adr = win32com.client.Dispatch("Dscbo1.StockAdR")
        self.obj_ads = win32com.client.Dispatch("Dscbo1.StockAdS")
        self.obj_adkr = win32com.client.Dispatch("Dscbo1.StockAdKR")
        self.obj_adks = win32com.client.Dispatch("Dscbo1.StockAdKS")
        self.obj_chart = win32com.client.Dispatch("CpSysDib.StockChart")
    
    def subscribe_kosdaq_status(self):
        """코스닥 실시간 등락 현황 구독 시작"""
        handler = win32com.client.WithEvents(self.obj_adks, KOSDAQStatusHandler)
        handler.set_params(self.obj_adks)
        self.obj_adks.Subscribe()
        self.event_adks = handler
        print("[시스템] 코스닥 실시간 현황 구독 시작")

    def unsubscribe_kosdaq_status(self):
        """구독 해지"""
        if self.event_adks:
            self.obj_adks.Unsubscribe()
            self.event_adks = None
            print("[시스템] 코스닥 실시간 현황 구독 해지")
    
    def subscribe_kospi_status(self):
        """거래소 실시간 등락 현황 구독 시작"""
        handler = win32com.client.WithEvents(self.obj_ads, KOSPIStatusHandler)
        handler.set_params(self.obj_ads)
        self.obj_ads.Subscribe()
        self.event_ads = handler
        print("[시스템] 실시간 시장 현황 구독 시작 (KOSPI)")

    def unsubscribe_kospi_status(self):
        """구독 해지"""
        if self.event_ads:
            self.obj_ads.Unsubscribe()
            self.event_ads = None
            print("[시스템] 실시간 시장 현황 구독 해지")
            
    def get_chart_data(self, 
                       stk_code, 
                       target_count, 
                       chart_type='D', 
                       cycle=1, 
                       req_type='2'):
        """
        갯수 기준으로 차트 데이터를 조회하여 리스트로 반환합니다.
        code: 종목코드 (주식 'A...', 업종 'U...', ELW 'J...')
        target_count: 총 요청할 데이터 개수
        chart_type: 'D'(일), 'W'(주), 'M'(월), 'm'(분), 'T'(틱)
        cycle: 주기 (분차트에서 5분봉이면 5 입력)
        req_type: '1'(기간), '2'(개수) - 여기선 개수 기준 기본
        """
        # 1. 요청 필드 설정 (0:날짜, 1:시간, 2:시가, 3:고가, 4:저가, 5:종가, 8:거래량)
        # 중요: GetDataValue는 필드값의 오름차순 인덱스로 반환됨
        fields = [0, 1, 2, 3, 4, 5, 8] 
        
        self.obj_chart.SetInputValue(0, stk_code)
        self.obj_chart.SetInputValue(1, req_type)    # '1':기간, '2':개수
        self.obj_chart.SetInputValue(4, target_count)     # 요청 개수
        self.obj_chart.SetInputValue(5, fields)           # 필드 배열
        self.obj_chart.SetInputValue(6, chart_type)  # 차트 구분
        self.obj_chart.SetInputValue(7, cycle)            # 주기
        self.obj_chart.SetInputValue(9, '1')         # 수정 주가 반영

        results = []
        current_count = 0

        while current_count < target_count:
            # 2. 데이터 요청
            ret = self.obj_chart.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = self.obj_chart.GetHeaderValue(3) # 실제 수신 개수
            field_cnt = self.obj_chart.GetHeaderValue(1)  # 요청한 필드 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    'date': self.obj_chart.GetDataValue(0, i),   # 날짜(0)
                    'time': self.obj_chart.GetDataValue(1, i),   # 시간(1)
                    'open': self.obj_chart.GetDataValue(2, i),   # 시가(2)
                    'high': self.obj_chart.GetDataValue(3, i),   # 고가(3)
                    'low': self.obj_chart.GetDataValue(4, i),    # 저가(4)
                    'close': self.obj_chart.GetDataValue(5, i),  # 종가(5)
                    'vol': self.obj_chart.GetDataValue(6, i),    # 거래량(8) - 인덱스는 오름차순임
                }
                results.append(item)
            
            current_count += recv_count
            print(f"현재 {current_count}개 수신 완료...")

            # 5. 연속 데이터 유무 확인
            if not self.obj_chart.Continue or current_count >= target_count:
                break
            
            # TR 과부하 방지 (연속 요청 시 짧은 휴식)
            time.sleep(0.1)

        return results
    
    def get_kosdaq_status(self):
        """[StockAdKR] 코스닥 시장의 현재 등락 종목 수 및 거래대금을 가져옵니다."""
        ret = self.obj_adkr.BlockRequest()
        
        if ret != 0:
            print(f"코스닥 현황 조회 실패 (에러코드: {ret})")
            return None

        # 코스닥은 인덱스 5, 6이 바로 거래량과 거래대금입니다.
        result = {
            'market': 'KOSDAQ',
            'rising': self.obj_adkr.GetHeaderValue(0),        # 상승 종목 수
            'upper_limit': self.obj_adkr.GetHeaderValue(1),   # 상한 종목 수
            'unchanged': self.obj_adkr.GetHeaderValue(2),     # 보합 종목 수
            'falling': self.obj_adkr.GetHeaderValue(3),       # 하락 종목 수
            'lower_limit': self.obj_adkr.GetHeaderValue(4),   # 하한 종목 수
            
            # 단위 보정 (천주 -> 주, 백만 -> 원)
            'total_vol': self.obj_adkr.GetHeaderValue(5) * 1000,
            'total_amt': self.obj_adkr.GetHeaderValue(6) * 1000000,
        }
        
        # 코스닥 ADR(등락비율) 계산
        if result['falling'] > 0:
            result['adr'] = round((result['rising'] / result['falling'] * 100), 2)
        else:
            result['adr'] = 100.0
            
        return result    

    def get_kospi_status(self):
        """거래소의 현재 등락 종목 수 및 지수 정보를 가져옵니다."""
        # 이 오브젝트는 SetInputValue가 필요 없습니다.
        ret = self.obj_adr.BlockRequest()
        
        if ret != 0:
            print(f"시장 현황 조회 실패 (에러코드: {ret})")
            return None

        # 데이터 추출 및 단위 보정
        result = {
            'rising': self.obj_adr.GetHeaderValue(0),        # 상승 종목 수
            'upper_limit': self.obj_adr.GetHeaderValue(1),   # 상한 종목 수
            'unchanged': self.obj_adr.GetHeaderValue(2),     # 보합 종목 수
            'falling': self.obj_adr.GetHeaderValue(3),       # 하락 종목 수
            'lower_limit': self.obj_adr.GetHeaderValue(4),   # 하한 종목 수
            
            'index': round(self.obj_adr.GetHeaderValue(5), 2),      # 현재 지수
            'index_diff': round(self.obj_adr.GetHeaderValue(6), 2), # 지수 대비
            
            # 주의: 단위 변환 필요
            'total_vol': self.obj_adr.GetHeaderValue(7) * 1000,      # 총 거래량 (천주 -> 주)
            'total_amt': self.obj_adr.GetHeaderValue(9) * 1000000,   # 총 거래대금 (백만 -> 원)
        }
        
        # 시장 심리 지표 (ADR) 계산
        # 상승 종목 수 / 하락 종목 수
        total_move = result['rising'] + result['falling']
        result['adr'] = round((result['rising'] / result['falling'] * 100), 2) if result['falling'] > 0 else 100.0
        
        return result
    
    def get_single_quote(self, 
                         stock_code:str='',
                         mkt_type:str='K'):
        """주식종목의 현재가에 관련된 데이터(10차 호가 포함)

        Args:
            stock_code (str, optional): 종목코드.
            mkt_type (str, optional):  거래소구분[Default:'K'] 'A'전체, 'K'KRX, 'N'NXT.

        Returns:
            _type_: _description_
        """
        self.obj_mst.SetInputValue(0, stock_code)
        self.obj_mst.SetInputValue(1, mkt_type)
        
        ret = self.obj_mst.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None
        
        return {
            'code': self.obj_mst.GetHeaderValue(0),        # 종목코드
            'name': self.obj_mst.GetHeaderValue(1),        # 종목명
            'time': self.obj_mst.GetHeaderValue(4),        # 시간
            'high_limit': self.obj_mst.GetHeaderValue(8),  # 상한가
            'low_limit': self.obj_mst.GetHeaderValue(9),   # 하한가
            'prev_close': self.obj_mst.GetHeaderValue(10), # 전일종가
            'current': self.obj_mst.GetHeaderValue(11),    # 현재가
            'diff': self.obj_mst.GetHeaderValue(12),       # 전일대비
            'open': self.obj_mst.GetHeaderValue(13),       # 시가
            'high': self.obj_mst.GetHeaderValue(14),       # 고가
            'low': self.obj_mst.GetHeaderValue(15),        # 저가
            'ask': self.obj_mst.GetHeaderValue(16),        # 매도호가
            'bid': self.obj_mst.GetHeaderValue(17),        # 매수호가
            'volume': self.obj_mst.GetHeaderValue(18),     # 누적거래량
            'amount': self.obj_mst.GetHeaderValue(19),     # 누적거래대금
            'per': self.obj_mst.GetHeaderValue(28),        # PER
            'eps': self.obj_mst.GetHeaderValue(20),        # EPS
            'status': self.obj_mst.GetHeaderValue(68),     # 거래정지여부 ('Y'/'N')
            'vi_base': self.obj_mst.GetHeaderValue(80),    # 정적VI 발동 예상기준가
        }

    def get_multi_quotes(self, code_list):
        """주식복수종목에대해간단한내용을일괄조회요청하고수신한다"""
         # 대비구분코드 매핑
        map_diff_status = {
            1: '상한', 2: '상승', 3: '보합', 4: '하한', 5: '하락',
            6: '기세상한', 7: '기세상승', 8: '기세하한', 9: '기세하락'
        }
        # 장구분플래그 매핑
        map_market_status = {
            '0': '장외', '1': '동시호가', '2': '장중'
        }
        if not code_list:
            print("조회할 종목코드가 없습니다.")
            return False

        if len(code_list) > 110:
            print("최대 조회 가능 종목수는 110개입니다. (현재: {}개)".format(len(code_list)))
            code_list = code_list[:110]

        # 1. 입력 데이터 설정 (종목코드들을 하나의 문자열로 결합)
        # 예: "A005930A000660"
        codes_str = "".join(code_list)
        self.obj_mst_m.SetInputValue(0, codes_str)

        # 2. 데이터 요청
        ret = self.obj_mst_m.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None
        
        count = self.obj_mst_m.GetHeaderValue(0)
        
        results = []
        for i in range(count):
            diff_code = self.obj_mst_m.GetDataValue(3, i)
            market_flag = self.obj_mst_m.GetDataValue(8, i)
            
            item = {
                'code': self.obj_mst_m.GetDataValue(0, i),          # 종목코드
                'name': self.obj_mst_m.GetDataValue(1, i),          # 종목명
                'diff': self.obj_mst_m.GetDataValue(2, i),          # 대비
                'diff_status': map_diff_status.get(diff_code, diff_code), # 대비구분
                'current': self.obj_mst_m.GetDataValue(4, i),       # 현재가
                'ask': self.obj_mst_m.GetDataValue(5, i),           # 매도호가
                'bid': self.obj_mst_m.GetDataValue(6, i),           # 매수호가
                'volume': self.obj_mst_m.GetDataValue(7, i),        # 거래량
                'market_status': map_market_status.get(market_flag, market_flag), # 장구분
                'exp_price': self.obj_mst_m.GetDataValue(9, i),     # 예상체결가
                'exp_diff': self.obj_mst_m.GetDataValue(10, i),     # 예상체결가 전일대비
                'exp_vol': self.obj_mst_m.GetDataValue(11, i),      # 예상체결수량
            }
            results.append(item)
            
        return results

    def get_hoga_detail(self, 
                        code_list:list, 
                        market_type:str='K'):
        """[StockMst2] 주식복수종목에대해일괄조회를요청하고수신한다"""
        map_status = {
            '1': '상한', '2': '상승', '3': '보합', '4': '하한', '5': '하락',
            '6': '기세상한', '7': '기세상승', '8': '기세하한', '9': '기세하락'
        }
        
        if not code_list:
            print("조회할 종목코드가 없습니다.")
            return False

        if len(code_list) > 110:
            print(f"최대 조회 가능 종목수는 110개입니다. (현재: {len(code_list)}개)")
            code_list = code_list[:110]

        # 1. 입력 데이터 설정 (구분자 ',' 사용)
        codes_str = ",".join(code_list)
        self.obj_mst_2.SetInputValue(0, codes_str)
        self.obj_mst_2.SetInputValue(1, market_type) # char 타입이므로 ord() 사용

        # 2. 데이터 요청
        ret = self.obj_mst_2.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return None
        
        count = self.obj_mst_2.GetHeaderValue(0) # 수신된 종목 수
        
        results = []
        for i in range(count):
            status_code = self.obj_mst_2.GetDataValue(5, i)      # 상태구분
            market_flag = self.obj_mst_2.GetDataValue(25, i)     # 동시호가구분
            exp_status_code = self.obj_mst_2.GetDataValue(28, i) # 예상체결상태구분
            
            item = {
                'code': self.obj_mst_2.GetDataValue(0, i),           # 종목코드
                'name': self.obj_mst_2.GetDataValue(1, i),           # 종목명
                'time': self.obj_mst_2.GetDataValue(2, i),           # 시간(HHMM)
                'current': self.obj_mst_2.GetDataValue(3, i),        # 현재가
                'diff': self.obj_mst_2.GetDataValue(4, i),           # 전일대비
                'status': map_status.get(status_code, status_code), # 상태
                'open': self.obj_mst_2.GetDataValue(6, i),           # 시가
                'high': self.obj_mst_2.GetDataValue(7, i),           # 고가
                'low': self.obj_mst_2.GetDataValue(8, i),            # 저가
                'ask': self.obj_mst_2.GetDataValue(9, i),            # 매도호가
                'bid': self.obj_mst_2.GetDataValue(10, i),           # 매수호가
                'volume': self.obj_mst_2.GetDataValue(11, i),        # 거래량(1주 단위)
                'amount': self.obj_mst_2.GetDataValue(12, i),        # 거래대금(원 단위)
                'total_ask_rem': self.obj_mst_2.GetDataValue(13, i), # 총매도잔량
                'total_bid_rem': self.obj_mst_2.GetDataValue(14, i), # 총매수잔량
                'listed_stock': self.obj_mst_2.GetDataValue(17, i),  # 상장주식수
                'strength': self.obj_mst_2.GetDataValue(21, i),      # 체결강도
                'market_type': '동시호가' if market_flag == '1' else '장중',
                'exp_price': self.obj_mst_2.GetDataValue(26, i),     # 예상체결가
                'exp_status': map_status.get(exp_status_code, exp_status_code) # 예상상태
            }
            results.append(item)
            
        return results