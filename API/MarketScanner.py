import path_finder
import win32com.client
from API.CpAPI import CreonAPI
import time

class MarketScanner:
    def __init__(self):
        # CreonAPI 인스턴스 생성 (객체 초기화 및 연결 확인)
        self.api = CreonAPI()
        
        # 최종 선별된 종목들이 저장될 통합 캐시
        self.integrated_cache = {
            '7043': [],  # CpSvrNew7043: 등락, 신고가, 강세, 반등 등
            '7034': [],  # CpSvr7034: 큰손(4천만 이상) 수급 비중
            '7049': []   # CpSvr7049: 거래량/거래대금 상위 랭킹
        }
        self.update_integrated_selection()
        
    def _wait_for_request_limit(self):
        """
        [방어 코드] 시세 조회(LT_NONTRADE_REQUEST) 남은 횟수를 확인하고,
        제한에 도달했다면 남은 시간만큼 대기합니다.
        """
        # 1: 시세 및 종목 정보 일반 요청 (15초당 60건 제한)
        remain_count = self.api.obj_cybos.GetLimitRemainCount(1)
        
        if remain_count <= 2:  # 여유 있게 2건 이하로 남았을 때 대기
            # 재계산까지 남은 시간 (ms)
            wait_time = self.api.obj_cybos.LimitRequestRemainTime
            
            if wait_time > 0:
                print(f"⏳ 요청 제한 감지: {wait_time / 1000:.2f}초 대기 후 재개합니다...")
                time.sleep(wait_time / 1000 + 0.1)  # 0.1초 마진 추가
    
    # --- [전략별 종목 선별 메서드] ---
    def get_5d_breakout_leaders(self):
        """전략 1: 5일 신고가 돌파 + 거래대금 상위 (스캘핑용)"""
        return self._fetch_market_movement(
            criteria='6', period='1', sort_by=61, vol_filter='3'
        )

    def get_intraday_strength_stocks(self):
        """전략 2: 당일 시가 대비 강세 (5%~15% 상승)"""
        return self._fetch_market_movement(
            criteria='2', period='0', rate_start=5, rate_end=15, vol_filter='4', sort_by=21
        )

    def get_bottom_bounce_stocks(self):
        """전략 3: 당일 저점 대비 반등 (낙폭 과대 반등)"""
        return self._fetch_market_movement(
            criteria='2', period='2', vol_filter='2', sort_by=21
        )
    def get_continuous_up_stocks(self):
        """전략 4: 연속 상승 일수 상위 (sort_by=31)
        3~4일 이상 꾸준히 오르며 '달리는 말'이 된 종목을 선별합니다.
        """
        return self._fetch_market_movement(
            criteria='2',      # 상승 종목
            sort_by=31,        # 연속일수 상위순
            vol_filter='3'     # 10만주 이상 (유동성 확인)
        )
    
    def get_major_buy_dominance(self):
        """
        코스피('1')와 코스닥('2')에서 4,000만 원('4') 이상 체결 건 중 
        매수 우위 종목을 수집하여 반환합니다.
        """
        combined_results = []
        # 코스피(1), 코스닥(2) 순차 조회
        for m_id in ['1', '2']:
            try:
                # 기존에 만든 _fetch_buy_dominance_ratio 호출
                data = self._fetch_buy_dominance_ratio(market=m_id, amount='4')
                if data:
                    combined_results.extend(data)
            except Exception as e:
                market_name = "KOSPI" if m_id == '1' else "KOSDAQ"
                print(f"{market_name} 수급 데이터 수집 중 오류: {e}")
                
        return combined_results
    
    def get_high_volatility_stocks(self):
        """전략 5: 시고저 대비율 상위 (sort_by=41)
        당일 변동폭이 커서 스캘핑 타점이 많이 나오는 '꿈틀대는' 종목을 선별합니다.
        """
        return self._fetch_market_movement(
            criteria='2',      # 상승 종목
            sort_by=41,        # 시고저폭 상위순
            vol_filter='3'     # 10만주 이상
        )
        
    # --- [통합 데이터 수집 및 캐시 저장 메서드] ---

    def update_integrated_selection(self):
        """
        호출하는 API 오브젝트(TR)별로 리스트를 분리하여 캐시를 업데이트합니다.
        """
        # 캐시 초기화
        for key in self.integrated_cache:
            self.integrated_cache[key] = []
            
        # 1. CpSvrNew7043 관련 전략들 (시장 등락/가격 데이터)
        tr_7043_strategies = [
            ('5D_BREAKOUT', self.get_5d_breakout_leaders),
            ('INTRADAY_STRENGTH', self.get_intraday_strength_stocks),
            ('BOTTOM_BOUNCE', self.get_bottom_bounce_stocks),
            ('CONTINUOUS_UP', self.get_continuous_up_stocks),
            ('HIGH_VOLATILITY', self.get_high_volatility_stocks)
        ]

        for tag, method in tr_7043_strategies:
            stocks = method()
            if stocks:
                for stock in stocks:
                    stock['strategy_tag'] = tag
                    self.integrated_cache['7043'].append(stock)

        # 2. CpSvr7034 관련 전략 (큰손 수급 데이터)
        big_hand_stocks = self.get_major_buy_dominance()
        if big_hand_stocks:
            for stock in big_hand_stocks:
                # 비중 계산 전처리
                total = stock['buy_count'] + stock['sell_count']
                stock['buy_ratio'] = round((stock['buy_count'] / total * 100), 2) if total > 0 else 0
                stock['strategy_tag'] = 'BUY_DOMINANCE_40M'
                self.integrated_cache['7034'].append(stock)

        # 3. CpSvr7049 관련 데이터 (필요 시 추가)
        # self.integrated_cache['7049'] = self._fetch_volume_rank(market='4', selection='A')

        print(f"✅ 통합 업데이트 완료 (7043: {len(self.integrated_cache['7043'])}건, 7034: {len(self.integrated_cache['7034'])}건)")
        return self.integrated_cache
               
    # --- [데이터 수집부: 내부 호출용] ---
    def _fetch_market_movement(self, 
                       market: str = '0', 
                       criteria: str = '6', 
                       date_type: str = '1', 
                       sort_by: int = 61, 
                       admin: str = '1', 
                       vol_filter: str = '0', 
                       period: str = '1', 
                       rate_start: int = 0, 
                       rate_end: int = 0):
        """
        [CpSysDib.CpSvrNew7043] 거래소, 코스닥 등락현황 데이터 요청
        :param market: 0 - (char) 시장구분 ('0':전체, '1':거래소, '2':코스닥, '3':프리보드)
        :param criteria: 1 - (char) 선택기준구분 ('1':상한, '2':상승, '3':보합, '4':하락, '5':하한, '6':신고, '7':신저, '8':상한후하락, '9':하한후상승)
        :param date_type: 2 - (char) 기준일자구분 ('1':당일, '2':전일) *상한, 상승, 보합, 하한, 하락일 때만 유효
        :param sort_by: 3 - (short) 순서구분 (11:코드상위, 21:대비율상위, 51:거래량상위, 61:거래대금상위 등) *51~62는 신고/신저 시만 유효
        :param admin: 4 - (char) 관리구분 ('1':관리제외, '2':관리포함)
        :param vol_filter: 5 - (char) 거래량구분 ('0':전체, '1':1만, '2':5만, '3':10만, '4':50만, '5':100만주 이상)
        :param period: 6 - (char) 기간구분 (신고/신저 시: '1':5일, '2':20일, '3':60일... / 상승/하락 시: '0':시가대비...)
        :param rate_start: 7 - (short) 등락률시작 (상승, 하락인 경우만 유효)
        :param rate_end: 8 - (short) 등락률끝 (상승, 하락인 경우만 유효)
        """
        obj = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")
        
        # [INPUT] 명세서 규격에 따른 입력값 설정
        obj.SetInputValue(0, market)      # 시장구분
        obj.SetInputValue(1, criteria)    # 선택기준구분
        
        # 주의: 선택기준구분이 상한, 상승, 보합, 하한, 하락(1,2,3,4,5)일 경우에만 기준일자구분(2) 설정
        if criteria in ['1', '2', '3', '4', '5']:
            obj.SetInputValue(2, date_type)
            
        # [수정 요청 부분] 51, 52, 61, 62는 신고/신저가('6', '7')일 때만 가능
        if criteria not in ['6', '7'] and sort_by in [51, 52, 61, 62]:
            sort_by = 21  # 조건이 안 맞으면 전일대비율상위(21)로 강제 변경
        obj.SetInputValue(3, sort_by)       # 순서구분
        
        obj.SetInputValue(4, admin)       # 관리구분
        obj.SetInputValue(5, vol_filter)  # 거래량구분
        
        # 주의: 신고, 신저, 상승, 하락인 경우만 기간구분(6) 설정
        if criteria in ['6', '7', '2', '4']:
            obj.SetInputValue(6, period)
            
        # 주의: 상승(2), 하락(4)인 경우만 등락률 시작/끝(7, 8) 설정
        if criteria in ['2', '4']:
            obj.SetInputValue(7, rate_start)
            obj.SetInputValue(8, rate_end)
        
        self._wait_for_request_limit()
        obj.BlockRequest()
        
        recv_count = obj.GetHeaderValue(0)

        
        # [DATA] 명세서의 모든 Type(0~10) 수집
        data = []
        for i in range(recv_count):
            item = {
                'code': obj.GetDataValue(0, i),      # 0 - 종목코드
                'name': obj.GetDataValue(1, i),      # 1 - 종목명
                'price': obj.GetDataValue(2, i),     # 2 - 현재가
                'diff_flag': obj.GetDataValue(3, i), # 3 - 대비플래그
                'diff': obj.GetDataValue(4, i),      # 4 - 대비
                'diff_rate': round(obj.GetDataValue(5, i), 2), # 5 - 대비율
                'volume': obj.GetDataValue(6, i),    # 6 - 거래량
                'high_price': obj.GetDataValue(7, i),# 7 - 신고가(가격)
                'high_diff': obj.GetDataValue(8, i), # 8 - 신고가대비
                'high_rate': obj.GetDataValue(9, i), # 9 - 신고가대비율
                'amount': obj.GetDataValue(10, i)    # 10 - 거래대금
            }
            data.append(item)
        return data
    
    def _fetch_buy_dominance_ratio(self, 
                                market: str = '1', 
                                size: str = '4', 
                                amount: str = '1', 
                                criteria: str = '1', 
                                exchange: str = 'K'):
        """
        명세서 규격에 맞춰 단일 요청을 수행합니다.
        :param market: 0 - (char) 시장구분 ('1':거래소, '2':코스닥)
        :param size: 1 - (char) 세부종목분류 ('1':소형, '2':중형, '3':대형, '4':전체)
        :param amount: 2 - (char) 건별금액분류 ('1':1천만, '2':2천만, '3':3천만, '4':4천만, '5':1억 이상)
        :param criteria: 3 - (char) 조회기준 ('1':매수상위, '2':매도상위)
        :param exchange: 4 - (char) 거래소구분 ('A':전체, 'K':KRX, 'N':NXT)
        """
        obj = win32com.client.Dispatch("CpSysDib.CpSvr7034")

        # [INPUT] 명세서 규격에 따른 5가지 입력값 설정
        obj.SetInputValue(0, ord(market))    # 시장구분 ('1', '2')
        obj.SetInputValue(1, ord(size))      # 세부종목분류 ('1'~'4')
        obj.SetInputValue(2, ord(amount))    # 건별금액분류 ('1'~'5')
        obj.SetInputValue(3, ord(criteria))  # 조회기준 ('1', '2')
        obj.SetInputValue(4, ord(exchange))  # 거래소구분 ('A', 'K', 'N')
        
        self._wait_for_request_limit()
        obj.BlockRequest()
            # [HEADER] 조회 건수 확인
        recv_count = obj.GetHeaderValue(0)
        recv_mkt_flag = obj.GetHeaderValue(1)
        data_list = []

        for i in range(recv_count):
            # 명세서에 정의된 Type 0 ~ 8 데이터 추출
            item = {
                'code': obj.GetDataValue(0, i),         # (string) 종목코드
                'name': obj.GetDataValue(1, i),         # (string) 종목명
                'price': obj.GetDataValue(2, i),        # (ulong) 현재가
                'diff_flag': obj.GetDataValue(3, i),    # (char) 전일대비 Flag
                'diff_price': obj.GetDataValue(4, i),   # (long) 전일대비
                'volume': obj.GetDataValue(5, i),       # (ulong) 거래량
                'buy_count': obj.GetDataValue(6, i),    # (ulong) 매수체결건수
                'sell_count': obj.GetDataValue(7, i),   # (ulong) 매도체결건수
                'diff': obj.GetDataValue(8, i),         # (long) 대비
                'market_type': recv_mkt_flag            # (char) 시장구분
            }
            data_list.append(item)

        return data_list
            
    def _fetch_volume_rank(self, 
                      market: str = '1', 
                      selection: str = 'V', 
                      manage_category: str = 'Y', 
                      pref_stock: str = 'Y'):
        """
        [CpSysDib.CpSvr7049] 당일 거래량/거래대금 상위종목 데이터를 요청합니다.
        
        :param market: 0 - (string) 시장 구분 ("1":거래소, "2":코스닥, "4":전체)
        :param selection: 1 - (string) 선택 구분 ("V":거래량, "A":거래대금, "U":상승률, "D":하락률)
        :param manage_category: 2 - (string) 관리 구분("Y", "N")
        :param pref_stock: 3 - (string) 우선주 구분("Y", "N")
        """
        obj = win32com.client.Dispatch("CpSysDib.CpSvr7049")

        # [INPUT] 명세서 규격에 따른 4가지 입력값 설정
        obj.SetInputValue(0, market)           # 시장 구분
        obj.SetInputValue(1, selection)        # 선택 구분
        obj.SetInputValue(2, manage_category)  # 관리 구분
        obj.SetInputValue(3, pref_stock)       # 우선주 구분

        # [REQUEST] 데이터 요청 (Blocking Mode)
        obj.BlockRequest()

        # [OUTPUT] 헤더 데이터 (수신 개수) 확인
        recv_count = obj.GetHeaderValue(0)
        data_list = []

        for i in range(recv_count):
            # 명세서에 정의된 Type 0 ~ 7 데이터 추출
            item = {
                'rank': obj.GetDataValue(0, i),          # (short) 순위
                'code': obj.GetDataValue(1, i),          # (string) 종목코드
                'name': obj.GetDataValue(2, i),          # (string) 종목명
                'price': obj.GetDataValue(3, i),         # (long) 현재가
                'diff_price': obj.GetDataValue(4, i),    # (long) 전일대비
                'diff_rate': obj.GetDataValue(5, i),     # (float) 전일대비율
                'volume': obj.GetDataValue(6, i),        # (long) 거래량
                'amount': obj.GetDataValue(7, i)         # (long) 거래대금 (KOSPI:만원, KOSDAQ:천원)
            }
            data_list.append(item)

        return data_list    

if __name__ == "__main__":
    scanner = MarketScanner()
    
    # 데이터 수집 및 캐시 업데이트 실행
    cache = scanner.update_integrated_selection()

    print("\n" + "="*70)
    print(" [실시간 시장 통합 선별 리포트] ".center(60))
    print("="*70)

    # 1. TR 7043: 시장 등락 및 가격 돌파 섹션 (Breakout, Strength 등)
    print(f"\n📡 [TR 7043] 가격 모멘텀 선별 종목 ({len(cache['7043'])}건)")
    print("-" * 70)
    if not cache['7043']:
        print("  현재 선별된 가격 돌파 종목이 없습니다.")
    else:
        # 가독성을 위해 상위 15개만 출력하거나 전체 출력
        for stock in cache['7043']:
            tag = stock.get('strategy_tag', 'N/A')
            print(f"[{tag:18}] {stock['name']}({stock['code']})")
            # print(f"    └ 현재가: {stock['price']:,}원 | 대비율: {stock['diff_rate']}% | 거래대금: {stock['amount']:,}만")
            # print("-" * 50)

    print("\n" + "*"*70)

    # 2. TR 7034: 4,000만 원 이상 수급 우위 섹션 (Buy Dominance)
    print(f"\n💰 [TR 7034] 4,000만 원 이상 수급 집중 종목 ({len(cache['7034'])}건)")
    print("-" * 70)
    if not cache['7034']:
        print("  현재 수급 우위 종목이 포착되지 않았습니다.")
    else:
        # 매수 비중이 높은 순서대로 정렬해서 보여주면 더 좋습니다.
        sorted_7034 = sorted(cache['7034'], key=lambda x: x['buy_ratio'], reverse=True)
        
        for stock in sorted_7034:
            # 비중이 70% 이상인 경우 강조 표시
            highlight = "🔥" if stock['buy_ratio'] >= 70 else "  "
            print(f"{highlight} [비중 {stock['buy_ratio']}%] {stock['name']}({stock['code']})")
            # print(f"    └ 대량매수: {stock['buy_count']}건 | 대량매도: {stock['sell_count']}건 | 현재가: {stock['price']:,}원")
            # print("-" * 50)

    print("\n" + "="*70)
    print(" [스캔 종료] ".center(60))
    print("="*70)
        
            
    ####################################################################
    # # 1. 조회할 시장 설정 (ID, 시장명)
    # # '1': 거래소(KOSPI), '2': 코스닥(KOSDAQ)
    # target_markets = [('1', 'KOSPI'), ('2', 'KOSDAQ')]

    # for m_id, m_name in target_markets:
    #     print("="*80)
    #     print(f" [{m_name}] 4,000만 원 이상 대량 체결 매수 우위 종목 (Top 10)")
    #     print("="*80)
        
    #     # 파라미터: 시장(m_id), 금액='4'(4,000만 원 이상)
    #     # 메서드 내부에서 ord() 처리가 되어 있어야 에러가 나지 않습니다.
    #     dominance_results = scanner._fetch_buy_dominance_ratio(market=m_id, amount='4')

    #     # 데이터가 없는 경우 대비
    #     if not dominance_results:
    #         print(f" {m_name} 시장에서 조건에 맞는 데이터가 없습니다.")
    #         continue

    #     for stock in dominance_results[:10]:  # 시장별 상위 10개 출력
    #         # 매수/매도 합계 건수 계산
    #         total_cnt = stock['buy_count'] + stock['sell_count']
            
    #         # 매수 비중(Dominance Ratio) 계산
    #         buy_ratio = (stock['buy_count'] / total_cnt * 100) if total_cnt > 0 else 0
            
    #         print(f"[{stock['name']}({stock['code']})] 현재가: {stock['price']:,}원")
    #         print(f"   ▶ 대량매수: {stock['buy_count']}건 | 대량매도: {stock['sell_count']}건")
    #         print(f"   ▶ 큰손 매수 비중: {buy_ratio:.2f}%")
    #         print("-" * 40)
        
    #     print("\n" * 2) # 시장 간 간격 띄우기

    # 2. 코스닥(2) 시장 데이터도 확인하고 싶다면 아래 주석 해제 후 사용
    """
    print("\n" + "="*80)
    print(" [KOSDAQ] 1억 이상 대량 체결 매수 우위 종목 (Top 10)")
    print("="*80)
    kosdaq_dominance = scanner._fetch_buy_dominance_ratio(market='2', amount='5')
    for stock in kosdaq_dominance[:10]:
        total_cnt = stock['buy_count'] + stock['sell_count']
        buy_ratio = (stock['buy_count'] / total_cnt * 100) if total_cnt > 0 else 0
        print(f"[{stock['name']}] 비중: {buy_ratio:.2f}% | 매수건수: {stock['buy_count']}")
    """
    
    #######################################################################################
    # # -------------------------------------------------------------------------
    # # 전략 1: [신고가 주도주] 20일 신고가 돌파 + 거래대금 상위
    # # -------------------------------------------------------------------------
    # print("="*60)
    # print("전략 1: [신고가 주도주] 5일 신고가 + 거래대금 상위")
    # print("="*60)
    
    # # 결과를 바로 변수에 담습니다.
    # breakout_leaders = scanner._fetch_market_movement(
    #     market='0',      # 전체시장
    #     criteria='6',    # 신고가
    #     sort_by=61,      # 거래대금 상위순
    #     vol_filter='3',  # 10만주 이상
    #     period='1'       # 20일 신고가
    # )
    
    # for stock in breakout_leaders:
    #     print(f"[{stock['name']}] 현재가: {stock['price']} | 대비율: {stock['diff_rate']}% | 거래대금: {stock['amount']}만")


    # # -------------------------------------------------------------------------
    # # 전략 2: [당일 강세주] 시가 대비 5%~15% 상승 + 거래량 급증
    # # -------------------------------------------------------------------------
    # print("\n" + "="*60)
    # print("전략 2: [당일 강세주] 시가 대비 강세 (5%~15%)")
    # print("="*60)
    
    # intraday_strength = scanner._fetch_market_movement(
    #     market='0',      
    #     criteria='2',    # 상승 종목
    #     sort_by=21,      # 대비율 상위순
    #     vol_filter='4',  # 50만주 이상
    #     period='0',      # 시가 대비 상승
    #     rate_start=5,    
    #     rate_end=15      
    # )
    
    # for stock in intraday_strength:
    #     print(f"[{stock['name']}] 시가대비: {stock['diff_rate']}% | 현재가: {stock['price']} | 거래량: {stock['volume']}")


    # # -------------------------------------------------------------------------
    # # 전략 3: [낙폭 과대 반등] 저점 대비 반등 시도
    # # -------------------------------------------------------------------------
    # print("\n" + "="*60)
    # print("전략 3: [낙폭 과대] 당일 저점 대비 반등 시도")
    # print("="*60)
    
    # bottom_bounce = scanner._fetch_market_movement(
    #     market='0',      
    #     criteria='2',    # 상승(반등) 시도
    #     sort_by=21,      
    #     vol_filter='2',  # 5만주 이상
    #     period='2'       # 저가 대비 상승 (아래꼬리 확인)
    # )
    
    # for stock in bottom_bounce:
    #     print(f"[{stock['name']}] 저점대비반등: {stock['diff_rate']}% | 현재가: {stock['price']} | 거래량: {stock['volume']}")