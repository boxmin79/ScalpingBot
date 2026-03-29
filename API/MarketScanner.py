import win32com.client

class MarketScanner:
    def __init__(self):
        # 거래량/거래대금 상위종목 API 모듈 연결
        # 1. 거래대금 상위 (CpTopVolume)
        self.obj_top_vol = win32com.client.Dispatch("CpSysDib.CpSvr7049") 
        # 2. 신고가/등락 현황 (CpSvrNew7043)
        self.obj_breakout = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")
        # 3. 큰손 매수비중 (CpSvr7034)
        self.obj_whale = win32com.client.Dispatch("CpSysDib.CpSvr7034")

    def get_whale_ratio(self, market='1', size='4', amount='4', criteria='1'):
        """
        market: '1'거래소, '2'코스닥
        size: '1'소형, '2'중형, '3'대형, '4'전체
        amount: '1'1천만, '2'2천만, '3'3천만, '4'4천만, '5'1억 이상
        criteria: '1'매수비중 상위, '2'매도비중 상위
        """
        self.obj_whale.SetInputValue(0, market)    # 시장 구분
        self.obj_whale.SetInputValue(1, size)      # 세부종목 분류
        self.obj_whale.SetInputValue(2, amount)    # 건별 금액 분류 (큰손 기준)
        self.obj_whale.SetInputValue(3, criteria)  # 조회 기준 (매수 상위)
        
        self.obj_whale.BlockRequest()

        count = self.obj_whale.GetHeaderValue(0)
        results = []

        for i in range(count):
            item = {
                'code': self.obj_whale.GetDataValue(0, i),      # 종목코드
                'name': self.obj_whale.GetDataValue(1, i),      # 종목명
                'price': self.obj_whale.GetDataValue(2, i),     # 현재가
                'diff_flag': self.obj_whale.GetDataValue(3, i), # 대비 플래그
                'diff': self.obj_whale.GetDataValue(4, i),      # 전일대비
                'volume': self.obj_whale.GetDataValue(5, i),    # 거래량
                'buy_cnt': self.obj_whale.GetDataValue(6, i),   # 매수체결건수 (지정금액 이상)
                'sell_cnt': self.obj_whale.GetDataValue(7, i),  # 매도체결건수 (지정금액 이상)
            }
            # 매수/매도 건수 비율 계산 (수급 강도 분석용)
            total_cnt = item['buy_cnt'] + item['sell_cnt']
            item['buy_ratio'] = (item['buy_cnt'] / total_cnt * 100) if total_cnt > 0 else 0
            
            results.append(item)
            
        return results
    
    def get_breakout_list(self, market='0', criteria='6', sort_by=61, volume_filter='3', period='1'):
        """
        market: '0'전체, '1'거래소, '2'코스닥
        criteria: '1'상한, '2'상승, '3'보합, '4'하락, '5'하한, 6'신고가, '7'신저가
        sort_by: 21(대비율순), 51(거래량순), 61(거래대금순) -> 51, 61은 신고/신저가일 때만 가능
        volume_filter: '0'전체, '1'1만, '3'10만주 이상
        period: 신고가일 때 '1':5일, '2':20일, '5':52주
        """
        # 1. 입력값 설정
        self.obj_breakout.SetInputValue(0, market)        # 시장구분
        self.obj_breakout.SetInputValue(1, criteria)      # 선택기준 (6:신고가)
        if criteria in ['1', '2', '3', '4', '5']:
            self.obj_breakout.SetInputValue(2, '1')           # 당일 기준
        self.obj_breakout.SetInputValue(2, '0')           # 전일 기준
        self.obj_breakout.SetInputValue(3, sort_by)       # 순서 (61:거래대금순)
        self.obj_breakout.SetInputValue(4, '1')           # 관리종목 제외
        self.obj_breakout.SetInputValue(5, volume_filter) # 거래량 필터
        self.obj_breakout.SetInputValue(6, period)        # 기간 (2:20일)

        # 2. 데이터 요청
        self.obj_breakout.BlockRequest()

        # 3. 데이터 수신 및 파싱
        count = self.obj_breakout.GetHeaderValue(0)
        results = []

        for i in range(count):
            # 공통 데이터
            item = {
                'code': self.obj_breakout.GetDataValue(0, i),
                'name': self.obj_breakout.GetDataValue(1, i),
                'price': self.obj_breakout.GetDataValue(2, i),
                'diff_rate': self.obj_breakout.GetDataValue(5, i),
                'volume': self.obj_breakout.GetDataValue(6, i),
            }

            # 선택기준이 '신고가(6)' 또는 '신저가(7)'일 때만 거래대금(index 10) 추출 가능
            if criteria in ['6', '7']:
                item['amount'] = self.obj_breakout.GetDataValue(10, i)  # 거래대금
            
            results.append(item)

        return results
    
    def get_top_volume_list(self, 
                            market_type='4', 
                            rank_type='A',
                            admin_issue_yn='Y',
                            pref_stock_yn='Y'
                     ):
        """
        거래량/거래대금 상위 종목 데이터를 요청하고 수신합니다.
        market_type: '1'-거래소, '2'-코스닥, '4'-전체
        rank_type: 'V'-거래량, 'A'-거래대금, , "U"-상승률, "D"-하락률
        admin_issue_yn: 'Y'-관리종목 제외, 'N'-관리종목 포함
        pref_stock_yn: 'Y'-우선주 제외, 'N'-우선주 포함
        """
        self.obj_top_vol.SetInputValue(0, market_type) # 시장구분
        self.obj_top_vol.SetInputValue(1, rank_type)   # 상위순기준 (거래량/거래대금)
        self.obj_top_vol.SetInputValue(2, admin_issue_yn) # 관리종목 제외
        self.obj_top_vol.SetInputValue(3, pref_stock_yn)  # 우선주 제외
        
        self.obj_top_vol.BlockRequest()
        
        count = self.obj_top_vol.GetHeaderValue(0) # 수신 데이터 개수
        print(f"조회된 상위 종목 수: {count}")
        
        results = []
        for i in range(count):
            rank = self.obj_top_vol.GetDataValue(0, i) # 순위
            code = self.obj_top_vol.GetDataValue(1, i) # 종목코드
            name = self.obj_top_vol.GetDataValue(2, i) # 종목명
            price = self.obj_top_vol.GetDataValue(3, i) # 현재가
            diff = self.obj_top_vol.GetDataValue(4, i) # 전일대비
            diff_rate = self.obj_top_vol.GetDataValue(5, i) # 전일대비율
            volume = self.obj_top_vol.GetDataValue(6, i) # 거래량
            amount = self.obj_top_vol.GetDataValue(7, i) # 거래대금
            
            
            results.append({'rank': rank, 
                            'code': code, 
                            'name': name, 
                            'price': price, 
                            'diff': diff, 
                            'diff_rate': round(diff_rate, 2), 
                            'volume': volume, 
                            'amount': amount})
        return results
    
            
if __name__ == "__main__":
    obj = MarketScanner()
    