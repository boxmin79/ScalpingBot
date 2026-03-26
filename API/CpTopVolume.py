import win32com.client

class CpTopVolume:
    def __init__(self):
        # 거래량/거래대금 상위종목 API 모듈 연결
        self.obj = win32com.client.Dispatch("CpSysDib.CpSvr7049")

    def get_top_list(self, 
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
        self.obj.SetInputValue(0, market_type) # 시장구분
        self.obj.SetInputValue(1, rank_type)   # 상위순기준 (거래량/거래대금)
        self.obj.SetInputValue(2, admin_issue_yn) # 관리종목 제외
        self.obj.SetInputValue(3, pref_stock_yn)  # 우선주 제외
        
        self.obj.BlockRequest()
        
        count = self.obj.GetHeaderValue(0) # 수신 데이터 개수
        print(f"조회된 상위 종목 수: {count}")
        
        results = []
        for i in range(count):
            rank = self.obj.GetDataValue(0, i) # 순위
            code = self.obj.GetDataValue(1, i) # 종목코드
            name = self.obj.GetDataValue(2, i) # 종목명
            price = self.obj.GetDataValue(3, i) # 현재가
            diff = self.obj.GetDataValue(4, i) # 전일대비
            diff_rate = self.obj.GetDataValue(5, i) # 전일대비율
            volume = self.obj.GetDataValue(6, i) # 거래량
            amount = self.obj.GetDataValue(7, i) # 거래대금
            
            
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
    obj = CpTopVolume()
    data = obj.get_top_list()
    for item in data:
        print(item)