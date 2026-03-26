import win32com.client

class CpStockMst2:
    """
    Dscbo1.StockMst2 기능을 포함하는 클래스
    설명: 최대 110개 종목에 대해 상세 정보를 일괄 조회합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockMst2")
        
        # 상태구분 및 예상체결상태 매핑 ('char' 타입 대응)
        self.map_status = {
            '1': '상한', '2': '상승', '3': '보합', '4': '하한', '5': '하락',
            '6': '기세상한', '7': '기세상승', '8': '기세하한', '9': '기세하락'
        }
        
    def request_data(self, code_list, market_type='K'):
        """
        다수의 종목코드를 입력받아 데이터를 요청합니다.
        code_list: 종목코드 리스트 (예: ['A005930', 'A000660']) / 최대 110개
        market_type: 'A':전체, 'K':KRX, 'N':NXT (기본값 'K')
        """
        if not code_list:
            print("조회할 종목코드가 없습니다.")
            return False

        if len(code_list) > 110:
            print(f"최대 조회 가능 종목수는 110개입니다. (현재: {len(code_list)}개)")
            code_list = code_list[:110]

        # 1. 입력 데이터 설정 (구분자 ',' 사용)
        codes_str = ",".join(code_list)
        self.obj.SetInputValue(0, codes_str)
        self.obj.SetInputValue(1, ord(market_type)) # char 타입이므로 ord() 사용

        # 2. 데이터 요청
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return False
        return True

    def get_data_list(self):
        """수신된 데이터를 리스트로 반환합니다."""
        count = self.obj.GetHeaderValue(0) # 수신된 종목 수
        
        results = []
        for i in range(count):
            status_code = self.obj.GetDataValue(5, i)      # 상태구분
            market_flag = self.obj.GetDataValue(25, i)     # 동시호가구분
            exp_status_code = self.obj.GetDataValue(28, i) # 예상체결상태구분
            
            item = {
                'code': self.obj.GetDataValue(0, i),           # 종목코드
                'name': self.obj.GetDataValue(1, i),           # 종목명
                'time': self.obj.GetDataValue(2, i),           # 시간(HHMM)
                'current': self.obj.GetDataValue(3, i),        # 현재가
                'diff': self.obj.GetDataValue(4, i),           # 전일대비
                'status': self.map_status.get(status_code, status_code), # 상태
                'open': self.obj.GetDataValue(6, i),           # 시가
                'high': self.obj.GetDataValue(7, i),           # 고가
                'low': self.obj.GetDataValue(8, i),            # 저가
                'ask': self.obj.GetDataValue(9, i),            # 매도호가
                'bid': self.obj.GetDataValue(10, i),           # 매수호가
                'volume': self.obj.GetDataValue(11, i),        # 거래량(1주 단위)
                'amount': self.obj.GetDataValue(12, i),        # 거래대금(원 단위)
                'total_ask_rem': self.obj.GetDataValue(13, i), # 총매도잔량
                'total_bid_rem': self.obj.GetDataValue(14, i), # 총매수잔량
                'listed_stock': self.obj.GetDataValue(17, i),  # 상장주식수
                'strength': self.obj.GetDataValue(21, i),      # 체결강도
                'market_type': '동시호가' if market_flag == '1' else '장중',
                'exp_price': self.obj.GetDataValue(26, i),     # 예상체결가
                'exp_status': self.map_status.get(exp_status_code, exp_status_code) # 예상상태
            }
            results.append(item)
            
        return results

# --- 사용 예시 ---
if __name__ == "__main__":
    mst2 = CpStockMst2()
    
    # 관심 종목 리스트
    target_codes = ["A005930", "A000660", "A035420"]
    
    if mst2.request_data(target_codes):
        data = mst2.get_data_list()
        
        print(f"일괄 조회 결과 (총 {len(data)}종목):")
        print("-" * 60)
        for row in data:
            print(f"[{row['name']}] 현재가: {row['current']:,}원 | 대비: {row['diff']} ({row['status']})")
            print(f"   거래량: {row['volume']:,}주 | 거래대금: {row['amount']:,}원 | 체결강도: {row['strength']}%")
            print("-" * 60)