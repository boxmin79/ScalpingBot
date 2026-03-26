import win32com.client

class CpStockMstM:
    """
    Dscbo1.StockMstM 기능을 포함하는 클래스
    설명: 최대 110개의 주식 종목 정보를 일괄 조회합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockMstM")
        
        # 대비구분코드 매핑
        self.map_diff_status = {
            1: '상한', 2: '상승', 3: '보합', 4: '하한', 5: '하락',
            6: '기세상한', 7: '기세상승', 8: '기세하한', 9: '기세하락'
        }
        # 장구분플래그 매핑
        self.map_market_status = {
            '0': '장외', '1': '동시호가', '2': '장중'
        }

    def request_data(self, code_list):
        """
        다수의 종목코드를 입력받아 데이터를 요청합니다.
        code_list: 종목코드 리스트 (예: ['A005930', 'A000660', ...]) / 최대 110개
        """
        if not code_list:
            print("조회할 종목코드가 없습니다.")
            return False

        if len(code_list) > 110:
            print("최대 조회 가능 종목수는 110개입니다. (현재: {}개)".format(len(code_list)))
            code_list = code_list[:110]

        # 1. 입력 데이터 설정 (종목코드들을 하나의 문자열로 결합)
        # 예: "A005930A000660"
        codes_str = "".join(code_list)
        self.obj.SetInputValue(0, codes_str)

        # 2. 데이터 요청
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return False
        return True

    def get_data_list(self):
        """수신된 복수 종목 데이터를 리스트로 반환합니다."""
        count = self.obj.GetHeaderValue(0) # 수신된 종목 수
        
        results = []
        for i in range(count):
            diff_code = self.obj.GetDataValue(3, i)
            market_flag = self.obj.GetDataValue(8, i)
            
            item = {
                'code': self.obj.GetDataValue(0, i),          # 종목코드
                'name': self.obj.GetDataValue(1, i),          # 종목명
                'diff': self.obj.GetDataValue(2, i),          # 대비
                'diff_status': self.map_diff_status.get(diff_code, diff_code), # 대비구분
                'current': self.obj.GetDataValue(4, i),       # 현재가
                'ask': self.obj.GetDataValue(5, i),           # 매도호가
                'bid': self.obj.GetDataValue(6, i),           # 매수호가
                'volume': self.obj.GetDataValue(7, i),        # 거래량
                'market_status': self.map_market_status.get(market_flag, market_flag), # 장구분
                'exp_price': self.obj.GetDataValue(9, i),     # 예상체결가
                'exp_diff': self.obj.GetDataValue(10, i),     # 예상체결가 전일대비
                'exp_vol': self.obj.GetDataValue(11, i),      # 예상체결수량
            }
            results.append(item)
            
        return results

# --- 사용 예시 ---
if __name__ == "__main__":
    mst_m = CpStockMstM()
    
    # 조회할 종목 리스트 (최대 110개)
    codes = ["A005930", "A000660", "A035420", "A035720"] 
    
    if mst_m.request_data(codes):
        data = mst_m.get_data_list()
        
        print(f"일괄 조회 결과 (총 {len(data)}종목):")
        print("-" * 50)
        for row in data:
            print(f"[{row['name']}] 현재가: {row['current']:,} | 대비: {row['diff']} ({row['diff_status']}) | 거래량: {row['volume']:,}")