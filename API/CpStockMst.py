import win32com.client

class CpStockMst:
    """
    Dscbo1.StockMst 기능을 포함하는 클래스
    설명: 주식 종목의 현재가 정보와 10차 호가 데이터를 조회합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockMst")

    def request_data(self, code):
        """
        특정 종목의 현재가 데이터를 요청합니다.
        code: 종목코드 (예: 'A005930')
        """
        self.obj.SetInputValue(0, code)
        self.obj.SetInputValue(1, ord('K')) # 거래소 구분: 'K' (KRX)
        
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return False
        return True

    def get_header_data(self):
        """헤더에서 주요 현재가 정보를 추출하여 딕셔너리로 반환합니다."""
        # 누적거래대금 단위 처리 (코스피: 만원, 코스닥: 천원)
        # 소속구분(45번)을 확인하여 단위를 조정하는 로직을 추가할 수 있습니다.
        
        return {
            'code': self.obj.GetHeaderValue(0),        # 종목코드
            'name': self.obj.GetHeaderValue(1),        # 종목명
            'time': self.obj.GetHeaderValue(4),        # 시간
            'high_limit': self.obj.GetHeaderValue(8),  # 상한가
            'low_limit': self.obj.GetHeaderValue(9),   # 하한가
            'prev_close': self.obj.GetHeaderValue(10), # 전일종가
            'current': self.obj.GetHeaderValue(11),    # 현재가
            'diff': self.obj.GetHeaderValue(12),       # 전일대비
            'open': self.obj.GetHeaderValue(13),       # 시가
            'high': self.obj.GetHeaderValue(14),       # 고가
            'low': self.obj.GetHeaderValue(15),        # 저가
            'ask': self.obj.GetHeaderValue(16),        # 매도호가
            'bid': self.obj.GetHeaderValue(17),        # 매수호가
            'volume': self.obj.GetHeaderValue(18),     # 누적거래량
            'amount': self.obj.GetHeaderValue(19),     # 누적거래대금
            'per': self.obj.GetHeaderValue(28),        # PER
            'eps': self.obj.GetHeaderValue(20),        # EPS
            'status': self.obj.GetHeaderValue(68),     # 거래정지여부 ('Y'/'N')
            'vi_base': self.obj.GetHeaderValue(80),    # 정적VI 발동 예상기준가
        }

    def get_quote_data(self):
        """10차 호가(매도/매수 잔량) 데이터를 리스트로 반환합니다."""
        quotes = []
        # 10차 호가이므로 0~9 인덱스 사용
        for i in range(10):
            item = {
                'index': i + 1,
                'ask_price': self.obj.GetDataValue(0, i),  # 매도호가
                'bid_price': self.obj.GetDataValue(1, i),  # 매수호가
                'ask_rem': self.obj.GetDataValue(2, i),    # 매도잔량
                'bid_rem': self.obj.GetDataValue(3, i),    # 매수잔량
                'ask_diff': self.obj.GetDataValue(4, i),   # 매도잔량대비
                'bid_diff': self.obj.GetDataValue(5, i),   # 매수잔량대비
            }
            quotes.append(item)
        return quotes

# --- 사용 예시 ---
if __name__ == "__main__":
    mst = CpStockMst()
    
    if mst.request_data("A005930"): # 삼성전자
        header = mst.get_header_data()
        quotes = mst.get_quote_data()
        
        print(f"[{header['name']}] 현재가: {header['current']:,}원")
        print(f"전일대비: {header['diff']} | 거래량: {header['volume']:,}")
        print("-" * 30)
        
        # 10차 호가 중 상위 3개만 출력
        print("최우선 호가 잔량:")
        for q in quotes[:3]:
            print(f"  {q['index']}차: 매도 {q['ask_price']} ({q['ask_rem']}) | 매수 {q['bid_price']} ({q['bid_rem']})")