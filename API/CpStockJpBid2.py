import win32com.client

class CpStockJpBid2:
    """
    Dscbo1.StockJpBid2 기능을 포함하는 클래스
    설명: 주식 종목의 1차~10차 매도/매수 호가 및 잔량 데이터를 조회합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockJpBid2")

    def request_data(self, code, market_type='K'):
        """
        특정 종목의 호가 데이터를 요청합니다.
        code: 종목코드 (예: 'A005930')
        market_type: 'A':전체, 'K':KRX, 'N':NXT
        """
        self.obj.SetInputValue(0, code)
        self.obj.SetInputValue(1, ord(market_type)) # char 타입이므로 ord() 사용
        
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"조회 실패 (에러코드: {ret})")
            return False
        return True

    def get_header_data(self):
        """호가창의 헤더(요약) 정보를 추출하여 반환합니다."""
        return {
            'code': self.obj.GetHeaderValue(0),           # 종목코드
            'time': self.obj.GetHeaderValue(3),           # 시각
            'total_ask_rem': self.obj.GetHeaderValue(4),  # 총매도잔량
            'total_bid_rem': self.obj.GetHeaderValue(6),  # 총매수잔량
            'mid_price': self.obj.GetHeaderValue(13),     # 중간가격
            'market_kind': self.obj.GetHeaderValue(16),   # 거래소구분
        }

    def get_bid_ask_list(self):
        """1차부터 10차까지의 매도/매수 호가 및 잔량 리스트를 반환합니다."""
        count = self.obj.GetHeaderValue(1) # 고정값 10
        results = []
        
        for i in range(count):
            item = {
                'level': i + 1,
                'ask_price': self.obj.GetDataValue(0, i), # 매도호가
                'bid_price': self.obj.GetDataValue(1, i), # 매수호가
                'ask_rem': self.obj.GetDataValue(2, i),   # 매도잔량
                'bid_rem': self.obj.GetDataValue(3, i),   # 매수잔량
                'ask_diff': self.obj.GetDataValue(4, i),  # 매도잔량대비
                'bid_diff': self.obj.GetDataValue(5, i),  # 매수잔량대비
            }
            results.append(item)
        return results

# --- 사용 예시 ---
if __name__ == "__main__":
    jp_bid = CpStockJpBid2()
    
    # 삼성전자 호가 조회
    if jp_bid.request_data("A005930"):
        header = jp_bid.get_header_data()
        quotes = jp_bid.get_bid_ask_list()
        
        print(f"[{header['code']}] 조회 시각: {header['time']}")
        print(f"총매도잔량: {header['total_ask_rem']:,} | 총매수잔량: {header['total_bid_rem']:,}")
        print(f"중간가격: {header['mid_price']:,}")
        print("-" * 50)
        
        # 상위 5차 호가까지 출력
        print(f"{'차수':<4} | {'매도호가':<8} | {'매도잔량':<8} | {'매수호가':<8} | {'매수잔량':<8}")
        for q in quotes[:5]:
            print(f"{q['level']:<5} | {q['ask_price']:>8,} | {q['ask_rem']:>8,} | {q['bid_price']:>8,} | {q['bid_rem']:>8,}")