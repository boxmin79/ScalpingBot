import win32com.client

class CpStockBid:
    """
    Dscbo1.StockBid 기능을 포함하는 클래스
    설명: 주식 종목의 시간대별 체결 데이터를 요청하고 수신합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("Dscbo1.StockBid")

    def get_tick_data(self, code, count=80, mode='C', search_time=""):
        """
        시간대별 체결 데이터를 조회하여 리스트로 반환합니다.
        code: 종목코드
        count: 요청 개수 (최대 80개)
        mode: 'C' (체결가 비교 방식 - Default), 'H' (호가 비교 방식)
        search_time: 검색 시작 시간 (예: "0910")
        """
        # 1. 입력 데이터 설정
        self.obj.SetInputValue(0, code)
        self.obj.SetInputValue(2, count)        # 요청개수
        self.obj.SetInputValue(3, ord(mode))    # 체결비교방식 ('C' or 'H')
        if search_time:
            self.obj.SetInputValue(4, search_time) # 시간검색
        self.obj.SetInputValue(5, ord('K'))     # 거래소구분: 'K' (KRX)

        results = []
        
        while True:
            # 2. 데이터 요청
            ret = self.obj.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = self.obj.GetHeaderValue(2) # 실제 수신 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    'time': self.obj.GetDataValue(0, i),         # 시각 (HHMM)
                    'time_s': self.obj.GetDataValue(9, i),       # 시각 (초)
                    'current': self.obj.GetDataValue(4, i),      # 현재가
                    'diff': self.obj.GetDataValue(1, i),         # 전일대비
                    'volume': self.obj.GetDataValue(5, i),       # 누적거래량
                    'tick_vol': self.obj.GetDataValue(6, i),     # 순간체결량
                    'side': '매수' if self.obj.GetDataValue(7, i) == '1' else '매도', # 체결상태
                    'strength': round(self.obj.GetDataValue(8, i), 2), # 체결강도
                    'market_flag': '장중' if self.obj.GetDataValue(10, i) == '2' else '예상',
                }
                results.append(item)

            # 연속 데이터 유무 확인 (Paging)
            # 수신 개수가 요청한 count보다 작거나 Continue가 False면 종료
            if not self.obj.Continue or len(results) >= count:
                break
                
            # 추가 조회가 필요한 경우 (필요 시 로직 확장 가능)
            break 
            
        return results

# --- 사용 예시 ---
if __name__ == "__main__":
    bid_mgr = CpStockBid()
    
    # 삼성전자 최근 20개 체결 내역 조회 (호가비교 방식)
    ticks = bid_mgr.get_tick_data("A005930", count=20, mode='H')

    print(f"\n최근 체결 내역 (조회수: {len(ticks)})")
    print(f"{'시간':<10} | {'현재가':<8} | {'체결량':<6} | {'구분':<4} | {'체결강도':<6}")
    print("-" * 50)
    for t in ticks:
        print(f"{t['time']:04d}{t['time_s']:02d} | {t['current']:>8,} | {t['tick_vol']:>6,} | {t['side']:<4} | {t['strength']:>6}%")