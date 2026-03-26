import win32com.client
import time

class CpStockChart:
    """
    CpSysDib.StockChart 기능을 포함하는 클래스
    설명: 주식, 업종, ELW의 일/주/월/분/틱 차트 데이터를 조회합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.StockChart")

    def get_chart_data(self, code, target_count, chart_type='D', cycle=1, req_type='2'):
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
        
        self.obj.SetInputValue(0, code)
        self.obj.SetInputValue(1, ord(req_type))    # '1':기간, '2':개수
        self.obj.SetInputValue(4, target_count)     # 요청 개수
        self.obj.SetInputValue(5, fields)           # 필드 배열
        self.obj.SetInputValue(6, ord(chart_type))  # 차트 구분
        self.obj.SetInputValue(7, cycle)            # 주기
        self.obj.SetInputValue(9, ord('1'))         # 수정 주가 반영

        results = []
        current_count = 0

        while current_count < target_count:
            # 2. 데이터 요청
            ret = self.obj.BlockRequest()
            if ret != 0:
                print(f"조회 실패 (에러코드: {ret})")
                break

            # 3. 헤더 정보 확인
            recv_count = self.obj.GetHeaderValue(3) # 실제 수신 개수
            field_cnt = self.obj.GetHeaderValue(1)  # 요청한 필드 개수
            
            # 4. 데이터 추출
            for i in range(recv_count):
                item = {
                    'date': self.obj.GetDataValue(0, i),   # 날짜(0)
                    'time': self.obj.GetDataValue(1, i),   # 시간(1)
                    'open': self.obj.GetDataValue(2, i),   # 시가(2)
                    'high': self.obj.GetDataValue(3, i),   # 고가(3)
                    'low': self.obj.GetDataValue(4, i),    # 저가(4)
                    'close': self.obj.GetDataValue(5, i),  # 종가(5)
                    'vol': self.obj.GetDataValue(6, i),    # 거래량(8) - 인덱스는 오름차순임
                }
                results.append(item)
            
            current_count += recv_count
            print(f"현재 {current_count}개 수신 완료...")

            # 5. 연속 데이터 유무 확인
            if not self.obj.Continue or current_count >= target_count:
                break
            
            # TR 과부하 방지 (연속 요청 시 짧은 휴식)
            time.sleep(0.1)

        return results

# --- 사용 예시 ---
if __name__ == "__main__":
    chart_mgr = CpStockChart()
    
    # 삼성전자(A005930) 최근 100개의 '일봉' 데이터 요청
    # 'D': 일봉, 1: 1일 주기
    data = chart_mgr.get_chart_data("A005930", 100, 'D', 1)

    print(f"\n총 {len(data)}개의 데이터를 가져왔습니다.")
    print("-" * 60)
    # 최근 5일치만 출력
    for row in data[:5]:
        print(f"날짜: {row['date']} | 시가: {row['open']:,} | 고가: {row['high']:,} | "
              f"저가: {row['low']:,} | 종가: {row['close']:,} | 거래량: {row['vol']:,}")