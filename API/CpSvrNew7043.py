import win32com.client

class CpSvr7043:
    """
    [CpSysDib.CpSvrNew7043] 거래소/코스닥 등락 및 신고/신저가 종목 조회
    시장의 주도주(상승률 상위, 신고가 돌파 등)를 빠르게 필터링합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")

    def get_data(self, market='0', criteria='6', sort_by=61, volume_filter='3', period='1'):
        """
        market: '0'전체, '1'거래소, '2'코스닥
        criteria: '1'상한, '2'상승, '3'보합, '4'하락, '5'하한, 6'신고가, '7'신저가
        sort_by: 21(대비율순), 51(거래량순), 61(거래대금순) -> 51, 61은 신고/신저가일 때만 가능
        volume_filter: '0'전체, '1'1만, '3'10만주 이상
        period: 신고가일 때 '1':5일, '2':20일, '5':52주
        """
        # 1. 입력값 설정
        self.obj.SetInputValue(0, market)        # 시장구분
        self.obj.SetInputValue(1, criteria)      # 선택기준 (6:신고가)
        if criteria in ['1', '2', '3', '4', '5']:
            self.obj.SetInputValue(2, '1')           # 당일 기준
        self.obj.SetInputValue(2, '0')           # 전일 기준
        self.obj.SetInputValue(3, sort_by)       # 순서 (61:거래대금순)
        self.obj.SetInputValue(4, '1')           # 관리종목 제외
        self.obj.SetInputValue(5, volume_filter) # 거래량 필터
        self.obj.SetInputValue(6, period)        # 기간 (2:20일)

        # 2. 데이터 요청
        self.obj.BlockRequest()

        # 3. 데이터 수신 및 파싱
        count = self.obj.GetHeaderValue(0)
        results = []

        for i in range(count):
            # 공통 데이터
            item = {
                'code': self.obj.GetDataValue(0, i),
                'name': self.obj.GetDataValue(1, i),
                'price': self.obj.GetDataValue(2, i),
                'diff_rate': self.obj.GetDataValue(5, i),
                'volume': self.obj.GetDataValue(6, i),
            }

            # 선택기준이 '신고가(6)' 또는 '신저가(7)'일 때만 거래대금(index 10) 추출 가능
            if criteria in ['6', '7']:
                item['amount'] = self.obj.GetDataValue(10, i)  # 거래대금
            
            results.append(item)

        return results

# 테스트 실행 코드
if __name__ == "__main__":
    helper = CpSvr7043()
    # 코스닥('2'), 신고가('6'), 거래대금순(61), 20일 신고가('2') 조회
    data = helper.get_data()
    print(data)
    
    # print(f"조회된 종목 수: {len(data)}")
    # for stock in data:  # 상위 10개만 출력
    #     print(f"[{stock['code']}] {stock['name']} | 등락: {stock['diff_rate']:.2f}% | 대금: {stock.get('amount', 0):,}백만원")