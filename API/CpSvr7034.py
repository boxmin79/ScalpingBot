import win32com.client

class CpSvr7034:
    """
    [CpSysDib.CpSvr7034] 매수체결비중 상위종목 조회
    특정 금액 이상의 '큰손' 매수세가 집중되는 종목을 포착합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.CpSvr7034")

    def get_data(self, market='1', size='4', amount='4', criteria='1'):
        """
        market: '1'거래소, '2'코스닥
        size: '1'소형, '2'중형, '3'대형, '4'전체
        amount: '1'1천만, '2'2천만, '3'3천만, '4'4천만, '5'1억 이상
        criteria: '1'매수비중 상위, '2'매도비중 상위
        """
        self.obj.SetInputValue(0, market)    # 시장 구분
        self.obj.SetInputValue(1, size)      # 세부종목 분류
        self.obj.SetInputValue(2, amount)    # 건별 금액 분류 (큰손 기준)
        self.obj.SetInputValue(3, criteria)  # 조회 기준 (매수 상위)
        
        self.obj.BlockRequest()

        count = self.obj.GetHeaderValue(0)
        results = []

        for i in range(count):
            item = {
                'code': self.obj.GetDataValue(0, i),      # 종목코드
                'name': self.obj.GetDataValue(1, i),      # 종목명
                'price': self.obj.GetDataValue(2, i),     # 현재가
                'diff_flag': self.obj.GetDataValue(3, i), # 대비 플래그
                'diff': self.obj.GetDataValue(4, i),      # 전일대비
                'volume': self.obj.GetDataValue(5, i),    # 거래량
                'buy_cnt': self.obj.GetDataValue(6, i),   # 매수체결건수 (지정금액 이상)
                'sell_cnt': self.obj.GetDataValue(7, i),  # 매도체결건수 (지정금액 이상)
            }
            # 매수/매도 건수 비율 계산 (수급 강도 분석용)
            total_cnt = item['buy_cnt'] + item['sell_cnt']
            item['buy_ratio'] = (item['buy_cnt'] / total_cnt * 100) if total_cnt > 0 else 0
            
            results.append(item)
            
        return results
    
if __name__ == "__main__":
    obj = CpSvr7034()
    data = obj.get_data()
    print(data)
    # print(f"조회된 종목 수: {len(data)}")
    # print("-" * 105)  # 간격에 맞춰 구분선 길이를 조절했습니다.

    # # 헤더 출력 (왼쪽 정렬 : < , 오른쪽 정렬 : >)
    # header = (f"{'종목코드':<10}{'종목명':<16}{'현재가':>10}{'대비':>8}"
    #         f"{'거래량':>12}{'매수건수':>8}{'매도건수':>8}{'매수비율':>10}")
    # print(header)
    # print("-" * 105)

    # for item in data:
    #     # f-string 포맷팅 설명:
    #     # :<10 -> 10칸 차지하고 왼쪽 정렬
    #     # :>10, -> 10칸 차지하고 오른쪽 정렬 + 천 단위 콤마
    #     # :>9.2f -> 9칸 차지하고 소수점 2자리까지 표시
        
    #     line = (f"{item['code']:<10}  "
    #             f"{item['name']:<16}  "   # 한글 명칭은 터미널 환경에 따라 정렬이 어긋날 수 있음
    #             f"{item['price']:>10,}  "
    #             f"{item['diff']:>8,}  "
    #             f"{item['volume']:>12,}  "
    #             f"{item['buy_cnt']:>8,}  "
    #             f"{item['sell_cnt']:>8,}  "
    #             f"{item['buy_ratio']:>9.2f}%")
    #     print(line)

    # print("-" * 105)