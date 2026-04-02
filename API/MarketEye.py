import win32com.client
import time

class MarketEye:
    """
    [CpSysDib.MarketEye] 여러 종목의 다양한 데이터를 한 번에 조회하는 클래스
   
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.MarketEye")

    def get_market_data(self, codes:list=[], field_ids:list=None):
        """
        주식,지수,선물/옵션등의여러종목의필요항목들을한번에수신합니다.
        
        Args:
            codes (list): 종목코드 리스트 (최대 200개)
            field_ids (list): 조회할 필드 ID 리스트 (최대 64개)
        """
        # 20:총상장주식수(ulonglong)
        # 67 PER(float)
        # 70 EPS(float)
        # 71 자본금(ulonglong)-(백만)
        # 72 액면가(ushort)
        # 73 배당률(float)
        # 74 배당수익률(float)
        # 75 부채비율(float)
        # 76 유보율(float)
        # 77 ROE(float) - 자기자본손이익률
        # 78 매출액증가율(float)
        # 79 경상이익증가율(float)
        # 80 순이익증가율(float)
        # 86 매출액(ulonglong) - 단위:백만
        # 87 경상이익(ulonglong) - 단위:원
        # 88 당기순이익(ulonglong) - 단위:원
        # 89 BPS(ulong) - 주당순자산
        # 90 영업이익증가율(float)
        # 91 영업이익(ulonglong) - 단위:원
        # 92 매출액영업이익률(float)
        # 93 매출액경상이익률(float)
        # 94 이자보상비율(float)
        # 95 결산년월(ulong) - yyyymm
        # 96 분기BPS(ulong) - 분기주당순자산
        # 97 분기매출액증가율(float)
        # 98 분기영업이액증가율(float)
        # 99 분기경상이익증가율(float)
        # 100 분기순이익증가율(float)
        # 101 분기매출액(ulonglong) - 단위:백만
        # 102 분기영업이익(ulonglong) - 단위:원
        # 103 분기경상이익(ulonglong) - 단위:원
        # 104 분기당기순이익(ulonglong) - 단위:원
        # 105 분개매출액영업이익률(float)
        # 106 분기매출액경상이익률(float)
        # 107 분기ROE(float) - 자기자본순이익률
        # 108 분기이자보상비율(float)
        # 109 분기유보율(float)
        # 110 분기부채비율(float)
        # 111 최근분기년월(ulong) - yyyymm
        # 123 SPS(ulong) 주당 매출액
        # 124 CFPS(ulong) 주당 현금흐름
        # 125 EBITDA(ulong)
        
        if not field_ids:
            field_ids = [0, 20, 67, 70, 71, 72, 73, 74, 76, 77, 78, 79, 80, 86, 87,
                         88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102,
                         103, 104, 105, 106, 107, 108, 109, 110, 111, 123, 124, 125]
            
        # 1. 입력값 설정
        self.obj.SetInputValue(0, field_ids) # 필드 배열
        self.obj.SetInputValue(1, codes)     # 종목코드 배열
        self.obj.SetInputValue(2, '2') # 체결 방식 '1' 체결가비교방식, 2'호가비교방식(default)
        self.obj.SetInputValue(3, ord('K')) # 거래소구분 'A' 전체, 'K' KRX, 'N' NXT
        
        # 2. 데이터 요청 (Request/Reply)
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"❌ MarketEye 요청 실패 (에러코드: {ret})")
            return []

        # 3. 헤더 정보 확인
        field_count = self.obj.GetHeaderValue(0) # 요청한 필드 개수
        field_array = self.obj.GetHeaderValue(1) # 요청한 필드 배열
        stock_count = self.obj.GetHeaderValue(2) # 수신된 종목 개수
        market_flag = self.obj.GetHeaderValue(3) # 거래소 구분
        
        # ⚠️ 중요: GetDataValue의 type(필드 인덱스)은 요청한 field_ids의 
        # 값이 아니라, field_ids를 오름차순으로 정렬했을 때의 순서입니다.
        sorted_fields = sorted(field_ids)
        results = []
        # for fd_idx, col_idx in zip(sorted_fields, field_array):
            # print(fd_idx, col_idx)
            
        # 4. 데이터 파싱
        for i in range(stock_count):
            stock_data = {}
            for f_idx, f_id in enumerate(sorted_fields):
                # GetDataValue(필드순서, 종목순서)
                stock_data[f_id] = self.obj.GetDataValue(f_idx, i)
            results.append(stock_data)

        return results, field_array

# 사용 예시
if __name__ == "__main__":
    manager = MarketEye()
    # 종목: SK하이닉스, 삼성전자 / 필드: 0(코드), 17(명칭), 4(현재가), 24(체결강도)
    target_codes = ["A000660", "A005930"]
    # target_fields = [0, 17, 4, 24, 11] # 11: 거래대금
    
    data = manager.get_market_data(target_codes)
    for item in data:
        print(f"종목코드: {item[0]}")

    
    ### 가져올 수 있는 데이터 ###
    
    