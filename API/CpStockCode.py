import win32com.client

class CpStockCode:
    """주식 코드 조회 및 호가 단위 계산을 담당하는 클래스"""
    
    def __init__(self):
        # CpUtil.CpStockCode API 모듈 연결
        self.obj = win32com.client.Dispatch("CpUtil.CpStockCode")

    def get_name_by_code(self, code):
        """종목코드(A005930)를 입력받아 종목명을 반환합니다."""
        return self.obj.CodeToName(code)

    def get_code_by_name(self, name):
        """
        종목명을 입력받아 종목코드를 반환합니다.
        [주의] 중복된 종목명이 있을 경우 정확하지 않을 수 있습니다.
        """
        return self.obj.NameToCode(name)

    def get_full_code_by_code(self, code):
        """종목코드를 입력받아 FullCode(표준코드)를 반환합니다."""
        return self.obj.CodeToFullCode(code)

    def get_name_by_full_code(self, full_code):
        """FullCode를 입력받아 종목명을 반환합니다."""
        return self.obj.FullCodeToName(full_code)

    def get_code_by_full_code(self, full_code):
        """FullCode를 입력받아 종목코드(A로 시작하는 Short Code)를 반환합니다."""
        return self.obj.FullCodeToCode(full_code)

    def get_index_by_code(self, code):
        """종목코드를 입력받아 내부 Index를 반환합니다."""
        return self.obj.CodeToIndex(code)

    def get_count(self):
        """전체 종목 코드의 개수를 반환합니다."""
        return self.obj.GetCount()

    def get_data(self, data_type, index):
        """
        해당 인덱스의 종목 데이터를 구합니다.
        data_type: 0-종목코드, 1-종목명, 2-FullCode
        """
        return self.obj.GetData(data_type, index)

    def get_price_unit(self, code, base_price, direction_up=True):
        """
        주식/ETF/ELW의 호가 단위를 계산하여 반환합니다. (매우 유용)
        code: 종목코드
        base_price: 기준 가격 (현재가 혹은 주문 예정가)
        direction_up: True(상승 호가 단위), False(하락 호가 단위)
        """
        return self.obj.GetPriceUnit(code, base_price, direction_up)

    # --- 추가 편의 기능: 전체 종목 리스트 가져오기 ---
    def get_all_stock_list(self):
        """현재 CYBOS에 등록된 모든 종목(코드, 이름)을 딕셔너리 리스트로 반환합니다."""
        count = self.get_count()
        stock_list = []
        for i in range(count):
            code = self.get_data(0, i)
            name = self.get_data(1, i)
            stock_list.append({'code': code, 'name': name})
        return stock_list

# --- 테스트 및 사용 예시 ---
if __name__ == "__main__":
    code_mgr = CpStockCode()
    
    # 1. 삼성전자 코드 찾기
    samsung_code = code_mgr.get_code_by_name("삼성전자")
    print(f"삼성전자 코드: {samsung_code}")
    
    # 2. 코드 정보를 통해 이름 찾기
    name = code_mgr.get_name_by_code("A000660")
    print(f"A000660 종목명: {name}")
    
    # 3. 호가 단위 계산 (중요: 주문 시 가격 설정에 활용)
    # 삼성전자가 70,000원일 때, 한 호가 위 가격은?
    unit = code_mgr.get_price_unit(samsung_code, 70000, True)
    print(f"삼성전자 70,000원 기준 호가 단위: {unit}원")
    print(f"다음 매수 호가: {70000 + unit}원")