import win32com.client

class CpCodeManager:
    """
    CpUtil.CpCodeMgr의 모든 기능을 포함하는 통합 클래스
    설명: 각종 코드 정보 및 종목 리스트를 조회합니다.
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpCodeMgr")

    # --- [1] 기본 명칭 및 검색 기능 ---
    def code_to_name(self, code):
        """종목코드에 해당하는 주식/선물/옵션 종목명을 반환"""
        return self.obj.CodeToName(code)

    def get_group_name(self, code):
        """관심종목(700~799) 및 업종코드에 해당하는 명칭 반환"""
        return self.obj.GetGroupName(code)

    def get_industry_name(self, code):
        """증권전산 업종코드에 해당하는 명칭 반환"""
        return self.obj.GetIndustryName(code)

    def get_member_name(self, code):
        """거래원코드(회원사)에 해당하는 명칭 반환"""
        return self.obj.GetMemberName(code)

    # --- [2] 종목 상세 정보 (주식) ---
    def get_stock_margin_rate(self, code):
        """주식 매수 증거금률 반환"""
        return self.obj.GetStockMarginRate(code)

    def get_stock_meme_min(self, code):
        """주식 매매 거래 단위 주식수 반환"""
        return self.obj.GetStockMemeMin(code)

    def get_stock_industry_code(self, code):
        """증권전산 업종코드 반환"""
        return self.obj.GetStockIndustryCode(code)

    def get_stock_market_kind(self, code):
        """소속부 반환 (1:거래소, 2:코스닥, 3:K-OTC, 4:KRX, 5:KONEX)"""
        return self.obj.GetStockMarketKind(code)

    def get_stock_control_kind(self, code):
        """감리구분 반환 (0:정상, 1:주의, 2:경고, 3:위험예고, 4:위험, 5:경고예고)"""
        return self.obj.GetStockControlKind(code)

    def get_overheating(self, code):
        """단기과열 구분 반환 (0:해당없음, 1:지정예고, 2:지정, 3:해제연기)"""
        return self.obj.GetOverHeating(code)

    def get_stock_supervision_kind(self, code):
        """관리구분 반환 (0:일반, 1:관리)"""
        return self.obj.GetStockSupervisionKind(code)

    def get_stock_status_kind(self, code):
        """주식상태 반환 (0:정상, 1:거래정지, 2:거래중단)"""
        return self.obj.GetStockStatusKind(code)

    def get_stock_capital(self, code):
        """자본금 규모 반환 (0:제외, 1:대, 2:중, 3:소)"""
        return self.obj.GetStockCapital(code)

    def get_stock_section_kind(self, code):
        """부구분코드 반환 (1:주권, 10:ETF, 17:ETN 등)"""
        return self.obj.GetStockSectionKind(code)

    def get_stock_lac_kind(self, code):
        """락구분코드 반환 (0:정상, 1:권리락, 2:배당락 등)"""
        return self.obj.GetStockLacKind(code)

    def get_stock_listed_date(self, code):
        """상장일 반환 (YYYYMMDD)"""
        return self.obj.GetStockListedDate(code)

    def get_listing_stock_count(self, code):
        """상장주식수 반환 (단위: 천주)"""
        return self.obj.GetListingStock(code)

    # --- [3] 가격 관련 정보 ---
    def get_stock_prices(self, code):
        """종목의 주요 가격 정보들 반환 (상한/하한/액면/기준가/전일종가 등)"""
        return {
            'max': self.obj.GetStockMaxPrice(code),
            'min': self.obj.GetStockMinPrice(code),
            'par': self.obj.GetStockParPrice(code),
            'std': self.obj.GetStockStdPrice(code),
            'open': self.obj.GetStockYdOpenPrice(code),
            'high': self.obj.GetStockYdHighPrice(code),
            'low': self.obj.GetStockYdLowPrice(code),
            'close': self.obj.GetStockYdClosePrice(code)
        }

    # --- [4] 각종 종목 리스트 (배열 반환) ---
    def get_stock_list_by_market(self, market_kind):
        """시장별 종목 리스트 (1:KOSPI, 2:KOSDAQ 등)"""
        return self.obj.GetStockListByMarket(market_kind)

    def get_industry_list(self):
        """증권전산 업종 코드 리스트 반환"""
        return self.obj.GetIndustryList()

    def get_group_code_list(self, group_code):
        """관심종목/업종코드에 해당하는 종목 리스트 반환"""
        return self.obj.GetGroupCodeList(group_code)

    def get_index_code_list(self, big, mid, small):
        """지수 코드 리스트 반환 (대/중/소분류)"""
        return self.obj.GetIndexCodeList(big, mid, small)

    def get_mini_future_list(self): return self.obj.GetMiniFutureList()
    def get_mini_option_list(self): return self.obj.GetMiniOptionList()
    def get_member_list(self): return self.obj.GetMemberList()
    def get_nxt_stock_all_list(self): return self.obj.GetNxtStockAllList()

    # --- [5] 파생상품 및 해외선물 ---
    def get_fo_trade_unit(self, code):
        """파생상품 거래단위 반환"""
        return self.obj.GetFOTradeUnit(code)

    def get_ov_fut_info(self, code):
        """해외선물 상세 정보 반환"""
        return {
            'name': self.obj.OvFutCodeToName(code),
            'exch': self.obj.OvFutGetExchCode(code),
            'last_date': self.obj.OvFutGetLastTradeDate(code),
            'tick_unit': self.obj.GetTickUnit(code),
            'tick_value': self.obj.GetTickValue(code)
        }

    def get_ov_fut_all_codes(self):
        """해외선물 전체 코드 리스트 반환"""
        return self.obj.OvFutGetAllCodeList()

    def get_stock_future_list(self):
        """주식선물 전체 코드 리스트 반환"""
        return self.obj.GetStockFutureList()

    # --- [6] 유무 여부 확인 (Boolean) ---
    def is_spac(self, code): return self.obj.IsSPAC(code)
    def is_credit_enable(self, code): return self.obj.IsStockCreditEnable(code)
    def is_big_listing(self, code): return self.obj.IsBigListingStock(code)
    def is_arrg_sby(self, code): return self.obj.IsStockArrgSby(code) # 정리매매
    def is_low_liquidity(self, code): return self.obj.IsLowLiquidity(code) # 초저유동성
    def is_invest_danger(self, code): return self.obj.IsStockInvestDangerCompany(code) # 투자환기
    def is_stock_ioi(self, code): return self.obj.IsStockIoi(code) # ETN/ETF

    # --- [7] 기타 설정 및 시간 ---
    def get_market_times(self):
        """장 시작/종료 시간 반환 (900, 1530 등)"""
        return {
            'start': self.obj.GetMarketStartTime(),
            'end': self.obj.GetMarketEndTime()
        }

    def reload_port_data(self):
        """관심종목 데이터 갱신"""
        self.obj.ReLoadPortData()