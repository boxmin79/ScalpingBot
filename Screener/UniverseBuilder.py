import os
import pandas as pd
from datetime import datetime, timedelta, date
import path_finder
from API.CpAPI import CreonAPI
from API.MarketEye import MarketEye
from API.MarketDataManager import MarketDataManager
from Util.FileManager import FileManager
from Util.FloatingDataManager import FloatingDataManager

class UniverseBuilder:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        self.api = CreonAPI()
        self.market_eye = MarketEye()
        self.mdm = MarketDataManager()
        self.file_mgr = FileManager()
        self.fdm = FloatingDataManager()
        
        # 저장 경로
        self.file_path = self.cfg.CODE_DIR / "scalping_universe.json"
        self.universe_data = []

    def build_universe(self):
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 백테스트 최적화 유니버스 구축 시작...")
        
        end_date, start_date = self._get_target_dates()
        pre_filtered_list, drop_reasons = self._get_basic_stock_list()
        
        final_universe = []
        # 필드: 0(코드), 20(주식수), 77(ROE), 92(영익률), 94(이자보상), 75(부채), 76(유보), 4(현재가)
        target_fields = [0, 20, 77, 92, 94, 75, 76, 4]
                
        for i in range(0, len(pre_filtered_list), 200):
            chunk_codes = pre_filtered_list[i:i+200]
            market_data, _ = self.market_eye.get_market_data(chunk_codes, target_fields)
            
            for data in market_data:
                # 1. 재무 건전성 및 최소 시총(1,600억) 필터
                frd = self.fdm.get_data(data[0]) 
                
                # 🔥 [안전장치 추가] 데이터가 없거나 유동주식수가 None인 경우 스킵
                if frd is None or frd.get('유동주식수') is None or frd.get('유동주식비율') is None:
                    # 데이터가 없는 종목은 보수적으로 '거래 강도 미달' 또는 '재무 미달'로 간주하여 탈락시킵니다.
                    drop_reasons['finance'] += 1
                    continue
                
                floating_ratio = frd['유동주식비율']
                
                if data[4] * frd['유동주식수'] < 160_000_000_000:
                    drop_reasons['finance'] += 1
                    continue
                
                # 2. 유동비율 데이터 로드 및 컷 (품절주 제외)
                if floating_ratio is None or floating_ratio < 15.0:
                    drop_reasons['finance'] += 1
                    continue
                
                # 3. 백테스트 기반 시총별 차등 회전율/거래대금 필터
                activity_data = self._check_activity_filter(data, floating_ratio, end_date, start_date)
                if activity_data:
                    final_universe.append(activity_data)
                else:
                    drop_reasons['turnover'] += 1

        self._print_summary(drop_reasons, final_universe)
        self.universe_data = final_universe
        return self.universe_data

    def _is_financially_sound(self, data):
        """재무 건전성과 최소 시총 기준(1,600억)을 확인합니다."""
        # 시총 계산 (현재가 * 상장주식수)
        market_cap = data[4] * self._get_actual_listed_shares(data[0], data[20])
        
        # [수정] 백테스트 결과 기반: 1,600억 미만 그룹은 노이즈가 많아 제외
        if market_cap < 160_000_000_000: 
            return False 
        
        # 기본 재무 컷 (ROE, 영익률 등)
        is_quality = data[77] >= 5 and data[92] >= 5 and data[94] >= 1
        is_stable = data[75] <= 200 and data[76] >= 500
        return is_quality and is_stable

    def _check_activity_filter(self, data, floating_ratio, end_date, start_date):
        """백테스트 데이터 기반 시총별 차등 필터링"""
        code = data[0]
        listed_shares = self._get_actual_listed_shares(code, data[20])
        floating_shares = listed_shares * (floating_ratio / 100.0)
        market_cap = data[4] * listed_shares
        
        chart_data = self.mdm.get_chart_data(stk_code=code, req_type='1', end_date=int(end_date), start_date=int(start_date), target_count=60)
        if not chart_data or len(chart_data) < 20: return None

        avg_amt_20 = sum([day['amt'] for day in chart_data[:20]]) / 20
        avg_vol_20 = sum([day['vol'] for day in chart_data[:20]]) / 20
        turnover_20 = (avg_vol_20 / floating_shares) * 100 if floating_shares > 0 else 0
        
        # [수정] 백테스트 기반 최적 유니버스 필터 기준 (거래대금 하한 100억)
        # 1. 공통 최소 거래대금 100억
        min_amt = 10_000_000_000 
        
        # 2. 시총 구간별 최소 회전율 설정 (백테스트 Group별 Turnover 데이터 반영)
        if market_cap >= 10_000_000_000_000:        # 10조 이상 (대형주)
            min_turnover = 1.0
        elif market_cap >= 2_000_000_000_000:       # 2조 ~ 10조
            min_turnover = 2.0
        elif market_cap >= 700_000_000_000:        # 7천억 ~ 2조
            min_turnover = 4.0
        elif market_cap >= 300_000_000_000:        # 3천억 ~ 7천억
            min_turnover = 6.0
        else:                                      # 1.6천억 ~ 3천억 (중소형 활황주)
            min_turnover = 9.0

        if avg_amt_20 >= min_amt and turnover_20 >= min_turnover:
            return {
                "code": code,
                "name": self.api.obj_code_mgr.CodeToName(code),
                "market_cap": int(market_cap),
                "avg_amt_20": int(avg_amt_20),
                "avg_turnover_20": round(turnover_20, 3),
                "floating_ratio": floating_ratio
            }
        return None

    # --- 기존 보조 함수들 (유지) ---
    def _get_target_dates(self):
        now = datetime.now()
        curr_time = now.hour * 100 + now.minute
        end_date = (now - timedelta(days=1)).strftime('%Y%m%d') if curr_time < 1530 else now.strftime('%Y%m%d')
        start_date = (now - timedelta(days=100)).strftime('%Y%m%d')
        return end_date, start_date

    def _get_basic_stock_list(self):
        codes = list(self.api.obj_code_mgr.GetStockListByMarket(1)) + list(self.api.obj_code_mgr.GetStockListByMarket(2))
        pre_filtered = []
        reasons = {'section': 0, 'spac': 0, 'pref': 0, 'status': 0, 'control': 0, 'liquidity': 0, 'finance': 0, 'turnover': 0}
        for code in codes:
            if self.api.obj_code_mgr.GetStockSectionKind(code) != 1: reasons['section'] += 1; continue
            if self.api.obj_code_mgr.IsSpac(code): reasons['spac'] += 1; continue
            if code[-1] != '0': reasons['pref'] += 1; continue
            if self.api.obj_code_mgr.GetStockStatusKind(code) != 0: reasons['status'] += 1; continue
            if self.api.obj_code_mgr.GetStockControlKind(code) != 0: reasons['control'] += 1; continue
            if self.api.obj_code_mgr.IsLowLiquidity(code): reasons['liquidity'] += 1; continue
            pre_filtered.append(code)
        return pre_filtered, reasons

    def _get_actual_listed_shares(self, code, raw_shares):
        if self.api.obj_code_mgr.IsBigListingStock(code):
            return raw_shares * 1000
        return raw_shares

    def _print_summary(self, drop_reasons, final_universe):
        print("-" * 45)
        print(f"🎯 [최종 유니버스]: {len(final_universe)}개 (백테스트 최적화 적용)")
        print(f"- 탈락(재무/시총): {drop_reasons['finance']} / 탈락(거래강도): {drop_reasons['turnover']}")
        print("-" * 45)

    def save_universe(self):
        if not self.universe_data: return False
        return self.file_mgr.save(self.universe_data, self.file_path)
    
if __name__ == "__main__":
    ub = UniverseBuilder()
    ub.build_universe()
    ub.save_universe()
    # print(ub.universe_data)