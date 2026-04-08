import os
import pandas as pd
from bs4 import BeautifulSoup
import path_finder
from datetime import datetime, date
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
        
        # 저장 파일명 변경 (백테스트용 데이터임을 명시)
        self.file_path = r"C:\Users\realb\Documents\TradingBot\BackTest\backtest_base_universe.json"
        self.universe_data = []

    def _get_fnguide_floating_data(self, code):
        """컴퍼니가이드에서 유동주식수와 유동비율을 추출합니다."""
        try:
            clean_code = code if code.startswith('A') else f"A{code}"
            url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?gicode={clean_code}"
            
            tables = pd.read_html(url, header=0)
            snapshot_table = tables[0]
            
            target_row = snapshot_table[snapshot_table.iloc[:, 0].str.contains('유동주식수', na=False)]
            
            if not target_row.empty:
                raw_val = target_row.iloc[0, 1]
                # "주식수 / 비율" 형태 파싱
                shares_str, ratio_str = raw_val.split('/')
                
                # 콤마 제거 및 수치화
                floating_shares = int(shares_str.replace(',', '').strip())
                floating_ratio = float(ratio_str.strip())
                
                return floating_shares, floating_ratio
            
            return None, None
        except Exception as e:
            print(f"[{code}] FnGuide 크롤링 실패: {e}")
            return None, None
    
    def build_universe(self):
        """기본 재무와 유동성 데이터만 수집하여 유니버스를 구축합니다."""
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 백테스트 기초 유니버스 구축 시작...")
        
        # 1. 상장 상태 기준 기초 명단 필터링 (스팩, 우선주, 정지 종목 등 제외)
        pre_filtered_list, drop_reasons = self._get_basic_stock_list()
        
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 2단계: 재무 필터 및 유동 데이터 수집 (대상: {len(pre_filtered_list)}종목)...")
        
        final_universe = []
        # 필드: 0(코드), 20(주식수), 77(ROE), 92(영익률), 94(이자보상), 75(부채), 76(유보), 4(현재가)
        target_fields = [0, 20, 77, 92, 94, 75, 76, 4]
                
        for i in range(0, len(pre_filtered_list), 200):
            chunk_codes = pre_filtered_list[i:i+200]
            market_data, _ = self.market_eye.get_market_data(chunk_codes, target_fields)
            
            for data in market_data:
                code = data[0]
                
                # 2. 기본적인 재무 건전성 필터 (최소한의 필터만 남김)
                if not self._is_financially_sound(data):
                    drop_reasons['finance'] += 1
                    continue
                
                ###########################################################
                # 3. 유동주식 데이터 수집 (FnGuide)
                # --- 🎯 3. 유동주식 데이터 가져오기 (수정된 부분) ---
                # fdm의 캐시(dict)에서 즉시 데이터를 찾습니다. (O(1) 속도)
                f_data = self.fdm.get_data(code)
                
                # [CASE 1] 마스터 데이터 파일에 종목 코드 자체가 없는 경우 -> 즉시 종료
                if f_data is None:
                    print(f"\n❌ [중단] {code} ({self.api.obj_code_mgr.CodeToName(code)}) 종목의 유동주식 데이터가 없습니다.")
                    print("👉 FloatingDataManager를 먼저 실행하여 마스터 데이터를 최신화하세요.")
                    return None  # 또는 sys.exit()를 사용하여 프로그램 전체 종료 가능
                
                # [CASE 2] 데이터는 있으나 크롤링 실패 등으로 값이 None인 경우 -> 기존처럼 스킵
                if f_data.get('유동주식수') is None:
                    drop_reasons['finance'] += 1
                    continue
                    
                f_shares = f_data['유동주식수']
                f_ratio = f_data['유동주식비율'] 
                ###########################################################

                # 실제 상장주식수 보정
                listed_shares = self._get_actual_listed_shares(code, data[20])
                market_cap = data[4] * listed_shares
                
                # 백테스트에 필요한 모든 기초 정보를 저장
                final_universe.append({
                    "code": code,
                    "name": self.api.obj_code_mgr.CodeToName(code),
                    "market": "KOSPI" if self.api.obj_code_mgr.GetStockMarketKind(code) == 1 else "KOSDAQ",
                    "market_cap": market_cap,
                    "listed_shares": listed_shares,
                    "floating_shares": f_shares,
                    "floating_ratio": f_ratio,
                    "current_price": data[4]
                })

        self._print_summary(drop_reasons, final_universe)
        self.universe_data = final_universe
        return self.universe_data

    def _get_basic_stock_list(self):
        """기본적인 종목 상태를 필터링합니다."""
        codes = list(self.api.obj_code_mgr.GetStockListByMarket(1)) + list(self.api.obj_code_mgr.GetStockListByMarket(2))
        pre_filtered = []
        reasons = {'section': 0, 'spac': 0, 'pref': 0, 'status': 0, 'control': 0, 'liquidity': 0, 'finance': 0}

        for code in codes:
            if self.api.obj_code_mgr.GetStockSectionKind(code) != 1: reasons['section'] += 1; continue # 주권만
            if self.api.obj_code_mgr.IsSpac(code): reasons['spac'] += 1; continue
            if code[-1] != '0': reasons['pref'] += 1; continue # 보통주만
            if self.api.obj_code_mgr.GetStockStatusKind(code) != 0: reasons['status'] += 1; continue # 정상종목만
            if self.api.obj_code_mgr.GetStockControlKind(code) != 0: reasons['control'] += 1; continue # 관리/투자유의 제외
            pre_filtered.append(code)
            
        return pre_filtered, reasons

    def _is_financially_sound(self, data):
        """최소한의 상장 적격성 및 재무 기준만 확인합니다."""
        # 시총 300억 미만 컷 (너무 작은 종목 제외)
        market_cap = data[4] * data[20]
        if market_cap < 30_000_000_000: return False 
        
        # 완전 자본잠식이나 심각한 부실만 필터링 (기준 완화)
        is_stable = data[75] <= 500 and data[76] >= 100 # 부채 500% 이하, 유보율 100% 이상
        return is_stable

    def _get_actual_listed_shares(self, code, raw_shares):
        if self.api.obj_code_mgr.IsBigListingStock(code):
            return raw_shares * 1000
        return raw_shares
    
    def _print_summary(self, drop_reasons, final_universe):
        print("-" * 45)
        print(f"🗑️ [필터링 탈락 요약]")
        print(f"- 종목상태/제외대상 : {sum(list(drop_reasons.values())[:6])}개")
        print(f"- 최소 재무기준 미달 : {drop_reasons['finance']}개")
        print("-" * 45)
        print(f"🎯 [백테스트 기초 유니버스]: {len(final_universe)}개")
    
    def save_universe(self):
        if not self.universe_data:
            print("❌ 저장할 데이터가 없습니다.")
            return False
        return self.file_mgr.save(self.universe_data, self.file_path)

if __name__ == "__main__":
    builder = UniverseBuilder()
    universe = builder.build_universe()
    builder.save_universe()