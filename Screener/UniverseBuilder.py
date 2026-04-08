import os
import requests
import pandas as pd
from bs4 import BeautifulSoup
import path_finder
from datetime import datetime, timedelta, date
from API.CpAPI import CreonAPI
from API.MarketEye import MarketEye  # MarketEye 임포트
from API.MarketDataManager import MarketDataManager  # MarketDataManager 임포트
from Util.FileManager import FileManager
from Util.FloatingDataManager import FloatingDataManager


class UniverseBuilder:
    def __init__(self):
        # 경로 설정
        self.cfg = path_finder.get_cfg()
        
        # 1. API 통합 객체 초기화 (이 안에서 Cybos 연결 체크도 자동으로 수행됨)
        self.api = CreonAPI()
        self.market_eye = MarketEye()
        self.mdm = MarketDataManager() # MarketDataManager 초기화
                
        # 2. 파일 매니저 초기화
        self.file_mgr = FileManager()
        
        # 3. 데이터 저장 경로 설정 (프로젝트 루트 안의 Data 폴더)
        self.file_path = self.cfg.CODE_DIR / "scalping_universe.json"
        self.universe_data = []
        
        self.fdm = FloatingDataManager()
        
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

    def _get_fnguide_floating_data(self, code):
        """
        컴퍼니가이드에서 유동주식수와 유동비율을 추출합니다.
        code: 'A005930' 형식
        """
        try:
            # 크레온 코드는 'A'가 붙어있을 수도 있고 아닐 수도 있으므로 전처리
            clean_code = code if code.startswith('A') else f"A{code}"
            url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?gicode={clean_code}"
            
            # 성능을 위해 pandas.read_html 사용 (속도가 빠름)
            tables = pd.read_html(url, header=0)
            snapshot_table = tables[0]  # 보통 첫 번째 테이블(IFRS 연결/개별 요약 위)에 위치
            
            # '유동주식수/비율' 행 찾기 (보통 '발행주식수' 근처에 위치함)
            # FnGuide 스냅샷 테이블 구조상 '유동주식수/비율' 텍스트를 포함하는 행 추출
            target_row = snapshot_table[snapshot_table.iloc[:, 0].str.contains('유동주식수', na=False)]
            
            if not target_row.empty:
                # 데이터 형식 예: "4,402,406,120 / 74.45"
                raw_val = target_row.iloc[0, 1]
                shares, ratio = raw_val.split('/')
                return float(ratio.strip())
            
            return None
        except Exception as e:
            print(f"[{code}] FnGuide 크롤링 실패: {e}")
            return None
    
    def build_universe(self):
        """기능별로 분리된 함수들을 호출하여 최종 유니버스를 구축합니다."""
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 유니버스 구축 시작...")
        
        # 1. 날짜 및 기초 명단 확보
        end_date, start_date = self._get_target_dates()
        pre_filtered_list, drop_reasons = self._get_basic_stock_list()
        
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 2단계: 상세 필터링 진행 (대상: {len(pre_filtered_list)}종목)...")
        
        final_universe = []
        # 필드: 0(코드), 20(주식수), 77(ROE), 92(영익률), 94(이자보상), 75(부채), 76(유보), 4(현재가)
        target_fields = [0, 20, 77, 92, 94, 75, 76, 4]
                
        # 200종목씩 끊어서 상세 검사
        for i in range(0, len(pre_filtered_list), 200):
            chunk_codes = pre_filtered_list[i:i+200]
            market_data, _ = self.market_eye.get_market_data(chunk_codes, target_fields)
            
            for data in market_data:
                # 2. 재무 및 시총 필터링
                if not self._is_financially_sound(data):
                    drop_reasons['finance'] += 1
                    continue
                
                # 2. 유동주식비율 필터링 (재무 통과 종목만 FnGuide에서 크롤링)
                # 매번 크롤링하면 느리므로, 여기서 실시간으로 가져옵니다.
                # floating_ratio = self._get_fnguide_floating_data(data[0])
                
                frd = self.fdm.get_data(data[0])
                floating_ratio = frd['유동주식비율']
                
                if floating_ratio is None or floating_ratio < 15.0: # 유동비율 15% 미만 품절주 제외
                    drop_reasons['finance'] += 1
                    continue
                
                # 3. 거래 활성도 필터링 (위에서 구한 floating_ratio 활용)
                activity_data = self._check_activity_filter(data, floating_ratio, end_date, start_date)
                if activity_data:
                    final_universe.append(activity_data)
                else:
                    drop_reasons['turnover'] += 1

        self._print_summary(drop_reasons, final_universe)
        self.universe_data = final_universe
        return self.universe_data

    # --- 세부 헬퍼 함수들 ---

    def _get_target_dates(self):
        """조회 기준일과 시작일을 계산합니다."""
        now = datetime.now()
        curr_time = now.hour * 100 + now.minute
        end_date = (now - timedelta(days=1)).strftime('%Y%m%d') if curr_time < 1530 else now.strftime('%Y%m%d')
        start_date = (now - timedelta(days=100)).strftime('%Y%m%d')
        return end_date, start_date

    def _get_basic_stock_list(self):
        """기본적인 종목 상태(스팩, 우선주, 정지 등)를 걸러냅니다."""
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

    def _is_financially_sound(self, data):
        """재무 건전성과 최소 시총 기준을 확인합니다."""
        # 필드 인덱스: 77(ROE), 92(영익률), 94(이자보상), 75(부채), 76(유보), 4(가), 20(수)
        market_cap = data[4] * data[20]
        if market_cap < 50_000_000_000: return False # 시총 500억 미만 컷
        
        is_quality = data[77] >= 5 and data[92] >= 5 and data[94] >= 1
        is_stable = data[75] <= 200 and data[76] >= 500
        return is_quality and is_stable

    def _check_activity_filter(self, data, floating_ratio, end_date, start_date):
        """
        시총별 차등 거래량/회전율 필터를 적용합니다.
        유동비율을 적용한 정밀 회전율 계산
        """
        code = data[0]
        raw_shares = data[20]
        # 🔥 [수정] 실제 주식수 계산
        listed_shares = self._get_actual_listed_shares(code, raw_shares)
        
        # 실제 유통되는 주식수 계산
        floating_shares = listed_shares * (floating_ratio / 100.0)
        
        market_cap = data[4] * listed_shares
        
        chart_data = self.mdm.get_chart_data(stk_code=code, req_type='1', end_date=int(end_date), start_date=int(start_date), target_count=100)
        
        if not chart_data or len(chart_data) < 60:
            return None

        # 20일 평균 계산
        avg_vol_20 = sum([day['vol'] for day in chart_data[:20]]) / 20
        avg_amt_20 = sum([day['amt'] for day in chart_data[:20]]) / 20
        
        # 2. 🔥 [추가] 60일 평균 거래량 계산 (RealtimeManager의 분당 평균 기준용)
        # RealtimeManager는 이 값을 390(장 운영 시간)으로 나누어 '분당 평균'을 구함
        avg_vol_60 = sum([day['vol'] for day in chart_data[:60]]) / 60
        
        # 유동주식수 대비 회전율 계산
        turnover_20 = (avg_vol_20 / floating_shares) * 100 if floating_shares > 0 else 0
        
        # 🔥 5단계 세분화 필터링
        if market_cap >= 100_000_000_000_000:       # 100조 이상 (Mega)
            min_amt, min_turnover = 500_000_000_000, 0.3
        elif market_cap >= 10_000_000_000_000:      # 10조 ~ 100조 (Large)
            min_amt, min_turnover = 100_000_000_000, 0.5
        elif market_cap >= 1_000_000_000_000:       # 1조 ~ 10조 (Mid-Large)
            min_amt, min_turnover = 8_000_000_000, 0.7  # 달바글로벌(87억) 포함
        elif market_cap >= 500_000_000_000:         # 5천억 ~ 1조 (Mid)
            min_amt, min_turnover = 5_000_000_000, 1.5
        else:                                       # 5천억 미만 (Small)
            min_amt, min_turnover = 2_000_000_000, 3.0
            

        if avg_amt_20 >= min_amt and turnover_20 >= min_turnover:
            return {
                "code": code,
                "name": self.api.obj_code_mgr.CodeToName(code),
                "market": "KOSPI" if self.api.obj_code_mgr.GetStockMarketKind(code) == 1 else "KOSDAQ",
                "market_cap": round(market_cap / 100_000_000, 1),
                "avg_amt_20": round(avg_amt_20 / 100_000_000, 1),
                "avg_vol_60": int(avg_vol_60), # 🎯 60일 평균 거래량 추가
                "floating_ratio": floating_ratio,
                "avg_turnover_20": round(turnover_20, 3)
            }
        return None

    def _get_actual_listed_shares(self, code, raw_shares):
        """API 수신 단위(1단위 vs 1000단위)를 판별하여 실제 주식수를 반환합니다."""
        # 20억 주 이상 대형 상장주인 경우 1000을 곱해줌
        if self.api.obj_code_mgr.IsBigListingStock(code):
            return raw_shares * 1000
        return raw_shares
    
    def _print_summary(self, drop_reasons, final_universe):
        """결과 요약을 출력합니다."""
        print("-" * 45)
        print(f"🗑️ [필터링 탈락 요약]")
        print(f"- 종목상태/저유동성 등 : {drop_reasons['section']+drop_reasons['spac']+drop_reasons['pref']+drop_reasons['status']+drop_reasons['control']+drop_reasons['liquidity']}개")
        print(f"- 재무/시총 미달       : {drop_reasons['finance']}개")
        print(f"- 거래 활성도 미달     : {drop_reasons['turnover']}개")
        print("-" * 45)
        print(f"🎯 [최종 유니버스]: {len(final_universe)}개")
    
    def save_universe(self):
        """구축된 유니버스를 FileManager를 통해 JSON으로 자동 저장합니다."""
        if not self.universe_data:
            print("❌ 저장할 데이터가 없습니다. 먼저 build_universe()를 실행하세요.")
            return False
            
        # 확장자가 .json이므로 FileManager가 알아서 json 형식으로 저장합니다.
        return self.file_mgr.save(self.universe_data, self.file_path)

    def load_universe(self):
        """저장된 유니버스를 읽어오되, 오늘 생성된 파일이 아니면 새로 구축합니다."""
        # 파일 수정 시간 확인
        if self.file_path.exists():
            file_date = date.fromtimestamp(os.path.getmtime(self.file_path))
            today = date.today()
            
            if file_date == today:
                loaded_data = self.file_mgr.load(self.file_path)
                if loaded_data:
                    print(f"[시스템] ✅ 오늘 생성된 유니버스 로드 완료 ({len(loaded_data)}종목)")
                    self.universe_data = loaded_data
                    return self.universe_data
            else:
                print("[시스템] 📅 날짜가 지난 유니버스 파일입니다. 새로 갱신합니다.")

        # 파일이 없거나 오늘 자 데이터가 아니면 새로 생성
        self.build_universe()
        self.save_universe()
        return self.universe_data

# --- 단독 실행 테스트용 ---
if __name__ == "__main__":
    builder = UniverseBuilder()
    universe = builder.build_universe()
    # 첫 번째 종목 예시 출력
    if universe:
        print(f"샘플 데이터: {universe[0]}")
    builder.save_universe()
    
    