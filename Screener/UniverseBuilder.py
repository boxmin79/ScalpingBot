import path_finder
from datetime import datetime, timedelta
from API.CpAPI import CreonAPI
from API.MarketEye import MarketEye  # MarketEye 임포트
from API.MarketDataManager import MarketDataManager  # MarketDataManager 임포트
from Util.FileManager import FileManager

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

    def build_universe(self):
        """장 시작 전 불량 종목(스팩, 우선주, 거래정지 등)을 걸러내고 순수 보통주 명단을 만듭니다."""
        
       # 1. 기준 날짜(종료일) 및 시작일 계산
        now = datetime.now()
        curr_time = now.hour * 100 + now.minute
        
        # 장 마감(15:30) 기준 종료일 설정
        if curr_time < 1530:
            end_date = (now - timedelta(days=1)).strftime('%Y%m%d')
        else:
            end_date = now.strftime('%Y%m%d')
            
        # 시작일 설정 (넉넉하게 100일 전)
        start_date = (now - timedelta(days=100)).strftime('%Y%m%d')
        
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 1단계: 기본 필터링 시작...")
        
        # CreonAPI 객체 내부의 obj_code_mgr 사용
        kospi_codes = self.api.obj_code_mgr.GetStockListByMarket(1)
        kosdaq_codes = self.api.obj_code_mgr.GetStockListByMarket(2)
        all_codes = list(kospi_codes) + list(kosdaq_codes)
        
        pre_filtered_list = []
        drop_reasons = {
            'section': 0, 
            'spac': 0, 
            'pref': 0, 
            'status': 0, 
            'control': 0, 
            'liquidity': 0,
            'finance': 0, # ★ 초저유동성 필터 카운트 추가
            'turnover': 0
            }

        for code in all_codes:
            # 1. 주권(1)이 아닌 파생/복합 상품 제외 (ETF, ETN 등)
            if self.api.obj_code_mgr.GetStockSectionKind(code) != 1:
                drop_reasons['section'] += 1
                continue
            
            # 2. 스팩(SPAC) 제외
            if self.api.obj_code_mgr.IsSpac(code):
                drop_reasons['spac'] += 1
                continue
            
            # 3. 우선주 제외 (끝자리가 '0'이 아닌 종목)
            if code[-1] != '0':
                drop_reasons['pref'] += 1
                continue
            
            # 4. 거래정지/중단 제외 (0: 정상)
            if self.api.obj_code_mgr.GetStockStatusKind(code) != 0:
                drop_reasons['status'] += 1
                continue
            
            # 5. 투자주의/경고/위험/관리 종목 제외 (0: 정상)
            if self.api.obj_code_mgr.GetStockControlKind(code) != 0:
                drop_reasons['control'] += 1
                continue
            
            # 6. ★ 초저유동성 종목 제외 (단일가 매매 대상)
            # IsLowLiquidity(code): 초저유동성 종목이면 True 반환
            if self.api.obj_code_mgr.IsLowLiquidity(code):
                drop_reasons['liquidity'] += 1
                continue
            
            pre_filtered_list.append(code)
        
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 2단계: 재무 필터링 시작 (대상: {len(pre_filtered_list)}종목)...")    
        
        # [STEP 2] MarketEye를 이용한 재무 필터링 (200종목씩 끊어서 요청)
        final_universe = []
        # 조회할 필드: 0(코드), 20(상장주식수), 77(ROE), 92(영업이익률), 94(이자보상비율)
        target_fields = [0, 20, 77, 92, 94, 75, 76, 4]
        
        for i in range(0, len(pre_filtered_list), 200):
            chunk_codes = pre_filtered_list[i:i+200]
            # MarketEye.py의 get_market_data 호출
            market_data, _ = self.market_eye.get_market_data(chunk_codes, target_fields)
            
            for data in market_data:
                code = data[0]
                roe = data[77]
                op_margin = data[92]
                int_coverage = data[94]
                debt_ratio = data[75]      # 부채비율
                reserve_ratio = data[76]   # 유보율
                current_price = data[4]    # 현재가
                listed_shares = data[20]   # 상장주식수
                # 시가총액 계산 (단위: 원)
                market_cap = current_price * listed_shares

                # 1. 재무 건전성 필터
                is_quality = roe >= 5 and op_margin >= 5 and int_coverage >= 1
                is_stable = debt_ratio <= 200 and reserve_ratio >= 500
                is_liquid = market_cap >= 50_000_000_000 # 시총 500억 이상

                if is_quality and is_stable and is_liquid:
                    # 2. ★ 평균 회전율 필터 (MarketDataManager 사용)
                    # 20일치 일봉 거래량 데이터를 가져옵니다.
                    
                    # 🎯 2. 계산된 target_date를 get_chart_data에 전달
                    # int()로 형변환하여 전달 (API 요구 포맷)
                    chart_data = self.mdm.get_chart_data(
                        stk_code=code, 
                        req_type='1', 
                        end_date=int(end_date), 
                        start_date=int(start_date),
                        target_count=100  # 함수 내부의 데이터 수집 루프 종료 조건으로 사용됨
                        )
                    
                    # 데이터가 정상적으로 수신되었는지 확인
                    if chart_data and len(chart_data) >= 60:
                        # 🎯 실제 거래대금(amt) 필드를 사용하여 평균 계산 (가장 정확함)
                        vol_20_list = [day['vol'] for day in chart_data[:20]]
                        vol_60_list = [day['vol'] for day in chart_data[:60]]
                        amt_20_list = [day['amt'] for day in chart_data[:20]] # 실제 거래대금 리스트
                        
                        avg_vol_20 = sum(vol_20_list) / 20
                        avg_vol_60 = sum(vol_60_list) / 60
                        avg_amt_20 = sum(amt_20_list) / 20 # 20일 평균 실제 거래대금
                        
                        # [필터] 평균 거래대금이 50억 미만이면 탈락 (실제 수치 기반)
                        if avg_amt_20 < 2_000_000_000:
                            drop_reasons['liquidity'] += 1
                            continue
                        
                        # 20일 평균 회전율 계산
                        turnover_ratio_20 = (avg_vol_20 / listed_shares) * 100
                        
                        # 기준: 20일 평균 회전율 0.2% 이상 (시장 상황에 따라 조정 가능)
                        if turnover_ratio_20 >= 0.2:
                            name = self.api.obj_code_mgr.CodeToName(code)
                            market_kind = self.api.obj_code_mgr.GetStockMarketKind(code)
                            market_name = "KOSPI" if market_kind == 1 else "KOSDAQ"

                            final_universe.append({
                                "code": code,
                                "name": name,
                                "market": market_name,
                                "market_cap": round(market_cap / 100_000_000, 1), # 억 단위 저장
                                "roe": round(roe, 2),
                                "avg_vol_20": int(avg_vol_20),  # 🎯 추가: 20일 평균 거래량
                                "avg_vol_60": int(avg_vol_60),  # 🎯 추가: 60일 평균 거래량
                                "avg_turnover_20": round(turnover_ratio_20, 3)
                            })
                        else:
                            drop_reasons['turnover'] += 1
                    else:
                        drop_reasons['turnover'] += 1
                else:
                    drop_reasons['finance'] += 1

        print("-" * 45)
        print(f"🗑️ [필터링 탈락 요약]")
        print(f"- 종목구분/정지/주의 등  : {sum(drop_reasons.values()) - drop_reasons['finance'] - drop_reasons['turnover']}개") # 
        print(f"- 재무미달 (수익/안정)   : {drop_reasons['finance']}개")
        print(f"- 회전율미달 (활동성)    : {drop_reasons['turnover']}개")
        print("-" * 45)
        print(f"🎯 [최종 유니버스]: {len(final_universe)}개")
        
        self.universe_data = final_universe
        return self.universe_data
    
    def save_universe(self):
        """구축된 유니버스를 FileManager를 통해 JSON으로 자동 저장합니다."""
        if not self.universe_data:
            print("❌ 저장할 데이터가 없습니다. 먼저 build_universe()를 실행하세요.")
            return False
            
        # 확장자가 .json이므로 FileManager가 알아서 json 형식으로 저장합니다.
        return self.file_mgr.save(self.universe_data, self.file_path)

    def load_universe(self):
        """저장된 유니버스 JSON 파일을 읽어옵니다. (장중 봇 재시작 시 활용)"""
        loaded_data = self.file_mgr.load(self.file_path)
        
        # 데이터가 None이 아니고, 내용이 비어있지 않은지 확인 (순수 파이썬 리스트 검증)
        if loaded_data and len(loaded_data) > 0:
            self.universe_data = loaded_data
            print(f"[시스템] ✅ 기존 유니버스 로드 완료 (총 {len(self.universe_data)}종목)")
            return self.universe_data
        else:
            print("[시스템] ⚠️ 유니버스 파일이 없거나 비어있습니다. 새로 구축을 시도합니다.")
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
    
    