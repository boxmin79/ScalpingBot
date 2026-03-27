import path_finder
from datetime import datetime
from API.CpAPI import CreonAPI
from Util.FileManager import FileManager

class UniverseBuilder:
    def __init__(self):
        # 경로 설정
        self.cfg = path_finder.get_cfg()
        
        # 1. API 통합 객체 초기화 (이 안에서 Cybos 연결 체크도 자동으로 수행됨)
        self.api = CreonAPI()
        
        # 2. 파일 매니저 초기화
        self.file_mgr = FileManager()
        
        # 3. 데이터 저장 경로 설정 (프로젝트 루트 안의 Data 폴더)
        self.file_path = self.cfg.CODE_DIR / "scalping_universe.json"
        self.universe_data = []

    def build_universe(self):
        """장 시작 전 불량 종목(스팩, 우선주, 거래정지 등)을 걸러내고 순수 보통주 명단을 만듭니다."""
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 1단계: 스캘핑 유니버스 구축 시작...")

        # CreonAPI 객체 내부의 obj_code_mgr 사용
        kospi_codes = self.api.obj_code_mgr.GetStockListByMarket(1)
        kosdaq_codes = self.api.obj_code_mgr.GetStockListByMarket(2)
        all_codes = list(kospi_codes) + list(kosdaq_codes)
        
        clean_data = []
        drop_reasons = {'section': 0, 'spac': 0, 'pref': 0, 'status': 0, 'control': 0}

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

            # 6. 시장 구분 가져오기 (1: 코스피, 2: 코스닥)
            market_kind = self.api.obj_code_mgr.GetStockMarketKind(code)
            # print(market_kind)
            market_name = "KOSPI" if market_kind == 1 else "KOSDAQ"

            # 생존 종목 추출 시 시장 정보(market) 추가
            name = self.api.obj_code_mgr.CodeToName(code)
            # print(name)
            clean_data.append({"code": code, "name": name, "market": market_name})

        print("-" * 45)
        print(f"🗑️ [1차 필터링 탈락 요약]")
        print(f"- 파생/복합상품 (ETF 등) : {drop_reasons['section']}개")
        print(f"- 스팩 (SPAC)            : {drop_reasons['spac']}개")
        print(f"- 우선주                 : {drop_reasons['pref']}개")
        print(f"- 거래정지/중단          : {drop_reasons['status']}개")
        print(f"- 주의/경고/위험         : {drop_reasons['control']}개")
        print("-" * 45)
        print(f"🎯 [최종 생존 스캘핑 타깃]: {len(clean_data)}개")
        
        self.universe_data = clean_data
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
    import logging
    # FileManager에서 출력하는 로그를 보기 위해 기본 로깅 설정
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    
    builder = UniverseBuilder()
    
    # 1. 새롭게 유니버스 구축하고 저장하기 (매일 아침 08:30 실행 목적)
    builder.build_universe()
    builder.save_universe()
    
    # 2. 잘 저장되었는지 다시 로드해보기 테스트
    print("\n--- 로드 테스트 ---")
    builder.load_universe()
    