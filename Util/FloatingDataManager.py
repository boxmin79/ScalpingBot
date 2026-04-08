import os
import pandas as pd
import path_finder
from datetime import date
from concurrent.futures import ThreadPoolExecutor, as_completed
from API.CpAPI import CreonAPI
from Util.FileManager import FileManager

class FloatingDataManager:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        self.fm = FileManager()
        self.api = CreonAPI()  # 종목 리스트 수집을 위한 API 초기화
        
        # 파일 경로 설정 (기존 필터와 분리된 유동주식 전용 마스터 데이터)
        self.cache_file = self.cfg.CODE_DIR / "floating_master_data.json"
        
        # 내부 관리는 dict 형태(O(1) 검색용), 저장은 요청하신 list[dict] 형태
        self.floating_cache = self._load_cache()

    def _load_cache(self):
        """파일을 로드하여 {코드: 데이터} 형태의 딕셔너리로 변환합니다."""
        if not self.cache_file.exists():
            return {}
        
        raw_list = self.fm.load(self.cache_file)
        if not raw_list:
            return {}
        
        # 리스트 형태를 검색 효율을 위해 딕셔너리로 변환
        return {item['code']: item for item in raw_list}

    def _save_cache(self):
        """딕셔너리를 요청하신 [{'code':...}, ...] 리스트 형태로 변환하여 저장합니다."""
        save_list = list(self.floating_cache.values())
        return self.fm.save(save_list, self.cache_file)

    def _get_all_market_codes(self):
        """크레온 API를 통해 코스피(1), 코스닥(2)의 모든 종목 코드를 가져옵니다."""
        kospi = list(self.api.obj_code_mgr.GetStockListByMarket(1))
        kosdaq = list(self.api.obj_code_mgr.GetStockListByMarket(2))
        return kospi + kosdaq

    def update_data(self, codes=None, force_full=False, max_workers=10):
        """
        데이터 업데이트 메인 로직
        :param codes: 수집할 종목 코드 리스트 (None이면 전체 시장 코드 수집)
        :param force_full: True이면 기존 캐시를 무시하고 해당 코드들을 무조건 재수집
        :param max_workers: 멀티스레드 작업자 수
        """
        # 1. 대상 코드 확정 (인자가 있으면 사용, 없으면 전체 수집)
        target_codes = codes if codes is not None else self._get_all_market_codes()
        
        if force_full:
            print(f"[FloatingMgr] ⚠️ 지정된 {len(target_codes)}종목에 대해 강제 재수집을 시작합니다.")
            needed_codes = target_codes
            # 전체 시장 강제 업데이트일 경우에만 캐시 전체 초기화
            if codes is None:
                self.floating_cache = {}
        else:
            # 2. 기존 캐시와 대조하여 없는 코드(신규)만 추출
            # (데이터가 None으로 저장된 종목도 '필드로 존재'하므로 중복 수집에서 제외됨)
            needed_codes = [c for c in target_codes if c not in self.floating_cache]
            print(f"[FloatingMgr] 🔎 기존 데이터 확인 완료. (신규 추가 필요: {len(needed_codes)}종목)")

        if not list(needed_codes):
            print("[FloatingMgr] ✅ 업데이트할 새로운 종목이 없습니다.")
            return

        print(f"[FloatingMgr] 🔄 수집 시작 (대상: {len(needed_codes)} / Thread: {max_workers})")
        
        results_count = 0
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_code = {executor.submit(self._fetch_fnguide, code): code for code in needed_codes}
            
            for future in as_completed(future_to_code):
                code, res = future.result()
                # 데이터가 있든 없든(None 포함 딕셔너리) 캐시에 저장하여 다음 루프에서 제외
                if res:
                    self.floating_cache[code] = res
                    results_count += 1
                
                if results_count % 50 == 0:
                    print(f"[FloatingMgr] 진행 중... ({results_count}/{len(needed_codes)})")

        self._save_cache()
        print(f"[FloatingMgr] ✨ 저장 완료! (현재 총 {len(self.floating_cache)}종목 데이터 관리 중)")

    def _fetch_fnguide(self, code):
        """FnGuide 크롤링 후 데이터가 없으면 None 값이 담긴 딕셔너리 반환 (중복 수집 방지용)"""
        try:
            clean_code = code if code.startswith('A') else f"A{code}"
            url = f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?gicode={clean_code}"
            
            # 네트워크 지연 등을 고려해 pandas 대신 requests + lxml(또는 bs4) 조합이 더 빠를 수 있지만, 
            # 기존 구조 유지를 위해 read_html 사용
            tables = pd.read_html(url, header=0)
            if not tables:
                return code, {'code': code, '유동주식수': None, '유동주식비율': None}

            snapshot_table = tables[0]
            target_row = snapshot_table[snapshot_table.iloc[:, 0].str.contains('유동주식수', na=False)]
            
            if not target_row.empty:
                raw_val = target_row.iloc[0, 1]
                shares_str, ratio_str = raw_val.split('/')
                
                return code, {
                    'code': code,
                    '유동주식수': int(shares_str.replace(',', '').strip()),
                    '유동주식비율': float(ratio_str.strip())
                }
            
            # 유동주식수 항목이 없는 경우
            return code, {'code': code, '유동주식수': None, '유동주식비율': None}

        except Exception:
            # 페이지 오류, 네트워크 단절 등 모든 예외 상황 처리
            return code, {'code': code, '유동주식수': None, '유동주식비율': None}

    def get_data(self, code):
        """개별 종목 데이터 반환"""
        return self.floating_cache.get(code)

# --- 실행 테스트 ---
if __name__ == "__main__":
    mgr = FloatingDataManager()
    
    # 1. 없는 데이터만 추가하고 싶을 때 (일반적인 경우)
    # mgr.update_data(force_full=False)
    
    # 2. 전체를 새로 싹 긁고 싶을 때 (주기적인 전체 갱신 시)
    # mgr.update_data(force_full=True)
    fd = mgr.get_data('A005930')
    print(fd)