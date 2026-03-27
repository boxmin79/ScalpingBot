import path_finder
import os
import sys
from datetime import datetime
from API.CpTopVolume import CpTopVolume
from Util.FileManager import FileManager

class DynamicScreener:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        
        # 1. 기존에 만들어두신 거래대금 상위 조회 객체 초기화
        self.top_volume_api = CpTopVolume()
        
        # 2. 파일 매니저 및 1단계 유니버스 로드
        self.file_mgr = FileManager()
        self.universe_path = self.cfg.CODE_DIR / "scalping_universe.json"
        # 유니버스 리스트를 탐색 속도가 빠른 Set(집합) 형태로 저장합니다. (O(1) 속도)
        self.universe_dict = self._load_universe_dict()

    def _load_universe_dict(self):
        """유니버스를 { 'A005930': {'name': '삼성전자', 'market': 'KOSPI'}, ... } 형태의 딕셔너리로 로드"""
        loaded_data = self.file_mgr.load(self.universe_path)
        
        if loaded_data and len(loaded_data) > 0:
            # 리스트를 딕셔너리로 변환 (검색 속도 극대화 및 시장 정보 포함)
            univ_dict = {item['code']: item for item in loaded_data}
            print(f"[시스템] 유니버스 딕셔너리 로드 완료: {len(univ_dict)} 종목 대기 중")
            return univ_dict
        else:
            print("❌ 유니버스 파일이 없거나 비어있습니다. UniverseBuilder를 먼저 실행하세요.")
            return {}

    def get_dynamic_threshold(self):
        """장 초반(09:00~09:10) 거래대금 허들 완화"""
        now = datetime.now()
        if now.hour == 9 and now.minute <= 10:
            return 5000000000   # 50억
        else:
            return 20000000000  # 200억

    def run_screener(self, top_n=20):
        """
        CpTopVolume 데이터를 받아와서 1단계 유니버스와 교집합을 구합니다.
        """
        if not self.universe_dict:
            return []

        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 2단계: 하이브리드 주도주 포착 시작...")
        
        # 1. 회원님의 CpTopVolume을 이용해 서버 공식 랭킹 데이터를 가져옵니다.
        # (앞서 보여주신 {'rank': 1, 'code': 'A005930', ...} 형태의 리스트가 반환된다고 가정)
        raw_top_data = self.top_volume_api.get_top_list() 
        # print(raw_top_data)
        
        min_amount = self.get_dynamic_threshold()
        filtered_targets = []

        # 2. 필터링 시작
        for stock in raw_top_data:
            code = stock.get('code', '')
            name = stock.get('name', '')
            diff_rate = float(stock.get('diff_rate', 0.0))
            raw_amount = int(stock.get('amount', 0))
            # [핵심 1] 유니버스에 없는 종목 무조건 버리기 & 시장 정보 가져오기
            univ_info = self.universe_dict.get(code)
            # print(f"unive_info: {univ_info}")
            if not univ_info:
                continue
                
            # [핵심 2] ★★★ 시장별 거래대금 단위 보정 ★★★
            market = univ_info.get('market', 'KOSPI')
            
            if market == 'KOSPI':
                actual_amount = raw_amount * 10000  # 코스피: 만원 -> 원
            else: # 'KOSDAQ'
                actual_amount = raw_amount * 1000   # 코스닥: 천원 -> 원
            
            # [핵심 3] 등락률 및 실제 거래대금(actual_amount) 조건 검사
            if 3.0 <= diff_rate <= 20.0:
                # print(f'diff_rate // code: {code}, name: {name}, diff_rate: {diff_rate}, actual_amount: {actual_amount}')
                if actual_amount >= min_amount:
                    # print(f'actual_amount 통과 : {code}, {name}')
                    # 통과한 종목은 계산된 실제 금액과 이름을 업데이트해서 넣습니다.
                    stock['actual_amount'] = actual_amount
                    stock['name'] = univ_info['name'] 
                    # print(f'actual_amount 통과 : {stock}')
                    filtered_targets.append(stock)

            if len(filtered_targets) >= top_n:
                break

        return filtered_targets


# --- 실행 테스트 ---
if __name__ == "__main__":
    screener = DynamicScreener()
    targets = screener.run_screener(top_n=20)
    for target in targets:
        print(f"rank: {target['rank']}, code: {target['code']}, name: {target['name']}, diff_rate: {target['diff_rate']}, amount: {target['actual_amount']}")