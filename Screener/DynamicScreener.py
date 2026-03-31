import path_finder
import os
import sys
from datetime import datetime
from API.MarketScanner import MarketScanner
from API.MarketDataManager import MarketDataManager # MarketDataManager 임포트
from Util.FileManager import FileManager

class DynamicScreener:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        
        # API 객체들 초기화
        self.mrk_scanner = MarketScanner()
        self.mdm = MarketDataManager() # 👈 실시간 데이터 보완용 추가
        
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
        """시간대별로 거래대금 허들을 계단식으로 조정"""
        now = datetime.now()
        curr_time = now.hour * 100 + now.minute

        if curr_time <= 910:
            return 5000000000   # 09:10 이전: 50억 (시가 형성기)
        elif curr_time <= 940:
            return 10000000000  # 09:10 ~ 09:40: 100억 (주도주 결정기)
        elif curr_time <= 1030:
            return 15000000000  # 09:40 ~ 10:30: 150억 (추세 확인기)
        else:
            return 20000000000  # 10:30 이후: 200억 (안정기)

    def run_screener(self):
        """
        [개선된 3단계 로직]
        1. 7043/7034를 통해 수급 및 가격 전략 후보군을 1차로 뽑습니다.
        2. 후보군 중 유니버스에 포함된 종목들만 추려 StockMst2로 정확한 '거래대금'을 조회합니다.
        3. 실제 거래대금(원 단위)과 수급 비중을 최종 비교하여 반환합니다.
        """
        # [1단계] 기초 데이터 수집
        cache = self.mrk_scanner.update_integrated_selection()
        whale_pass_dict = {
            s['code']: s for s in cache['7034'] 
            if s.get('buy_ratio', 0) >= 65.0
        }

        # [2단계] 유니버스에 있고 수급이 통과된 1차 후보 리스트업
        # [2단계] 전략별 히트(중복) 수 계산
        hit_counts = {}
        initial_candidates = []
        for stock_7043 in cache['7043']:
            code = stock_7043['code']
            if code in self.universe_dict and code in whale_pass_dict:
                initial_candidates.append(code)
                hit_counts[code] = hit_counts.get(code, 0) + 1 # 카운팅 수행
        
        # 중복 제거
        initial_candidates = list(set(initial_candidates))
        
        if not initial_candidates:
            return []

        # [3단계] StockMst2를 통해 정확한 거래대금(원 단위) 가져오기
        # 한 번에 최대 110개까지 조회 가능하므로 스캘핑 후보군 처리에 충분함
        accurate_data = self.mdm.get_hoga_detail(initial_candidates)
        if not accurate_data:
            return []

        final_candidates = []
        min_amount_threshold = self.get_dynamic_threshold()

        for data in accurate_data:
            code = data['code']
            # StockMst2의 amount는 시장 불문 '원' 단위임
            actual_amount = data['amount'] 
            
            # 최종 거래대금 허들 체크
            if actual_amount < min_amount_threshold:
                continue

            # 최종 리스트 구성
            whale_info = whale_pass_dict[code]
            final_candidates.append({
                'code': code,
                'name': data['name'],
                'price': data['current'],
                'diff_rate': round(((data['current'] - data['open']) / data['open'] * 100), 2) if data['open'] > 0 else 0,
                'amount_억': round(actual_amount / 100000000, 1),
                'buy_ratio': whale_info['buy_ratio'],
                'strength': data['strength'], # StockMst2에서 제공하는 체결강도 추가
                'hit_count': hit_counts[code], # 👈 다시 추가!
            })

        # 정렬: 큰손 비중 순
        final_candidates.sort(key=lambda x: x['buy_ratio'], reverse=True)
        return final_candidates


# --- 실행 테스트 ---
if __name__ == "__main__":
    # 1. 스캐너 객체 생성
    screener = DynamicScreener()
    
    # 2. 6단계 필터링(가격전략+수급+대금+중복제거+유니버스) 실행
    final_targets = screener.run_screener()
    
    # 3. 결과 리스트 출력
    print("\n" + "=" * 85)
    print(f" 🎯 6단계 필터링 합격 종목 리스트 (총 {len(final_targets)}개)")
    print("-" * 85)
    print(f"{'종목명':<16}{'현재가':>10}{'등락률':>9}{'거래대금(억)':>14}{'큰손비중':>10}{'전략중복':>8}")
    print("-" * 85)
    
    if not final_targets:
        print(" 현재 조건(수급 65% 이상 & 거래대금 허들)을 만족하는 종목이 없습니다.")
    else:
        for t in final_targets:
            # 2개 이상의 전략 태그에 중복으로 걸린 종목은 강조 표시 (🚀)
            
            print(f"{t['name']:<14}{t['price']:>12,}{t['diff_rate']:>9.2f}%"
                  f"{t['amount_억']:>14.1f}억{t['buy_ratio']:>11.1f}%")
    
    print("=" * 85)