import path_finder
import os
import sys
from datetime import datetime
from API.MarketScanner import MarketScanner
from Util.FileManager import FileManager

class DynamicScreener:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        
        # API 객체들 초기화
        self.mrk_scanner = MarketScanner()
        
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

        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🔍 2단계: 입체적 주도주 스캐닝 시작...")
        
        # 1. 회원님의 CpTopVolume을 이용해 서버 공식 랭킹 데이터를 가져옵니다.
        # (앞서 보여주신 {'rank': 1, 'code': 'A005930', ...} 형태의 리스트가 반환된다고 가정)
        raw_top_data = self.mrk_scanner.get_top_volume_list()
        # print(raw_top_data)
        
        # [B] 20일 신고가 돌파 리스트 (거래대금 순 정렬)
        breakout_list = self.mrk_scanner.get_breakout_list(market='0', criteria='6', sort_by=61, period='2')
        breakout_codes = {s['code'] for s in breakout_list} # 빠른 검색을 위한 set
        
        # [C] 큰손 매수 집중 리스트 (4천만원 이상, 코스닥 중심)
        whale_list = []
        kospi_whale_list = self.mrk_scanner.get_whale_ratio(market='1', amount='4', criteria='1')
        kosdaq_whale_list = self.mrk_scanner.get_whale_ratio(market='2', amount='4', criteria='1'
                                                    )
        whale_list.extend(kospi_whale_list)
        whale_list.extend(kosdaq_whale_list)
        
        whale_info = {s['code']: s['buy_ratio'] for s in whale_list}
        
        min_amount = self.get_dynamic_threshold()
        filtered_targets = []

        # 2. 필터링 시작
        for stock in raw_top_data:
            code = stock.get('code', '')
            
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
            if not (3.0 <= diff_rate <= 20.0 and actual_amount >= min_amount):
                continue
                

            # [점수화] 신고가 돌파 여부와 큰손 매수비중 가중치 부여
            score = 0
            is_breakout = code in breakout_codes
            whale_buy_ratio = whale_info.get(code, 50.0) # 정보 없으면 기본 50%
            
            if is_breakout: score += 50  # 신고가 돌파 시 큰 가산점
            if whale_buy_ratio >= 70: score += 30 # 큰손 매수 압도적일 때 가산점
            
            stock.update({
                'name': univ_info['name'],
                'actual_amount': actual_amount,
                'is_breakout': is_breakout,
                'whale_buy_ratio': whale_buy_ratio,
                'score': score
            })
            filtered_targets.append(stock)

        # 3. 최종 점수(Score) 순으로 정렬하여 가장 유망한 종목 상단 배치
        filtered_targets.sort(key=lambda x: x['score'], reverse=True)

        return filtered_targets[:top_n]


# --- 실행 테스트 ---
if __name__ == "__main__":
    screener = DynamicScreener()
    targets = screener.run_screener(top_n=20)
    # print(targets)
    print(f"조회된 종목 수: {len(targets)}")
    print("-" * 105)
    header = (f"{'종목코드':<10}{'종목명':<16}{'현재가':>10}{'대비':>8}{'대비율':>8}"
            f"{'거래량':>12}{'신고가':>10}{'매수비율':>10}{'점수':>10}")
    print(header)
    print("-" * 105)
    for target in targets:
        line = (f"{target['code']:<10}  "
                f"{target['name']:<16}  "
                f"{target['price']:>10,}  "
                f"{target['diff']:>8,}  "
                f"{target['diff_rate']:>8.2f}%  "
                f"{target['volume']:>12,}  "
                f"{target['is_breakout']:>10}  "
                f"{round(target['whale_buy_ratio'], 2):>10}  "
                f"{target['score']:>10}")
        print(line)
    print("-" * 105)    


