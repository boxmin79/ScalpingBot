import sys
import os
import json # 🎯 JSON 로드를 위해 추가
# 프로젝트 루트 경로 추가 (API, Screener 등 임포트 가능하게 설정)
sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))

import time
from datetime import datetime
from API.MarketDataManager import MarketDataManager
from API.MarketScanner import MarketScanner
from API.CpAPI import CreonAPI
from API.MarketEye import MarketEye

class StrategyTester:
    def __init__(self):
        self.api = CreonAPI()
        self.mdm = MarketDataManager()
        self.scanner = MarketScanner()
        self.market_eye = MarketEye()
        # 🎯 유니버스 파일 경로 설정
        self.universe_path = r"C:\Users\realb\Documents\TradingBot\data\code\scalping_universe.json"
        self.test_results = []

    def load_universe_from_file(self):
        """저장된 scalping_universe.json 파일에서 종목 코드 리스트를 로드합니다."""
        print(f"📂 [파일 로드] 유니버스 파일 읽기 시작: {self.universe_path}")
        
        if not os.path.exists(self.universe_path):
            print(f"❌ 파일을 찾을 수 없습니다: {self.universe_path}")
            return []

        try:
            with open(self.universe_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 리스트 내 딕셔너리에서 'code' 필드만 추출
            codes = [item['code'] for item in data]
            print(f"✅ 유니버스 파일 로드 완료: {len(codes)}종목")
            return codes
        except Exception as e:
            print(f"❌ 파일 읽기 중 오류 발생: {e}")
            return []

    def get_test_targets(self):
        """오늘 상승한 종목과 파일에서 읽은 유니버스를 비교하여 최종 테스트 대상 선정"""
        # 1. 마켓스캐너로 오늘 강세(거래대금 상위 또는 상승률 5%~15%) 종목 추출
        rising_stocks = self.scanner.get_intraday_strength_stocks()
        rising_codes = {s['code']: s['name'] for s in rising_stocks}
        
        # 2. 🎯 재무 전수조사 대신 저장된 파일에서 우량 유니버스 획득
        financial_universe = self.load_universe_from_file()
        
        # 3. 교집합 추출 (재무 우량주 중 오늘 실제로 변동성이 생긴 종목)
        test_targets = {code: name for code, name in rising_codes.items() if code in financial_universe}
        
        print(f"🎯 [필터링 완료] 최종 테스트 대상: {len(test_targets)}종목 (우량 유니버스 내 오늘 강세주)")
        return test_targets

    def run_backtest(self):
        """파일 기반 유니버스를 대상으로 15배 거래량 로직 검증"""
        targets = self.get_test_targets()
        
        if not targets:
            print("⚠️ 테스트할 대상 종목이 없습니다. 시장 상황을 확인하거나 유니버스 파일을 확인하세요.")
            return

        print(f"\n🚀 [백테스트] 1분봉 기반 15배 거래량 폭발 로직 검증 시작...")
        print("-" * 75)

        for code, name in targets.items():
            # 1. 일봉 데이터 수집: 최근 60일 평균 거래량(1분 단위 환산) 계산
            daily = self.mdm.get_chart_data(code, req_type='2', target_count=61, chart_type='D')
            if len(daily) < 60: continue
            
            # 오늘(index 0)을 제외한 최근 60일 평균 거래량 / 390분
            avg_vol_1m = (sum(d['vol'] for d in daily[1:]) / 60) / 390
            
            # 2. 분봉 데이터 수집: 오늘 1분봉 전체(약 380개) 분석
            minutes = self.mdm.get_chart_data(code, req_type='2', target_count=380, chart_type='m', cycle=1)
            if not minutes: continue
            minutes.reverse() # 과거 -> 현재 순 정렬

            for i, candle in enumerate(minutes):
                # 3. 전략 조건 검증: 1분 거래량 > (평균 * 15배) & 양봉
                if candle['vol'] > (avg_vol_1m * 15) and candle['close'] > candle['open']:
                    
                    # 4. 성과 추적: 신호 발생 10분 후의 종가 수익률 확인
                    post_idx = min(i + 10, len(minutes) - 1)
                    entry_p = candle['close']
                    exit_p = minutes[post_idx]['close']
                    
                    # 제비용(수수료+세금 약 0.23%) 반영 수익률 계산
                    profit = round(((exit_p - entry_p) / entry_p * 100) - 0.23, 2)
                    
                    status = "✅ 성공" if profit > 0 else "❌ 실패"
                    print(f"[{name}] {candle['time']} | {round(candle['vol']/avg_vol_1m, 1)}배 폭발 | 10분수익: {profit}% {status}")

            time.sleep(0.05) # API 요청 제한 방지용 짧은 휴식

if __name__ == "__main__":
    tester = StrategyTester()
    tester.run_backtest()