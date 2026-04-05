import path_finder
import json
import pandas as pd
import time
from API.MarketDataManager import MarketDataManager
from API.MarketEye import MarketEye

class UniverseCodeAnalyzer:
    def __init__(self, json_path):
        self.mdm = MarketDataManager()
        self.me = MarketEye()
        self.universe_path = json_path

    def get_codes_from_json(self):
        """JSON 파일에서 종목코드 리스트만 추출"""
        with open(self.universe_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return [item['code'] for item in data]

    def get_fresh_metrics(self, codes):
        """
        MarketEye 필드 정보를 활용해 최신 지표 수신
        4:현재가, 20:상장주식수, 75:부채비율, 77:ROE, 17:종목명
        """
        field_ids = [0, 17, 4, 20, 75, 77]
        results, _ = self.me.get_market_data(codes, field_ids)
        
        metrics = {}
        for res in results:
            code = res[0]
            # 시가총액 계산 (단위: 억)
            # 20억주 이상 종목 예외처리는 생략(단순계산)
            m_cap = (res[4] * res[20]) / 100000000 
            
            metrics[code] = {
                'name': res[17],
                'm_cap': round(m_cap, 2),
                'debt_ratio': res[75],
                'roe': res[77]
            }
        return metrics

    def analyze(self, multiplier=20):
        codes = self.get_codes_from_json()
        # API 부하 방지를 위해 50종목씩 끊어서 지표 수신
        fresh_info = self.get_fresh_metrics(codes)
        
        final_stats = []
        
        for code in codes:
            if code not in fresh_info: continue
            print(f"📊 {info['name']}({code}) 분석 중... (시총: {info['m_cap']}억)")
            info = fresh_info[code]

            # 1. 60일 평균 거래대금 계산
            daily = self.mdm.get_chart_data(code, req_type='2', target_count=60, chart_type='D')
            if not daily: continue
            avg_min_amt = (sum(d['amt'] for d in daily) / 60) / 380

            # 2. 1분봉 데이터 시뮬레이션 (최근 1일치)
            minutes = self.mdm.get_chart_data(code, req_type='2', target_count=380, chart_type='m')
            minutes.reverse()

            signals = 0
            wins = 0
            for i, candle in enumerate(minutes):
                if candle['amt'] >= avg_min_amt * multiplier:
                    signals += 1
                    entry_p = candle['close']
                    # 10분 이내 1.5% 익절 목표
                    for j in range(i+1, min(i+11, len(minutes))):
                        if (minutes[j]['high'] - entry_p) / entry_p >= 0.015:
                            wins += 1
                            break
            
            win_rate = (wins / signals * 100) if signals > 0 else 0
            
            final_stats.append({
                'code': code,
                'name': info['name'],
                'm_cap': info['m_cap'],
                'debt_ratio': info['debt_ratio'],
                'roe': info['roe'],
                'win_rate': round(win_rate, 2),
                'signals': signals
            })
            time.sleep(0.05) # 속도 조절

        return pd.DataFrame(final_stats)

# 실행
if __name__ == "__main__":
    path = r"C:\Users\realb\Documents\TradingBot\data\code\scalping_universe.json"
    analyzer = UniverseCodeAnalyzer(path)
    result_df = analyzer.analyze(multiplier=20)
    
    print("\n[분석 결과 상위 10종목]")
    print(result_df.sort_values(by='win_rate', ascending=False).head(10))