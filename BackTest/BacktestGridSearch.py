import pandas as pd
import numpy as np
import json
from pathlib import Path
from tqdm import tqdm
import path_finder
from Util.FileManager import FileManager

class BacktestGridSearch:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        self.fm = FileManager()
        
        # 경로 설정
        self.universe_path = Path(r"C:\Users\realb\Documents\TradingBot\BackTest\backtest_base_universe.json")
        self.daily_dir = self.cfg.CHART_DIR / 'daily'
        self.minute_dir = self.cfg.CHART_DIR / 'minute'
        self.result_dir = self.cfg.BACKTEST_DIR
        self.result_dir.mkdir(parents=True, exist_ok=True)

        # 그리드 서치 파라미터
        self.avg_days_list = [20, 60]
        self.multiples = [10, 20, 30, 40, 50]
        self.strengths = [100, 110, 120, 130]

    def run(self):
        universe = self.fm.load(self.universe_path)
        if not universe:
            print("❌ 유니버스 파일을 찾을 수 없습니다.")
            return

        summary_results = []

        print(f"🚀 그리드 서치 백테스트 시작 (대상: {len(universe)}종목)")
        
        for item in tqdm(universe, desc="백테스트 진행 중"):
            code = item['code']
            name = item['name']
            
            # 1. 데이터 로드 (Parquet)
            daily_file = self.daily_dir / f"{code}.parquet"
            minute_file = self.minute_dir / f"{code}.parquet"
            
            if not daily_file.exists() or not minute_file.exists():
                continue

            try:
                df_daily = pd.read_parquet(daily_file, engine='fastparquet')
                df_min = pd.read_parquet(minute_file, engine='fastparquet')
                
                # 2. 체결강도 계산 (cum_buy / cum_sell)
                # 분모가 0인 경우를 대비해 1e-9(epsilon) 추가 혹은 replace 처리
                df_min['strength'] = (df_min['cum_buy_vol'] / df_min['cum_sell_vol'].replace(0, np.nan)) * 100
                df_min['strength'] = df_min['strength'].fillna(100) # 데이터가 없으면 100%로 간주
            
                # 데이터 전처리: 시간순 정렬
                df_daily = df_daily.sort_values(by='date')
                df_min = df_min.sort_values(by=['date', 'time'])
                
                # 2. 그리드 서치 실행
                stock_perf = self._grid_search_stock(code, df_daily, df_min)
                
                if stock_perf:
                    # 종목별 상세 결과 저장
                    self.fm.save(stock_perf, self.result_dir / f"{code}.json")
                    
                    # 요약 데이터 추가 (파라미터별 평균 수익률 등)
                    for res in stock_perf:
                        res['name'] = name
                        summary_results.append(res)
                        
            except Exception as e:
                print(f"Error processing {code}: {e}")
                continue

        # 3. 전체 요약 파일 생성
        self._save_summary(summary_results)

    def _grid_search_stock(self, code, df_daily, df_min):
        results = []
        
        # 일봉 기준 평균 분당 거래량 미리 계산
        # 한국 시장 정규장 시간: 09:00 ~ 15:30 (390분, 동시호가 제외 시 약 380분)
        MINUTES_PER_DAY = 380 
        
        for avg_day in self.avg_days_list:
            # 20일/60일 이동평균 거래량 산출
            df_daily[f'avg_vol_{avg_day}'] = df_daily['vol'].rolling(window=avg_day).mean()
            
            for mul in self.multiples:
                for strength in self.strengths:
                    # 해당 파라미터 조합에 대한 백테스트 로직
                    perf = self._backtest_logic(df_daily, df_min, avg_day, mul, strength, MINUTES_PER_DAY)
                    if perf:
                        results.append(perf)
        return results

    def _backtest_logic(self, df_daily, df_min, avg_day, mul, strength_threshold, minutes_per_day):
        # 일봉 날짜와 매칭하여 기준 거래량 병합
        df_target = pd.merge(df_min, df_daily[['date', f'avg_vol_{avg_day}']], on='date', how='left')
        
        # 임계치 설정: (평균 일일 거래량 / 장중 분) * 멀티플
        df_target['vol_threshold'] = (df_target[f'avg_vol_{avg_day}'] / minutes_per_day) * mul
        
        # 시그널 포착 (거래량 돌파 & 체결강도 조건)
        # ※ 주의: 분봉 데이터에 'strength'(체결강도) 컬럼이 포함되어 있어야 합니다.
        if 'strength' not in df_target.columns:
            # 체결강도 데이터가 없는 경우를 대비한 가상 로직 (필요시 수정)
            df_target['strength'] = 100 

        signals = df_target[
            (df_target['vol'] > df_target['vol_threshold']) & 
            (df_target['strength'] >= strength_threshold)
        ].copy()

        if signals.empty:
            return None

        # 성과 측정 (간단 예시: 시그널 발생 10분 후 수익률 평균)
        # 실제 백테스트 시에는 익절/손절 로직을 추가해야 합니다.
        returns = []
        for idx in signals.index:
            entry_price = df_target.loc[idx, 'close']
            # 10분 후 데이터 확인 (index + 10)
            exit_idx = idx + 10
            if exit_idx < len(df_target):
                exit_price = df_target.loc[exit_idx, 'close']
                returns.append((exit_price / entry_price - 1) * 100)

        if not returns:
            return None

        return {
            "avg_day": avg_day,
            "multiple": mul,
            "strength_threshold": strength_threshold,
            "signal_count": len(signals),
            "avg_return": round(np.mean(returns), 4),
            "win_rate": round(len([r for r in returns if r > 0]) / len(returns) * 100, 2)
        }

    def _save_summary(self, summary_results):
        """파라미터 조합별로 전체 종목의 성과를 평균내어 요약 리포트 저장"""
        df_summary = pd.DataFrame(summary_results)
        final_report = df_summary.groupby(['avg_day', 'multiple', 'strength_threshold']).agg({
            'avg_return': 'mean',
            'win_rate': 'mean',
            'signal_count': 'sum'
        }).reset_index()
        
        report_path = self.result_dir / "grid_search_summary.csv"
        final_report.to_csv(report_path, index=False, encoding='utf-8-sig')
        print(f"✅ 요약 리포트 저장 완료: {report_path}")

if __name__ == "__main__":
    tester = BacktestGridSearch()
    tester.run()