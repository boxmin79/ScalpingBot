import path_finder
import pandas as pd
from datetime import datetime, timedelta
from tqdm import tqdm
# 기존 프로젝트 구조 임포트
from API.CpAPI import CreonAPI
from API.MarketDataManager import MarketDataManager
from Util.FileManager import FileManager

class ChartCollector:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        self.api = CreonAPI()
        self.mdm = MarketDataManager()
        self.fm = FileManager()
        
        # 저장 경로 설정 및 생성
        self.daily_base_path = self.cfg.CHART_DIR / 'daily'
        self.minute_base_path = self.cfg.CHART_DIR / 'minute'
        
        self.daily_base_path.mkdir(parents=True, exist_ok=True)
        self.minute_base_path.mkdir(parents=True, exist_ok=True)

    def collect_all(self, universe_list):
        """유니버스 목록을 순회하며 차트 수집 실행"""
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 🚀 차트 데이터 수집 시작 (총 {len(universe_list)}종목)")
        
        for item in tqdm(universe_list, desc="수집 진행률"):
            code = item['code']
            
            # 1. 일봉 수집 (3년치)
            self._collect_daily(code)
            
            # 2. 분봉 수집 (2년치)
            self._collect_minute(code)

    def _collect_daily(self, code):
        """일봉 3년치 수집 및 Parquet 저장"""
        save_path = self.daily_base_path / f"{code}.parquet"
        
        # 3년 전 날짜 계산
        start_date = (datetime.now() - timedelta(days=3*365)).strftime('%Y%m%d')
        end_date = datetime.now().strftime('%Y%m%d')
        
        # 필드: 날짜, 시가, 고가, 저가, 종가, 거래량, 거래대금
        data = self.mdm.get_chart_data(stk_code=code, req_type='1', 
                                      end_date=int(end_date), 
                                      start_date=int(start_date))
        
        if data:
            df = pd.DataFrame(data)
            # FileManager의 경로 관리 기능을 활용하여 저장
            df.to_parquet(save_path, engine='fastparquet', compression='snappy')

    def _collect_minute(self, code):
        """분봉 2년치 수집 (API 제한 회피를 위한 루프 처리)"""
        save_path = self.minute_base_path / f"{code}.parquet"
        
        all_data = []
        target_start_date = int((datetime.now() - timedelta(days=2*365)).strftime('%Y%m%d'))
        last_date = int(datetime.now().strftime('%Y%m%d'))
        
        # 2년치 데이터를 채울 때까지 과거로 거슬러 올라가며 요청
        while True:
            # '2': 개수 기준으로 요청 (최대치인 5000개씩 요청)
            chunk = self.mdm.get_chart_data(stk_code=code, req_type='2',
                                            chart_type='m',                                             
                                          end_date=last_date, 
                                          target_count=5000
                                          )
            
            if not chunk:
                break
                
            all_data.extend(chunk)
            
            # 수집된 가장 오래된 날짜 확인
            oldest_record_date = chunk[-1]['date']
            
            # 목표 기간(2년)에 도달했거나 더 이상 데이터가 없으면 중단
            if oldest_record_date <= target_start_date or len(chunk) < 5000:
                break
                
            # 다음 요청을 위한 날짜 업데이트
            last_date = oldest_record_date

        if all_data:
            df = pd.DataFrame(all_data)
            # 정확한 2년치 커팅 및 중복 제거
            df = df[df['date'] >= target_start_date]
            df.drop_duplicates(subset=['date', 'time'], inplace=True)
            
            # Parquet 저장 (엔진: fastparquet)
            df.to_parquet(save_path, engine='fastparquet', compression='snappy')

# --- 단독 실행 로직 ---
if __name__ == "__main__":
    # 1. 파일 매니저를 통한 유니버스 로드
    fm = FileManager()
    universe = fm.load(r'C:\Users\realb\Documents\TradingBot\BackTest\backtest_base_universe.json')
    
    if universe:
        # 2. 수집기 실행
        collector = ChartCollector()
        collector.collect_all(universe)
    else:
        print("❌ 유니버스 파일을 찾을 수 없거나 데이터가 비어 있습니다.")