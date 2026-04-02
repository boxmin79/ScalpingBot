import win32com.client
import json
import os
import time
import sys
from datetime import datetime

class StandaloneTesterPro:
    def __init__(self):
        # 1. API 오브젝트 초기화
        self.obj_chart = win32com.client.Dispatch("CpSysDib.StockChart")
        self.obj_cybos = win32com.client.Dispatch("CpUtil.CpCybos")
        
        # 2. 유니버스 파일 경로
        self.universe_path = r"C:\Users\realb\Documents\TradingBot\data\code\scalping_universe.json"
        
    def check_request_limit(self):
        """대신증권 TR 요청 제한(15초당 60회)을 체크하고 대기합니다."""
        # type 0: 시세 조회용 TR 제한
        remain_count = self.obj_cybos.GetLimitRemainCount(0) 
        
        if remain_count <= 0:
            # 밀리초 단위를 초 단위로 변환 후 0.5초 여유 대기
            sleep_time = (self.obj_cybos.LimitRequestRemainTime / 1000) + 1
            print(f"⏳ 요청 제한 도달. {sleep_time:.1f}초 대기 후 재개합니다...")
            time.sleep(max(sleep_time, 0.1))

    def fetch_chart_data(self, code, req_type, count, chart_kind):
        """요청 제한을 체크하며 데이터를 수집합니다."""
        self.check_request_limit() # 🎯 요청 전 제한 확인
        
        # 필드 설정 (0:날짜, 1:시간, 2~5:OHLC, 8:거래량, 9:거래대금, 62:매도누적, 63:매수누적)
        # 필드는 오름차순으로 정렬되어 수신됩니다.
        fields = [0, 1, 2, 3, 4, 5, 8, 9, 62, 63] 
        
        self.obj_chart.SetInputValue(0, code)
        self.obj_chart.SetInputValue(1, ord(req_type))    # '2': 개수 조회
        self.obj_chart.SetInputValue(4, count)           # 요청 개수
        self.obj_chart.SetInputValue(5, fields)          # 필드 배열
        self.obj_chart.SetInputValue(6, ord(chart_kind)) # 'D': 일봉, 'm': 분봉
        self.obj_chart.SetInputValue(9, ord('1'))        # 수정주가 사용
        
        ret = self.obj_chart.BlockRequest()
        # 🎯 [추가] 매 요청이 끝난 후 강제로 0.2초의 휴식 시간을 가집니다.
        # 이렇게 하면 15초 동안 최대 60회를 꽉 채우지 않고 약 45~50회 정도로 조절되어 훨씬 안정적입니다.
        time.sleep(0.2)
        
        if ret != 0: return None
        
        recv_count = self.obj_chart.GetHeaderValue(3) 
        results = []
        for i in range(recv_count):
            # 필드 인덱스: 0(날짜), 1(시간), 2(시), 3(고), 4(저), 5(종), 6(거래량), 7(대금), 8(62), 9(63)
            item = {
                'time': self.obj_chart.GetDataValue(1, i),
                'open': self.obj_chart.GetDataValue(2, i),
                'close': self.obj_chart.GetDataValue(5, i),
                'vol': self.obj_chart.GetDataValue(6, i),
                'sell_vol_cum': self.obj_chart.GetDataValue(8, i), # 62번 필드
                'buy_vol_cum': self.obj_chart.GetDataValue(9, i)   # 63번 필드
            }
            results.append(item)
        return results

    def run_test(self):
        # 유니버스 파일 로드
        if not os.path.exists(self.universe_path):
            print("❌ 유니버스 파일을 찾을 수 없습니다.")
            return
            
        with open(self.universe_path, 'r', encoding='utf-8') as f:
            universe = json.load(f)
            
        print(f"🚀 {len(universe)}종목 정밀 백테스트 시작 (TR 제한 관리 가동)")
        print("-" * 85)

        for item in universe:
            code, name = item['code'], item['name']
            
            # 1. 일봉 평균 거래량 계산 (최근 60일)
            daily = self.fetch_chart_data(code, '2', 61, 'D')
            if not daily or len(daily) < 61: continue
            avg_vol_1m = (sum(d['vol'] for d in daily[1:]) / 60) / 390
            
            # 2. 분봉 데이터 수동 분석 (오늘치)
            minutes = self.fetch_chart_data(code, '2', 380, 'm')
            if not minutes: continue
            minutes.reverse() # 시간 순 정렬
            
            for i, candle in enumerate(minutes):
                # 🎯 필터 1: 거래량 15배 폭발 여부
                vol_multiple = candle['vol'] / avg_vol_1m if avg_vol_1m > 0 else 0
                
                if vol_multiple >= 15.0:
                    # 🎯 필터 2: 분봉 내 순수 체결강도 계산 (현재 누적 - 이전 누적)
                    if i > 0:
                        b_vol = candle['buy_vol_cum'] - minutes[i-1]['buy_vol_cum']
                        s_vol = candle['sell_vol_cum'] - minutes[i-1]['sell_vol_cum']
                    else:
                        b_vol, s_vol = candle['buy_vol_cum'], candle['sell_vol_cum']
                        
                    intensity = (b_vol / s_vol * 100) if s_vol > 0 else 100
                    
                    # 🎯 필터 3: 양봉 + 체결강도 110% 이상
                    if candle['close'] > candle['open'] and intensity >= 110:
                        # 수익률 검증 (10분 후 매도)
                        target_idx = min(i + 10, len(minutes) - 1)
                        profit = ((minutes[target_idx]['close'] - candle['close']) / candle['close'] * 100) - 0.23
                        
                        status = "✅ 성공" if profit > 0 else "❌ 실패"
                        print(f"[{name}] {candle['time']} | {vol_multiple:.1f}배 | 강도:{intensity:.1f}% | 수익:{profit:.2f}% {status}")
            
        print("-" * 85)
        print("✅ 모든 종목 테스트 완료")

if __name__ == "__main__":
    StandaloneTesterPro().run_test()