import path_finder
import time
import sys
import os
import pythoncom
from datetime import datetime
from Screener.UniverseBuilder import UniverseBuilder
from Screener.DynamicScreener import DynamicScreener
from Signal.RealtimeManager import RealtimeManager
from API.CpAPI import CreonAPI

class TradingBot:
    def __init__(self):
        print(f"==================================================")
        print(f"   🚀 Scalping Bot v1.0 시스템 가동 시작")
        print(f"   일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"==================================================")
        
        # 0. 서버 연결 상태 확인
        self.api = CreonAPI()
        
        # 2. 연결 상태 확인 (obj_cybos.IsConnect 사용)
        # IsConnect가 1이면 정상, 0이면 연결 끊김입니다.
        if self.api.obj_cybos.IsConnect == 0:
            print("❌ CYBOS 연결 실패. 프로그램을 종료합니다.")
            sys.exit()

        # 각 엔진 초기화
        self.builder = UniverseBuilder()
        self.screener = DynamicScreener()
        self.manager = None # 실시간 감시는 타깃이 정해진 후 시작
        
    def get_refresh_interval(self):
        """현재 시간에 따라 최적의 종목 갱신 주기를 반환합니다."""
        now = datetime.now()
        curr_time = now.hour * 100 + now.minute # 예: 9시 5분 -> 905

        if 900 <= curr_time < 930:
            return 120  # 2분 (장 초반 초정밀 감시)
        elif 930 <= curr_time < 1030:
            return 300  # 5분 (추세 확인)
        elif 1430 <= curr_time < 1530:
            return 300  # 5분 (장 마감 수급)
        else:
            return 600  # 10분 (횡보 구간)
        
    def run(self):
        """봇 메인 루프 실행"""
        
        # 1단계: 청정 유니버스 구축 (장 시작 전 1회 실행)
        # 이미 파일이 있다면 load_universe()를 통해 가져옵니다.
        print("\n[Step 1] 종목 유니버스 확인 중...")
        universe = self.builder.load_universe()
        if not universe:
            print("❌ 유니버스 구축 실패.")
            return

        # 무한 루프: 주기적으로 종목을 갱신하며 무한 감시
        while True:
            try:
                # 갱신 주기 결정
                interval = self.get_refresh_interval()
                # [Step 2] 현재 시점의 주도주 20개 포착
                print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🔍 주도주 탐색 및 리스트 갱신 시작...")
                targets = self.screener.run_screener(top_n=20)
                
                if not targets:
                    print("⚠️ 조건에 맞는 종목이 없습니다. 30초 후 재시도합니다.")
                    time.sleep(30)
                    continue

                # [Step 3] 실시간 감시 시작
                # 기존에 감시 중이던 종목이 있다면 구독 해제(Unsubscribe)
                if self.manager:
                    self.manager.stop_monitoring()
                
                # 새 종목으로 매니저 생성
                self.manager = RealtimeManager(targets)
                self.manager.start_subscribing() # 이 함수는 구독만 하고 바로 리턴해야 함

                print(f"✅ 새 타깃 {len(targets)}종목 감시 시작 (갱신 주기: {interval//60}분)")

                # 💡 지정된 시간 동안 이벤트를 수신하며 대기
                self.wait_and_monitor(interval)

            except KeyboardInterrupt:
                self.stop()
                break
            except Exception as e:
                print(f"❌ 운영 중 예외 발생: {e}")
                time.sleep(10)
    
    def wait_and_monitor(self, duration):
        """정해진 시간 동안 윈도우 메시지를 펌핑하며 실시간 데이터를 받습니다."""
        start_time = time.time()
        while time.time() - start_time < duration:
            # 실시간 이벤트를 처리하기 위한 핵심 함수
            pythoncom.PumpWaitingMessages()
            time.sleep(0.01)
            
    def stop(self):
        """시스템 종료 및 자원 해제"""
        if self.manager:
            self.manager.stop_monitoring()
        print("\n==================================================")
        print(f"   👋 시스템 종료 완료 (종료시간: {datetime.now().strftime('%H:%M:%S')})")
        print(f"==================================================")

if __name__ == "__main__":
    bot = TradingBot()
    bot.run()