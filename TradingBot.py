import path_finder
import time
import sys
import os
from datetime import datetime
from Screener.UniverseBuilder import UniverseBuilder
from Screener.DynamicScreener import DynamicScreener
from Signal.RealtimeManager import RealtimeManager
from API.CpCybos import CpCybos # 연결 체크용

class TradingBot:
    def __init__(self):
        print(f"==================================================")
        print(f"   🚀 Scalping Bot v1.0 시스템 가동 시작")
        print(f"   일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"==================================================")
        
        # 0. 서버 연결 상태 확인
        self.cybos = CpCybos()
        if not self.cybos.check_connection():
            print("❌ CYBOS 연결 실패. 프로그램을 종료합니다.")
            sys.exit()

        # 각 엔진 초기화
        self.builder = UniverseBuilder()
        self.screener = DynamicScreener()
        self.manager = None # 실시간 감시는 타깃이 정해진 후 시작

    def run(self):
        """봇 메인 루프 실행"""
        
        # 1단계: 청정 유니버스 구축 (장 시작 전 1회 실행)
        # 이미 파일이 있다면 load_universe()를 통해 가져옵니다.
        print("\n[Step 1] 종목 유니버스 확인 중...")
        universe = self.builder.load_universe()
        if not universe:
            print("❌ 유니버스 구축 실패.")
            return

        # 2단계: 주도주 스크리닝 (감시할 TOP 20 추출)
        print("\n[Step 2] 현재 시장 주도주 포착 중...")
        targets = self.screener.run_screener(top_n=20)
        if not targets:
            print("⚠️ 포착된 주도주가 없습니다. 잠시 후 다시 시도하세요.")
            return

        # 3단계: 실시간 감시 엔진 가동
        # RealtimeManager는 내부에서 무한 루프(PumpWaitingMessages)를 돌립니다.
        print("\n[Step 3] 실시간 타점 감시 엔진 진입")
        self.manager = RealtimeManager(targets)
        
        try:
            self.manager.start_monitoring()
        except Exception as e:
            print(f"❌ 가동 중 오류 발생: {e}")
        finally:
            self.stop()

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