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
from API.AccountManager import AccountManager # 🎯 AccountManager 임포트

class TradingBot:
    def __init__(self):
        print(f"==================================================")
        print(f"   🚀 Scalping Bot v1.0 시스템 가동 시작")
        print(f"   일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"==================================================")
        
        # 0. CreonAPI 통합 객체 생성
        self.api = CreonAPI()
        
        # 1. 서버 연결 상태 확인
        if self.api.obj_cybos.IsConnect == 0:
            print("❌ CYBOS 연결 실패. 프로그램을 종료합니다.")
            sys.exit()

        # 2. 주문 서비스 초기화 (TradeInit)
        init_status = self.api.obj_trade_util.TradeInit(0)
        if init_status != 0:
            print(f"❌ 주문 서비스 초기화 실패 (에러코드: {init_status})")
            sys.exit()

        # 3. 계좌 정보 자동 추출
        # CpTdUtil을 통해 접속된 첫 번째 계좌와 주식 상품 코드를 가져옵니다.
        self.acc_no = self.api.obj_trade_util.AccountNumber[0]
        self.acc_flag = self.api.obj_trade_util.GoodsList(self.acc_no, 1)[0]
        
        # 4. AccountManager 초기화 및 자산 확인
        # 추출한 계좌 정보를 AccountManager에 주입합니다.
        self.account = AccountManager(self.acc_no, self.acc_flag)
        
        print(f"✅ API 연결 및 계좌 설정 완료")
        print(f"   - 접속 계좌: {self.acc_no} (상품코드: {self.acc_flag})")
        
        # [자동화 기능] 가동 직후 현재 예수금 및 D+2 결제 예정금액 확인
        deposit_data = self.account.get_expected_deposit()
        if deposit_data:
            print(f"   - 현재 예수금: {deposit_data['current_deposit']:,}원")
            print(f"   - D+2 예정예수금: {deposit_data['d2_deposit']:,}원 (실질 매매 가능금)")

        # 엔진 초기화
        self.builder = UniverseBuilder()
        self.screener = DynamicScreener()
        self.manager = None 
        
    def get_refresh_interval(self):
        """현재 시간에 따라 최적의 종목 갱신 주기를 반환합니다."""
        now = datetime.now()
        curr_time = now.hour * 100 + now.minute 

        if 900 <= curr_time < 930:
            return 120  # 2분
        elif 930 <= curr_time < 1030:
            return 300  # 5분
        else:
            return 600  # 10분
        
    def run(self):
        """봇 메인 루프 실행"""
        universe = self.builder.load_universe()
        if not universe:
            print("❌ 유니버스 구축 실패.")
            return

        while True:
            try:
                # 통신 제한 횟수 체크
                remain_count = self.api.obj_cybos.GetLimitRemainCount(1)
                if remain_count < 5:
                    wait_time = self.api.obj_cybos.LimitRequestRemainTime
                    time.sleep(wait_time / 1000)

                interval = self.get_refresh_interval()
                
                # [Step 2] 주도주 포착
                print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🔍 타깃 종목 스캐닝...")
                targets = self.screener.run_screener(top_n=20)
                
                if not targets:
                    time.sleep(30)
                    continue

                # [Step 3] 실시간 감시 엔진 교체
                if self.manager:
                    self.manager.stop_monitoring()
                
                # RealtimeManager에 식별된 계좌 정보를 전달하여 매매 준비
                self.manager = RealtimeManager(targets, self.acc_no, self.acc_flag)
                self.manager.start_subscribing() 

                self.wait_and_monitor(interval)

            except KeyboardInterrupt:
                self.stop()
                break
            except Exception as e:
                print(f"❌ 시스템 예외 발생: {e}")
                time.sleep(10)
    
    def wait_and_monitor(self, duration):
        start_time = time.time()
        while time.time() - start_time < duration:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.01)
            
    def stop(self):
        if self.manager:
            self.manager.stop_monitoring()
        print(f"\n🛑 시스템 종료 (시간: {datetime.now().strftime('%H:%M:%S')})")

if __name__ == "__main__":
    bot = TradingBot()
    bot.run()