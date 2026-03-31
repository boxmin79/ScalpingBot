import path_finder
import time
import sys
import pythoncom
from datetime import datetime
from Screener.UniverseBuilder import UniverseBuilder
from Screener.DynamicScreener import DynamicScreener
from Signal.RealtimeManager import RealtimeManager
from API.CpAPI import CreonAPI
from API.AccountManager import AccountManager # 🎯 AccountManager 임포트
from Util.AsyncLogger import AsyncLogger

class TradingBot:
    def __init__(self):
        print(f"==================================================")
        print(f"   🚀 Scalping Bot v1.0 시스템 가동 시작")
        print(f"   일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"==================================================")
        
        # 0. CreonAPI 통합 객체 생성
        self.api = CreonAPI()
        self.logger = AsyncLogger()
        
        
        # 1. 서버 연결 상태 확인
        if self.api.obj_cybos.IsConnect == 0:
            self.logger.info("❌ CYBOS 연결 실패. 프로그램을 종료합니다.")
            self.logger.stop()
            sys.exit()

        # 2. 주문 서비스 초기화 (TradeInit)
        init_status = self.api.obj_trade_util.TradeInit(0)
        if init_status != 0:
            self.logger.info(f"❌ 주문 서비스 초기화 실패 (에러코드: {init_status})")
            self.logger.stop()
            sys.exit()

        # 3. 계좌 정보 자동 추출
        # CpTdUtil을 통해 접속된 첫 번째 계좌와 주식 상품 코드를 가져옵니다.
        self.acc_no = self.api.obj_trade_util.AccountNumber[0]
        self.acc_flag = self.api.obj_trade_util.GoodsList(self.acc_no, 1)[0]
        print(f"✅ 주문 준비 완료: 계좌({self.acc_no}), 상품코드({self.acc_flag})")
        
        # 4. AccountManager 초기화 및 자산 확인
        self.account = AccountManager(self.acc_no, self.acc_flag)
        
        # 🎯 메서드로 분리된 실잔고 동기화 호출
        self.trade_budget = 0  # 🎯 초기값 선언
        self.update_budget() # 초기 예산 설정
        self.initial_positions = self.sync_account_positions()
            
        # 엔진 초기화
        self.builder = UniverseBuilder()
        self.screener = DynamicScreener()
        self.manager = None 
        
    def update_budget(self):
        """D+2 예수금을 확인하여 1회 매매 한도(1/10)를 갱신합니다."""
        deposit_data = self.account.get_expected_deposit()
        # print(deposit_data)
        if deposit_data and deposit_data['d2_deposit'] > 0:
            # 예산을 1/10로 나누고 수수료를 고려하여 9.8% 수준으로 안전하게 책정  
            self.trade_budget = int(deposit_data['d2_deposit'] * 0.098)
            return True
        else:
            self.logger.info("⚠️ 매매 가능 예수금이 부족하거나 정보를 가져올 수 없습니다.")
            return False
        
    def sync_account_positions(self):
        """
        [교체] 현재 계좌의 실제 잔고를 긁어와서 봇 포맷으로 변환합니다.
        """
        print(f"📦 실잔고 동기화 중...")
        
        # AccountManager의 get_balance_data를 사용하여 실잔고 획득
        summary, stocks = self.account.get_balance_data()
        
        positions = {}
        for s in stocks:
            # AccountManager가 반환한 데이터 필드를 봇 포맷에 매핑
            positions[s['code']] = {
                'name': s['name'],
                'buy_price': s['buy_price'],
                'qty': s['total_qty'],
                'entry_time': time.time()  # 동기화 시점 기준으로 설정
            }
        
        if positions:
            for code, pos in positions.items():
                print(f"   > 발견된 보유 종목: {pos['name']}({code}) | {pos['qty']}주")
        else:
            print("   > 현재 보유 중인 종목이 없습니다.")
            
        return positions
        
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
            self.logger.info("❌ 유니버스 구축 실패.")
            return

        while True:
            try:
                # 🎯 [추가] 장 마감 시각 체크 (15:20)
                now = datetime.now()
                exit_time = now.replace(hour=15, minute=20, second=0, microsecond=0)
                
                if now >= exit_time:
                    self.logger.info(f"\n[{now.strftime('%H:%M:%S')}] 📢 장 마감 시간이 되어 포지션을 정리합니다.")
                    if self.manager:
                        # RealtimeManager 내부에 구현한 전량 매도 함수 호출
                        self.manager.force_exit_all()
                        # 모든 주문이 나갈 시간을 잠시 준 뒤 종료
                        self.wait_and_monitor(5)
                    self.stop()
                    break
                
                # [Step 2] 타깃 종목 스캐닝 전 예산 최신화
                if not self.update_budget():
                    self.logger.info("⚠️ 예수금 확인 불가. 이전 예산을 유지하거나 스킵합니다.")
                
                self.logger.info(f"\n[{now.strftime('%H:%M:%S')}] 🔍 스캐닝 (한도: {self.trade_budget:,}원)")
                targets = self.screener.run_screener()
                
                # 통신 제한 횟수 체크
                remain_count = self.api.obj_cybos.GetLimitRemainCount(1)
                if remain_count < 5:
                    wait_time = self.api.obj_cybos.LimitRequestRemainTime
                    time.sleep(wait_time / 1000)

                interval = self.get_refresh_interval()
                
                # [Step 2] 주도주 포착
                print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🔍 타깃 종목 스캐닝...")
                targets = self.screener.run_screener()
                
                if not targets:
                    # 🚨 time.sleep(30) 삭제
                    # 🎯 [수정] 보유 종목 매도를 위해 메시지 펌프는 돌아가야 함
                    self.wait_and_monitor(30) 
                    continue

                # [Step 3] 실시간 감시 엔진 교체
                if self.manager:
                    # 기존 포지션 정보를 유지하면서 감시 대상만 업데이트하려면
                    # stop_monitoring 대신 내부 targets만 교체하는 로직이 필요할 수 있습니다.
                    # 여기서는 단순 교체 방식을 유지하되 포지션을 인자로 넘겨줄 수 있습니다.
                    prev_positions = self.manager.positions
                    self.manager.stop_monitoring()
                else:
                    # 🎯 [수정] 첫 실행 시에는 초기화 때 긁어온 실잔고를 사용
                    prev_positions = self.initial_positions
                
                # 🎯 매니저를 생성할 때 self.trade_budget을 인자로 넘겨줍니다.
                self.manager = RealtimeManager(targets, self.acc_no, self.acc_flag, self.trade_budget, self.logger)
                
                # 🎯 중요: 새로 생성된 매니저에 기존 포지션 정보를 복사해줘야 매도가 가능함
                self.manager.positions = prev_positions
                
                # 체결 알림 콜백 연결
                self.manager.om.set_callback(self.manager.on_order_confirmed)
                
                self.manager.start_subscribing() 
        
                self.wait_and_monitor(interval)

            except KeyboardInterrupt:
                self.stop()
                break
            except Exception as e:
                self.logger.info(f"❌ 시스템 예외 발생: {e}")
                time.sleep(10)
    
    def wait_and_monitor(self, duration):
        start_time = time.time()
        while time.time() - start_time < duration:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.01)
            
            
    def stop(self):
        if self.manager:
            self.manager.stop_monitoring()
        self.logger.info(f"\n🛑 시스템 종료 (시간: {datetime.now().strftime('%H:%M:%S')})")
        # 메시지가 전송될 때까지 잠시 대기 (테스트용)
        time.sleep(2)
        self.logger.stop()
        sys.exit()


if __name__ == "__main__":
    bot = TradingBot()
    bot.run()