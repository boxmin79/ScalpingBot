import path_finder
import time
import sys
import pythoncom
from datetime import datetime
from Screener.UniverseBuilder import UniverseBuilder
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
        self.set_initial_budget() # 초기 1회 실행
        self.initial_positions = self.sync_account_positions()
            
        # 엔진 초기화
        self.builder = UniverseBuilder()
        self.manager = None 
    
    def set_initial_budget(self):
        """장 시작 시 총 자산의 1/10을 매매 한도로 고정합니다."""
        deposit_data = self.account.get_expected_deposit()
        if deposit_data and deposit_data['d2_deposit'] > 0:
            # 전체 자산의 약 1/10(9.8%)로 고정
            self.initial_deposit = deposit_data['d2_deposit']
            self.trade_budget = int(self.initial_deposit * 0.098)
            print(f"💰 예산 확정: 총 자산 {self.initial_deposit:,}원 -> 종목당 {self.trade_budget:,}원 고정")
            return True
        return False
       
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
                self.logger.info(f"   > 발견된 보유 종목: {pos['name']}({code}) | {pos['qty']}주", send_tg=False)
        else:
            self.logger.info("   > 현재 보유 중인 종목이 없습니다.")
            
        return positions
                
    def run(self):
        """봇 메인 루프 실행"""
        
        universe = self.builder.load_universe()
        if not universe:
            self.logger.info("❌ 유니버스 구축 실패.")
            return
        
        # 리스트에서 종목 코드만 추출하여 타깃 설정
        targets = universe
        self.logger.info(f"✅ 유니버스 {len(targets)}종목 로드 완료. 실시간 감시를 시작합니다.")
        
        # 🎯 [수정] 루프 진입 전 초기화 필수
        last_budget_update = time.time()
        last_sync_time = time.time()
        
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
                
                # 🎯 5분(300초)마다 실잔고 강제 동기화
                if time.time() - last_sync_time > 300:
                    current_real_positions = self.sync_account_positions()
                    if self.manager:
                        # 로컬에만 있고 서버에 없는 종목 제거 및 수량 업데이트
                        self.manager.positions = current_real_positions
                        self.logger.info("🔄 [시스템] 실잔고 동기화 완료")
                    last_sync_time = time.time()
                    
                if time.time() - last_budget_update > 300: 
                    if self.update_budget():
                        if self.manager:
                            self.manager.trade_budget = self.trade_budget
                        last_budget_update = time.time()
                    
                if self.manager is None:
                    # 첫 실행 시에만 생성
                    self.manager = RealtimeManager(targets, self.acc_no, self.acc_flag, self.trade_budget, self.logger)
                    # [추가] OrderManager의 체결 이벤트를 RealtimeManager가 받도록 연결
                    self.manager.positions = self.initial_positions # 초기 실잔고 이식
                    self.manager.om.set_callback(self.manager.on_order_confirmed)
                    self.manager.start_subscribing()
                
                # 메시지 펌프 유지 (실시간 이벤트 수신을 위해 필수)
                self.wait_and_monitor(10)

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