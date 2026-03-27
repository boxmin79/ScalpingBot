import time
import pythoncom
import logging
from Screener.ScalpingScreener import ScalpingScreener
from API.RealtimeManager import RealtimeManager
# from API.CpTopVolume import CpTopVolume (실제 데이터 요청용)

def setup_logger():
    logger = logging.getLogger("TradingBot")
    logger.setLevel(logging.INFO)
    ch = logging.StreamHandler()
    formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt='%H:%M:%S')
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    return logger

class TradingBot:
    def __init__(self):
        self.logger = setup_logger()
        self.screener = ScalpingScreener()
        # 실시간 데이터가 들어올 때마다 실행될 함수(self.on_realtime_data)를 매니저에 넘김
        self.rt_manager = RealtimeManager(self.on_realtime_data)
        self.last_screen_time = 0

    def on_realtime_data(self, data):
        """이벤트 핸들러를 통해 수신된 실시간 데이터를 분석하여 매수/매도 시그널을 생성합니다."""
        if data['type'] == 'orderbook':
            ask_size = data['total_ask_size']
            bid_size = data['total_bid_size']
            
            # 스캘핑 시그널 로직: 매도 잔량이 매수 잔량보다 2배 이상 두껍게 쌓여있을 때 (돌파 준비)
            if bid_size > 0 and (ask_size / bid_size) >= 2.0:
                # 여기서 OrderManager의 buy_market()을 호출하는 로직으로 연결!
                self.logger.info(f"🔥 [돌파 임박 시그널] {data['code']} | 매도잔량: {ask_size} / 매수잔량: {bid_size}")

    def run(self):
        self.logger.info("트레이딩 봇 시작...")
        try:
            while True:
                current_time = time.time()
                
                # 3분(180초)마다 스크리너를 실행하여 종목 리스트 갱신 (장 초반엔 간격을 더 줄여도 좋습니다)
                if current_time - self.last_screen_time > 180:
                    self.logger.info("--- 감시 종목 리스트 갱신 중 ---")
                    
                    # 1. CpTopVolume 등을 호출해 실제 데이터를 가져옴 (여기선 더미 리스트로 가정)
                    # raw_data = CpTopVolume().get_data() 
                    raw_data = [] # TODO: 위에서 주신 리스트 형태의 데이터를 여기에 넣습니다.
                    
                    # 2. 스크리너를 통해 조건에 맞는 20개 압축
                    target_codes = self.screener.filter_targets(raw_data)
                    
                    # 3. 실시간 매니저에 전달하여 구독 상태 업데이트
                    self.rt_manager.update_targets(target_codes)
                    
                    self.last_screen_time = current_time

                # COM 이벤트 수신을 위한 필수 메시지 펌프 (메인 루프를 블로킹하지 않음)
                pythoncom.PumpWaitingMessages()
                
                # CPU 과부하 방지
                time.sleep(0.01) 

        except KeyboardInterrupt:
            self.logger.info("봇을 안전하게 종료합니다.")
        finally:
            self.rt_manager.unsubscribe_all()

if __name__ == "__main__":
    bot = TradingBot()
    bot.run()