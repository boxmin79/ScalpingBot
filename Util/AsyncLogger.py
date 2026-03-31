import path_finder
import logging
import logging.handlers
import queue
import threading
import time
from datetime import datetime  # 날짜 처리를 위해 추가
from Util.TelegramBot import TelegramBot #

class AsyncLogger:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        # 1. 로그 메시지를 담을 큐 생성
        self.log_queue = queue.Queue(-1)
        
        # 1. 오늘 날짜 가져오기 (예: 2026-03-31)
        today = datetime.now().strftime("%Y-%m-%d")
            
        # 3. 파일명에 날짜 적용
        self.logs_path = self.cfg.LOGS_DIR / f"trading_bot_{today}.log"
        
        # 2. 실제 로그를 처리할 핸들러들 정의
        # 콘솔 출력용
        console_handler = logging.StreamHandler()
        # 파일 저장용
        file_handler = logging.FileHandler(self.logs_path, encoding='utf-8')
        
        formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s')
        console_handler.setFormatter(formatter)
        file_handler.setFormatter(formatter)

        # 3. 큐 리스너 설정 (백그라운드에서 핸들러들을 실행)
        self.listener = logging.handlers.QueueListener(
            self.log_queue, 
            console_handler, 
            file_handler
        )
        
        # 4. 메인 로거 설정 (큐 핸들러만 연결)
        self.logger = logging.getLogger("TradingBot")
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(logging.handlers.QueueHandler(self.log_queue))
        
        # 텔레그램 봇 연동
        self.tel_bot = TelegramBot()
        
        # 리스너 시작
        self.listener.start()

    def info(self, msg, send_tg=True):
        self.logger.info(msg)
        if send_tg:
            # 텔레그램 전송도 별도 스레드로 돌려야 메인 루프가 안 멈춥니다.
            threading.Thread(target=self.tel_bot.send_message, args=(f"{msg}",), daemon=True).start()

    def error(self, msg):
        self.logger.error(msg)
        threading.Thread(target=self.tel_bot.send_message, args=(f"🚨 **[ERROR]** {msg}",), daemon=True).start()

    def stop(self):
        """프로그램 종료 시 리스너 정지"""
        self.info("프로그램 종료")
        self.listener.stop()
        
if __name__ == "__main__":
    logger = AsyncLogger()
    logger.info("테스트 메시지")
    # 메시지가 전송될 때까지 잠시 대기 (테스트용)
    time.sleep(2)
    logger.stop()