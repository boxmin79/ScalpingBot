import path_finder
import logging
import logging.handlers
import queue
import threading
from datetime import datetime
from Util.TelegramBot import TelegramBot
import os

# 🎯 텔레그램 전송 전용 핸들러 정의
class TelegramLogHandler(logging.Handler):
    def __init__(self, tel_bot):
        super().__init__()
        self.tel_bot = tel_bot

    def emit(self, record):
        # 🎯 레코드에 'send_tg' 속성이 있거나 레벨이 ERROR 이상일 때만 전송
        send_tg = getattr(record, 'send_tg', False)
        if send_tg or record.levelno >= logging.ERROR:
            msg = self.format(record)
            # 이미 리스너 쓰레드 내부이므로 여기서 직접 호출해도 메인 루프에 영향 없음
            self.tel_bot.send_message(msg)

class AsyncLogger:
    def __init__(self):
        self.cfg = path_finder.get_cfg()
        self.log_queue = queue.Queue(-1)
        today = datetime.now().strftime("%Y-%m-%d")
        self.logs_path = self.cfg.LOGS_DIR / f"trading_bot_{today}.log"
        
        # 1. 텔레그램 봇 초기화
        self.tel_bot = TelegramBot()
        self.tele_thread = threading.Thread(target=self.tel_bot.listen, daemon=True)
        self.tele_thread.start()
        
        # 2. 핸들러 설정
        console_handler = logging.StreamHandler()
        file_handler = logging.FileHandler(self.logs_path, encoding='utf-8')
        telegram_handler = TelegramLogHandler(self.tel_bot) # 🎯 텔레그램 핸들러 추가
        
        formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s')
        console_handler.setFormatter(formatter)
        file_handler.setFormatter(formatter)
        telegram_handler.setFormatter(logging.Formatter('%(message)s')) # 텔레그램은 메시지만

        # 3. 큐 리스너에 모든 핸들러 등록
        self.listener = logging.handlers.QueueListener(
            self.log_queue, 
            console_handler, 
            file_handler,
            telegram_handler
        )
        
        # 4. 로거 설정
        self.logger = logging.getLogger("TradingBot")
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(logging.handlers.QueueHandler(self.log_queue))
        
        self.listener.start()

        # AsyncLogger.py 에 추가할 예시 메서드
    def log_trade_summary(self, trade_data):
        """
        trade_data: dict 형태의 매매 결과
        형식: ID, 종목, 시총, M, I, 진입가, 매도가, 수익률, MAE, MFE, 사유 등
        """
        import csv
        file_path = "Data/trade_summary.csv"
        file_exists = os.path.isfile(file_path)
        
        with open(file_path, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=trade_data.keys())
            if not file_exists:
                writer.writeheader()
            writer.writerow(trade_data)
    
    def info(self, msg, send_tg=True):
        # 🎯 extra를 통해 텔레그램 전송 여부를 핸들러에 전달
        self.logger.info(msg, extra={'send_tg': send_tg})

    def error(self, msg):
        # ERROR 레벨은 핸들러 설정에 의해 자동으로 텔레그램 전송됨
        self.logger.error(f"🚨 **[ERROR]** {msg}")

    def stop(self):
        self.info("시스템 종료 중...", send_tg=True)
        self.tel_bot.stop() # 봇의 stop 메서드 호출
        self.listener.stop()