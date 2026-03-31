import os
import requests
from dotenv import load_dotenv

# ... (경로 설정 로직) ...
load_dotenv(dotenv_path=r"C:\Users\realb\Documents\TradingBot\.env")

class TelegramBot:
    def __init__(self):
        self.token = os.getenv("TELEGRAM_API")
        # 1. 먼저 .env에서 CHAT_ID를 가져옵니다.
        self.chat_id = os.getenv("CHAT_ID")

        # 2. 만약 .env에 없다면, 그때 get_chat_id()를 실행합니다.
        if not self.chat_id:
            print("⚠️ .env에 CHAT_ID가 없어 직접 조회를 시작합니다.")
            self.chat_id = self.fetch_chat_id_from_api()
        else:
            print(f"✅ 설정 파일에서 Chat ID를 로드했습니다: {self.chat_id}")

    def fetch_chat_id_from_api(self):
        """최근 메시지를 보낸 사용자의 ID를 API로 조회"""
        try:
            url = f"https://api.telegram.org/bot{self.token}/getUpdates"
            response = requests.get(url).json()
            if response.get("result"):
                chat_id = response["result"][-1]["message"]["chat"]["id"]
                print(f"🔎 조회된 chat_id: {chat_id}")
                return chat_id
            return None
        except Exception as e:
            print(f"❌ 텔레그램 chat_id 조회 실패: {e}")
            return None

    def send_message(self, message):
        """메시지 전송"""
        if not self.token or not self.chat_id:
            print("🚫 메시지를 보낼 수 없습니다. 토큰 또는 Chat ID를 확인하세요.")
            return

        try:
            url = f"https://api.telegram.org/bot{self.token}/sendMessage"
            # chat_id가 .env에서 읽어오면 문자열(str)일 수 있으므로 그대로 사용하거나 int로 변환
            data = {"chat_id": self.chat_id, "text": message, "parse_mode": "Markdown"}
            requests.post(url, data=data)
        except Exception as e:
            print(f"❌ 메시지 전송 오류: {e}")
            
if __name__ == "__main__":
    bot = TelegramBot()
    bot.send_message("테스트 메시지입니다.")