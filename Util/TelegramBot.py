import path_finder
import os
import requests
import time
import threading
import pythoncom  # 🎯 추가: COM 초기화를 위해 필요
from dotenv import load_dotenv
from API.AccountManager import AccountManager

load_dotenv()

class TelegramBot:
    def __init__(self):
        self.is_running = True # 종료 플래그 추가
        
        # 🎯 수정: get_cfg 뒤에 () 추가
        self.cfg = path_finder.get_cfg() 
        self.token = os.getenv("TELEGRAM_API")
        self.chat_id = os.getenv("CHAT_ID")
        self.last_update_id = 0
                
        if not self.chat_id:
            print("⚠️ .env에 CHAT_ID가 없어 직접 조회를 시작합니다.")
            self.chat_id = self.fetch_chat_id_from_api()
        else:
            print(f"✅ Chat ID 로드 완료: {self.chat_id}")
            
        self.manager = None # 🎯 TradingBot에서 manager를 주입받을 변수 추가

    def send_message(self, message):
        """메시지 전송 (Markdown 지원)"""
        if not self.chat_id: return
        try:
            url = f"https://api.telegram.org/bot{self.token}/sendMessage"
            data = {"chat_id": self.chat_id, "text": message, "parse_mode": "Markdown"}
            # 🎯 팁: 로그용 전송은 timeout을 짧게 잡는 것이 좋습니다.
            requests.post(url, data=data, timeout=5)
        except Exception as e:
            print(f"❌ 텔레그램 전송 오류: {e}")

    def listen(self):
        """사용자의 메시지를 실시간으로 감시 (별도 쓰레드에서 실행됨)"""
        # 🎯 중요: 별도 쓰레드에서 COM 객체(AccountManager)를 사용하기 위한 초기화
        pythoncom.CoInitialize() 
        print("🚀 텔레그램 봇 리스닝 시작...")
        
        # 🎯 2. [추가] 반드시 COM 환경이 구축된 쓰레드 내부에서 매니저를 생성해야 합니다!
        self.am = AccountManager()
        
        while self.is_running:
            try:
                url = f"https://api.telegram.org/bot{self.token}/getUpdates?offset={self.last_update_id + 1}&timeout=30"
                response = requests.get(url, timeout=35).json()

                if response.get("result"):
                    for update in response["result"]:
                        self.last_update_id = update["update_id"]
                        if "message" in update and "text" in update["message"]:
                            self.handle_command(update["message"]["text"])
                
                time.sleep(1)
            except Exception as e:
                print(f"❌ 리스닝 오류: {e}")
                time.sleep(5)
        
        # 리스닝 종료 시 COM 해제 (생략 가능하지만 깔끔한 종료를 위해)
        pythoncom.CoUninitialize()
    
    def stop(self):
        self.is_running = False
        print("🛑 텔레그램 봇 리스닝을 중단합니다.")
        
    def handle_command(self, command):
        # 양끝 공백 제거 및 텍스트 정규화
        cmd = command.strip()

        # 🎯 [추가] 봇이 메시지를 정상적으로 수신했는지 터미널에서 확인하기 위한 로그
        print(f"📨 [텔레그램 수신] 입력된 명령어: {cmd}")
        
        if cmd in ["/start", "시작", "start"]:
            self.send_message("🤖 트레이딩 봇 가동! 명령어를 입력하세요.")
            
        elif cmd in ["/매매손익", "매매손익", "손익", "수익률"]:
            # "매매손익" 또는 "손익"이라고만 보내도 실행됨
            self.send_internal_profit_report() # 🎯 수정된 내부 리포트 함수 호출
        
        # 🎯 '잔고평가' 명령어 추가
        elif cmd in ["잔고평가", "잔고", "현황"]:
            self.send_balance_report()
            
        elif "계좌" in cmd:
            # "내 계좌 보여줘" 같이 단어가 포함만 되어도 실행
            self.send_message("💰 계좌 정보를 조회합니다...")
            
        else:
            self.send_message(f"❓ '{cmd}'은(는) 등록되지 않은 명령어입니다.")

    def send_profit_loss_report(self):
        """매매 손익 리포트 전송 (소수점 보존 및 개별 항목 정밀 출력)"""
        summary, details = self.am.get_profit_loss_data()
        if not summary:
            self.send_message("❌ 데이터를 가져오지 못했습니다.")
            return

        # 1. 헬퍼 함수 정의 (타입 에러 방지용)
        def to_money_header(val):
            """헤더용: 천 원 -> 원 단위 변환 (문자열 대응)"""
            try:
                return int(float(str(val).replace(',', '').strip())) * 1000
            except:
                return 0

        def to_int_pure(val):
            """상세 내역용: 원 단위 그대로 정수 변환"""
            try:
                return int(float(str(val).replace(',', '').strip()))
            except:
                return 0

        def to_float(val):
            """수익률용: 소수점 유실 방지 (float 강제 형변환)"""
            try:
                if val is None or str(val).strip() == "": return 0.0
                return float(str(val).replace(',', ''))
            except:
                return 0.0

        # 2. 요약 데이터 가공
        realized_pl = to_money_header(summary.get('total_realized_pl'))
        eval_pl = to_money_header(summary.get('total_eval_pl'))
        # API 제공 수익률을 사용하되, float로 변환하여 절삭 오류 방지
        total_yield = to_float(summary.get('total_yield'))

        msg = "📊 *금일 매매 손익 리포트*\n"
        msg += "--------------------------\n"
        msg += f"💰 *실현 손익:* {realized_pl:,}원\n"
        msg += f"📈 *평가 손익:* {eval_pl:,}원\n"
        msg += f"📉 *당일 수익률:* {total_yield:.2f}%\n"
        msg += "--------------------------\n\n"

        # 3. 상세 내역 출력
        if not details:
            msg += "매매 내역이 없습니다."
        else:
            for item in details:
                # 개별 종목 수익률 및 손익 추출
                item_yield = to_float(item.get('yield'))
                item_pl = to_int_pure(item.get('realized_pl'))
                
                # 수익률에 따른 이모지 설정 (빨강: 상승, 파랑: 하락)
                emoji = "🔴" if item_yield > 0 else "🔵" if item_yield < 0 else "⚪"
                
                msg += f"{emoji} *{item.get('name', '알 수 없음')}*\n"
                msg += f"   - 실현: {item_pl:,}원 ({item_yield:.2f}%)\n"
        
        self.send_message(msg)
    
    def send_internal_profit_report(self):
        """내부 로직 기반 리포트 전송"""
        if not self.manager:
            self.send_message("❌ 매니저가 연결되지 않았습니다. (봇 가동 전)")
            return

        data = self.manager.get_internal_report_data()
        
        msg = "📊 *프로그램 내부 매매 리포트*\n"
        msg += "--------------------------\n"
        msg += f"💰 *실현 손익:* {data['realized_pl']:,}원\n"
        msg += f"📈 *평가 손익:* {data['eval_pl']:,}원\n"
        msg += f"📉 *당일 평균 수익률:* {data['total_yield']:.2f}%\n"
        msg += "--------------------------\n\n"

        if not data['details']:
            msg += "오늘의 매매 내역이 아직 없습니다."
        else:
            for item in data['details']:
                emoji = "🔴" if item['yield'] > 0 else "🔵"
                msg += f"{emoji} *{item['name']}*\n"
                msg += f"   - 실현: {item['pl']:,}원 ({item['yield']:.2f}%)\n"
        
        self.send_message(msg)
        
    def send_balance_report(self):
        """실시간 잔고 및 보유 종목 현황 전송 (안전한 형변환 적용)"""
        summary, stocks = self.am.get_balance_data()
        
        if not summary:
            self.send_message("❌ 잔고 데이터를 가져오지 못했습니다.")
            return

        # 헬퍼 함수: 금액용 (원 단위 그대로 사용, 소수점만 제거)
        def to_int(val):
            try:
                if val is None or str(val).strip() == "": return 0
                return int(float(str(val).replace(',', '')))
            except:
                return 0

        # 헬퍼 함수: 수익률용 (소수점 유지)
        def to_float(val):
            try:
                if val is None or str(val).strip() == "": return 0.0
                return float(str(val).replace(',', ''))
            except:
                return 0.0

        # 1. 계좌 요약 정보 추출
        total_eval = to_int(summary.get('Total_Eval_Amt'))
        d2_deposit = to_int(summary.get('D2_Deposit'))
        total_pl = to_int(summary.get('Total_Profit_Loss'))
        total_yield = to_float(summary.get('Total_Yield'))

        msg = "🏢 *실시간 잔고 평가 현황*\n"
        msg += "--------------------------\n"
        msg += f"💰 *총 평가금액:* {total_eval:,}원\n"
        msg += f"💵 *D+2 예상예수금:* {d2_deposit:,}원\n"
        msg += f"📊 *총 평가손익:* {total_pl:,}원\n"
        msg += f"📈 *총 수익률:* {total_yield:.2f}%\n"
        msg += "--------------------------\n\n"

        # 2. 보유 종목 상세 내역
        if not stocks:
            msg += "현재 보유 중인 종목이 없습니다."
        else:
            for s in stocks:
                # 안전하게 데이터 추출
                s_name = s.get('name', '알 수 없음')
                s_code = s.get('code', '000000')
                s_yield = to_float(s.get('yield'))
                s_qty = to_int(s.get('total_qty'))
                s_pl = to_int(s.get('profit_loss'))
                s_buy_price = to_int(s.get('buy_price'))

                # 수익률 이모지 설정
                emoji = "🔴" if s_yield > 0 else "🔵" if s_yield < 0 else "⚪"
                
                msg += f"{emoji} *{s_name}* ({s_code})\n"
                msg += f"   - 보유: {s_qty:,}주\n"
                msg += f"   - 손익: {s_pl:,}원 ({s_yield:.2f}%)\n"
                msg += f"   - 매입가: {s_buy_price:,}원\n\n"

        self.send_message(msg)
        
if __name__ == "__main__":
    # 1. 봇 객체 생성 
    # 생성자에서 .env 로드, Chat ID 확인, AccountManager 초기화가 수행됩니다.
    test_bot = TelegramBot()

    # 2. 메시지 전송 테스트
    # 프로그램이 실행되자마자 텔레그램으로 메시지가 오는지 확인하세요.
    print("--- 전송 테스트 시작 ---")
    test_bot.send_message("🚀 **텔레그램 봇 단독 테스트 시작**\n이제 `/매매손익` 또는 아무 메시지나 보내보세요!")
    print("✅ 테스트 메시지를 전송했습니다.")

    # 3. 메시지 수신(Listening) 테스트
    # 이 함수는 무한 루프이므로 실행 시 프로그램이 종료되지 않고 대기 상태가 됩니다.
    # 터미널에서 Ctrl+C를 누르면 종료할 수 있습니다.
    try:
        print("--- 리스닝 테스트 시작 ---")
        test_bot.listen()
    except KeyboardInterrupt:
        print("\n👋 사용자에 의해 테스트가 종료되었습니다.")
    except Exception as e:
        print(f"🚨 테스트 중 오류 발생: {e}")