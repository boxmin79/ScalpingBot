import path_finder
import os
import csv
import json
import time
import pythoncom
from datetime import datetime
from collections import deque
from API.RealtimeDataManager import RealtimeDataManager
from API.OrderManager import OrderManager  # 🎯 OrderManager 임포트
from API.AccountManager import AccountManager # 🎯 AccountManager 임포트


class RealtimeManager:
    def __init__(self, target_list, acc_no, acc_flag, trade_budget, logger):
        self.cfg = path_finder.get_cfg()
        # ... (기존 초기화 로직 동일) ...
        self.logger = logger
        self.acc_no = acc_no
        self.acc_flag = acc_flag
        self.trade_budget = trade_budget
        self.om = OrderManager()
        self.am = AccountManager()
        
        self.rdm = RealtimeDataManager(callback_func=self.on_realtime_data)
        self.targets = target_list
        self.prev_strength = {t['code']: 0.0 for t in self.targets}
        
        self.avg_vol_1m = {t['code']: t.get('avg_vol_60', 0) / 390 for t in self.targets}
        self.vol_windows = {t['code']: deque() for t in self.targets}
        
        self.buy_signals = {}
        self.sold_codes = set() 
        self.positions = {}
        self.max_positions = 10
        self.orderbook_state = {} # 🎯 호가창 데이터를 담을 딕셔너리
        self.is_exiting = {} 
        self.subscribed_codes = set()
        
        # 설정값 튜닝
        self.ts_activation_pct = 0.5
        self.ts_callback_pct = 0.3
        self.hard_stop_loss = -1.2
        
        # 🎯 [추가] 체결강도 상한선 (이상 수치 차단)
        self.strength_limit = 1000.0
        
        self.trade_summary_path = self.cfg.DATA_DIR / "trade_summary.csv"
        
        self._init_logging()
        
    def _init_logging(self):
        """로그 폴더 및 CSV 헤더 초기화 (중복 제거 통합본)"""
        # if not os.path.exists(self.cfg.DATA_DIR): 
        #     os.makedirs(self.cfg.DATA_DIR)
        # path_config 에서 생성하므로 주석처리
            
        if not os.path.exists(self.trade_summary_path):
            with open(self.trade_summary_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow([
                    'Signal_ID', 'Code', 'Name', 'Expected_Entry', 'Actual_Entry', 
                    'Slippage_Rate', 'Exit_Price', 'Return_Rate', 'MAE', 'MFE', 
                    'Exit_Reason', 'Hold_Duration'
                ])
            
    def start_subscribing(self):
        """[업데이트] 160여 종목에 대해 현재가와 호가창 동시 감시"""
        if not self.targets: return
        self.om.subscribe_conclusion()
        
        for t in self.targets:
            code = t['code']
            self.buy_signals[code] = False
            # 🎯 'jpbidcnld'(호가창)를 추가하여 160 x 2 = 320개 모듈 사용
            self.rdm.start_monitoring(code, types=['cur', 'jpbidcnld']) 
            self.subscribed_codes.add(code)
            
        self.logger.info(f"📡 [정밀 감시] {len(self.targets)}종목 수급+호가 데이터 분석 시작")
        
    def on_realtime_data(self, code, data):
        """RealtimeDataManager에서 넘어오는 모든 콜백의 최상위 진입점"""
        # 호가창 데이터 처리
        if 'total_ask_vol' in data:
            self.process_orderbook(code, data)
            return
            
        # 1. 포지션 관리 및 MAE/MFE 업데이트
        if code in self.positions:
            pos = self.positions[code]
            curr_price = data.get('current', data.get('cur', 0))
            
            # MAE/MFE 추적 (최고/최저가 갱신)
            pos['max_price'] = max(pos.get('max_price', curr_price), curr_price)
            pos['min_price'] = min(pos.get('min_price', curr_price), curr_price)
            
            # 매도 조건 체크
            self.manage_exit(code, curr_price)
            return # 이미 보유 중인 종목은 매수 로직(아래)을 안 탐

        # 2. 보유 중이 아니라면 틱 데이터 분석 (매수 신호 탐색)
        self.process_tick(code, data)

    def force_exit_all(self):
        """보유 중인 모든 포지션 강제 매도"""
        if not self.positions:
            return

        # 딕셔너리를 순회하며 매도 주문 집행
        # 순회 중 삭제 에러 방지를 위해 list로 감싸서 key를 가져옵니다.
        for code in list(self.positions.keys()):
            pos = self.positions[code]
            self.logger.info(f"⏰ [장 마감 강제매도] {pos['name']}({code})")
            
            # 현재가 정보를 실시간 데이터 매니저 등에서 가져와야 함 (가장 최근가 기준)
            # 여기서는 편의상 시장가에 가까운 호가나 마지막 수신가로 가정
            # 만약 시장가(03) 주문이 가능하다면 가격을 0으로 넣고 호가구분을 변경하세요.
            self.om.request_new_order(
                self.acc_no, self.acc_flag, code, pos['qty'], 0, order_type="1", hoga_flag="03"
            )
            # 주의: 여기서 바로 del 하지 말고, on_order_confirmed에서 체결 확인 후 삭제하는 것이 안전합니다.
    
    def process_orderbook(self, code, data):
        """호가창 데이터를 분석하여 상태 저장"""
        # 총 매도잔량과 총 매수잔량 추출
        total_ask_vol = data.get('total_ask_vol', 0)
        total_bid_vol = data.get('total_bid_vol', 0)
        
        if total_ask_vol > 0 and total_bid_vol > 0:
            self.orderbook_state[code] = {
                'total_ask_vol': total_ask_vol,
                'total_bid_vol': total_bid_vol,
                'ratio': total_ask_vol / total_bid_vol # 매도/매수 비율
            }

    def process_tick(self, code, data):
        curr_price = data.get('current', 0)
        tick_vol = data.get('tick_vol', 0)
        strength = data.get('strength', 0.0) # 현재 체결강도
        
        # 2. 거래량 누적 (최근 10초 슬라이딩 윈도우)
        now = time.time()
        self.vol_windows[code].append((now, tick_vol))
        
        # 10초가 지난 데이터는 제거
        while self.vol_windows[code] and now - self.vol_windows[code][0][0] > 10:
            self.vol_windows[code].popleft()

        # 2. 체결강도 가속도 계산
        # 현재 체결강도 - 직전 틱 체결강도
        accel = strength - self.prev_strength.get(code, strength)
        self.prev_strength[code] = strength
        
        # 3. 통합 전략 분석
        self.analyze_combined_signal(code, data, strength, accel)
    
    def analyze_combined_signal(self, code, data, strength, accel):
        """거래량 폭발 + 체결강도 가속도 + 호가창 잔량 통합 분석"""
        price = data.get('current', 0)
        if price <= 0: return
        
        # 1. 거래량 폭발 계산 (10초 -> 1분 예측)
        ten_sec_vol = sum(v for t, v in self.vol_windows[code])
        predicted_1m_vol = ten_sec_vol * 6
        base_vol_1m = self.avg_vol_1m.get(code, 0)
        
        if base_vol_1m <= 0: return
        vol_multiple = predicted_1m_vol / base_vol_1m
        
        # 2. 호가창 필터 확인
        ob = self.orderbook_state.get(code)
        # 매도잔량이 매수잔량보다 최소 1.2배는 많아야 상방 에너지가 있다고 판단
        is_orderbook_valid = ob and ob['ratio'] >= 1.2 
        
        # 🎯 [핵심 필터 적용]
        # - 거래량 15배 돌파
        # - 체결강도 110% 이상이며 1000% 이하 (이상치 제거)
        # - 가속도 1.5 이상
        # - 매도잔량 우위 (호가창 필터)
        if (vol_multiple > 15.0 and 
            110.0 <= strength <= self.strength_limit and 
            accel >= 1.5 and 
            is_orderbook_valid):
            
            self.logger.info(f"🚀 [진짜 수급 포착] {data.get('name')} | "
                            f"폭발:{vol_multiple:.1f}배 | 강도:{strength:.1f}% | "
                            f"호가비율:{ob['ratio']:.2f}")
            self.execute_buy(code, data.get('name'), price)
    
    def execute_buy(self, code, name, price):
        """매수 주문 실행 (서버 조회 없이 로컬 예산 기반으로 즉시 집행)"""
        # 1. 방어 로직
        if price <= 0 or self.trade_budget <= 0 or len(self.positions) >= self.max_positions:
            return

        # 🎯 [수정] 서버 통신(get_buyable_data)을 생략하고 로컬 예산으로 수량 계산
        # 수수료/세금을 고려하여 예산의 99%만 사용하도록 안전장치
        safe_budget = self.trade_budget * 0.99
        final_buy_qty = int(safe_budget // price)

        if final_buy_qty <= 0:
            self.logger.info(f"⚠️ [매수 스킵] {name} | 예산 대비 주가 너무 높음 (현재가: {price:,}원)")
            return

        # 재매수 방지 로직 (sold_codes에 있으면 패스)
        if code in self.sold_codes:
            return

        self.logger.info(f"🔥 [매수 요청] {name}({code}) | 수량: {final_buy_qty}주")
        
        # 시장가(03) 주문 실행
        try:
            # 주문 유형: "2" (매수), "03" (시장가)
            self.om.request_new_order(self.acc_no, self.acc_flag, code, final_buy_qty, 0, order_type="2", hoga_flag="03")
            
           # MAE/MFE 추적을 위해 임시 포지션 정보 바로 생성
            signal_id = f"{datetime.now().strftime('%H%M%S')}_{code}"
            self.positions[code] = {
                'signal_id': signal_id,
                'name': name,
                'qty': final_buy_qty,
                'expected_entry_price': price, # 실제 매수가(buy_price)로 쓸 기준
                'actual_entry_price': price,   # 체결 전까지는 예상가와 동일하게 세팅
                'entry_time': time.time(),
                'max_price': price,
                'min_price': price,
                'is_concluded': False
            }
            self.buy_signals[code] = True 
        except Exception as e:
            self.logger.error(f"❌ 주문 중 오류 발생: {e}")

    # RealtimeManager.py 내 on_order_confirmed 수정
    def on_order_confirmed(self, concl_data):
        """OrderManager로부터 실제 체결 정보를 전달받음"""
        code = concl_data['code']
        if code in self.positions:
            pos = self.positions[code]
            pos['actual_entry_price'] = concl_data['actual_price']
            pos['is_concluded'] = True
            
            # 슬리피지 계산
            slippage = (pos['actual_entry_price'] - pos['expected_entry_price']) / pos['expected_entry_price']
            self.logger.info(f"✅ {pos['name']} 체결 완료 | 슬리피지: {slippage:.4%}")
                       
    def manage_exit(self, code, curr_price):
        """트레일링 스탑 및 손절 관리"""
        if self.is_exiting.get(code, False):
            return

        pos = self.positions[code]
        # 에러 수정: 'buy_price' 대신 'actual_entry_price' 사용
        buy_price = pos.get('actual_entry_price', pos['expected_entry_price'])
        
        fee_tax_rate = 0.23
        current_profit = ((curr_price - buy_price) / buy_price * 100) - fee_tax_rate
        highest_profit = ((pos['max_price'] - buy_price) / buy_price * 100) - fee_tax_rate

        sell_reason = None

        if highest_profit >= self.ts_activation_pct:
            if current_profit <= (highest_profit - self.ts_callback_pct):
                sell_reason = f"TS(최고 {highest_profit:.2f}% 대비 {self.ts_callback_pct}% 하락)"
        elif current_profit <= self.hard_stop_loss:
            sell_reason = f"손절(기준선 {self.hard_stop_loss}%)"

        if sell_reason:
            actual_balance = self.am.get_present_balance() 
            actual_qty = actual_balance.get(code, {}).get('qty', 0)
            
            sell_qty = min(pos['qty'], actual_qty) if actual_qty > 0 else pos['qty']
            
            if sell_qty <= 0:
                self.logger.error(f"❌ [매도 실패] {pos['name']} 서버 잔고 없음")
                self.write_final_log(code, curr_price, "에러-잔고없음")
                return

            self.logger.info(f"💰 [매도 실행] {pos['name']} | {sell_reason} | 수량: {sell_qty}주")
            self.is_exiting[code] = True
            self.om.request_new_order(self.acc_no, self.acc_flag, code, sell_qty, 0, order_type="1", hoga_flag="03")
            
            # 매도 주문 즉시 로그를 기록하고 포지션 해제
            self.write_final_log(code, curr_price, sell_reason)
    
    def write_final_log(self, code, exit_price, reason):
        """매매 종료 시 최종 결과 CSV 기록"""
        pos = self.positions.get(code)
        if not pos: return

        entry_price = pos.get('actual_entry_price', pos['expected_entry_price'])
        ret_rate = (exit_price - entry_price) / entry_price
        mae = (pos['min_price'] - entry_price) / entry_price
        mfe = (pos['max_price'] - entry_price) / entry_price
        duration = time.time() - pos['entry_time']
        slippage = (entry_price / pos['expected_entry_price']) - 1

        with open(self.trade_summary_path, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow([
                pos['signal_id'], code, pos['name'], pos['expected_entry_price'], 
                entry_price, round(slippage*100, 4),
                exit_price, round(ret_rate*100, 2), round(mae*100, 2), round(mfe*100, 2),
                reason, round(duration, 1)
            ])
        
        self.logger.info(f"📊 {pos['name']} 매도 완료 ({reason}) | 수익률: {ret_rate:.2%}")
        del self.positions[code]
        if code in self.is_exiting:
            del self.is_exiting[code]
            
    def get_tick_size(self, price):
        if price < 2000: return 1
        if price < 5000: return 5
        if price < 20000: return 10
        if price < 50000: return 50
        if price < 200000: return 100
        if price < 500000: return 500
        return 1000

    def stop_monitoring(self):
        self.rdm.stop_all()
        # 🎯 [추가] 실시간 체결 알림 해제 (메모리 누수 방지)
        self.om.unsubscribe_conclusion()
    
    # RealtimeManager.py에 추가할 메서드

    def sync_balance_with_server(self):
        """서버의 실제 잔고를 가져와 로컬 positions를 동기화합니다."""
        server_positions = self.am.get_present_balance() # AccountManager의 잔고 조회 함수
        
        for code, info in server_positions.items():
            if code in self.positions:
                # 수량이 다를 경우 서버 데이터 기준으로 업데이트
                if self.positions[code]['qty'] != info['qty']:
                    self.logger.warning(f"⚠️ [잔고 불일치 수정] {info['name']}: 로컬({self.positions[code]['qty']}) -> 서버({info['qty']})")
                    self.positions[code]['qty'] = info['qty']
            else:
                # 로컬에는 없는데 서버에 있는 경우 (예: 프로그램 재시작 시)
                self.positions[code] = {
                    'name': info['name'],
                    'buy_price': info['buy_price'],
                    'highest_price': info['current_price'],
                    'qty': info['qty'],
                    'entry_time': time.time()
                }
      