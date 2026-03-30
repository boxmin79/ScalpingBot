import path_finder
import time
import pythoncom
from datetime import datetime
from collections import deque
from API.RealtimeDataManager import RealtimeDataManager
from API.OrderManager import OrderManager  # 🎯 OrderManager 임포트

class RealtimeManager:
    def __init__(self, target_list, acc_no, acc_flag):
        # 🎯 계좌 정보 및 주문 매니저 초기화
        self.acc_no = acc_no
        self.acc_flag = acc_flag
        self.om = OrderManager()
        
        self.rdm = RealtimeDataManager(callback_func=self.on_realtime_data)
        self.targets = target_list
        
        # 상태 관리 변수
        self.buy_signals = {}      # 신호 발생 여부
        self.positions = {}        # 🎯 현재 보유 종목 상태 {code: {'buy_price': 0, 'qty': 0}}
        self.orderbook_state = {}
        self.volume_history = {t['code']: deque(maxlen=20) for t in self.targets}
        self.prev_strength = {t['code']: 0.0 for t in self.targets}

    def start_subscribing(self):
        """TradingBot에서 호출: 구독 시작"""
        if not self.targets: return
        print(f"   📡 {len(self.targets)}종목 실시간 감시 및 매매 엔진 가동...")
        
        for t in self.targets:
            code = t['code']
            self.buy_signals[code] = False
            self.orderbook_state[code] = {'total_ask_vol': 0, 'total_bid_vol': 0, 'spread': 9999, 'is_dense': False}
            self.rdm.start_monitoring(code, types=['cur', 'jpbidcnld'])

    def on_realtime_data(self, data):
        """데이터 수신 콜백"""
        code = data.get('code')
        dtype = data.get('type')
        
        if dtype == 'jpbidcnld':
            self.process_orderbook(code, data)
        elif dtype == 'cur':
            self.process_tick(code, data)

    def process_orderbook(self, code, data):
            asks = data.get('asks', [])
            bids = data.get('bids', [])
            ask_vols = data.get('ask_vols', [])
            
            if asks and bids:
                spread = asks[0] - bids[0]
                is_dense = all(v > 0 for v in ask_vols[:3])
                self.orderbook_state[code] = {
                    'total_ask_vol': data['total_ask_vol'],
                    'total_bid_vol': data['total_bid_vol'],
                    'spread': spread,
                    'is_dense': is_dense
                }

    def process_tick(self, code, data):
        curr_price = data.get('current', 0)
        
        # 1. 매도(수익률) 관리: 이미 보유 중인 종목인 경우
        if code in self.positions:
            self.manage_exit(code, curr_price)
            return

        # 2. 매수 타점 분석
        curr_strength = data.get('strength', 0.0)
        delta = curr_strength - self.prev_strength.get(code, 0.0)
        self.prev_strength[code] = curr_strength

        self.analyze_entry(code, data, curr_strength, delta)

    def analyze_entry(self, code, data, strength, delta):
        """매수 타점 분석 및 주문 집행"""
        if self.buy_signals.get(code, False): return
        
        price = data.get('current', 0)
        vol = data.get('tick_vol', 0)
        buy_sell = data.get('side') # '1': 매수체결, '2': 매도체결
        
        # 거래량 Spike 체크
        history = self.volume_history.get(code)
        avg_vol = sum(history) / len(history) if len(history) >= 5 else 999999
        history.append(vol)
        
        state = self.orderbook_state.get(code)
        if not state or state['total_bid_vol'] == 0: return

        # 전략 필터 (호가밀도, 강도 가속도, 거래량 폭발)
        if state['spread'] <= self.get_tick_size(price) and state['is_dense']:
            if strength >= 110.0 and delta >= 1.0:
                if buy_sell == '1' and (vol >= 1000 and vol >= avg_vol * 5.0):
                    if state['total_ask_vol'] >= (state['total_bid_vol'] * 1.5):
                        # 🎯 실제 매수 주문 집행
                        self.execute_buy(code, data.get('name'), price)

    def execute_buy(self, code, name, price):
        """매수 주문 실행"""
        # 스캘핑을 위해 고정 수량(예: 10주) 또는 금액 기준으로 계산 필요
        buy_qty = 10 
        
        print(f"\n🔥 [매수 집행] {name}({code}) | 가격: {price:,} | 수량: {buy_qty}")
        # OrderManager의 request_new_order 호출
        order_no = self.om.request_new_order(self.acc_no, self.acc_flag, code, buy_qty, price, order_type="2")
        
        if order_no:
            self.buy_signals[code] = True
            self.positions[code] = {
                'name': name,
                'buy_price': price,
                'qty': buy_qty,
                'entry_time': time.time()
            }

    def manage_exit(self, code, curr_price):
        """실시간 익절/손절 관리 (매도 로직)"""
        pos = self.positions[code]
        buy_price = pos['buy_price']
        profit_rate = (curr_price - buy_price) / buy_price * 100

        # 🎯 스캘핑 매도 원칙 (예: +1.5% 익절, -1.0% 손절)
        is_take_profit = profit_rate >= 1.5
        is_stop_loss = profit_rate <= -1.0
        
        if is_take_profit or is_stop_loss:
            reason = "익절" if is_take_profit else "손절"
            print(f"💰 [{reason} 실행] {pos['name']} | 수익률: {profit_rate:.2f}% | 가격: {curr_price:,}")
            
            # 🎯 실제 매도 주문 집행
            order_no = self.om.request_new_order(
                self.acc_no, self.acc_flag, code, pos['qty'], curr_price, order_type="1"
            )
            
            if order_no:
                del self.positions[code] # 포지션 제거

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