import path_finder
import time
import pythoncom
from datetime import datetime
from collections import deque
from API.RealtimeDataManager import RealtimeDataManager
from API.OrderManager import OrderManager  # 🎯 OrderManager 임포트
from API.AccountManager import AccountManager # 🎯 AccountManager 임포트


class RealtimeManager:
    def __init__(self, target_list, acc_no, acc_flag, trade_budget, logger):
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
        
    def on_realtime_data(self, data):
        code = data.get('code')
        dtype = data.get('type')
        
        if dtype == 'jpbidcnld':
            self.process_orderbook(code, data)
        elif dtype == 'cur':
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
        
        # 1. 매도 관리
        if code in self.positions:
            self.manage_exit(code, curr_price)
            return

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
                   
    # def analyze_volume_spike(self, code, data):
    #     """거래량 예측 돌파 전략 분석"""
    #     if len(self.positions) >= self.max_positions: return
    #     if self.buy_signals.get(code, False): return
        
    #     price = data.get('current', 0)
    #     diff = data.get('diff', 0) # 전일 대비 등락
        
    #     # [조건 1] 상승 중인 종목인가?
    #     if diff <= 0: return

    #     # [조건 2] 10초 누적 거래량 계산
    #     ten_sec_vol = sum(v for t, v in self.vol_windows[code])
        
    #     # [조건 3] 1분 거래량 예측 (10초 * 6)
    #     predicted_1m_vol = ten_sec_vol * 6
        
    #     # [조건 4] 60일 평균 분당 거래량과 비교 (예: 15배 초과 시)
    #     base_vol_1m = self.avg_vol_1m.get(code, 0)
    #     if base_vol_1m <= 0: return
        
    #     multi_level = 15.0 # 배수 설정 (10~20배 사이 권장)
        
    #     if predicted_1m_vol > (base_vol_1m * multi_level):
    #         # 호가창 잔량 확인 (매도잔량이 매수잔량보다 많아야 위로 쏠 에너지가 있음)
    #         state = self.orderbook_state.get(code)
    #         if state and state['total_ask_vol'] > state['total_bid_vol']:
    #             self.logger.info(f"🚀 [폭발 포착] {data.get('name')} | 예측1분:{int(predicted_1m_vol)} > 평균1분:{int(base_vol_1m)} ({multi_level}배 돌파)")
    #             self.execute_buy(code, data.get('name'), price)
                
    # def analyze_entry(self, code, data, strength, delta):
    #     """매수 타점 분석 및 주문 집행"""
        
    #     if len(self.positions) >= self.max_positions:
    #         return
        
    #     if self.buy_signals.get(code, False): return
        
    #     price = data.get('current', 0)
    #     vol = data.get('tick_vol', 0)
    #     buy_sell = data.get('side') # '1': 매수체결, '2': 매도체결
        
    #     # 거래량 Spike 체크
    #     history = self.volume_history.get(code)
    #     avg_vol = sum(history) / len(history) if len(history) >= 5 else 999999
    #     history.append(vol)
        
    #     state = self.orderbook_state.get(code)
    #     if not state or state['total_bid_vol'] == 0: return

    #     # 전략 필터 (호가밀도, 강도 가속도, 거래량 폭발)
    #     # print(f"code: {code} | {state['spread']} | {state['is_dense']} | {strength} | {round(delta,2)} | {vol} | {round(avg_vol, 2)}")
    #     if state['spread'] <= self.get_tick_size(price) and state['is_dense']:
    #         if strength >= 110.0 and delta >= 1.0:
    #             if buy_sell == '1' and (vol >= 1000 and vol >= avg_vol * 5.0):
    #                 if state['total_ask_vol'] >= (state['total_bid_vol'] * 1.5):
    #                     # 🎯 실제 매수 주문 집행
    #                     self.execute_buy(code, data.get('name'), price)

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

        self.logger.info(f"🔥 [매수 요청] {name}({code}) | 수량: {final_buy_qty}주 | 예상금액: {final_buy_qty * price:,}원")
        
        # 시장가(03) 주문 실행
        try:
            # 주문 유형: "2" (매수), "03" (시장가)
            self.om.request_new_order(self.acc_no, self.acc_flag, code, final_buy_qty, 0, order_type="2", hoga_flag="03")
            
            # 임시 포지션 등록 (체결 확인 전까지 중복 주문 방지용)
            self.buy_signals[code] = True 
            
            # positions 업데이트는 on_order_confirmed에서 처리됨
        except Exception as e:
            self.logger.error(f"❌ 주문 중 오류 발생: {e}")

    # RealtimeManager.py 내 on_order_confirmed 수정
    def on_order_confirmed(self, data):
        if data['status'] == 'CONCLUDED':
            code = data['stock_code']
            exec_qty = data['volume']
            exec_price = data['price']
            
            if data['side'] == 'BUY':
                if code in self.positions:
                    pos = self.positions[code]
                    
                    # 🎯 가중 평균 단가 및 수량 합산 (정확한 로직)
                    current_total_cost = pos['buy_price'] * pos['qty']
                    new_fill_cost = exec_price * exec_qty
                    
                    pos['qty'] += exec_qty # 기존 수량에 더하기
                    pos['buy_price'] = (current_total_cost + new_fill_cost) / pos['qty']
                    
                    # 추가 매수 시 최고가는 현재 체결가와 기존 최고가 중 큰 것으로 갱신
                    pos['highest_price'] = max(pos.get('highest_price', 0), exec_price)
                    
                    self.logger.info(f"✅ [매수 추가체결] {data['name']} | +{exec_qty}주 | 총: {pos['qty']}주 | 평단: {pos['buy_price']:,.0f}원")
                else:
                    # 신규 진입
                    self.positions[code] = {
                        'name': data['name'],
                        'buy_price': exec_price,
                        'highest_price': exec_price,
                        'qty': exec_qty,
                        'entry_time': time.time()
                    }
                    self.logger.info(f"✅ [매수 신규체결] {data['name']} | {exec_qty}주 | 평단가: {exec_price:,.0f}원")
                    
            elif data['side'] == 'SELL':
                if code in self.positions:
                    # 🎯 핵심: 전체 수량에서 체결된 만큼만 뺍니다.
                    self.positions[code]['qty'] -= exec_qty
                    
                    self.logger.info(f"✅ [매도 체결] {data['name']} | -{exec_qty}주 (남은 수량: {self.positions[code]['qty']}주)")
                    
                    # 🎯 남은 수량이 0 이하일 때만 포지션에서 완전히 삭제
                    if self.positions[code]['qty'] <= 0:
                        # 🎯 [재매수 금지] 매도 완료 시 당일 금지 목록에 추가
                        self.sold_codes.add(code)
                        self.logger.info(f"🏁 [매도 완료] {data['name']} 포지션 종료")
                        del self.positions[code]
                        if code in self.is_exiting:
                            del self.is_exiting[code]
                    
    def manage_exit(self, code, curr_price):
        """트레일링 스탑 및 손절 관리"""
        if code not in self.positions or self.is_exiting.get(code, False):
            return

        pos = self.positions[code]
        buy_price = pos['buy_price']
        
        # 🎯 highest_price가 없으면 현재가로 초기화 (KeyError 방지)
        if 'highest_price' not in pos:
            pos['highest_price'] = curr_price

        # 1. 최고가 갱신
        if curr_price > pos['highest_price']:
            pos['highest_price'] = curr_price

        # 2. 수익률 계산 (제비용 0.23% 반영)
        fee_tax_rate = 0.23
        current_profit = ((curr_price - buy_price) / buy_price * 100) - fee_tax_rate
        highest_profit = ((pos['highest_price'] - buy_price) / buy_price * 100) - fee_tax_rate

        # 🎯 3. 매도 판단 로직
        sell_reason = None

        # A. 트레일링 스탑 (수익 보전)
        # 설정한 활성화 수익률(0.5%)을 넘긴 적이 있고, 최고가 수익률 대비 지정된 폭(0.3%)만큼 하락했을 때
        if highest_profit >= self.ts_activation_pct:
            if current_profit <= (highest_profit - self.ts_callback_pct):
                sell_reason = f"TS(최고 {highest_profit:.2f}% 대비 {self.ts_callback_pct}% 하락)"

        # B. 하드 손절 (방어선)
        if current_profit <= self.hard_stop_loss:
            sell_reason = f"손절(기준선 {self.hard_stop_loss}%)"

        if sell_reason:
            # 🎯 주문 전 실제 서버 잔고 수량 확인 (동기화 실패 대비)
            actual_balance = self.am.get_present_balance() # AccountManager에 구현된 함수라 가정
            actual_qty = actual_balance.get(code, {}).get('qty', 0)
            
            sell_qty = min(pos['qty'], actual_qty) if actual_qty > 0 else pos['qty']
            
            if sell_qty <= 0:
                self.logger.error(f"❌ [매도 실패] {pos['name']} 서버 잔고 없음 (로컬:{pos['qty']}주)")
                del self.positions[code] # 잔고가 없으므로 포지션 삭제
                return

            self.logger.info(f"💰 [매도 실행] {pos['name']} | {sell_reason} | 수량: {sell_qty}주")
            self.is_exiting[code] = True
            self.om.request_new_order(self.acc_no, self.acc_flag, code, sell_qty, 0, order_type="1", hoga_flag="03")
            
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
                
    # RealtimeManager.py 내부 추가 및 수정

    # def update_targets(self, new_target_list):
    #     """TradingBot에서 호출: 기존 포지션은 유지하며 감시 대상 종목만 교체"""
    #     print(f"🔄 감시 종목 업데이트 중... (기존 포지션 유지)")

    #     # 🎯 1. 수정: rdm.monitored_codes 대신 직접 만든 subscribed_codes 사용!
    #     current_subscribed_codes = set(self.subscribed_codes) 
        
    #     # 2. 새로운 타깃 코드 목록 (보유 종목 포함)
    #     new_target_codes = {t['code'] for t in new_target_list}
    #     holding_codes = set(self.positions.keys())
        
    #     # 최종 유지해야 할 전체 코드 (타깃 + 현재 보유 중인 종목)
    #     final_codes = new_target_codes | holding_codes
        
    #     # 3. 삭제 대상: 현재 구독 중이지만 final_codes에 없는 종목
    #     to_unsubscribe = current_subscribed_codes - final_codes
    #     for code in to_unsubscribe:
    #         self.rdm.stop_monitoring(code)
    #         self.subscribed_codes.discard(code) # 🎯 내 리스트에서도 삭제
    #         if code in self.buy_signals: del self.buy_signals[code]
            
    #     # 4. 추가 대상: final_codes에 있지만 아직 구독 중이 아닌 종목
    #     to_subscribe = final_codes - current_subscribed_codes
    #     for code in to_subscribe:
    #         # 종목명 찾기 (new_target_list 또는 positions에서)
    #         name = next((t['name'] for t in new_target_list if t['code'] == code), "Unknown")
    #         if name == "Unknown" and code in self.positions:
    #             name = self.positions[code]['name']
                
    #         # 상태 변수 초기화
    #         self.buy_signals[code] = False
    #         self.orderbook_state[code] = {'total_ask_vol': 0, 'total_bid_vol': 0, 'spread': 9999, 'is_dense': False}
    #         if code not in self.volume_history:
    #             self.volume_history[code] = deque(maxlen=20)
            
    #         # 실제 구독 시작
    #         self.rdm.start_monitoring(code, types=['cur', 'jpbidcnld'])
    #         self.subscribed_codes.add(code) # 🎯 [추가] 새 종목도 내 리스트에 등록
    #         print(f"   ➕ 새 감시 추가: {name}({code})")

    #     # 5. 내부 targets 리스트 갱신
    #     self.targets = new_target_list