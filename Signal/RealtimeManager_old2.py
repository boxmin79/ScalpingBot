import path_finder
import time
import pythoncom
from datetime import datetime
from collections import deque
from API.RealtimeDataManager import RealtimeDataManager
from API.OrderManager import OrderManager  # 🎯 OrderManager 임포트
from API.AccountManager import AccountManager # 🎯 AccountManager 임포트


class RealtimeManager:
    # 🎯 trade_budget 인자 추가
    def __init__(self, target_list, acc_no, acc_flag, trade_budget, logger):
        self.logger = logger
        # 🎯 계좌 정보 및 주문 매니저 초기화
        self.acc_no = acc_no
        self.acc_flag = acc_flag
        self.trade_budget = trade_budget # 🎯 전달받은 예산 저장
        self.om = OrderManager()
        self.am = AccountManager()  # 🎯 추가: 계좌 관리 객체 저장
        
        self.rdm = RealtimeDataManager(callback_func=self.on_realtime_data)
        self.targets = target_list
        
        # 🎯 체결강도 이력 관리를 위한 딕셔너리
        self.prev_strength = {t['code']: 0.0 for t in self.targets}
        
        # 🎯 [전략 수정] 60일 평균 분당 거래량 계산 (390분 기준)
        self.avg_vol_1m = {}
        for t in self.targets:
            # UniverseBuilder에서 넘겨준 avg_vol_60 사용
            self.avg_vol_1m[t['code']] = t.get('avg_vol_60', 0) / 390

        # 실시간 거래량 누적을 위한 윈도우 (최근 10초)
        self.vol_windows = {t['code']: deque() for t in self.targets}
        
        self.buy_signals = {}
        self.sold_codes = set() # 🎯 [추가] 당일 매도 완료 종목 리스트 (재매수 금지용)
        self.positions = {}
        self.max_positions = 10
        self.orderbook_state = {}
        self.is_exiting = {} 
        self.subscribed_codes = set()
        
        # 트레일링 스탑 설정값
        self.ts_activation_pct = 0.5   # 감시 시작 수익률 (0.5% 이상 수익 시 작동)
        self.ts_callback_pct = 0.3     # 하락 허용 폭 (최고가 대비 0.3% 하락 시 매도)
        self.hard_stop_loss = -1.2     # 최소 방어선 (손절선)
        
    def start_subscribing(self):
        """[최적화] cur 모듈 1개만 사용하여 316종목 전체 감시"""
        if not self.targets: return
        
        # 1. 시계열 알림 구독
        self.om.subscribe_conclusion()
        
        # 2. 유니버스 전체 구독 (316 x 1 = 316 < 400 한도 통과)
        for t in self.targets:
            code = t['code']
            self.buy_signals[code] = False
            # 호가창 데이터는 사용하지 않으므로 초기화 생략
            self.rdm.start_monitoring(code, types=['cur']) # 🎯 cur만 사용
            self.subscribed_codes.add(code)
            
        self.logger.info(f"📡 [단일모듈 감시] {len(self.targets)}종목 현재가/체결 데이터 수집 시작")
        
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
        """거래량 폭발 + 체결강도 가속도 기반 분석 (호가 필터 제외)"""
        price = data.get('current', 0)
        # diff = data.get('diff_val', 0) # RealtimeHandler의 필드명 확인
        
        # # [필터 1] 상승 중인 종목인가?
        # if diff <= 0: return
        
        # [필터 2] 거래량 폭발 (10초 누적 -> 1분 예측)
        ten_sec_vol = sum(v for t, v in self.vol_windows[code])
        predicted_1m_vol = ten_sec_vol * 6
        base_vol_1m = self.avg_vol_1m.get(code, 0)
        
        if base_vol_1m <= 0: return
        vol_multiple = predicted_1m_vol / base_vol_1m
        
        # [필터 3] 체결강도 및 가속도 (15배 돌파 & 강도 110% & 가속도 1.5)
        # 호가창 잔량 필터는 제거하여 실행 속도와 종목 수 확보
        if vol_multiple > 15.0 and strength >= 110.0 and accel >= 1.5:
            self.logger.info(f"🔥 [강력 신호] {data.get('name')} | "
                            f"거래폭발:{vol_multiple:.1f}배 | "
                            f"강도:{strength:.1f}%(↑{accel:.1f})")
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

    def on_order_confirmed(self, data):
        """실시간 체결 알림 처리 (부분 체결 대응 버전)"""
        if data['status'] == 'CONCLUDED':
            code = data['stock_code']
            exec_qty = data['volume'] # 이번에 체결된 수량
            exec_price = data['price'] # 이번 체결 가격
            
            if data['side'] == 'BUY':
                if code in self.positions:
                    self.positions[code] = {
                        'name': data['name'],
                        'buy_price': exec_price,
                        'highest_price': exec_price, # 🎯 트레일링 스탑용 최고가 초기화
                        'qty': exec_qty,
                        'entry_time': time.time()
                    }
                    # 🎯 가중 평균 단가 계산 (기존 금액 + 새 금액) / 전체 수량
                    pos = self.positions[code]
                    
                    current_total_cost = pos['buy_price'] * pos['qty']
                    new_fill_cost = exec_price * exec_qty
                    
                    total_qty = pos['qty'] + exec_qty
                    new_avg_price = (current_total_cost + new_fill_cost) / total_qty
                    
                    # 데이터 업데이트
                    pos['qty'] = total_qty
                    pos['buy_price'] = new_avg_price
                    
                    self.logger.info(f"✅ [매수 추가체결] {data['name']} | +{exec_qty}주 | 새 평단가: {new_avg_price:,.0f}원")
                else:
                    # 첫 진입
                    self.positions[code] = {
                        'name': data['name'],
                        'buy_price': exec_price,
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
            self.logger.info(f"💰 [매도 실행] {pos['name']} | {sell_reason} | 현재수익: {current_profit:.2f}%")
            self.is_exiting[code] = True
            self.om.request_new_order(self.acc_no, self.acc_flag, code, pos['qty'], 0, order_type="1", hoga_flag="03")
            
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