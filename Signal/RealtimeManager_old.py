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
        
        # 상태 관리 변수
        self.buy_signals = {}      # 신호 발생 여부
        self.positions = {}        # 🎯 현재 보유 종목 상태 {code: {'buy_price': 0, 'qty': 0}}
        self.max_positions = 10
        self.orderbook_state = {}
        self.volume_history = {t['code']: deque(maxlen=20) for t in self.targets}
        self.prev_strength = {t['code']: 0.0 for t in self.targets}
        
        self.is_exiting = {} # 🎯 [추가] 매도 진행 상태 체크용
        self.subscribed_codes = set() # 🎯 [추가] 현재 구독 중인 종목 코드를 직접 관리!
        

    def start_subscribing(self):
        """TradingBot에서 호출: 구독 시작"""
        if not self.targets and not self.positions: return
        print(f"   📡 {len(self.targets)}종목 실시간 감시 및 매매 엔진 가동... 현재시간")
        
        # 🎯 [추가] 보유 종목(positions)이 targets에 없다면 감시 대상에 강제 추가
        target_codes = [t['code'] for t in self.targets]
        for code, info in self.positions.items():
            if code not in target_codes:
                self.targets.append({'code': code, 'name': info['name']})
                self.volume_history[code] = deque(maxlen=20)
                self.prev_strength[code] = 0.0
                print(f"   📌 보유종목 감시 유지: {info['name']}({code})")

        print(f"   📡 {len(self.targets)}종목 실시간 감시 및 매매 엔진 가동...")
        self.om.subscribe_conclusion()
        
        for t in self.targets:
            code = t['code']
            self.buy_signals[code] = False
            self.orderbook_state[code] = {'total_ask_vol': 0, 'total_bid_vol': 0, 'spread': 9999, 'is_dense': False}
            self.rdm.start_monitoring(code, types=['cur', 'jpbidcnld'])
            
            # 🎯 [추가] 내 구독 리스트에도 등록
            self.subscribed_codes.add(code)

    def on_realtime_data(self, data):
        """데이터 수신 콜백"""
        # 🎯 마감 시간 체크 (예: 오후 3시 20분부터 강제 청산) 중복기능으로 주석처리
        # now = datetime.now().time()
        # if now >= datetime.strptime("15:20:00", "%H:%M:%S").time():
        #     self.force_exit_all()
        #     return # 마감 이후에는 신규 분석 중단
        
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
        
        if len(self.positions) >= self.max_positions:
            return
        
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
        # print(f"code: {code} | {state['spread']} | {state['is_dense']} | {strength} | {round(delta,2)} | {vol} | {round(avg_vol, 2)}")
        if state['spread'] <= self.get_tick_size(price) and state['is_dense']:
            if strength >= 110.0 and delta >= 1.0:
                if buy_sell == '1' and (vol >= 1000 and vol >= avg_vol * 5.0):
                    if state['total_ask_vol'] >= (state['total_bid_vol'] * 1.5):
                        # 🎯 실제 매수 주문 집행
                        self.execute_buy(code, data.get('name'), price)

    def execute_buy(self, code, name, price):
        """매수 주문 실행 (현금 100% 수량 검증 버전)"""
        # 1. 예산 및 종목 수 방어
        if price == 0 or self.trade_budget <= 0 or len(self.positions) >= self.max_positions:
            return

        # 2. AccountManager를 통해 실제 현금 매수 가능 수량 조회
        # 제공해주신 AccountManager.get_buyable_data 사용
        buyable_info = self.am.get_buyable_data(code, price=price, quote_type='03', query_type='2')
        
        if not buyable_info:
            self.logger.error(f"❌ [{name}] 매수 가능 수량 조회 실패")
            return

        # 🎯 3. 실제 내 현금으로만 살 수 있는 수량 (오타 반영: cach_buyable_qty)
        max_cash_qty = buyable_info.get('cash_buyable_qty', 0)
        
        # 🎯 4. 봇의 예산 설정값(1/10)과 실제 현금 가능 수량 중 작은 값을 선택
        budget_qty = int(self.trade_budget // price)
        final_buy_qty = min(budget_qty, max_cash_qty)

        if final_buy_qty <= 0:
            self.logger.info(f"⚠️ [매수 스킵] {name} | 현금 잔고 부족 (가능수량: {max_cash_qty}주)")
            return

        self.logger.info(f"🔥 [매수 요청] {name}({code}) | 수량: {final_buy_qty}주 (현용한도: {max_cash_qty}주)")
        
        # 시장가(03) 주문 실행 (미수 방지를 위해 수량 엄격 제한)
        self.om.request_new_order(self.acc_no, self.acc_flag, code, final_buy_qty, 0, order_type="2", hoga_flag="03")

    # RealtimeManager.py 내 on_order_confirmed 메서드 수정

    def on_order_confirmed(self, data):
        """실시간 체결 알림 처리 (부분 체결 대응 버전)"""
        if data['status'] == 'CONCLUDED':
            code = data['stock_code']
            exec_qty = data['volume'] # 이번에 체결된 수량
            exec_price = data['price'] # 이번 체결 가격
            
            if data['side'] == 'BUY':
                if code in self.positions:
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
                        self.logger.info(f"🏁 [매도 완료] {data['name']} 포지션 종료")
                        del self.positions[code]
                        if code in self.is_exiting:
                            del self.is_exiting[code]
                    
    def manage_exit(self, code, curr_price):
        """실시간 익절/손절 관리 (수수료 및 세금 0.23% 반영 버전)"""
        if code not in self.positions or self.is_exiting.get(code, False):
            return

        pos = self.positions[code]
        
        # 1. 제비용 설정 (수수료 왕복 + 세금)
        # 일반적인 온라인 수수료(0.015% * 2) + 거래세(0.18%~0.2%) ≈ 0.23%
        fee_tax_rate = 0.23 

        # 2. 단순 수익률 계산
        raw_profit_rate = (curr_price - pos['buy_price']) / pos['buy_price'] * 100
        
        # 3. 실질 수익률 계산 (단순 수익률 - 제비용)
        net_profit_rate = raw_profit_rate - fee_tax_rate

        # 4. 매도 판단 (실질 수익률 기준)
        # 실질적으로 내 주머니에 1.5%가 남을 때 익절, -1.0%일 때 손절
        if net_profit_rate >= 1.5 or net_profit_rate <= -1.0:
            self.logger.info(f"💰 [매도 실행] {pos['name']} | 실질수익: {net_profit_rate:.2f}% (수수료차감 전: {raw_profit_rate:.2f}%)")
            
            self.is_exiting[code] = True
            # 시장가(03) 매도 주문
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

    def update_targets(self, new_target_list):
        """TradingBot에서 호출: 기존 포지션은 유지하며 감시 대상 종목만 교체"""
        print(f"🔄 감시 종목 업데이트 중... (기존 포지션 유지)")

        # 🎯 1. 수정: rdm.monitored_codes 대신 직접 만든 subscribed_codes 사용!
        current_subscribed_codes = set(self.subscribed_codes) 
        
        # 2. 새로운 타깃 코드 목록 (보유 종목 포함)
        new_target_codes = {t['code'] for t in new_target_list}
        holding_codes = set(self.positions.keys())
        
        # 최종 유지해야 할 전체 코드 (타깃 + 현재 보유 중인 종목)
        final_codes = new_target_codes | holding_codes
        
        # 3. 삭제 대상: 현재 구독 중이지만 final_codes에 없는 종목
        to_unsubscribe = current_subscribed_codes - final_codes
        for code in to_unsubscribe:
            self.rdm.stop_monitoring(code)
            self.subscribed_codes.discard(code) # 🎯 내 리스트에서도 삭제
            if code in self.buy_signals: del self.buy_signals[code]
            
        # 4. 추가 대상: final_codes에 있지만 아직 구독 중이 아닌 종목
        to_subscribe = final_codes - current_subscribed_codes
        for code in to_subscribe:
            # 종목명 찾기 (new_target_list 또는 positions에서)
            name = next((t['name'] for t in new_target_list if t['code'] == code), "Unknown")
            if name == "Unknown" and code in self.positions:
                name = self.positions[code]['name']
                
            # 상태 변수 초기화
            self.buy_signals[code] = False
            self.orderbook_state[code] = {'total_ask_vol': 0, 'total_bid_vol': 0, 'spread': 9999, 'is_dense': False}
            if code not in self.volume_history:
                self.volume_history[code] = deque(maxlen=20)
            
            # 실제 구독 시작
            self.rdm.start_monitoring(code, types=['cur', 'jpbidcnld'])
            self.subscribed_codes.add(code) # 🎯 [추가] 새 종목도 내 리스트에 등록
            print(f"   ➕ 새 감시 추가: {name}({code})")

        # 5. 내부 targets 리스트 갱신
        self.targets = new_target_list