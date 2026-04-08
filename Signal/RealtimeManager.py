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
    def __init__(self, target_list, acc_no, acc_flag, trade_budget, logger, acc_manager=None):
        self.cfg = path_finder.get_cfg()
        # ... (기존 초기화 로직 동일) ...
        self.logger = logger
        self.acc_no = acc_no
        self.acc_flag = acc_flag
        self.trade_budget = trade_budget
        
        # 1. OrderManager 생성
        self.om = OrderManager() 
        # 2. 🎯 체결 콜백 연결 (이 부분이 핵심입니다)
        # OrderManager 내부에 콜백을 등록하는 메서드(예: set_callback)가 있다고 가정합니다.
        self.om.set_callback(self.on_order_confirmed)
        
        self.am = acc_manager if acc_manager else AccountManager(acc_no, acc_flag)
        # rdm에 구독하고 데이터가 넘어오면 self.on_realtime_data넣는다
        self.rdm = RealtimeDataManager(callback_func=self.on_realtime_data) 
        
        self.targets = target_list # 실시간 감시목록 대상 
        # print(self.targets)
        
        self.prev_strength = {t['code']: 0.0 for t in self.targets} # 체결강도
        self.avg_vol_1m = {t['code']: t.get('avg_vol_60', 0) / 390 for t in self.targets} # 60일 평균 분당 거래량
        self.vol_windows = {t['code']: deque() for t in self.targets} # 분당 거래량 윈도우
        
        self.buy_signals = {} # 매수신호
        self.sold_codes = set() # 매도 종목코드
        self.positions = {} # 포지션
        self.max_positions = 10 # 최대 포지션 수
        self.orderbook_state = {} # 🎯 호가창 데이터를 담을 딕셔너리
        self.is_exiting = {}  # 매도 진행 중 여부
        self.subscribed_codes = set() #실시간 감시중인 목록
        
        # 설정값 튜닝
        self.ts_activation_pct = 0.5 # 트레일링 스탑
        self.ts_callback_pct = 0.3 # 트레일링 스탑
        self.hard_stop_loss = -1.2 # 손절선
        
        # 🎯 [추가] 체결강도 상한선 (이상 수치 차단)
        self.strength_limit = 1000.0
        
        today_str = datetime.now().strftime("%Y%m%d")
        # 2. 파일 이름에 날짜 포함시키기
        self.trade_summary_path = self.cfg.DATA_DIR / f"trade_summary_{today_str}.jsonl"
            
    def start_subscribing(self):
        """[업데이트] 160여 종목에 대해 현재가와 호가창 동시 감시"""
        if not self.targets: return
        
        self.om.subscribe_conclusion() # 실시간 체결내역
        
        for t in self.targets:
            code = t['code']
            self.buy_signals[code] = False
            # 🎯 'jpbidcnld'(호가창)를 추가하여 160 x 2 = 320개 모듈 사용
            self.rdm.start_monitoring(code, types=['cur', 'jpbidcnld']) 
            self.subscribed_codes.add(code)
            
        self.logger.info(f"📡 [정밀 감시] {len(self.targets)}종목 수급+호가 데이터 분석 시작")
        
    def on_realtime_data(self, data):
        """RealtimeDataManager에서 넘어오는 모든 콜백의 최상위 진입점"""
         # 🎯 'jpbidcnld'(호가창)를 추가하여 160 x 2 = 320개 모듈 사용
        
        # 만약 data가 리스트나 딕셔너리라면 그 안에서 code를 꺼내 씁니다.
        code = data.get('code') # 또는 data[0] 등 데이터 구조에 맞게 수정
        if not code: return   
        
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
        price = data.get('current', 0)
        tick_vol = data.get('tick_vol', 0)
        strength = data.get('strength', 0.0) # 현재 체결강도
        
        # 2. 거래량 누적 (최근 10초 슬라이딩 윈도우)
        now = time.time()
        # 🎯 데이터 구조 변경: (시간, 거래량, 가격)
        self.vol_windows[code].append((now, tick_vol, price))
        
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
        # print(price)
        if price <= 0: return
        
        # 1. 거래량 폭발 계산 (10초 -> 1분 예측)
        ten_sec_vol = sum(v for t, v, p in self.vol_windows[code])
        predicted_1m_vol = ten_sec_vol * 6
        base_vol_1m = self.avg_vol_1m.get(code, 0)
        # print(base_vol_1m)
        if base_vol_1m <= 0: return
        vol_multiple = predicted_1m_vol / base_vol_1m
        
        # 2. 🎯 [다중 시간 가격 상승 체크]
        now = time.time()
        p_1s, p_3s, p_5s = None, None, None
        
        # 역순 탐색으로 각 시점의 가격 추출
        for ts, vol, p in reversed(self.vol_windows[code]):
            diff = now - ts
            if p_1s is None and diff >= 1.0: p_1s = p
            if p_3s is None and diff >= 3.0: p_3s = p
            if p_5s is None and diff >= 5.0: p_5s = p
            if diff > 10.0: break # 10초 넘어가면 중단
        
        # 상승 조건 정의 (하나라도 상승 중이면 True, 상황에 따라 and로 변경 가능)
        # 🎯 추천 로직: 3초 전 대비 상승을 기본으로 하되, 데이터 부족 시 1초 전 확인
        is_price_rising = False
        if p_3s: 
            is_price_rising = price > p_3s 
        elif p_1s: 
            is_price_rising = price > p_1s
        
        # 2. 호가창 필터 확인
        ob = self.orderbook_state.get(code)
        ratio = ob['ratio'] if ob else 0.0
        
        # 🎯 [핵심 필터 적용]
        # - 거래량 15배 돌파
        # - 체결강도 110% 이상이며 1000% 이하 (이상치 제거)
        # - 가속도 1.5 이상
        # - 매도잔량 우위 (호가창 필터)
        # 🎯 [통합 필터] 매수 조건 충족 시 모든 스냅샷 전달
        if (vol_multiple > 15.0 and is_price_rising and 
            110.0 <= strength <= self.strength_limit and 
            accel >= 1.5 and ratio >= 1.2):
            
            self.logger.info(f"🚀 [진짜 수급 포착] {data.get('name')} | "
                            f"폭발:{vol_multiple:.1f}배 | 강도:{strength:.1f}% | "
                            f"현재가:{price:,} | 3초전:{p_3s if p_3s else '미확인'} | "
                            f"호가비율:{ob['ratio']:.2f}")
            
            # 🔥 스냅샷 데이터를 묶어서 전달
            snapshots = {
                'entry_vol_multiple': vol_multiple,
                'entry_strength': strength,
                'entry_ob_ratio': ratio,
                'p_1s': p_1s,
                'p_3s': p_3s,
                'p_5s': p_5s,
                'entry_price': price
            }
            
            # 🔥 [수정] execute_buy 호출 시 멀티플과 체결강도 데이터를 함께 넘김
            self.execute_buy(code, data.get('name'), price, snapshots)
            
    def execute_buy(self, code, name, price, snapshots):
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
                'expected_qty': final_buy_qty, # 🎯 주문 넣은 '목표 수량'
                'qty': 0,                      # 🎯 실제 체결된 수량 (처음엔 0)
                'expected_entry_price': price, # 실제 매수가(buy_price)로 쓸 기준
                'actual_entry_price': price,   # 체결 전까지는 예상가와 동일하게 세팅
                'entry_time': time.time(),
                'max_price': price,
                'min_price': price,
                'is_concluded': False,
                'entry_vol_multiple': round(float(snapshots['entry_vol_multiple']), 2),
                'entry_strength': round(float(snapshots['entry_strength']), 1),
                'entry_ob_ratio': round(float(snapshots['entry_ob_ratio']), 2),
                'p_1s': snapshots['p_1s'],
                'p_3s': snapshots['p_3s'],
                'p_5s': snapshots['p_5s'],
            }
            self.buy_signals[code] = True 
        except Exception as e:
            self.logger.error(f"❌ 주문 중 오류 발생: {e}")

    # RealtimeManager.py 내 on_order_confirmed 수정
    def on_order_confirmed(self, concl_data):
        """OrderManager로부터 실제 체결 정보를 전달받음"""
        code = concl_data['code']
        side = concl_data.get('side') # '1': 매도, '2': 매수 (대신증권 기준 확인 필요)
        
        if code not in self.positions:
            return
        
        pos = self.positions[code]
        
        # --- [CASE 1] 매수 체결 ('2') ---
        if side == '2':
            old_qty = pos['qty']
            new_qty = concl_data['concluded_qty']
            
            # 평단가 가중 평균 계산
            if old_qty + new_qty > 0:
                curr_actual = pos['actual_entry_price'] if old_qty > 0 else float(concl_data['actual_price'])
                new_actual = float(concl_data['actual_price'])
                pos['actual_entry_price'] = ((curr_actual * old_qty) + (new_actual * new_qty)) / (old_qty + new_qty)
            
            pos['qty'] += new_qty
            pos['is_concluded'] = True
            
            # 슬리피지 및 체결 현황 로그
            expected = pos['expected_entry_price']
            actual = pos['actual_entry_price']
            slippage = (actual - expected) / expected * 100
            
            self.logger.info(f"✅ [매수체결] {pos['name']}({code}) {new_qty}주 | "
                            f"누적:{pos['qty']}/{pos['expected_qty']}주 | "
                            f"평단:{actual:,.0f} | 슬리피지:{slippage:.2f}%")

        # --- [CASE 2] 매도 체결 ('1') ---
        elif side == '1':
            self.logger.info(f"✨ [매도체결] {pos['name']}({code}) 전량 체결 완료")
            
            # 교착 상태 방지 플래그 해제
            if code in self.is_exiting:
                del self.is_exiting[code]
            
            # 당일 재매수 방지 목록 추가 및 포지션 제거
            self.sold_codes.add(code)
            del self.positions[code]

                               
    def manage_exit(self, code, curr_price):
        """트레일링 스탑 및 손절 관리"""
        if code not in self.positions:
            return
            
        # 🎯 매수 주문 중이거나 매도 주문이 이미 나간 경우 중복 실행 방지
        if self.is_exiting.get(code, False):
            return

        pos = self.positions[code]
        
        # 🎯 아직 단 1주도 체결되지 않았다면 매도 로직 패스
        if pos.get('qty', 0) == 0:
            return

        buy_price = pos.get('actual_entry_price', pos.get('expected_entry_price', curr_price))
        if 'max_price' not in pos: pos['max_price'] = curr_price
        
        fee_tax_rate = 0.23
        current_profit = ((curr_price - buy_price) / buy_price * 100) - fee_tax_rate
        highest_profit = ((pos['max_price'] - buy_price) / buy_price * 100) - fee_tax_rate

        sell_reason = None
        if highest_profit >= self.ts_activation_pct:
            if current_profit <= (highest_profit - self.ts_callback_pct):
                sell_reason = f"TS(최고 {highest_profit:.2f}% 대비 하락)"
        elif current_profit <= self.hard_stop_loss:
            sell_reason = f"손절(기준선 {self.hard_stop_loss}%)"

        if sell_reason:
            # 🎯 느린 잔고 조회 API 호출 제거. 로컬에 저장된 체결 수량을 즉시 매도!
            sell_qty = pos['qty'] 

            self.logger.info(f"💰 [매도 실행] {pos['name']} | {sell_reason} | 수량: {sell_qty}주")
            self.is_exiting[code] = True
            self.om.request_new_order(self.acc_no, self.acc_flag, code, sell_qty, 0, order_type="1", hoga_flag="03")
            self.write_final_log(code, curr_price, sell_reason)
    
    def sync_balance_with_server(self):
        """서버의 실제 잔고를 가져와 로컬 positions 동기화 (기존 데이터 보존)"""
        # 🎯 메서드명 수정 및 리스트 처리 로직으로 변경
        _, stocks = self.am.get_balance_data()
        
        server_codes = set() # 서버 목록
        for s in stocks:
            code = s['code']
            server_codes.add(code)
            
            if code in self.positions:
                # 봇이 이미 추적 중이면 '수량'만 몰래 갱신 (max_price 등은 보존)
                if self.positions[code]['qty'] != s['total_qty']:
                    self.logger.warning(f"⚠️ [잔고 불일치] {s['name']}: {self.positions[code]['qty']} -> {s['total_qty']}")
                    self.positions[code]['qty'] = s['total_qty']
            else:
                # 봇이 모르는 종목이 서버에 있으면 봇 포맷(모든 키 포함)에 맞춰 새로 등록
                self.positions[code] = {
                    'signal_id': f"SYNC_{code}",
                    'name': s['name'],
                    'qty': s['total_qty'],
                    'expected_entry_price': s['buy_price'],
                    'actual_entry_price': s['buy_price'],
                    'buy_price': s['buy_price'], 
                    'max_price': s['buy_price'], # 현재가가 없으므로 매입가로 대체
                    'min_price': s['buy_price'],
                    'entry_time': time.time(),
                    'is_concluded': True
                }
        # self.logger.info("🔄 [시스템] 실잔고 동기화 완료")


    def write_final_log(self, code, exit_price, reason):
        """매매 종료 시 최종 결과 JSONL 기록 (데이터 타입 보존)"""
        pos = self.positions.get(code)
        if not pos: return

        # 1. 수치 계산
        entry_price = pos.get('actual_entry_price', pos['expected_entry_price'])
        ret_rate = (exit_price - entry_price) / entry_price
        mae = (pos['min_price'] - entry_price) / entry_price
        mfe = (pos['max_price'] - entry_price) / entry_price
        duration = time.time() - pos['entry_time']
        slippage = (entry_price / pos['expected_entry_price']) - 1

        # 2. 저장할 데이터 구성 (딕셔너리 구조 유지)
        trade_data = {
            "date": datetime.now().strftime("%Y-%m-%d"), # 날짜 추가
            "signal_id": pos['signal_id'],
            "code": code,
            "name": pos['name'],
            # 🔥 진입 당시 수급/가격 스냅샷 기록
            "entry_ob_ratio": pos.get('entry_ob_ratio'),
            "entry_v_mult": pos.get('entry_vol_multiple'),
            "entry_str": pos.get('entry_strength'),
            "p_diff_1s": round((pos['expected_entry_price'] - pos['p_1s'])/pos['p_1s']*100, 2) if pos['p_1s'] else 0,
            "p_diff_3s": round((pos['expected_entry_price'] - pos['p_3s'])/pos['p_3s']*100, 2) if pos['p_3s'] else 0,
            "p_diff_5s": round((pos['expected_entry_price'] - pos['p_5s'])/pos['p_5s']*100, 2) if pos['p_5s'] else 0,
            "expected_entry": int(pos['expected_entry_price']),
            "actual_entry": float(entry_price),
            "slippage_rate": round(float(slippage * 100), 4),
            "exit_price": int(exit_price),
            "return_rate": round(float(ret_rate * 100), 2),
            "mae": round(float(mae * 100), 2),
            "mfe": round(float(mfe * 100), 2),
            "reason": str(reason),
            "hold_duration": round(float(duration), 1)
        }

        # 3. JSONL 형식으로 한 줄씩 추가 저장 (utf-8)
        # 파일 확장자를 .jsonl로 변경하는 것이 좋습니다.
        jsonl_path = self.trade_summary_path.with_suffix('.jsonl') 
        
        with open(jsonl_path, 'a', encoding='utf-8') as f:
            # ensure_ascii=False를 해야 한글 종목명이 깨지지 않고 저장됩니다.
            f.write(json.dumps(trade_data, ensure_ascii=False) + '\n')
        
        self.logger.info(f"📊 {pos['name']} 매도 완료 ({reason}) | 수익률: {ret_rate:.2%}")
        
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
    
## TODO : 실전 가동 전 마지막 팁 (로직상 조언)
# 단, 가동 후 쌓이는 로그(Data/trade_summary.csv)를 보실 때 이것 하나만 꼭 확인해 보세요.
# 부분 체결 후 급락 상황: 만약 100주 매수 주문을 넣었는데 10주만 체결된 상태에서 주가가 -1.2% 하락하면, 현재 로직은 가진 10주를 즉시 손절합니다.
# 그런데 대신증권 서버에는 아직 "나머지 90주 매수 대기(미체결)" 주문이 살아있을 수 있습니다. 
# 손절이 자주 나가는 하락장이라면, 추후 매도 주문을 넣기 전에 "해당 종목의 미체결 매수 주문 취소(CancelOrder)" 로직을 덧붙이는 것을 고려해 볼 수 있습니다.
# 멀티플 차등 적용 연구
# 재매수 로직 연구
