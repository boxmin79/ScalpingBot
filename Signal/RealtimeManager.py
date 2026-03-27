import path_finder
import time
import pythoncom
from datetime import datetime
from collections import deque # 상단에 추가
# API 폴더의 기능 임포트 (체결 & 호가)
from API.CpStockCur import CpStockCur
from API.CpStockJpBid import CpStockJpBid

# -----------------------------------------------------------------------------
# 1. 실시간 [호가창] 감시 래퍼 (CpStockJpBid 상속)
# -----------------------------------------------------------------------------
class ScalpingStockJpBid(CpStockJpBid):
    def __init__(self, code, name, engine):
        super().__init__()
        self.code = code
        self.name = name
        self.engine = engine

    def process_received(self):
        try:
            # [수정] 총 매도/매수 잔량 데이터 추가 추출
            sell_tot = self.obj.GetHeaderValue(2) # 총 매도 잔량
            buy_tot = self.obj.GetHeaderValue(3)  # 총 매수 잔량

            # 1. 1단계 스프레드(촘촘함) 확인
            ask_1 = self.obj.GetDataValue(0, 0) # 최우선 매도호가
            bid_1 = self.obj.GetDataValue(1, 0) # 최우선 매수호가
            spread = ask_1 - bid_1
            
            # 2. 상위 5단계 매도 잔량 밀도 확인
            is_dense = True
            for i in range(5):
                vol = self.obj.GetDataValue(2, i)
                if vol < 100: 
                    is_dense = False
                    break
            
            # [수정] 모든 호가창 정보를 엔진에 통합 전달
            self.engine.update_orderbook(self.code, sell_tot, buy_tot, spread, is_dense)
            
        except Exception as e:
            pass


# -----------------------------------------------------------------------------
# 2. 실시간 [체결창] 감시 래퍼 (CpStockCur 상속)
# -----------------------------------------------------------------------------
class ScalpingStockCur(CpStockCur):
    def __init__(self, code, name, engine):
        super().__init__()
        self.code = code
        self.name = name
        self.engine = engine
        self.prev_strength = 0.0  # 이전 틱의 체결강도 저장용

    def process_received(self):
        try:
            # 1. 원천 데이터 추출
            acc_sell_vol = self.obj.GetHeaderValue(15) # 누적 매도 체결량
            acc_buy_vol = self.obj.GetHeaderValue(16)  # 누적 매수 체결량
            
            # 2. 현재 체결강도 계산
            if acc_sell_vol > 0:
                curr_strength = (acc_buy_vol / acc_sell_vol) * 100
            else:
                curr_strength = 100.0

            # 3. 체결강도 기울기(Delta) 계산
            # 첫 데이터 수신 시에는 기울기를 0으로 처리
            delta_strength = 0.0
            if self.prev_strength > 0:
                delta_strength = curr_strength - self.prev_strength
            
            # 다음 계산을 위해 현재값을 이전값으로 저장
            self.prev_strength = curr_strength

            # 4. 기타 필요 데이터 추출
            time_str = self.obj.GetHeaderValue(18)
            price = self.obj.GetHeaderValue(13)
            vol = self.obj.GetHeaderValue(17)
            buy_sell = self.obj.GetHeaderValue(14)

            # 엔진으로 현재 강도와 기울기(Delta)를 함께 전달
            self.engine.analyze_tick(
                self.code, self.name, time_str, price, vol, 
                buy_sell, curr_strength, delta_strength
            )
            
        except Exception as e:
            print(f"❌ [데이터 처리 오류] {self.name}: {e}")


# -----------------------------------------------------------------------------
# 3. 실시간 감시 엔진 메인 클래스 (RealtimeManager)
# -----------------------------------------------------------------------------
class RealtimeManager:
    def __init__(self, target_list):
        self.targets = target_list
        self.subscribers = []        # 체결 및 호가 구독 객체 보관 (GC 방지)
        self.buy_signals = {}        # 중복 매수 방지
        
        # 💡 각 종목의 실시간 호가 상태를 기억하는 메모리 (Dict)
        self.orderbook_state = {}    
        # 💡 [추가] 종목별 최근 순간체결량 이력 저장 (최근 20틱)
        self.volume_history = {t['code']: deque(maxlen=20) for t in self.targets}

    def start_monitoring(self):
        if not self.targets:
            return

        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 🚀 3단계: 호가+체결 하이브리드 감시 가동 ({len(self.targets)}종목)\n")
        
        for t in self.targets:
            code = t['code']
            name = t['name']
            
            self.buy_signals[code] = False
            # [수정] 상태 구조 초기화 (누락 방지)
            self.orderbook_state[code] = {'sell_tot': 0, 'buy_tot': 0, 'spread': 9999, 'is_dense': False}
            
            # 1. 호가창(JpBid) 구독
            bid_obj = ScalpingStockJpBid(code=code, name=name, engine=self)
            bid_obj.subscribe(code)
            self.subscribers.append(bid_obj)
            
            # 2. 체결창(Cur) 구독
            cur_obj = ScalpingStockCur(code=code, name=name, engine=self)
            cur_obj.subscribe(code)
            self.subscribers.append(cur_obj)
            
        print("-" * 65)
        print("👀 [호가+체결] 입체적 감시 중... (매도벽 돌파 & 대량 체결 대기)")
        print("   [종료: Ctrl+C]")
        print("-" * 65)

        try:
            while True:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.01)
        except KeyboardInterrupt:
            self.stop_monitoring()

    def update_orderbook(self, code, sell_tot, buy_tot, spread, is_dense):
        """[수정] 호가창의 모든 지표를 한 번에 업데이트"""
        self.orderbook_state[code] = {
            'sell_tot': sell_tot,
            'buy_tot': buy_tot,
            'spread': spread,
            'is_dense': is_dense
        }
    
    def get_tick_size(self, price):
        """[추가] 한국 주식시장 호가 단위 계산 (스프레드 촘촘함 판별용)"""
        if price < 2000: return 1
        if price < 5000: return 5
        if price < 20000: return 10
        if price < 50000: return 50
        if price < 200000: return 100
        if price < 500000: return 500
        return 1000
        
    def analyze_tick(self, code, name, time_str, price, vol, buy_sell, strength, delta):
        """
        [최종 타점 분석] 체결 방아쇠가 당겨지면, 호가창 상태를 확인하고 매수합니다.
        """
        if self.buy_signals.get(code, False):
            return
        
        # 1. [데이터 업데이트] 최근 거래량 이력에 현재 체결량 추가
        history = self.volume_history.get(code)
        if history is None: return
        
        # 이전 평균 계산 (최소 5개 이상의 데이터가 쌓였을 때부터 작동)
        avg_vol = sum(history) / len(history) if len(history) >= 5 else 9999999
        history.append(vol) # 현재 체결량을 이력에 추가
        
        # --- [데이터 검증] 현재 호가창 상태 로드 ---
        state = self.orderbook_state.get(code)
        if not state or state['buy_tot'] == 0: return

        # --- [필터 1] 호가창 촘촘함 & 밀도 체크 ---
        # 스프레드가 1틱 이내여야 하고, 매도호가가 비어있지 않아야 함
        if state['spread'] > self.get_tick_size(price) or not state['is_dense']:
            return
        
        # --- [필터 2] 체결강도 및 가속도 체크 ---
        STRENGTH_THRESHOLD = 110.0
        ACCEL_THRESHOLD = 2.0  # 틱당 2%p 이상 급증 시

        if strength < STRENGTH_THRESHOLD or delta < ACCEL_THRESHOLD:
            return
        
        # 4. [필터 3] ★ 상대적 거래량 급증 체크 ★
        # 고정 수치(3000주)와 상대적 수치(평균 대비 5배)를 동시에 만족해야 함
        #보수적 매매: 10.0 (평균의 10배 이상 터질 때만 진입)
        # 공격적 매매: 3.0 (평균의 3배만 터져도 진입)
        VOLUME_MULTIPLIER = 5.0 # 평균 대비 5배
        
        is_volume_spike = (vol >= 3000) and (vol >= avg_vol * VOLUME_MULTIPLIER)
        
        # --- [필터 3] 대량 시장가 매수 및 매도벽 비율 체크 ---
        if buy_sell == ord('1') and is_volume_spike:
            sell_tot = state['sell_tot']
            buy_tot = state['buy_tot']
            
            # 매도 잔량이 매수 잔량보다 2배 이상 (강한 매도벽 존재)
            if sell_tot >= (buy_tot * 2):
                print(f"🔥 [매수 시그널] {time_str} | {name}({code}) | {price:,}원")
                print(f"   📊 [강도] {strength:.1f}% (▲{delta:.1f}) | [체결] {vol:,}주")
                print(f"   🧱 [벽] 매도 {sell_tot:,} vs 매수 {buy_tot:,} ({round(sell_tot/buy_tot, 1)}배)")
                
                self.buy_signals[code] = True
                
                # TODO: 주문 모듈(API/CpOrderManager.py) 호출
                
    def stop_monitoring(self):
        print("\n[시스템] 실시간 감시 해제 중...")
        for obj in self.subscribers:
            obj.unsubscribe()
        self.subscribers.clear()


if __name__ == "__main__":
    sample_targets = [
        {'code': 'A047040', 'name': '대우건설'},
        {'code': 'A263750', 'name': '펄어비스'}
    ]
    manager = RealtimeManager(sample_targets)
    manager.start_monitoring()