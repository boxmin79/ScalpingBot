import win32com.client
import pythoncom

class RealtimeHandler:
    def set_params(self, obj, name, callback):
        self.obj = obj
        self.name = name
        self.callback = callback

    def OnReceived(self):
        # 공통 헤더: 종목코드
        code = self.obj.GetHeaderValue(0)
        data = {'type': self.name, 'code': code}
        
        # 1. StockCur: 실시간 체결 및 현재가 정보 (가장 빈번함)
        if self.name == 'cur':
            data.update({
                'time': self.obj.GetHeaderValue(1),      # 시간 (HHMMSS)
                'diff_stat': self.obj.GetHeaderValue(2), # 대비부분 (1:상한, 2:상승...)
                'prev_close': self.obj.GetHeaderValue(3),# 전일종가
                'open': self.obj.GetHeaderValue(4),      # 시가
                'high': self.obj.GetHeaderValue(5),      # 고가
                'low': self.obj.GetHeaderValue(6),       # 저가
                'ask': self.obj.GetHeaderValue(7),       # 매도호가
                'bid': self.obj.GetHeaderValue(8),       # 매수호가
                'cum_vol': self.obj.GetHeaderValue(9),   # 누적거래량
                'cum_amt': self.obj.GetHeaderValue(10),  # 누적거래대금
                'current': self.obj.GetHeaderValue(13),  # 현재가
                'diff': self.obj.GetHeaderValue(14),     # 대비
                'tick_vol': self.obj.GetHeaderValue(17), # 순간 체결수량 (스캘핑 핵심)
                'side': chr(self.obj.GetHeaderValue(19)),# 체결구분 ('1':매도, '2':매수)
                'strength': self.obj.GetHeaderValue(24), # 체결강도
                'market_stat': chr(self.obj.GetHeaderValue(20)), # 장구분 ('1':동시호가, '2':장중)
            })

        # 2. StockJpBid / StockJpBid2: 실시간 10단계 호가 잔량
        elif self.name in ['hoga', 'hoga2']:
            data.update({
                'time': self.obj.GetHeaderValue(1),
                'total_ask_vol': self.obj.GetHeaderValue(5), # 총매도잔량
                'total_bid_vol': self.obj.GetHeaderValue(6), # 총매수잔량
            })
            # 10단계 호가 및 잔량 데이터 추출
            asks, bids, ask_vols, bid_vols = [], [], [], []
            for i in range(10):
                asks.append(self.obj.GetDataValue(0, i))     # 매도호가
                ask_vols.append(self.obj.GetDataValue(1, i)) # 매도잔량
                bids.append(self.obj.GetDataValue(2, i))     # 매수호가
                bid_vols.append(self.obj.GetDataValue(3, i)) # 매수잔량
            
            data.update({'asks': asks, 'ask_vols': ask_vols, 'bids': bids, 'bid_vols': bid_vols})
            
            if self.name == 'hoga2': # hoga2에만 있는 추가 정보
                data.update({
                    'total_ask_diff': self.obj.GetHeaderValue(7), # 총매도잔량 전일대비
                    'total_bid_diff': self.obj.GetHeaderValue(8), # 총매수잔량 전일대비
                })

        # 3. StockJpBidCnld: 실시간 호가 잔량 변화 (취소/정정 실시간 추적)
        elif self.name == 'hoga_cnld':
            data.update({
                'time': self.obj.GetHeaderValue(1),
                'side': chr(self.obj.GetHeaderValue(2)), # 1:매도, 2:매수
                'price': self.obj.GetHeaderValue(3),    # 변화 발생 호가
                'vol': self.obj.GetHeaderValue(4),      # 변화 수량
                'change_type': self.obj.GetHeaderValue(5) # 1:신규/정정, 2:취소 (허매수 감지용)
            })

        # 4. StockMemberCnld: 실시간 거래원 매매 정보
        elif self.name == 'member':
            # 거래원별 매수/매도 상위 데이터 (보통 5개 증권사)
            members = []
            for i in range(5):
                members.append({
                    'rank': i + 1,
                    'name': self.obj.GetDataValue(0, i),     # 증권사명
                    'buy_qty': self.obj.GetDataValue(1, i),  # 매수수량
                    'sell_qty': self.obj.GetDataValue(2, i)  # 매도수량
                })
            data.update({'members': members})

        # 5. StockBsccnsCnld: 실시간 기세 및 체결 (프로그램 매매/대량 체결 등)
        elif self.name == 'basis':
            data.update({
                'price': self.obj.GetHeaderValue(2),     # 체결가/기세가
                'vol': self.obj.GetHeaderValue(3),       # 체결량
                'time': self.obj.GetHeaderValue(4),      # 시간
                'side': chr(self.obj.GetHeaderValue(5))  # 체결구분
            })

        # 6. StockJpBidCnld (Dscbo1): 호가잔량 체결 (요청하신 리스트 중 하나)
        # ※ StockJpBidCnld와 기능이 유사하나 기초자산별 특성에 따라 사용
        elif self.name == 'hoga_exec':
            data.update({
                'time': self.obj.GetHeaderValue(1),
                'price': self.obj.GetHeaderValue(2),
                'vol': self.obj.GetHeaderValue(3),
                'type': self.obj.GetHeaderValue(4) # 체결/잔량변화 구분
            })

        # 최종 콜백 실행
        if self.callback:
            self.callback(data)

class RealtimeDataManager:
    def __init__(self, callback_func=None):
        self.callback = callback_func
        # 사용 가능한 모든 실시간 모듈 맵핑
        self.module_map = {
            'cur': "DsCbo1.StockCur",           # 체결 시세
            'hoga': "DsCbo1.StockJpBid",        # 호가 잔량
            'hoga2': "DsCbo1.StockJpBid2",      # 호가 잔량2 (더 상세)
            'hoga_cnld': "DsCbo1.StockJpBidCnld",# 호가 잔량 변화
            'member': "DsCbo1.StockMemberCnld", # 거래원 체결
            'basis': "DsCbo1.StockBsccnsCnld"   # 기세/체결 상세
        }
        self.active_subscriptions = {}

    def start_monitoring(self, code, types=['cur', 'hoga2']):
        """종목별 원하는 모듈을 골라서 구독 시작"""
        for t in types:
            if t not in self.module_map: continue
            
            key = f"{code}_{t}"
            if key in self.active_subscriptions: continue

            obj = win32com.client.Dispatch(self.module_map[t])
            obj.SetInputValue(0, code)
            
            handler = win32com.client.WithEvents(obj, RealtimeHandler)
            handler.set_params(obj, t, self.callback)
            obj.Subscribe()
            
            self.active_subscriptions[key] = (obj, handler)
        print(f"🚀 [실시간 모니터링 시작] {code} -> {types}")

    def stop_monitoring(self, code):
        """특정 종목의 모든 실시간 구독 해지"""
        keys_to_del = [k for k in self.active_subscriptions.keys() if k.startswith(code)]
        for k in keys_to_del:
            obj, _ = self.active_subscriptions[k]
            obj.Unsubscribe()
            del self.active_subscriptions[k]
        print(f"🛑 [실시간 모니터링 종료] {code}")

    def stop_all(self):
        """프로그램 종료 시 전체 구독 해지"""
        for k in list(self.active_subscriptions.keys()):
            obj, _ = self.active_subscriptions[k]
            obj.Unsubscribe()
        self.active_subscriptions.clear()