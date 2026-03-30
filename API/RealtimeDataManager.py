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
            # 매수/매도 누적 수량 가져오기
            vol_sell = self.obj.GetHeaderValue(15) # 누적매도체결수량
            vol_buy = self.obj.GetHeaderValue(16)  # 누적매수체결수량
            # 🎯 체결강도 계산 (분모가 0인 경우 예외 처리)
            strength = (vol_buy / vol_sell * 100) if vol_sell > 0 else 0
            
            data.update({
                'code': self.obj.GetHeaderValue(0),      # 종목코드
                'name': self.obj.GetHeaderValue(1),      # 종목명
                'diff_val': self.obj.GetHeaderValue(2),  # 전일대비 금액
                'time': self.obj.GetHeaderValue(3),      # 시간 (HHMM)
                'open': self.obj.GetHeaderValue(4),      # 시가
                'high': self.obj.GetHeaderValue(5),      # 고가
                'low': self.obj.GetHeaderValue(6),       # 저가
                'ask': self.obj.GetHeaderValue(7),       # 매도호가
                'bid': self.obj.GetHeaderValue(8),       # 매수호가
                'cum_vol': self.obj.GetHeaderValue(9),   # 누적거래량
                'cum_amt': self.obj.GetHeaderValue(10),  # 누적거래대금
                'current': self.obj.GetHeaderValue(13),  # 현재가 (또는 예상체결가)
                'side': chr(self.obj.GetHeaderValue(14)),# 체결상태 ('1':매수, '2':매도)
                'tick_vol': self.obj.GetHeaderValue(17), # 순간 체결수량
                'time_sec': self.obj.GetHeaderValue(18), # 시간(초)
                'expect_flag': chr(self.obj.GetHeaderValue(19)), # 예상체결가 구분 ('1':동시호가, '2':장중)
                'market_stat': chr(self.obj.GetHeaderValue(20)), # 장구분플래그 ('1':장전, '2':장중...)
                'diff_stat': chr(self.obj.GetHeaderValue(22)),   # 대비부호 ('1':상한, '2':상승...)
                'lp_rate': self.obj.GetHeaderValue(25),  # LP보유율 추가 권장
                'vol_sell': vol_sell,
                'vol_buy': vol_buy,
                'strength': round(strength, 2), # 소수점 2자리까지 반올림
            })

        # 2. StockJpBid: 실시간 10단계 호가 잔량
        elif self.name == 'jpbid':
            # 기본 정보 추출 (인덱스 수정)
            data.update({
                'time': self.obj.GetHeaderValue(1),           # 시간
                'total_ask_vol': self.obj.GetHeaderValue(23),  # 총매도잔량 (기존 5번은 1차잔량임)
                'total_bid_vol': self.obj.GetHeaderValue(24),  # 총매수잔량 (기존 6번은 1차잔량임)
                'side_over_ask': self.obj.GetHeaderValue(25),  # 시간외 총매도잔량
                'side_over_bid': self.obj.GetHeaderValue(26),  # 시간외 총매수잔량
            })

            # 10단계 호가 및 잔량 데이터 추출 (GetHeaderValue 사용 필수)
            asks, bids, ask_vols, bid_vols = [], [], [], []
            
            for i in range(10):
                # 명세서 규칙: 1~5차(인덱스 3~22), 6~10차(인덱스 27~46)
                if i < 5:
                    base = 3 + (i * 4)
                else:
                    base = 27 + ((i - 5) * 4)
                
                asks.append(self.obj.GetHeaderValue(base))      # 매도호가
                bids.append(self.obj.GetHeaderValue(base + 1))  # 매수호가
                ask_vols.append(self.obj.GetHeaderValue(base + 2)) # 매도잔량
                bid_vols.append(self.obj.GetHeaderValue(base + 3)) # 매수잔량
            
            data.update({
                'asks': asks, 
                'bids': bids, 
                'ask_vols': ask_vols, 
                'bid_vols': bid_vols,
                'mid_price': self.obj.GetHeaderValue(69) # 중간가격
            })

        # 3. StockJpBidCnld: 통합(KRX+NXT) 10단계 호가 잔량 실시간 수신
        elif self.name == 'jpbidcnld':
            # 1. 공통 헤더 정보 추출
            data.update({
                'time': self.obj.GetHeaderValue(1),           # 시간
                'total_volume': self.obj.GetHeaderValue(2),   # 거래량
                'total_ask_vol': self.obj.GetHeaderValue(23),  # 통합 총 매도잔량
                'total_bid_vol': self.obj.GetHeaderValue(24),  # 통합 총 매수잔량
                
                # KRX vs NXT 총잔량 비교 (차익거래/수급 분석용)
                'krx_total_ask': self.obj.GetHeaderValue(89),  # KRX 총 매도잔량
                'krx_total_bid': self.obj.GetHeaderValue(90),  # KRX 총 매수잔량
                'nxt_total_ask': self.obj.GetHeaderValue(114), # NXT 총 매도잔량
                'nxt_total_bid': self.obj.GetHeaderValue(115), # NXT 총 매수잔량
            })

            # 2. 통합 10단계 호가 및 잔량 추출 루프
            asks, bids, ask_vols, bid_vols = [], [], [], []
            for i in range(10):
                # 인덱스 규칙: 1~5차(3~22), 6~10차(27~46)
                base = (3 + i * 4) if i < 5 else (27 + (i - 5) * 4)
                
                asks.append(self.obj.GetHeaderValue(base))
                bids.append(self.obj.GetHeaderValue(base + 1))
                ask_vols.append(self.obj.GetHeaderValue(base + 2))
                bid_vols.append(self.obj.GetHeaderValue(base + 3))

            data.update({
                'asks': asks, 'bids': bids, 
                'ask_vols': ask_vols, 'bid_vols': bid_vols
            })

            # 3. LP 잔량 합계 (ETF/ELW 매매 시 필수)
            data.update({
                'lp_total_ask': self.obj.GetHeaderValue(67),
                'lp_total_bid': self.obj.GetHeaderValue(68)
            })

        # 4. StockMember: 실시간 거래원 매매 정보 (매도/매수 상위 5개사)
        elif self.name == 'member':
            # 헤더에서 실제 데이터 개수(보통 5개) 확인
            count = self.obj.GetHeaderValue(1)
            
            sell_list = [] # 매도 상위 거래원 정보
            buy_list = []  # 매수 상위 거래원 정보
            
            for i in range(count):
                # 매도 거래원 추출
                sell_list.append({
                    'rank': i + 1,
                    'name': self.obj.GetDataValue(0, i),    # 매도거래원 명
                    'qty': self.obj.GetHeaderValue(2) if i==0 else self.obj.GetDataValue(2, i) 
                    # 주의: 명세서상 인덱스 2번은 총매도수량입니다.
                })
                
                # 매수 거래원 추출
                buy_list.append({
                    'rank': i + 1,
                    'name': self.obj.GetDataValue(1, i),    # 매수거래원 명
                    'qty': self.obj.GetDataValue(3, i)      # 총매수수량
                })
            
            data.update({
                'time': self.obj.GetHeaderValue(2),         # 시각
                'sell_members': sell_list,
                'buy_members': buy_list
            })

        # 4. StockMemberCnld: 통합 실시간 거래원 매매 정보
        elif self.name == 'member_cnld':
            # 1. 헤더 정보 추출
            count = self.obj.GetHeaderValue(1)        # 수신 개수 (보통 5개)
            time = self.obj.GetHeaderValue(2)         # 시각
            par_value = self.obj.GetHeaderValue(3)    # 액면가

            sell_members = []
            buy_members = []

            # 2. 상위 거래원 상세 데이터 추출 (GetDataValue 사용)
            for i in range(count):
                # 매도 상위 거래원 정보 저장
                sell_members.append({
                    'rank': i + 1,
                    'name': self.obj.GetDataValue(0, i), # 매도거래원 명
                    'qty': self.obj.GetDataValue(2, i)   # 총매도수량
                })
                
                # 매수 상위 거래원 정보 저장
                buy_members.append({
                    'rank': i + 1,
                    'name': self.obj.GetDataValue(1, i), # 매수거래원 명
                    'qty': self.obj.GetDataValue(3, i)   # 총매수수량
                })

            # 3. 최종 데이터 업데이트
            data.update({
                'time': time,
                'par_value': par_value,
                'sell_members': sell_members,
                'buy_members': buy_members
            })

        # 5. StockBsccnsCnld: 통합(KRX+NXT) 실시간 시세 및 체결 정보
        elif self.name == 'cur_cnld':
            # 매수/매도 누적 수량을 통한 체결강도 계산 준비
            vol_sell = self.obj.GetHeaderValue(15) # 누적매도체결수량
            vol_buy = self.obj.GetHeaderValue(16)  # 누적매수체결수량
            strength = (vol_buy / vol_sell * 100) if vol_sell > 0 else 0

            data.update({
                'time': self.obj.GetHeaderValue(3),        # 시간 (HHMMSS)
                'current': self.obj.GetHeaderValue(13),     # 현재가
                'diff_val': self.obj.GetHeaderValue(2),     # 전일대비
                'open': self.obj.GetHeaderValue(4),         # 시가
                'high': self.obj.GetHeaderValue(5),         # 고가
                'low': self.obj.GetHeaderValue(6),          # 저가
                'ask': self.obj.GetHeaderValue(7),  # 1차 매도호가
                'bid': self.obj.GetHeaderValue(8),  # 1차 매수호가
                'tick_vol': self.obj.GetHeaderValue(17),    # 순간체결수량
                'cum_vol': self.obj.GetHeaderValue(9),      # 누적거래량
                
                # 🎯 통합 데이터 핵심: 집행거래소 구분 ('K': KRX, 'N': NXT)
                'exchange': chr(self.obj.GetHeaderValue(29)), 
                
                # 수급 분석 필드
                'side': chr(self.obj.GetHeaderValue(14)),   # 체결상태 ('1': 매수, '2': 매도)
                'strength': round(strength, 2),             # 계산된 체결강도
                'market_stat': chr(self.obj.GetHeaderValue(20)) # 장구분플래그
            })

        # 최종 콜백 실행
        if self.callback:
            self.callback(data)

class RealtimeDataManager:
    def __init__(self, callback_func=None):
        self.callback = callback_func
        # 사용 가능한 모든 실시간 모듈 맵핑
        self.module_map = {
            'cur': "DsCbo1.StockCur",           # 실시간 시세 Subscribe/Publish v
            'cur_cnld': "DsCbo1.StockBsccnsCnld",   # 실시간 시세 KRX/NXT통합 Subscribe/Publish
            'jpbid': "DsCbo1.StockJpBid",        # 호가 잔량 Subscribe/Publish v
            'jpbid2': "DsCbo1.StockJpBid2",      # 호가 잔량 Request/Reply
            'jpbidcnld': "DsCbo1.StockJpBidCnld",# 호가 잔량 KRX/NXT통합 Subscribe/Publish v
            'member_cnld': "DsCbo1.StockMemberCnld", # 거래원 체결 KRX/NXT통합  Subscribe/Unsubscribe v
            'member': "Dscbo1.StockMember", # 주식 거래원 Subscribe/Publish v
            'member1': "Dscbo1.StockMember1", # 주식 거래원 Request/Reply
        }
        self.active_subscriptions = {}

    def start_monitoring(self, code, types=['cur', 'jpbid'], callback=None):
        """종목별 원하는 모듈을 골라서 구독 시작"""
        # 🎯 [추가] 인자로 넘어온 callback이 없으면 생성자에서 받은 self.callback을 사용합니다.
        target_callback = callback if callback else self.callback
        
        for t in types:
            if t not in self.module_map: continue
            
            key = f"{code}_{t}"
            if key in self.active_subscriptions: continue

            obj = win32com.client.Dispatch(self.module_map[t])
            obj.SetInputValue(0, code)
            
            handler = win32com.client.WithEvents(obj, RealtimeHandler)
            # 🎯 [수정] target_callback을 전달합니다.
            handler.set_params(obj, t, target_callback)
            obj.Subscribe()
            
            self.active_subscriptions[key] = (obj, handler)
        # print(f"🚀 [실시간 모니터링 시작] {code} -> {types}")

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