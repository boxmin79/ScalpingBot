import win32com.client
import time

class MarketEye:
    """
    [CpSysDib.MarketEye] 여러 종목의 다양한 데이터를 한 번에 조회하는 클래스
   
    """
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.MarketEye")

    def get_market_data(self, codes:list=[], field_ids:list=None):
        """
        주식,지수,선물/옵션등의여러종목의필요항목들을한번에수신합니다.
        
        Args:
            codes (list): 종목코드 리스트 (최대 200개)
            field_ids (list): 조회할 필드 ID 리스트 (최대 64개)

        [CpSysDib.MarketEye] 여러 종목의 다양한 데이터를 한 번에 조회하는 서비스

        1. SetInputValue(type, value) 설정:
            0 - (long array) 조회할 필드 ID 배열 (최대 64개)
            1 - (string array) 종목코드 배열 (최대 200개)
            2 - (char) 체결비교방식 ('1':체결가, '2':호가(기본))
            3 - (char) 거래소구분 ('A':전체, 'K':KRX, 'N':NXT)

        2. 주요 필드 ID 목록:
            0:종목코드(string)
            1:시간( ulong) - hhmm
            2:대비부호(char)('0':판단불가/초기값, '1':상한, '2':상승, '3':보합, '4':하한, '5':하락, '6':기세상한, '7':기세상승, '8':기세하한, '9':기세하락)
            3:전일대비(long or float) - 주의) 반드시대비부호(2)와같이요청을하여야함
            4:현재가(long or float)
            5:시가(long or float)
            6:고가(long or float)
            7:저가(long or float)
            8:매도호가(long or float)
            9:매수호가(long or float)
            10:거래량( ulong)
            11:거래대금(ulonglong) - 단위:원
            12:장구분(char or empty)('0':장전, '1':동시호가, '2':장중)
            13:총매도호가잔량(ulong)
            14:총매수호가잔량(ulong)
            15:최우선매도호가잔량(ulong)
            16:최우선매수호가잔량(ulong)
            17:종목명(string)
            20:총상장주식수(ulonglong) 
                    2018/4/30 이후로 상장주식수 20억 기준으로 이상인 경우에는 천단위로
                    그 이외의 경우에는 일단위로 수신
                    상장주식수 20억 이상 여부 종목 확인은
                    CpUtil.CpCodeMgr IsBigListingStock(code)를 사용하면 됨.
            21:외국인보유비율(float)
            22:전일거래량(ulong)
            23:전일종가(long or float)
            24:체결강도(float)
            25:체결구분(char or empty)('1':매수, '2':매도)
            27:미결제약정(long)
            28:예상체결가(long)
            29:예상체결가대비(long) - 주의) 반드시예샹체결가대비부호(30)와같이요청을하여야함
            30:예상체결가대비부호(char or empty)('1':상한, '2':상승, '3':보합, '4':하한, '5':하락)
            31:예상체결수량(ulong)
            32:19일종가합(long or float)
            33:상한가(long or float)
            34:하한가(long or float)
            35:매매수량단위(ushort)
            36:시간외단일대비부호(char or empty)
            ('0':판단불가/초기값, '1':상한, '2':상승, '3':보합, '4':하한, '5':하락, '6':기세상한, '7':기세상승, '8':기세하한, '9':기세하락)
            37:시간외단일전일대비(long) - 주의) 반드시시간외단일대비부호(36)와같이요청을하여야함
            38:시간외단일현재가(long)
            39:시간외단일시가(long)
            40:시간외단일고가(long)
            41:시간외단일저가(long)
            42:시간외단일매도호가(long)
            43:시간외단일매수호가(long)
            44:시간외단일거래량(ulong)
            45:시간외단일거래대금(ulonglong) - 단위:원
            46:시간외단일총매도호가잔량(ulong)
            47:시간외단일총매수호가잔량(ulong)
            48:시간외단일최우선매도호가잔량(ulong)
            49:시간외단일최우선매수호가잔량(ulong)
            50:시간외단일체결강도(float)
            51:시간외단일체결구분(char or empty)('1':매수, '2':매도)
            53:시간외단일예상/실체결구분(char)('1':예상, '2':실체결)
            54:시간외단일예상체결가(long)
            55:시간외단일예상체결전일대비(long) - 주의) 반드시시간외예상체결대비부호(56)와같이요청을하여야함
            56:시간외단일예상체결대비부호(char or empty)('1':상한, '2':상승, '3':보합, '4':하한, '5':하락)
            57:시간외단일예상체결수량(ulong)
            59:시간외단일기준가(long)
            60:시간외단일상한가(long)
            61:시간외단일하한가(long)
            62:외국인순매매(long) - 단위:주
            63:52주최고가(long or float)
            64:52주최저가(long or float)
            65:연중주최저가(long or float)
            66:연중최저가(long or float)
            67:PER(float)
            68:시간외매수잔량(ulong)
            69:시간외매도잔량(ulong)
            70:EPS(ulong)
            71:자본금(ulonglong)- 단위:백만
            72:액면가(ushort)
            73:배당률(float)
            74:배당수익률(float)
            75:부채비율(float)
            76:유보율(float)
            77:ROE(float) -  자기자본순이익률
            78:매출액증가율(float)
            79:경상이익증가율(float)
            80:순이익증가율(float)
            81:투자심리(float) : 제공하지 않음
            82: VR(float)        : 제공하지 않음
            83:5일회전율(float) : 제공하지 않음
            84:4일종가합(ulong)
            85:9일종가합(ulong)
            86:매출액(ulonglong) - 단위:백만
            87:경상이익(ulonglong) - 단위:원
            88:당기순이익(ulonglog) - 단위:원
            89:BPS(ulong) - 주당순자산
            90:영업이익증가율(float)
            91:영업이익(ulonglong) - 단위:원
            92:매출액영업이익률(float)
            93:매출액경상이익률(float)
            94:이자보상비율(float)
            95:결산년월(ulong) - yyyymm
            96:분기BPS(ulong) - 분기주당순자산
            97:분기매출액증가율(float)
            98:분기영업이액증가율(float)
            99:분기경상이익증가율(float)
            100:분기순이익증가율(float)
            101:분기매출액(ulonglong) - 단위:백만
            102:분기영업이익(ulonglong) - 단위:원
            103:분기경상이익(ulonglong) - 단위:원
            104:분기당기순이익(ulonglong) - 단위:원
            105:분개매출액영업이익률(float)
            106:분기매출액경상이익률(float)
            107:분기ROE(float) - 자기자본순이익률
            108:분기이자보상비율(float)
            109:분기유보율(float)
            110:분기부채비율(float)
            111:최근분기년월(ulong) - yyyymm
            112:BASIS(float)
            113:현지날짜(ulong) - yyyymmdd
            114:국가명(string) - 해외지수국가명
            115:ELW이론가(ulong)
            116:프로그램순매수(long)
            117:당일외국인순매수잠정구분(char)('\0'(0):해당없음, '1':확정, '2':잠정)
            118:당일외국인순매수(long)
            119:당일기관순매수잠정구분(char)('\0'(0):해당없음, '1':확정, '2':잠정)
            120:당일기관순매수(long)
            121:전일외국인순매수(long)
            122:전일기관순매수(long)
            123:SPS(ulong)
            124:CFPS(ulong)
            125:EBITDA(ulong)
            126:신용잔고율(float)
            127:공매도수량(ulong)
            128:공매도일자(ulong)
            129:ELW e-기어링(float)
            130:ELW LP보유양(ulong)
            131:ELW LP보유율(float)
            132:ELW LP Moneyness(float)
            133:ELW LP Moneyness구분(char)('1':ITM, '2':OTM, ' ':해당없음)
            134:ELW 감마(float)
            135:ELW 기어링(float)
            136:ELW 내재변동성(float)
            137:ELW 델타(float)
            138:ELW 발행수량(ulong)
            139:ELW 베가(float)
            140:ELW 세타(float)
            141:ELW 손익분기율(float)
            142:ELW 역사적변동성(float)
            143:ELW 자본지지점(float)
            144:ELW 패리티(float)
            145:ELW 프리미엄(float)
            146:ELW 베리어(float)
            147:ELW 풀 종목명(string)
            148:파생상품 장상태구분값(short)
            149: 지수/주식선물 전일미결제약정(long)
            150: 베타계수(float)
            153: 59일 종가합(long)
            154: 119일 종가합(long)
            155: 당일 개인 순매수 잠정구분(char)('\0'(0):해당없음, '1':확정, '2':잠정)
            156:  당일 개인 순매수(long)
            157:  전일 개인 순매수(long)
            158:  5일 전 종가(long)
            159:  10일 전 종가(float)
            160:  20일 전 종가(long)
            161:  60일 전 종가(long)
            162:  120일 전 종가(long)
            163:   정적VI 발동 예상기준가(long)
            164:   정적VI 발동 예상상승가(long)
            165:   정적VI 발동 예상하락가(long)

        """

        
        if not field_ids:
            field_ids = [0, 20, 67, 70, 71, 72, 73, 74, 76, 77, 78, 79, 80, 86, 87,
                         88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102,
                         103, 104, 105, 106, 107, 108, 109, 110, 111, 123, 124, 125]
            
        # 1. 입력값 설정
        self.obj.SetInputValue(0, field_ids) # 필드 배열
        self.obj.SetInputValue(1, codes)     # 종목코드 배열
        self.obj.SetInputValue(2, '2') # 체결 방식 '1' 체결가비교방식, 2'호가비교방식(default)
        self.obj.SetInputValue(3, ord('K')) # 거래소구분 'A' 전체, 'K' KRX, 'N' NXT
        
        # 2. 데이터 요청 (Request/Reply)
        ret = self.obj.BlockRequest()
        if ret != 0:
            print(f"❌ MarketEye 요청 실패 (에러코드: {ret})")
            return []

        # 3. 헤더 정보 확인
        field_count = self.obj.GetHeaderValue(0) # 요청한 필드 개수
        field_array = self.obj.GetHeaderValue(1) # 요청한 필드 배열
        stock_count = self.obj.GetHeaderValue(2) # 수신된 종목 개수
        market_flag = self.obj.GetHeaderValue(3) # 거래소 구분
        
        # ⚠️ 중요: GetDataValue의 type(필드 인덱스)은 요청한 field_ids의 
        # 값이 아니라, field_ids를 오름차순으로 정렬했을 때의 순서입니다.
        sorted_fields = sorted(field_ids)
        results = []
        # for fd_idx, col_idx in zip(sorted_fields, field_array):
            # print(fd_idx, col_idx)
            
        # 4. 데이터 파싱
        for i in range(stock_count):
            stock_data = {}
            for f_idx, f_id in enumerate(sorted_fields):
                # GetDataValue(필드순서, 종목순서)
                stock_data[f_id] = self.obj.GetDataValue(f_idx, i)
            results.append(stock_data)

        return results, field_array

# 사용 예시
if __name__ == "__main__":
    manager = MarketEye()
    # 종목: SK하이닉스, 삼성전자 / 필드: 0(코드), 17(명칭), 4(현재가), 24(체결강도)
    target_codes = ["A000660", "A005930"]
    # target_fields = [0, 17, 4, 24, 11] # 11: 거래대금
    
    data = manager.get_market_data(target_codes)
    for item in data:
        print(f"종목코드: {item[0]}")

    
    ### 가져올 수 있는 데이터 ###
    
    