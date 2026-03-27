import win32com.client
import logging

class CreonAPI:
    """
    대신증권 크레온 Plus API 통합 관리 클래스
    
    _summary_
    이 클래스는 대신증권에서 제공하는 각 서비스별 COM 객체를 하나로 묶어 관리합니다.
    모든 객체는 별도의 메서드 호출 없이 인스턴스 속성(Attribute)으로 즉시 접근 가능합니다.
    """
    
    def __init__(self):
        self.logger = logging.getLogger("TradingBot.CreonAPI")
        
        try:
            # 1. 시스템 상태 관리 및 통신 제한 제어 객체
            self.obj_cybos = win32com.client.Dispatch("CpUtil.CpCybos")
            """
            [CpUtil.CpCybos 객체 상세 사용 설명서]
            
            크레온(CYBOS) API의 메인 연결 상태를 확인하고, 서버 통신 시 발생하는 
            조회(RQ) 및 구독(SB) 횟수 제한을 관리하여 프로그램이 서버로부터 차단되는 것을 방지합니다.

            ■ 주요 프로퍼티 (Property - 읽기 전용 속성)
            
            1. self.obj_cybos.IsConnect
               - 설명: API 프로그램과 증권사 서버 간의 통신 연결 상태를 확인합니다.
               - 반환값: 0 (연결 끊김), 1 (연결 정상)
               - [실전 팁]: 봇을 구동하기 전 반드시 가장 먼저 호출하여 1인지 확인해야 합니다. 
                 장중에도 주기적으로 체크하여 0이 되면 재연결 로직을 수행하도록 설계합니다.

            2. self.obj_cybos.ServerType
               - 설명: 현재 연결되어 있는 서버의 종류를 확인합니다.
               - 반환값: 0 (연결 끊김), 1 (CybosPlus 서버 - API 전용), 2 (HTS 보통 서버)
               - [실전 팁]: 정상적인 API 통신을 위해서는 1이 반환되어야 합니다.

            3. self.obj_cybos.LimitRequestRemainTime
               - 설명: 횟수 제한에 걸렸을 때, 요청 가능 횟수가 다시 리셋(재계산)되기까지 남은 시간을 알려줍니다.
               - 반환값: 남은 시간 (단위: 밀리초 ms) (예: 500 = 0.5초)
               - [실전 팁]: 통신 제한에 걸리기 직전, 이 반환값만큼 `time.sleep(남은시간 / 1000)`을 
                 주어 영구 정지를 피하는 스마트 대기 로직을 구현할 수 있습니다.


            ■ 주요 메서드 (Method - 동작 수행)

            1. self.obj_cybos.GetLimitRemainCount(limitType)
               - 설명: 특정 요청 타입에 대해 현재 남은 '호출 가능 횟수'를 반환합니다.
               - 파라미터 (limitType):
                 * 0 (LT_TRADE_REQUEST) : 주문 및 계좌 관련 일반 요청 (조회/RQ) 잔여 횟수
                 * 1 (LT_NONTRADE_REQUEST) : 시세 및 종목 정보 일반 요청 (조회/RQ) 잔여 횟수 
                                           (대신증권은 보통 15초당 60건 제한)
                 * 2 (LT_SUBSCRIBE) : 실시간 시세 구독 (SB) 등록 가능 잔여 횟수 (최대 400종목 제한)
               - 반환값: 남은 요청 개수 (정수형)
               - [실전 팁]: 스캘핑 스크리너(`CpTopVolume`, `MarketEye` 등)를 루프로 돌릴 때,
                 반드시 루프 안에서 `GetLimitRemainCount(1)`을 체크하여 값이 1~2 이하로 떨어지면 
                 위의 `LimitRequestRemainTime`만큼 대기하게 만드는 방어 코드가 필수적입니다.
            """
            
            # 2. 종목 정보 및 메타데이터 관리 객체
            self.obj_code_mgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
            """
            [CpUtil.CpCodeMgr 객체 상세 사용 설명서]
            
            크레온 API에서 '종목 사전' 역할을 하는 가장 중요한 객체 중 하나입니다.
            전체 상장 종목의 리스트를 뽑아오거나, 특정 종목의 상태(거래정지, 관리종목 등),
            ETF/스팩 여부 등을 확인하여 스캘핑에 적합한 '순수 보통주'만 걸러낼 때 필수적으로 사용됩니다.

            ■ 주요 메서드 (Method - 데이터 조회 및 변환)

            1. 기본 변환 기능
               - self.obj_code_mgr.CodeToName(code)
                 : 종목코드를 입력하면 종목명을 반환합니다. (예: "A005930" -> "삼성전자")
               - self.obj_code_mgr.NameToCode(name)
                 : 종목명을 입력하면 종목코드를 반환합니다. (예: "삼성전자" -> "A005930")

            2. 시장별 종목 리스트 추출
               - self.obj_code_mgr.GetStockListByMarket(market_kind)
                 : 특정 시장에 속한 전체 종목코드 리스트(Tuple)를 반환합니다.
                 * market_kind: 1 (거래소/KOSPI), 2 (코스닥/KOSDAQ), 3 (K-OTC), 4 (KRX), 5 (KONEX)

            3. 종목 필터링 및 상태 확인 (★ 스캘핑 필수 확인 요소)
               - self.obj_code_mgr.GetStockSectionKind(code)
                 : 부구분 코드를 반환하여 주식인지 ETF인지 판별합니다.
                 * 반환값: 1 (주권 - 일반주식), 10 (ETF), 17 (ETN) 등
                 * [실전 팁]: 스캘퍼는 보통 변동성이 큰 개별주를 노리므로 반환값이 '1'인 것만 챙깁니다.
                 
               - self.obj_code_mgr.IsSpac(code)
                 : 해당 종목이 스팩(SPAC)인지 확인합니다.
                 * 반환값: True (스팩), False (일반)
                 
               - self.obj_code_mgr.GetStockStatusKind(code)
                 : 주식의 거래 상태를 확인합니다.
                 * 반환값: 0 (정상), 1 (거래정지), 2 (거래중단)
                 * [실전 팁]: 0이 아닌 종목은 매매가 불가능하므로 감시 리스트에서 무조건 제외합니다.
                 
               - self.obj_code_mgr.GetStockControlKind(code)
                 : 감리(주의/경고) 구분을 반환합니다.
                 * 반환값: 0 (정상), 1 (주의), 2 (경고), 3 (위험예고), 4 (위험)
                 * [실전 팁]: 증거금 100%나 위험 종목을 피하려면 반환값이 0인 종목만 필터링합니다.

            4. 종목 상세 메타데이터 조회
               - self.obj_code_mgr.GetStockCapital(code)
                 : 자본금 규모를 반환합니다. (0: 제외, 1: 대형주, 2: 중형주, 3: 소형주)
               - self.obj_code_mgr.GetStockListedDate(code)
                 : 상장일을 반환합니다. (예: 20230511 - LONG 타입)
                 * [실전 팁]: 신규 상장주(상장일이 최근인 종목)만 따로 모아서 매매하는 전략에 유용합니다.
               - self.obj_code_mgr.GetStockFiscalMonth(code)
                 : 결산월을 반환합니다. (예: 12 -> 12월 결산법인)

            ■ [스캘핑 봇 실전 활용 가이드]
            장 시작 전(또는 프로그램 초기화 시), 
            1) `GetStockListByMarket`으로 코스피/코스닥 전체 코드를 받고,
            2) 루프를 돌면서 `GetStockSectionKind == 1` (일반주식)
            3) `IsSpac == False` (스팩 제외)
            4) `code[-1] == '0'` (우선주 제외, 끝자리가 0이어야 보통주)
            5) `GetStockStatusKind == 0` & `GetStockControlKind == 0` (정상 종목)
            위 조건들을 통과한 아주 '깨끗한' 종목 리스트(Universe)를 미리 만들어두고, 
            이후 시세 조회나 실시간 감시에 사용하는 것이 가장 정석적이고 안전한 방법입니다.
            """
            
            # 3. 주식 코드 및 인덱스 관리 객체
            self.obj_stock_code = win32com.client.Dispatch("CpUtil.CpStockCode")
            """
            [CpUtil.CpStockCode 객체 상세 사용 설명서]
            
            크레온 API에 등록된 모든 주식/선물/옵션 종목의 코드를 '배열(List)'처럼 
            인덱스(0, 1, 2...) 기반으로 관리하고 접근할 수 있게 해주는 객체입니다.
            'CpCodeMgr'이 특정 종목의 "상세 속성"을 본다면, 'CpStockCode'는 "전체 명단"을 훑어볼 때 유리합니다.

            ■ 주요 메서드 (Method - 데이터 조회)

            1. self.obj_stock_code.GetCount()
               - 설명: 현재 API 시스템에 등록된 전체 종목의 총 개수를 반환합니다.
               - 반환값: 총 종목 수 (정수형)
               - [실전 팁]: 전체 종목을 처음부터 끝까지 검색(Full Scan)해야 할 때 
                 `for i in range(self.obj_stock_code.GetCount()):` 형태로 루프를 돌릴 때 기준값이 됩니다.

            2. self.obj_stock_code.GetData(type, index)
               - 설명: 특정 인덱스(index)에 위치한 종목의 데이터를 반환합니다.
               - 파라미터 (type):
                 * 0 : 종목 코드 (예: "A005930")
                 * 1 : 종목 명 (예: "삼성전자")
                 * 2 : Full Code (표준코드, 예: "KR7005930003")
               - 반환값: 요청한 문자열(String) 데이터
               - [실전 팁]: `CpCodeMgr.CodeToName()`을 써도 되지만, 전체 종목 명칭 리스트를 
                 한 번에 딕셔너리로 구축해둘 때는 `GetData`를 루프로 돌리는 것이 훨씬 빠릅니다.

            3. self.obj_stock_code.CodeToIndex(code)
               - 설명: 특정 종목 코드의 인덱스 번호를 찾습니다.
               - 반환값: 인덱스 번호 (해당 코드가 없으면 -1 반환)
               - [실전 팁]: 사용자가 잘못된 종목 코드(예: 상장폐지된 코드)를 입력했는지 
                 검증(Validation)할 때, 이 메서드를 호출하여 `-1`이 나오면 걸러내는 식으로 활용합니다.

            ■ [스캘핑 봇 실전 활용 가이드]
            
            스캘핑 봇에서는 보통 시장별(코스피/코스닥)로 미리 걸러진 리스트를 얻기 위해 
            `CpCodeMgr.GetStockListByMarket`을 더 자주 사용합니다. 
            하지만 봇 초기화 단계에서 화면에 출력할 로그나, 빠른 종목명 검색용 캐시(Cache)를 
            만들 때는 `CpStockCode`를 활용하면 좋습니다.
            
            [예시: 초고속 코드-이름 매핑 딕셔너리 만들기]
            ```python
            # 봇 구동 시 1회만 실행하여 메모리에 올려둠
            code_to_name_dict = {}
            for i in range(self.obj_stock_code.GetCount()):
                code = self.obj_stock_code.GetData(0, i)
                name = self.obj_stock_code.GetData(1, i)
                code_to_name_dict[code] = name
            ```
            이후 장중에는 API를 호출할 필요 없이 이 딕셔너리(`code_to_name_dict`)에서 
            O(1) 속도로 종목명을 찾아오면 봇의 연산 속도(Latency)를 극대화할 수 있습니다.
            """
            
            # 4. 주문 유틸리티 및 계좌 정보 관리 객체
            self.obj_trade_util = win32com.client.Dispatch("CpTrade.CpTdUtil")
            """
            [CpTrade.CpTdUtil 객체 상세 사용 설명서]
            
            크레온 API에서 매수/매도 주문을 넣기 전 반드시 거쳐야 하는 초기화 및 
            계좌 정보(계좌번호, 상품번호) 관리 객체입니다.
            주문 시스템은 일반 시세 조회보다 보안이 엄격하므로, 이 객체를 통한 검증이 필수적입니다.

            ■ 주요 메서드 및 프로퍼티

            1. self.obj_trade_util.TradeInit(0)  [★ 주문 전 필수 관문 ★]
               - 설명: API 주문 서비스 초기화 작업을 수행합니다.
               - 파라미터: 0 (기본값)
               - 반환값: 
                 * 0 : 성공 (정상적으로 주문 가능한 상태)
                 * -1 : 오류 (초기화 실패, 통신 장애 등)
                 * 1 : 비밀번호 입력 필요 (HTS/MTS나 크레온 플러스(API)에서 통신비밀번호 미입력 상태)
                 * 2 : 계좌 비밀번호 불일치
                 * 3 : 취소됨
               - [실전 팁]: 봇을 시작할 때 반드시 이 반환값이 0인지 체크하는 로직이 있어야 합니다.
                 만약 0이 아니라면 프로그램에서 매수/매도 주문을 넣어도 서버에서 튕겨냅니다.

            2. self.obj_trade_util.AccountNumber (프로퍼티)
               - 설명: 로그인된 사용자의 보유 계좌 번호 목록을 배열(Tuple) 형태로 반환합니다.
               - 반환값 예시: ('123456789', '987654321')
               - [실전 팁]: 다수 계좌 보유자라면 보통 `account = self.obj_trade_util.AccountNumber[0]` 
                 형태로 첫 번째 메인 계좌를 선택하여 주문 객체(CpTd0311)에 전달합니다.

            3. self.obj_trade_util.GoodsList(account_number, filter)
               - 설명: 특정 계좌 번호(account_number)에 속한 상품 번호(보통 2자리 숫자) 목록을 반환합니다.
               - 파라미터:
                 * account_number: AccountNumber에서 얻은 계좌 번호
                 * filter: 필터 조건 (1: 주식/채권, 2: 선물/옵션 등)
               - 반환값 예시: ('01', '10')
               - [실전 팁]: 주식 스캘핑 시에는 일반적으로 `filter=1`을 주고 반환된 튜플의 
                 첫 번째 값(`GoodsList(acc, 1)[0]`)을 주식용 상품 번호로 사용합니다.

            ■ [스캘핑 봇 실전 활용 가이드]
            
            스캘핑 주문 모듈(`OrderManager`)을 만들 때 이 객체를 다음과 같이 활용하여 
            계좌 세팅을 자동화하는 것이 가장 깔끔합니다.

            ```python
            # 주문 초기화 체크
            init_status = self.obj_trade_util.TradeInit(0)
            if init_status != 0:
                print(f"❌ 주문 불가 상태입니다. (에러 코드: {init_status}) 크레온 플러스에서 계좌 비밀번호를 입력하세요.")
                return False
                
            # 계좌 및 상품 번호 자동 세팅
            self.account_no = self.obj_trade_util.AccountNumber[0]
            self.goods_no = self.obj_trade_util.GoodsList(self.account_no, 1)[0]
            
            print(f"✅ 주문 준비 완료: 계좌({self.account_no}), 상품코드({self.goods_no})")
            ```
            """
            
            # 5. 복수 종목 데이터 초고속 조회 객체 (스캘핑 스크리너 핵심)
            self.obj_market_eye = win32com.client.Dispatch("CpSysDib.MarketEye")
            """
            [CpSysDib.MarketEye 객체 상세 사용 설명서]
            
            여러 종목(최대 200개)의 다양한 시세 및 재무 데이터를 한 번의 API 통신(BlockRequest)으로 
            가져오는 초고속 조회 객체입니다. 스캘핑에서 실시간 감시 대상을 선정하는 '스크리너'를 
            만들 때 가장 필수적이고 강력한 도구입니다.

            ■ 주요 메서드 (Method)

            1. self.obj_market_eye.SetInputValue(type, value)
               - 설명: 서버에 어떤 데이터를 요청할지 조건을 설정합니다.
               - 파라미터 (type):
                 * 0 : 요청할 데이터의 '필드(Field) 번호 배열'을 설정합니다. (필수)
                       [자주 쓰는 필드 번호]
                       0(종목코드), 4(현재가), 10(거래량), 11(거래대금), 17(종목명),
                       20(총상장주식수), 22(시가총액 - 단, 부정확할 수 있어 계산 권장),
                       23(총매도호가잔량), 24(총매수호가잔량), 67(PER), 77(ROE), 
                       92(매출액영업이익률), 105(분기매출액영업이익률), 107(분기ROE)
                 * 1 : 조회할 '종목 코드 배열'을 설정합니다. (필수, 최대 200종목)
               - [실전 팁]: 200개가 넘는 종목을 조회할 때는 리스트를 200개 단위(Chunk)로 
                 쪼개서 여러 번 `BlockRequest`를 호출해야 합니다.

            2. self.obj_market_eye.BlockRequest()
               - 설명: 설정된 조건(`SetInputValue`)을 바탕으로 서버에 데이터를 요청합니다.
               - [실전 팁]: 동기(Synchronous) 방식이므로, 서버에서 응답이 올 때까지 
                 프로그램이 잠시 멈춥니다(Blocking). 너무 빈번하게 호출하면 통신 제한에 걸립니다.

            3. self.obj_market_eye.GetHeaderValue(type)
               - 설명: 수신된 데이터의 메타 정보(개수 등)를 반환합니다.
               - 파라미터 (type):
                 * 0 : 요청한 필드(Field)의 개수
                 * 1 : 필드 번호의 배열
                 * 2 : ★ 수신된 종목의 개수 (결과값을 루프 돌릴 때 기준이 됨)

            4. self.obj_market_eye.GetDataValue(field_index, row_index)
               - 설명: 실제 수신된 데이터를 추출합니다.
               - 파라미터:
                 * field_index : SetInputValue(0, ...)에서 설정한 필드 배열의 '순서(인덱스)' 
                                 (필드 번호 자체가 아님에 매우 주의!)
                 * row_index : 조회된 종목의 순서(인덱스) (0부터 GetHeaderValue(2)-1 까지)
               - 반환값: 해당 위치의 데이터 값

            ■ [스캘핑 봇 실전 활용 가이드]
            
            [예시: 3개 종목의 현재가, 거래량, 거래대금 한 번에 가져오기]
            ```python
            # 1. 요청 필드 설정: 0(코드), 4(현재가), 10(거래량), 11(거래대금)
            fields = [0, 4, 10, 11]
            codes = ['A005930', 'A000660', 'A035720'] # 삼성전자, SK하이닉스, 카카오
            
            self.obj_market_eye.SetInputValue(0, fields)
            self.obj_market_eye.SetInputValue(1, codes)
            
            # 2. 서버 요청
            self.obj_market_eye.BlockRequest()
            
            # 3. 데이터 추출
            count = self.obj_market_eye.GetHeaderValue(2) # 수신 종목 수 (3개)
            
            for i in range(count):
                code = self.obj_market_eye.GetDataValue(0, i)    # fields 배열의 0번째 인덱스(0:코드)
                price = self.obj_market_eye.GetDataValue(1, i)   # fields 배열의 1번째 인덱스(4:현재가)
                vol = self.obj_market_eye.GetDataValue(2, i)     # fields 배열의 2번째 인덱스(10:거래량)
                amount = self.obj_market_eye.GetDataValue(3, i)  # fields 배열의 3번째 인덱스(11:거래대금)
                print(f"{code} | 현재가: {price}, 거래량: {vol}, 거래대금: {amount}")
            ```
            이처럼 MarketEye는 스크리너가 수백 개의 종목 중에서 '오늘 거래대금이 터지는 놈'을 
            초고속으로 골라낼 때 절대적으로 필요한 1등 공신입니다.
            """
            
            # 연결 확인 (Cybos Plus 접속 여부)
            if self.obj_cybos.IsConnect == 0:
                self.logger.error("❌ 크레온 Plus가 연결되지 않았습니다. 관리자 권한으로 실행하세요.")
            else:
                self.logger.info("✅ 크레온 Plus API 객체 초기화 완료")
                
        except Exception as e:
            self.logger.error(f"❌ API 객체 초기화 중 오류 발생: {e}")

# --- 실전 활용 예시 ---
if __name__ == "__main__":
    api = CreonAPI()
    
    # 예: 연결 상태 확인 후 종목수 출력
    if api.obj_cybos.IsConnect == 1:
        total_count = api.obj_stock_code.GetCount()
        print(f"현재 상장된 총 종목 수: {total_count}")
        
        # 삼성전자 정보 한 줄 확인
        samsung_name = api.obj_code_mgr.CodeToName("A005930")
        print(f"종목명: {samsung_name}")