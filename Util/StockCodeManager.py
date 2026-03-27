import path_finder
import pandas as pd
import win32com.client
import os
from datetime import datetime
import json
# from pathlib import Path
import time

class CodeManager:
    def __init__(self, save=True):
        self.cfg = path_finder.get_cfg()
        self.obj_code_mgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
        self.obj_market_eye = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.obj_stock_chart = win32com.client.Dispatch("CpSysDib.StockChart")
        self.tickers_list_df = None
        self.save_enabled = save
        self.meta_path = str(self.cfg.CODE_LIST_PATH).replace('.csv', '_meta.json')
        
    def get(self):
        """
        종목 리스트를 가져오는 메인 메서드.
        메타 파일의 날짜를 확인하여 오늘 데이터가 아니면 새로 갱신합니다.
        """
        file_path = self.cfg.CODE_LIST_PATH
        
        # 1. CSV 파일과 메타 파일이 모두 존재하는지 확인
        if os.path.exists(file_path) and os.path.exists(self.meta_path):
            try:
                with open(self.meta_path, 'r', encoding='utf-8') as f:
                    meta_data = json.load(f)
                
                last_update = meta_data.get('last_update', '')
                today_date = datetime.now().strftime('%Y%m%d')

                # 2. 메타 파일의 날짜가 오늘인지 확인
                if last_update == today_date:
                    print(f"[시스템] 오늘 날짜({today_date})의 메타 정보를 확인했습니다. 데이터를 로드합니다.")
                    self.tickers_list_df = pd.read_csv(file_path, encoding='utf-8-sig')
                    return self.tickers_list_df
                else:
                    print(f"[시스템] 데이터가 최신이 아닙니다(마지막 갱신: {last_update}). 새로 갱신합니다.")
            except Exception as e:
                print(f"[경고] 메타 파일 읽기 오류: {e}. 새로 수집을 진행합니다.")
        else:
            print(f"[시스템] 기존 데이터 또는 메타 파일이 없습니다. 새로 생성합니다.")

        # 파일이 없거나 최신이 아니면 데이터 수집 실행
        self.update_tickers_list()
        return self.tickers_list_df
    
    def update_tickers_list(self):
        """데이터를 새로 수집하고 저장하는 로직"""
        df = pd.DataFrame()
        ticker_list = self._get_clean_tickers()
        
        if not ticker_list:
            return None

        df['tickers'] = ticker_list
        df['name'] = self._get_ticker_name(ticker_list)
        df['capital'] = self._get_stock_capital(ticker_list)
        df['market'] = self._get_stock_market(ticker_list)
        df['fiscal_month'] = self._get_stock_fiscal_month(ticker_list)
        df['listed_date'] = self._get_stock_listed_date(ticker_list)
        
        print("[시스템] MarketEye를 통해 재무 데이터(ROE, 영업이익률) 초고속 추출 시작...")
        financial_data = self._get_financial_metrics_marketeye(ticker_list)
        df['market_cap'] = financial_data['market_cap']
        df['listed_shares'] = financial_data['listed_shares']
        df['op_margin'] = financial_data['op_margin']
        df['op_margin_q'] = financial_data['op_margin_q']
        df['roe'] = financial_data['roe']
        df['roe_q'] = financial_data['roe_q']
        
        # 데이터 정제 및 필터링
        cols_to_numeric = ['market_cap', 'op_margin', 'op_margin_q', 'roe', 'roe_q']
        for col in cols_to_numeric:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        print("[시스템] 1차 재무/시총(500억↑) 필터링 진행 중...")
        cond1 = (df['op_margin_q'] > 3) & (df['roe_q'] > 5)
        cond2 = (df['op_margin'] > 5) & (df['roe'] > 8) & (df['op_margin_q'] > 0) & (df['roe_q'] > 0)
        self.tickers_list_df = df[(cond1 | cond2) & (df['market_cap'] >= 500)].copy()
        
        # 업데이트 날짜 추가
        self.tickers_list_df['update_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        if self.save_enabled:
            # 1. 종목 데이터는 CSV로 저장 (날짜 컬럼 제외)
            self.tickers_list_df.to_csv(self.cfg.CODE_LIST_PATH, index=False, encoding='utf-8-sig')
            
            # 2. 업데이트 날짜는 별도 JSON 파일로 저장
            today_str = datetime.now().strftime('%Y%m%d')
            full_time_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            meta_info = {
                "last_update": today_str,
                "full_timestamp": full_time_str,
                "count": len(self.tickers_list_df)
            }
            
            with open(self.meta_path, 'w', encoding='utf-8') as f:
                json.dump(meta_info, f, indent=4, ensure_ascii=False)
                
            print(f"[시스템] 종목 데이터 저장 완료: {self.cfg.CODE_LIST_PATH}")
            print(f"[시스템] 업데이트 날짜 저장 완료: {self.meta_path} ({full_time_str})")
                       
    def get_tickers_list_df(self):
        df = pd.DataFrame()
        ticker_list = self._get_clean_tickers()
        
        if not ticker_list:
            return df
        
        df['tickers'] = ticker_list
        df['name'] = self._get_ticker_name(ticker_list)
        df['capital'] = self._get_stock_capital(ticker_list)
        df['market'] = self._get_stock_market(ticker_list)
        df['fiscal_month'] = self._get_stock_fiscal_month(ticker_list)
        df['listed_date'] = self._get_stock_listed_date(ticker_list)
        
        # 💡 MarketEye를 통한 재무 데이터 대량 추출
        print("[시스템] MarketEye를 통해 재무 데이터(ROE, 영업이익률) 초고속 추출 시작...")
        financial_data = self._get_financial_metrics_marketeye(ticker_list)
        df['market_cap'] = financial_data['market_cap']
        df['listed_shares'] = financial_data['listed_shares']
        df['op_margin'] = financial_data['op_margin']      # 연간 영업이익률(매출액영업이익률)
        df['op_margin_q'] = financial_data['op_margin_q']  # 분기 영업이익률
        df['roe'] = financial_data['roe']                  # 연간 ROE
        df['roe_q'] = financial_data['roe_q']              # 분기 ROE
        
        # 숫자형 변환
        cols_to_numeric = ['market_cap', 'op_margin', 'op_margin_q', 'roe', 'roe_q']
        for col in cols_to_numeric:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # 2. [1차 필터] 재무 및 시가총액 필터링
        print("[시스템] 1차 재무/시총(500억↑) 필터링 진행 중...")
        cond1 = (df['op_margin_q'] > 3) & (df['roe_q'] > 5)
        cond2 = (df['op_margin'] > 5) & (df['roe'] > 8) & (df['op_margin_q'] > 0) & (df['roe_q'] > 0)
        df_final = df[(cond1 | cond2) & (df['market_cap'] >= 500)].copy()
        
        # 업데이트 날짜 추가
        self.tickers_list_df['update_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        if self.save_enabled:
            self.tickers_list_df.to_csv(self.cfg.CODE_LIST_PATH, index=False, encoding='utf-8-sig')
            print(f"[시스템] 최종 {len(self.tickers_list_df)}개 종목 CSV 저장 완료: {self.cfg.CODE_LIST_PATH}")

        
        return df_final
    
    def _get_financial_metrics_marketeye(self, ticker_list):
        # 0: 종목코드, 4: 현재가, 20: 총상장주식수
        # 92: 매출액영업이익률, 105: 분기매출액영업이익률, 77: ROE, 107: 분기ROE
        fields = [0, 4, 20, 92, 105, 77, 107]
        temp_dict = {}
        chunk_size = 200
        total_chunks = (len(ticker_list) // chunk_size) + 1
        
        for i in range(0, len(ticker_list), chunk_size):
            chunk = ticker_list[i : i + chunk_size]
            current_chunk = (i // chunk_size) + 1
            print(f"-> MarketEye 수신 중... ({current_chunk}/{total_chunks} 페이지)")
            
            self.obj_market_eye.SetInputValue(0, fields)
            self.obj_market_eye.SetInputValue(1, chunk)
            self.obj_market_eye.BlockRequest()
            time.sleep(0.3)
            
            count = self.obj_market_eye.GetHeaderValue(2)
            for j in range(count):
                code = self.obj_market_eye.GetDataValue(0, j)
                price = self.obj_market_eye.GetDataValue(1, j)
                shares = self.obj_market_eye.GetDataValue(2, j)
                
                # 💡 시가총액 계산 (현재가 * 상장주식수 / 1억)
                # 데이터가 None이거나 계산 불가일 경우를 대비한 안전 장치
                try:
                    market_cap = (int(price) * int(shares)) // 100000000
                except (ValueError, TypeError):
                    market_cap = 0

                temp_dict[code] = {
                    'listed_shares': shares,
                    'market_cap': market_cap,
                    'op_margin': self.obj_market_eye.GetDataValue(3, j),
                    'op_margin_q': self.obj_market_eye.GetDataValue(4, j),
                    'roe': self.obj_market_eye.GetDataValue(5, j),
                    'roe_q': self.obj_market_eye.GetDataValue(6, j)
                }

        # 결과 리스트화
        results = { 'listed_shares': [], 'market_cap': [], 'op_margin': [], 'op_margin_q': [], 'roe': [], 'roe_q': [] }
        for code in ticker_list:
            data = temp_dict.get(code, {'listed_shares': 0, 'market_cap': 0, 'op_margin': 0, 'op_margin_q': 0, 'roe': 0, 'roe_q': 0})
            for key in results:
                results[key].append(data[key])
            
        return results
    
    def _get_stock_listed_date(self, ticker_list):
        """
        value = object.GetStockListedDate ( code )

        code 에해당하는상장일을반환한다

        code : 주식코드
        반환값 : 상장일 (LONG)
        """
        # 1. 먼저 원본 숫자(LONG) 리스트를 쫙 뽑아옵니다.
        raw_dates = [self.obj_code_mgr.GetStockListedDate(ticker) for ticker in ticker_list]
        
        # 2. Pandas를 이용해 'YYYYMMDD' 형식을 날짜 타입으로 한 번에 강제 변환합니다.
        # errors='coerce'를 넣으면 0이나 이상한 값이 들어와도 에러를 뱉지 않고 자연스럽게 NaT(빈 값)로 처리해 줍니다.
        parsed_dates = pd.to_datetime(raw_dates, format='%Y%m%d', errors='coerce')
        
        return parsed_dates.tolist()
    
    def _get_stock_fiscal_month(self, ticker_list):
        """
        value = object.GetStockFiscalMonth ( code )

        code 에해당하는결산기반환한다.

        code : 주식코드
        반환값 : 결산기
        """  
        fiscal_month = []
        for ticker in ticker_list:
            fiscal_month.append(self.obj_code_mgr.GetStockFiscalMonth(ticker))
        return fiscal_month
        
        
    def _get_stock_market(self, ticker_list):
        """
        value = object.GetStockMarketKind ( code )

        code 에해당하는소속부를반환한다.

        code : 주식코드
        반환값 : 소속부
        typedefenum{
        [helpstring("구분없음")]CPC_MARKET_NULL= 0,
        [helpstring("거래소")]   CPC_MARKET_KOSPI= 1,
        [helpstring("코스닥")]   CPC_MARKET_KOSDAQ= 2,
        [helpstring("K-OTC")] CPC_MARKET_FREEBOARD= 3,
        [helpstring("KRX")]       CPC_MARKET_KRX= 4,
        [helpstring("KONEX")] CPC_MARKET_KONEX= 5,
        }CPE_MARKET_KIND;
        """
        market_list = []
        for ticker in ticker_list:
            val = self.obj_code_mgr.GetStockMarketKind(ticker)
            if val == 0:
                market = "구분없음"
            elif val == 1:
                market = "거래소"
            elif val == 2:
                market = "코스닥"
                
            market_list.append(market)
        return market_list
        
    def _get_stock_capital(self, ticker_list):
        """
        code에 해당하는 자본금 규모 구분을 반환한다.
        
        Args:
            code (str): 주식코드
            
        Returns:
            int: 자본금 규모 구분
                0: 제외 (CPC_CAPITAL_NULL)
                1: 대 (CPC_CAPITAL_LARGE)
                2: 중 (CPC_CAPITAL_MIDDLE)
                3: 소 (CPC_CAPITAL_SMALL)

        value = object.GetStockCapital ( code )
        """
        capital_list = []
        for ticker in ticker_list:
            val = self.obj_code_mgr.GetStockCapital(ticker)
            if val == 0:
                capital = "제외"
            elif val == 1:
                capital = "대"
            elif val == 2:
                capital = "중"
            elif val == 3:
                capital = "소"
            
            capital_list.append(capital)
        return capital_list
        
    def _get_ticker_name(self, ticker_list):
        """
        value = object.CodeToName( code )

        code 에해당하는주식/선물/옵션종목명을반환한다.

        code : 주식/선물/옵션코드

        반환값 : 주식/선물/옵션종목명

              Args:
                  ticker_list (_type_): _description_

              Returns:
                  _type_: _description_
        """
        name_list = []
        for ticker in ticker_list:
            name = self.obj_code_mgr.CodeToName(ticker)
            name_list.append(name)
        return name_list
        
      
    def _get_clean_tickers(self):
        print("[시스템] 코스피/코스닥 종목 리스트 추출 시작...")
        
        # 1. 원본 데이터가 제대로 들어오는지부터 확인!
        kospi_list = self.obj_code_mgr.GetStockListByMarket(1)
        kosdaq_list = self.obj_code_mgr.GetStockListByMarket(2)
        
        all_tickers = list(kospi_list) + list(kosdaq_list)
        
        print(f"-> [원본] 코스피: {len(kospi_list)}개 / 코스닥: {len(kosdaq_list)}개 / 총합: {len(all_tickers)}개")
        
        if len(all_tickers) == 0:
            print("[오류] 원본 데이터가 0개입니다! 사이보스 플러스를 완전히 종료 후 관리자 권한으로 재실행하세요.")
            return []

        clean_tickers = []
        
        # 2. 범인 찾기 카운터
        drop_section = 0
        drop_spac = 0
        drop_pref = 0
        drop_status = 0
        drop_control = 0

        for code in all_tickers:
            # 1. 일반 주식(주권)이 아닌 경우
            if self.obj_code_mgr.GetStockSectionKind(code) != 1: 
                drop_section += 1
                continue
                
            # 2. 스팩(SPAC)인 경우
            if self.obj_code_mgr.IsSpac(code): 
                drop_spac += 1
                continue
                
            # 3. 보통주가 아닌 경우 (코드 끝이 '0'이 아님)
            if code[-1] != '0': 
                drop_pref += 1
                continue
                
            # 4. 거래정지/중단 상태인 경우
            """
            value = object.GetStockStatusKind ( code )

            code 에해당하는주식상태를반환한다

            code : 주식코드
            반환값 : 관리구분
            typedefenum   {
            [helpstring("정상")]   CPC_STOCK_STATUS_NORMAL= 0,
            [helpstring("거래정지")]CPC_STOCK_STATUS_STOP= 1,
            [helpstring("거래중단")]CPC_STOCK_STATUS_BREAK= 2,
            }CPE_SUPERVISION_KIND;
            """
            if self.obj_code_mgr.GetStockStatusKind(code) != 0: 
                drop_status += 1
                continue
            
            # 5. 관리/주의 종목인 경우
            """
            value = object.GetStockControlKind ( code )

            code 에해당하는감리구분반환한다.

            code : 주식코드
            반환값 : 감리구분
            typedefenum {
            [helpstring("정상")]   CPC_CONTROL_NONE   = 0,
            [helpstring("주의")]   CPC_CONTROL_ATTENTION= 1,
            [helpstring("경고")]   CPC_CONTROL_WARNING= 2,
            [helpstring("위험예고")]CPC_CONTROL_DANGER_NOTICE= 3,
            [helpstring("위험")]   CPC_CONTROL_DANGER= 4,
            [helpstring("경고예고")]   CPC_CONTROL_WARNING_NOTICE= 5,
            }CPE_CONTROL_KIND;
            """
            if self.obj_code_mgr.GetStockControlKind(code) != 0: 
                drop_control += 1
                continue

            clean_tickers.append(code)

        print("-" * 40)
        print(f"🛠️ [기본 필터링 결과 요약]")
        print(f"- 주권 아님(ETF 등) 탈락: {drop_section}개")
        print(f"- 스팩(SPAC) 탈락: {drop_spac}개")
        print(f"- 우선주 탈락: {drop_pref}개")
        print(f"- 거래정지 탈락: {drop_status}개")
        print(f"- 관리/주의종목 탈락: {drop_control}개")
        print("-" * 40)
        print(f"[시스템] 1차 생존 순수 보통주: 총 {len(clean_tickers)}개")
        
        return clean_tickers

    
# --- 실행부 ---
if __name__ == "__main__":
    ticker_mgr = CodeManager(save=True)
    print(ticker_mgr.tickers_list_df)
    
