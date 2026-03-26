import win32com.client

class CpCybosStatus:
    """CYBOS의 각종 상태(연결, 서버 종류) 및 요청 제한을 확인하는 클래스"""
    
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpCybos")

    def is_connected(self):
        """CYBOS의 통신 연결 상태를 반환합니다. (0: 끊김, 1: 정상)"""
        return self.obj.IsConnect == 1

    def get_server_type(self):
        """
        연결된 서버 종류를 반환합니다.
        0: 연결끊김, 1: cybosplus 서버, 2: HTS 보통서버
        """
        return self.obj.ServerType

    def get_limit_remain_count(self, limit_type):
        """
        요청 제한까지 남은 개수를 반환합니다.
        limit_type: 
            0 (LT_TRADE_REQUEST): 주문/계좌 관련
            1 (LT_NONTRADE_REQUEST): 시세/조회 관련
            2 (LT_SUBSCRIBE): 실시간 시세(Subscribe) 관련
        """
        # 정수형 인자이므로 ord()를 쓰지 않습니다.
        return self.obj.GetLimitRemainCount(limit_type)

    def get_limit_remain_time(self, limit_type):
        """
        요청 제한이 해제될 때까지 남은 시간(ms)을 반환합니다.
        limit_type: 0 (주문/계좌), 1 (시세관련)
        """
        return self.obj.GetLimitRemainTime(limit_type)

    def get_total_limit_remain_time(self):
        """요청 개수를 재계산하기까지 남은 전체 시간(ms)을 반환합니다."""
        return self.obj.LimitRequestRemainTime

    def plus_disconnect(self):
        """Plus 연결 서비스를 종료합니다."""
        return self.obj.PlusDisconnect()

# --- 실시간 연결 끊김 이벤트 핸들러 ---
class CpCybosEvent:
    def OnDisConnect(self):
        """네트워크 장애 등으로 연결이 끊겼을 때 발생"""
        print("### [경고] CYBOS Plus와의 연결이 끊겼습니다. 프로그램을 안전하게 종료합니다. ###")
        # 여기에 로그 저장이나 알림 발송 로직을 추가할 수 있습니다.
        
