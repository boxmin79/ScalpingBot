class CpEvent:
    """대신증권 COM 객체의 실시간 수신 이벤트를 처리하는 공통 핸들러"""
    def set_params(self, callback_func):
        # 데이터가 들어오면 실행할 파이썬 함수를 저장합니다.
        self.callback_func = callback_func

    def OnReceived(self):
        # 대신증권 서버에서 이벤트가 발생하면 자동 호출됩니다.
        if hasattr(self, 'callback_func') and self.callback_func:
            self.callback_func()