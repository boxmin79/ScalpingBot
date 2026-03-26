class CpEvent:
    """실시간 이벤트 처리를 위한 공통 핸들러 클래스"""
    def set_params(self, client, name):
        self.client = client  # 호출한 객체 (콜백을 받을 객체)
        self.name = name

    def OnReceived(self):
        """데이터를 수신했을 때 발생하는 이벤트 [cite: 536, 1554, 1812]"""
        if hasattr(self.client, 'process_received'):
            self.client.process_received(self.name)
            