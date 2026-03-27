import json
import csv
import logging
from pathlib import Path

class FileManager:
    """
    [초경량 버전] 파이썬 내장 모듈만 사용하여 극강의 속도를 자랑하는 파일 매니저
    Pandas 의존성 제거
    """
    
    def __init__(self):
        self.logger = logging.getLogger("TradingBot.FileManager")

    def _ensure_dir(self, path: Path):
        if not path.parent.exists():
            path.parent.mkdir(parents=True, exist_ok=True)

    def save(self, data, file_path):
        """데이터(List of Dicts 또는 Dict)를 저장합니다."""
        path = Path(file_path)
        self._ensure_dir(path)
        ext = path.suffix.lower()

        try:
            # 1. JSON 저장 (가장 빠름, 타입 유지)
            if ext == '.json':
                with path.open('w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
            
            # 2. CSV 저장 (엑셀용, List of Dicts 형태만 지원)
            elif ext == '.csv':
                with path.open('w', newline='', encoding='utf-8-sig') as f:
                    if not data: return True # 빈 데이터면 패스
                    writer = csv.DictWriter(f, fieldnames=data[0].keys())
                    writer.writeheader()
                    writer.writerows(data)
            else:
                self.logger.error(f"❌ 지원하지 않는 확장자: {ext}")
                return False

            self.logger.debug(f"✅ 파일 저장 완료: {path.name}")
            return True

        except Exception as e:
            self.logger.error(f"❌ 저장 오류 ({path.name}): {e}")
            return False

    def load(self, file_path):
        """파일을 읽어 Python 객체(List of Dicts 등)로 반환합니다."""
        path = Path(file_path)
        if not path.exists():
            return None

        ext = path.suffix.lower()

        try:
            # 1. JSON 불러오기 (빛의 속도)
            if ext == '.json':
                with path.open('r', encoding='utf-8') as f:
                    return json.load(f)
            
            # 2. CSV 불러오기 (다시 List of Dicts로 완벽 복원)
            elif ext == '.csv':
                with path.open('r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    return [row for row in reader] # 리스트 컴프리헨션으로 즉시 변환

        except Exception as e:
            self.logger.error(f"❌ 로드 오류 ({path.name}): {e}")
            return None