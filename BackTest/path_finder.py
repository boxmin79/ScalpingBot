from pathlib import Path
import sys
import importlib
import os
from dotenv import load_dotenv # 패키지 설치 필요: pip install python-dotenv

def setup_project_path():
    current = Path(__file__).resolve()
    # 상위 폴더를 탐색하며 "requirements.txt"가 있는 루트를 찾음
    for parent in [current] + list(current.parents):
        if (parent / "requirements.txt").exists():
            root_path = parent
            # 1. 시스템 경로(sys.path)에 추가
            if str(root_path) not in sys.path:
                sys.path.insert(0, str(root_path))
            
            # 2. .env 파일이 있다면 시스템 환경변수로 등록
            env_path = root_path / ".env"
            if env_path.exists():
                load_dotenv(dotenv_path=env_path)
                # print(f"✅ [System] 환경변수 등록 완료: {env_path}")
            
            return root_path
    return None

# 파일이 불러와지는 시점에 자동으로 경로 설정 및 .env 등록 실행
PROJECT_ROOT = setup_project_path()

def get_cfg():
    """
    루트에 있는 config.py 모듈을 동적으로 가져와 반환합니다.
    """
    if PROJECT_ROOT is None:
        print("❌ 오류: 프로젝트 루트(config.py)를 찾을 수 없습니다.")
        return None
    
    try:
        import path_config
        importlib.reload(path_config)
        return path_config
    except ImportError:
        print("❌ 오류: config.py 파일을 찾을 수 없습니다.")
        return None

if __name__ == "__main__":
    # 테스트용: .env에 작성한 변수가 잘 나오는지 확인
    # 예: .env에 MY_SECRET=1234 가 있다면
    print(f"프로젝트 루트: {PROJECT_ROOT}")
    print(f"환경변수 테스트: {os.getenv('TELEGRAM_API')}")