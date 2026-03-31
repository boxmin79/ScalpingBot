from pathlib import Path

# BASE_DIR
BASE_DIR = BASE_DIR = Path(__file__).resolve().parent

# .env
ENV_PATH = BASE_DIR / '.env'

# DATA_DIR
DATA_DIR = BASE_DIR / 'data'
CODE_DIR = DATA_DIR / 'code'
CHART_DIR = DATA_DIR / 'chart'
INDEX_CHART_DIR = CHART_DIR / 'index'
STOCK_CHART_DIR = CHART_DIR / 'stock'
BACKTEST_DIR = DATA_DIR / 'backtest'
SNAPSHOT_DIR = DATA_DIR / 'snapshot'

# 파일 경로
CODE_LIST_PATH = CODE_DIR / 'code_list.csv'

# LOGS_DIR
LOGS_DIR = BASE_DIR / 'logs'

# API_DIR
API_DIR = BASE_DIR / 'API'

# 3. 폴더 자동 생성 (선택 사항이지만 강력 추천)
# 코드를 실행할 때 해당 폴더가 없다면 알아서 만들어줍니다.
# exist_ok=True 덕분에 이미 폴더가 있어도 에러가 나지 않습니다.
make_dirs = [DATA_DIR, CODE_DIR, 
             CHART_DIR, INDEX_CHART_DIR, INDEX_CHART_DIR,
             BACKTEST_DIR, SNAPSHOT_DIR, LOGS_DIR]
for directory in make_dirs:
    directory.mkdir(parents=True, exist_ok=True)   

if __name__ == '__main__':
    print(BASE_DIR)
    print(DATA_DIR)
    print(CODE_DIR)
    print(CHART_DIR)
    print(BACKTEST_DIR) 
    print(LOGS_DIR)