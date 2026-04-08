import requests
from bs4 import BeautifulSoup

# 1. 대상 설정: 삼성전자 (A005930)
code = "A064760"
url = f"https://comp.fnguide.com/SVO2/asp/SVD_Main.asp?gicode={code}"
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

print(f"🔍 '{code}'에서 '유동주식수/비율' 데이터 추출 중...")

try:
    res = requests.get(url, headers=headers, timeout=10)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, 'html.parser')
    
    # 2. '유동주식수/비율' 텍스트를 포함하는 th 태그를 직접 찾습니다.
    # replace(" ", "")를 통해 혹시 모를 공백까지 제거하고 비교합니다.
    target_th = None
    for th in soup.find_all('th'):
        if '유동주식수/비율' in th.get_text(strip=True).replace(" ", ""):
            target_th = th
            break

    if target_th:
        # th 바로 옆에 있는 td에서 데이터를 가져옵니다.
        td_val = target_th.find_next_sibling('td').get_text(strip=True)
        print(f"📊 수신 데이터: {td_val}")
        
        # 슬래시(/) 뒤의 비율만 추출 (예: "5,969,782,550 / 74.32" -> "74.32")
        floating_ratio = td_val.split('/')[-1].strip()
        print(f"✅ 최종 유동비율 확인: {floating_ratio}%")
    else:
        print("❌ 여전히 항목을 찾지 못했습니다. 상단 'Snapshot' 탭이 맞는지 확인이 필요합니다.")

except Exception as e:
    print(f"❗ 에러 발생: {e}")