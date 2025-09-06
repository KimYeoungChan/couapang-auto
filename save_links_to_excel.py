import pyperclip
import pandas as pd
from datetime import datetime
import os

# 클립보드에서 링크 가져오기
clipboard_content = pyperclip.paste()

# 링크들을 줄바꿈으로 분리 (여러 링크가 있을 경우)
links = [link.strip() for link in clipboard_content.split('\n') if link.strip()]

# 데이터프레임 생성
df = pd.DataFrame({
    '링크': links,
    '저장시간': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * len(links)
})

# 엑셀 파일로 저장
desktop_path = os.path.join(os.path.expanduser('~'), 'OneDrive', '바탕 화면')
file_path = os.path.join(desktop_path, f'links_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')

df.to_excel(file_path, index=False, engine='openpyxl')
print(f"링크가 저장되었습니다: {file_path}")
print(f"저장된 링크 개수: {len(links)}")