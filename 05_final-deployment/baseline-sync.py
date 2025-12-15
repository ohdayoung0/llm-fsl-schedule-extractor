import os
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
from datetime import datetime, timedelta
import time

# --- 1. VS Code 로컬 인증 설정 및 함수 임포트 ---
# 로컬 파일 기반 인증 라이브러리
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials 

# 다운로드한 JSON 파일
CLIENT_SECRET_FILE = '../client_secret.json' 
TOKEN_FILE = '../token.json'
SCOPES = ['https://www.googleapis.com/auth/calendar'] 
TIMEZONE = 'Asia/Seoul'

def authenticate_local():
    """로컬 환경에서 Google Calendar API 인증을 수행하고 서비스 객체를 반환합니다."""
    creds = None
    # 토큰 파일이 있다면 로드하여 재사용
    if os.path.exists('TOKEN_FILE'):
        creds = Credentials.from_authorized_user_file('TOKEN_FILE', SCOPES)
    
    # 토큰이 없거나 만료되었다면 재인증 실행
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            
        # 다음 실행을 위해 토큰을 'token.json' 파일로 저장
        with open('TOKEN_FILE', 'w') as token:
            token.write(creds.to_json())
            
    return build('calendar', 'v3', credentials=creds)

# --- 2. 이전 크롤링 및 정제 함수 정의 ---

# 2-1. 크롤링 함수
def crawl_notice_list(list_url):
    response = requests.get(list_url)
    response.raise_for_status() 
    soup = BeautifulSoup(response.text, 'html.parser')
    notice_rows = soup.select('tr.notice') 
    
    data = []
    for row in notice_rows:
        title_tag = row.select_one('a')
        if title_tag:
            title = title_tag.text.strip()
            full_url = title_tag['href'].replace('&amp;', '&')
            
            # 제목에서 마감일 패턴 추출
            date_match = re.search(r'~ *(\d{1,2}[./]\d{1,2}(?:\s+\d{1,2}:\d{2})?(?:\s+\S+)?)', title)
            deadline_str = '~' + date_match.group(1).strip() if date_match else None
            
            data.append({
                '제목': title,
                'URL': full_url,
                '마감일_원문': deadline_str,
                '주최기관': '정보 없음', 
                '장소': '온라인 또는 별도 명시 없음', 
                '대상': '전체 학생 또는 별도 명시 없음'
            })
    return pd.DataFrame(data)

# 2-2. 날짜 정제 함수
def format_deadline(date_str):
    if date_str is None or '정보 없음' in date_str or not date_str.strip():
        return None

    clean_date_str = date_str.replace('~', '').strip()
    now = datetime.now()
    current_year = now.year

    try: # 시간 정보 포함 형식 시도
        if re.search(r'\d{1,2}:\d{2}', clean_date_str):
            temp_date_str = clean_date_str.replace('/', '.')
            month_day_time_parts = temp_date_str.split(' ')
            month_day_str = month_day_time_parts[0]
            time_str = month_day_time_parts[1]

            deadline_month = int(month_day_str.split('.')[0])
            target_year = current_year + 1 if deadline_month < now.month else current_year
            full_date_str = f"{target_year}.{month_day_str} {time_str}"
            dt_object = datetime.strptime(full_date_str, '%Y.%m.%d %H:%M')
            return dt_object.isoformat()
    except ValueError:
        pass

    try: # 시간 정보 없는 형식 시도
        temp_date_str = clean_date_str.split(' ')[0].replace('/', '.') 
        month_day_str = temp_date_str

        deadline_month = int(month_day_str.split('.')[0])
        target_year = current_year + 1 if deadline_month < now.month else current_year
        full_date_str = f"{target_year}.{month_day_str}"
        dt_object = datetime.strptime(full_date_str, '%Y.%m.%d')
        # 마감일은 해당 날짜의 끝으로 설정 (23:59:59)
        return (dt_object + timedelta(days=1) - timedelta(seconds=1)).isoformat()
    except ValueError:
        return None

# --- 3. 캘린더 이벤트 생성 함수 ---

def create_calendar_event(row, service, calendar_id):
    """DataFrame의 행 데이터를 Google Calendar 이벤트로 생성합니다."""
    if pd.isna(row['마감일_정제']):
        return f"SKIP: {row['제목']} (날짜 정보 없음)"

    description = (
        f" 대상: {row['대상']}\n"
        f" 주최: {row['주최기관']}\n"
        f" 장소: {row['장소']}\n"
        f" 원본 링크: {row['URL']}"
    )
    
    # 날짜와 시간 설정
    start_time = row['마감일_정제']
    
    event = {
        'summary': f"[DEADLINE] {row['제목']}",
        'location': row['장소'],
        'description': description,
        'start': {
            'dateTime': start_time,
            'timeZone': TIMEZONE, # <--- [수정]: 시간대 필드 추가
        },
        'end': {
            'dateTime': start_time,
            'timeZone': TIMEZONE, # <--- [수정]: 시간대 필드 추가
        },
        'reminders': {
            'useDefault': False,
            'overrides': [{'method': 'popup', 'minutes': 60 * 24}],
        },
    }

    try:
        event = service.events().insert(calendarId=calendar_id, body=event).execute()
        return f"SUCCESS: {row['제목']} (ID: {event.get('id')})"
    except Exception as e:
        return f"FAILED: {row['제목']} (Error: {e})"


# --- 4. 메인 실행 로직 ---

if __name__ == '__main__':
    # 캘린더 ID 설정
    CALENDAR_ID = 'primary'
    LIST_URL = 'https://software.cbnu.ac.kr/index.php?mid=sub0401&category=8410'
    
    # 1. 로컬 인증 실행
    print("Google Calendar 로컬 인증을 시작합니다. 브라우저 창을 확인해 주세요.")
    # 인증 파일이 이미 있으므로 브라우저 창은 열리지 않고 바로 진행
    service = authenticate_local()
    print("인증 및 서비스 객체 생성 완료.")
    
    # 2. 크롤링 및 데이터 정제
    df_notices = crawl_notice_list(LIST_URL)
    df_notices['마감일_원문'] = df_notices['제목'].apply(lambda x: re.search(r'~ *(\d{1,2}[./]\d{1,2}(?:\s+\d{1,2}:\d{2})?(?:\s+\S+)?)', x).group(0) if re.search(r'~ *(\d{1,2}[./]\d{1,2}(?:\s+\d{1,2}:\d{2})?(?:\s+\S+)?)', x) else None)
    df_notices['마감일_정제'] = df_notices['마감일_원문'].apply(format_deadline)
    
    print(f"\n총 {len(df_notices)}개의 공지사항 데이터 정제 완료.")
    
    # 3. 캘린더 이벤트 생성
    print("\n--- Google Calendar 이벤트 생성 시작 ---")
    
    for index, row in df_notices.iterrows():
        result = create_calendar_event(row, service, CALENDAR_ID)
        print(result)
        time.sleep(0.5) # 서버 부하 방지
    
    print("--- Google Calendar 이벤트 생성 완료 ---")