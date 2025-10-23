#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
라포랩스 현대카드 OAuth 자동화 (GitHub Actions 버전)
Gmail → HTML 첨부파일 다운로드 → 보안메일 처리 → 스프레드시트 업데이트
"""

import os
import sys
import time
import zipfile
import pandas as pd
import base64
import pickle
import json
from pathlib import Path
import logging
from datetime import datetime

# 라이브러리 import (에러 발생시 설치 가이드 출력)
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    import gspread
except ImportError as e:
    print(f"❌ 필수 라이브러리가 설치되어 있지 않습니다: {e}")
    print("\n다음 명령어로 설치해주세요:")
    print("pip install selenium pandas google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client gspread openpyxl xlrd")
    input("설치 후 Enter를 눌러주세요...")
    sys.exit(1)

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('hyundai_automation.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class HyundaiCardBot:
    """현대카드 OAuth 자동화 봇 (GitHub Actions 호환)"""
    
    def __init__(self):
        # 라포랩스 설정
        self.AUTH_CODE = "8701718"
        self.SPREADSHEET_ID = "1Uu_8ccg-dFfYwqxi7QJiuWjH1Ow7FG-wJirp8PO_A14"
        self.SHEET_NAME = "현대카드보유내역_RAW"
        
        # OAuth 스코프
        self.SCOPES = [
            'https://www.googleapis.com/auth/gmail.readonly',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        # 다운로드 경로 (GitHub Actions 환경 대응)
        if os.environ.get('CI'):
            self.download_path = "/tmp/hyundai_auto"
        else:
            self.download_path = os.path.join(os.path.expanduser("~"), "Downloads", "hyundai_auto")
        
        os.makedirs(self.download_path, exist_ok=True)
        
        print("🏢 라포랩스 현대카드 자동화 봇 시작 (GitHub Actions 버전)")
        
    def authenticate(self):
        """OAuth 2.0 인증"""
        print("🔐 OAuth 인증 중...")
        
        creds = None
        
        # 기존 토큰 확인
        if os.path.exists('token.pickle'):
            try:
                with open('token.pickle', 'rb') as token:
                    creds = pickle.load(token)
                print("기존 토큰 로드 완료")
            except:
                print("기존 토큰 로드 실패, 새 인증 진행")
                if os.path.exists('token.pickle'):
                    os.remove('token.pickle')
        
        # 토큰 갱신 또는 새 인증
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    print("토큰 갱신 중...")
                    creds.refresh(Request())
                    print("✅ 토큰 갱신 완료")
                except:
                    print("토큰 갱신 실패, 새 인증 진행")
                    creds = None
            
            if not creds or not creds.valid:
                print("새 OAuth 인증 시작...")
                
                # 클라이언트 자격증명 파일 찾기
                client_files = ['client_secret.json', 'credentials.json', 'oauth_credentials.json']
                client_file = None
                
                for file in client_files:
                    if os.path.exists(file):
                        client_file = file
                        print(f"OAuth 파일 발견: {client_file}")
                        break
                
                if not client_file:
                    print("❌ OAuth 클라이언트 자격증명 파일을 찾을 수 없습니다!")
                    print("다음 파일 중 하나를 현재 폴더에 두세요:")
                    for file in client_files:
                        print(f"  - {file}")
                    return None
                
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(client_file, self.SCOPES)
                    creds = flow.run_local_server(port=0)
                    print("✅ 새 인증 완료")
                except Exception as e:
                    print(f"❌ OAuth 인증 실패: {e}")
                    return None
            
            # 토큰 저장
            try:
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)
                print("토큰 저장 완료")
            except Exception as e:
                print(f"토큰 저장 실패: {e}")
        
        print("✅ OAuth 인증 완료")
        return creds
    
    def find_hyundai_email(self, gmail_service):
        """현대카드 이메일 찾기"""
        try:
            print("📧 현대카드 이메일 검색 중...")
            
            # 검색 쿼리들
            queries = [
                'from:"현대카드 MY COMPANY" subject:"라포랩스 보유내역" has:attachment newer_than:14d',
                'from:"현대카드 MY COMPANY" subject:"보유내역" has:attachment newer_than:14d',
                'from:"현대카드 MY COMPANY" has:attachment newer_than:21d',
                'from:"현대카드" subject:"라포랩스 보유내역" has:attachment newer_than:14d',
                'from:"MY COMPANY" subject:"보유내역" has:attachment newer_than:21d'
            ]   
            
            all_messages = []
            
            for query in queries:
                try:
                    print(f"  검색: {query}")
                    results = gmail_service.users().messages().list(
                        userId='me',
                        q=query,
                        maxResults=10
                    ).execute()
                    
                    messages = results.get('messages', [])
                    all_messages.extend(messages)
                    print(f"  발견: {len(messages)}개")
                    
                except Exception as e:
                    print(f"  검색 오류: {e}")
                    continue
            
            if not all_messages:
                print("❌ 현대카드 이메일을 찾을 수 없습니다.")
                print("Gmail에서 다음을 확인해주세요:")
                print("- 현대카드에서 보낸 이메일")
                print("- HTML 첨부파일이 있는 이메일")
                print("- 최근 3주 이내의 이메일")
                return None
            
            # 중복 제거
            unique_messages = {msg['id']: msg for msg in all_messages}
            latest_id = list(unique_messages.keys())[0]
            
            print(f"✅ 이메일 발견: 총 {len(unique_messages)}개, 최신 선택")
            return latest_id
            
        except Exception as e:
            print(f"❌ 이메일 검색 실패: {e}")
            return None
    
    def download_html_attachment(self, gmail_service, message_id):
        """HTML 첨부파일 다운로드"""
        try:
            print("📥 HTML 첨부파일 다운로드 중...")
            
            # 메시지 상세 정보
            message = gmail_service.users().messages().get(
                userId='me',
                id=message_id,
                format='full'
            ).execute()
            
            # 제목 확인
            headers = message['payload'].get('headers', [])
            subject = ""
            sender = ""
            
            for header in headers:
                if header['name'].lower() == 'subject':
                    subject = header['value']
                elif header['name'].lower() == 'from':
                    sender = header['value']
            
            print(f"📧 제목: {subject}")
            print(f"📧 발신자: {sender}")
            
            # HTML 첨부파일 찾기
            def find_html_attachment(part):
                if part.get('filename') and part.get('filename').lower().endswith(('.html', '.htm')):
                    return part
                if 'parts' in part:
                    for subpart in part['parts']:
                        result = find_html_attachment(subpart)
                        if result:
                            return result
                return None
            
            html_attachment = find_html_attachment(message['payload'])
            
            if not html_attachment:
                print("❌ HTML 첨부파일을 찾을 수 없습니다.")
                return None
            
            filename = html_attachment.get('filename', 'secure_mail.html')
            attachment_id = html_attachment.get('body', {}).get('attachmentId')
            
            if not attachment_id:
                print("❌ 첨부파일 ID를 찾을 수 없습니다.")
                return None
            
            print(f"📄 HTML 파일: {filename}")
            
            # 첨부파일 다운로드
            attachment = gmail_service.users().messages().attachments().get(
                userId='me',
                messageId=message_id,
                id=attachment_id
            ).execute()
            
            # Base64 디코딩 및 저장
            file_data = base64.urlsafe_b64decode(attachment['data'])
            
            safe_filename = f"hyundai_secure_{int(time.time())}.html"
            local_path = os.path.join(self.download_path, safe_filename)
            
            with open(local_path, 'wb') as f:
                f.write(file_data)
            
            file_size = os.path.getsize(local_path)
            print(f"✅ HTML 다운로드 완료: {safe_filename} ({file_size} bytes)")
            
            return local_path
            
        except Exception as e:
            print(f"❌ HTML 다운로드 실패: {e}")
            return None
    
    def process_secure_email(self, html_file):
        """보안메일 처리 (Selenium Manager 사용)"""
        driver = None
        try:
            print("🔐 보안메일 처리 시작...")
            print("🤖 Selenium Manager가 자동으로 ChromeDriver를 관리합니다")
            
            # Chrome 설정
            chrome_options = Options()
            chrome_options.add_experimental_option("prefs", {
                "download.default_directory": self.download_path,
                "download.prompt_for_download": False,
            })
            chrome_options.add_argument("--window-size=1400,900")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            
            # GitHub Actions 환경에서는 headless 필수
            if os.environ.get('CI'):
                chrome_options.add_argument("--headless")
                print("💻 GitHub Actions 환경 감지 - Headless 모드 사용")
            
            # Selenium Manager가 자동으로 ChromeDriver 다운로드 및 관리
            driver = webdriver.Chrome(options=chrome_options)
            wait = WebDriverWait(driver, 20)
            
            print("✅ Chrome 브라우저 실행 완료")
            
            # HTML 파일 열기
            file_url = f"file://{os.path.abspath(html_file)}"
            print(f"📄 보안메일 열기: {os.path.basename(html_file)}")
            driver.get(file_url)
            
            time.sleep(5)
            
            # 현대카드 특별 처리: p2_temp 필드 클릭 후 password 필드로 전환
            print("🔍 현대카드 인증번호 입력 필드 처리 중...")
            
            try:
                # 1단계: p2_temp 필드 클릭 (플레이스홀더 필드)
                temp_input = driver.find_element(By.CSS_SELECTOR, "input[name='p2_temp']")
                if temp_input.is_displayed():
                    print("✅ 플레이스홀더 필드 발견: p2_temp")
                    temp_input.click()
                    time.sleep(2)  # JavaScript 실행 대기
                    
                    # 2단계: 실제 password 필드가 나타났는지 확인
                    try:
                        password_input = driver.find_element(By.CSS_SELECTOR, "input[name='p2']")
                        if password_input.is_displayed():
                            print("✅ 패스워드 필드로 전환 완료")
                            
                            # 3단계: 인증번호 입력
                            password_input.clear()
                            time.sleep(0.5)
                            password_input.send_keys(self.AUTH_CODE)
                            print(f"✅ 인증번호 입력 완료: {self.AUTH_CODE}")
                            
                            # 엔터키로 폼 제출
                            from selenium.webdriver.common.keys import Keys
                            password_input.send_keys(Keys.RETURN)
                            print("✅ 엔터키로 폼 제출 완료")
                            
                        else:
                            # password 필드가 숨겨져 있다면 JavaScript로 강제 입력
                            print("패스워드 필드가 숨겨져 있음, JavaScript로 입력 시도")
                            driver.execute_script(f"document.getElementsByName('p2')[0].value = '{self.AUTH_CODE}';")
                            print(f"✅ JavaScript로 인증번호 입력: {self.AUTH_CODE}")
                            
                            # 폼 제출
                            driver.execute_script("document.getElementById('decForm').submit();")
                            print("✅ JavaScript로 폼 제출 완료")
                            
                    except Exception as pwd_error:
                        print(f"패스워드 필드 처리 중 오류: {pwd_error}")
                        # 대체 방법: temp 필드에 직접 입력
                        temp_input.clear()
                        temp_input.send_keys(self.AUTH_CODE)
                        print(f"✅ 임시 필드에 인증번호 입력: {self.AUTH_CODE}")
                        
                        # 엔터키로 제출
                        from selenium.webdriver.common.keys import Keys
                        temp_input.send_keys(Keys.RETURN)
                        print("✅ 엔터키로 폼 제출 완료")
                
                else:
                    print("❌ p2_temp 필드를 찾을 수 없습니다.")
                    return None
                    
            except Exception as e:
                print(f"❌ 현대카드 입력 필드 처리 실패: {e}")
                print("일반적인 방법으로 재시도...")
                
                # 일반적인 방법으로 재시도
                selectors = [
                    "input[type='password']",
                    "input[name='p2']",
                    "input[name='p2_temp']",
                    "input[type='text']"
                ]
                
                auth_input = None
                for selector in selectors:
                    try:
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                        for element in elements:
                            if element.is_displayed() and element.is_enabled():
                                auth_input = element
                                print(f"✅ 입력 필드 발견: {selector}")
                                break
                        if auth_input:
                            break
                    except:
                        continue
                
                if not auth_input:
                    print("❌ 인증번호 입력 필드를 찾을 수 없습니다.")
                    return None
                
                # 입력 후 엔터키로 폼 제출
                try:
                    from selenium.webdriver.common.keys import Keys
                    auth_input.click()
                    time.sleep(1)
                    auth_input.clear()
                    time.sleep(1)
                    auth_input.send_keys(self.AUTH_CODE)
                    print(f"✅ 인증번호 입력: {self.AUTH_CODE}")
                    
                    # 엔터키로 폼 제출
                    auth_input.send_keys(Keys.RETURN)
                    print("✅ 엔터키로 폼 제출 완료")
                except Exception as input_error:
                    print(f"❌ 인증번호 입력 실패: {input_error}")
                    return None
            
            print("⏳ 폼 제출 후 페이지 로딩 대기...")
            time.sleep(10)
            
            # ZIP 다운로드 링크 찾기
            print("📦 ZIP 다운로드 링크 검색...")
            
            zip_selectors = [
                "a[href*='.zip']",
                "a[download*='.zip']",
                "a[href*='download']",
                "a"
            ]
            
            for selector in zip_selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        if element.is_displayed():
                            href = element.get_attribute('href') or ''
                            text = element.text.strip()
                            if '.zip' in href.lower() or '.zip' in text.lower() or '다운로드' in text:
                                try:
                                    element.click()
                                    print("✅ ZIP 다운로드 시작")
                                    break
                                except:
                                    driver.execute_script("arguments[0].click();", element)
                                    print("✅ ZIP 다운로드 시작 (JS)")
                                    break
                    else:
                        continue
                    break
                except:
                    continue
            
            # ZIP 다운로드 대기
            print("⏳ ZIP 다운로드 대기...")
            
            start_time = time.time()
            while (time.time() - start_time) < 60:
                time.sleep(3)
                
                zip_files = list(Path(self.download_path).glob("*.zip"))
                downloading = list(Path(self.download_path).glob("*.crdownload"))
                
                if zip_files and not downloading:
                    latest_zip = max(zip_files, key=lambda x: x.stat().st_mtime)
                    if latest_zip.stat().st_size > 0:
                        print(f"✅ ZIP 다운로드 완료: {latest_zip.name}")
                        return str(latest_zip)
                
                elapsed = int(time.time() - start_time)
                if elapsed % 15 == 0:
                    print(f"  대기 중... {elapsed}초")
            
            print("❌ ZIP 다운로드 타임아웃")
            return None
            
        except Exception as e:
            print(f"❌ 보안메일 처리 실패: {e}")
            import traceback
            traceback.print_exc()
            return None
        finally:
            if driver:
                try:
                    driver.quit()
                    print("🔒 브라우저 종료")
                except:
                    pass
    
    def extract_and_process_data(self, zip_file):
        """ZIP 압축해제 및 데이터 처리"""
        try:
            print(f"📦 ZIP 파일 처리: {os.path.basename(zip_file)}")
            
            zip_path = Path(zip_file)
            extract_path = zip_path.parent / f"{zip_path.stem}_extracted"
            
            # 기존 폴더 삭제
            if extract_path.exists():
                import shutil
                shutil.rmtree(extract_path)
            
            extract_path.mkdir()
            
            # 압축 해제
            try:
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    zip_ref.extractall(extract_path)
                print("✅ 압축 해제 완료")
            except:
                try:
                    # 한글 파일명 처리를 위한 대체 방법
                    import zipfile
                    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                        for member in zip_ref.namelist():
                            try:
                                zip_ref.extract(member, extract_path)
                            except:
                                # 파일명 문제시 임시 이름으로 저장
                                with zip_ref.open(member) as source:
                                    temp_name = f"temp_{len(member)}.tmp"
                                    with open(os.path.join(extract_path, temp_name), 'wb') as target:
                                        target.write(source.read())
                    print("✅ 압축 해제 완료 (대체 방법)")
                except Exception as e:
                    print(f"❌ 압축 해제 실패: {e}")
                    return None
            
            # 엑셀 파일 찾기
            excel_files = list(extract_path.rglob("*.xlsx")) + list(extract_path.rglob("*.xls"))
            
            if not excel_files:
                print("❌ 엑셀 파일을 찾을 수 없습니다.")
                return None
            
            excel_file = excel_files[0]
            print(f"📊 엑셀 파일: {excel_file.name}")
            
            # 엑셀 데이터 읽기
            xl = pd.ExcelFile(excel_file)
            sheet_names = xl.sheet_names
            print(f"📋 시트 목록: {sheet_names}")
            
            # 두 번째 시트 우선
            sheet_name = sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]
            print(f"선택된 시트: {sheet_name}")
            
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            print(f"✅ 데이터 읽기 완료: {len(df)}행 × {len(df.columns)}열")
            return df
            
        except Exception as e:
            print(f"❌ ZIP 처리 실패: {e}")
            return None
    
    def update_spreadsheet(self, gspread_client, data):
        """구글 스프레드시트 업데이트 (기본 버전)"""
        try:
            print("📝 구글 스프레드시트 업데이트...")
            
            spreadsheet = gspread_client.open_by_key(self.SPREADSHEET_ID)
            
            try:
                worksheet = spreadsheet.worksheet(self.SHEET_NAME)
            except:
                worksheet = spreadsheet.add_worksheet(title=self.SHEET_NAME, rows=1000, cols=26)
                print(f"새 워크시트 생성: {self.SHEET_NAME}")
            
            # 기존 데이터 삭제
            worksheet.clear()
            
            # 새 데이터 업로드
            headers = data.columns.tolist()
            values = data.values.tolist()
            
            # None 값을 빈 문자열로 변환
            clean_values = []
            for row in values:
                clean_row = ['' if pd.isna(cell) else str(cell) for cell in row]
                clean_values.append(clean_row)
            
            all_data = [headers] + clean_values
            
            worksheet.update('A1', all_data)
            
            print(f"✅ 스프레드시트 업데이트 완료: {len(data)}행")
            return True
            
        except Exception as e:
            print(f"❌ 스프레드시트 업데이트 실패: {e}")
            return False
    
    def run_automation(self):
        """전체 자동화 실행"""
        print("🚀 현대카드 자동화 시작!")
        print("="*60)
        
        start_time = time.time()
        
        try:
            # 1. OAuth 인증
            print("\n1️⃣ OAuth 인증...")
            creds = self.authenticate()
            if not creds:
                return False
            
            # 2. Google 서비스 생성
            print("\n2️⃣ Google 서비스 연결...")
            gmail_service = build('gmail', 'v1', credentials=creds)
            gspread_client = gspread.authorize(creds)
            
            # 3. 이메일 검색
            print("\n3️⃣ 현대카드 이메일 검색...")
            message_id = self.find_hyundai_email(gmail_service)
            if not message_id:
                return False
            
            # 4. HTML 다운로드
            print("\n4️⃣ HTML 첨부파일 다운로드...")
            html_file = self.download_html_attachment(gmail_service, message_id)
            if not html_file:
                return False
            
            # 5. 보안메일 처리
            print("\n5️⃣ 보안메일 처리...")
            zip_file = self.process_secure_email(html_file)
            if not zip_file:
                return False
            
            # 6. 데이터 처리
            print("\n6️⃣ 데이터 추출...")
            data = self.extract_and_process_data(zip_file)
            if data is None:
                return False
            
            # 7. 스프레드시트 업데이트
            print("\n7️⃣ 스프레드시트 업데이트...")
            success = self.update_spreadsheet(gspread_client, data)
            
            if success:
                elapsed = int(time.time() - start_time)
                print("\n" + "="*60)
                print("🎉 자동화 완료!")
                print(f"⏱️  소요시간: {elapsed}초")
                print(f"📊 데이터: {len(data)}행 × {len(data.columns)}열")
                print(f"🔗 링크: https://docs.google.com/spreadsheets/d/{self.SPREADSHEET_ID}")
                return True
            
            return False
            
        except Exception as e:
            print(f"❌ 자동화 실패: {e}")
            import traceback
            traceback.print_exc()
            return False

def main():
    """메인 실행 함수"""
    print("🏢 라포랩스 현대카드 OAuth 자동화")
    print("📧 Gmail → HTML → 보안메일 → Google Sheets")
    print("🤖 Selenium Manager 자동 ChromeDriver 관리")
    print("="*50)
    
    print("\n📋 확인사항:")
    print("✅ client_secret.json 파일")
    print("✅ Chrome 브라우저 (자동 버전 관리)")
    print("✅ 현대카드 HTML 첨부파일 이메일")
    
    # CI 환경이면 자동 실행
    if os.environ.get('CI'):
        print("\n🤖 CI 환경에서 자동 실행 중...")
    else:
        response = input("\n🚀 자동화를 시작하시겠습니까? (y/N): ").strip().lower()
        if response != 'y':
            print("자동화를 취소했습니다.")
            return
    
    try:
        bot = HyundaiCardBot()
        success = bot.run_automation()
        
        if success:
            print("\n🎊 자동화 성공!")
        else:
            print("\n😞 자동화 실패")
            print("hyundai_automation.log 파일을 확인해주세요.")
    
    except KeyboardInterrupt:
        print("\n⏹️ 사용자 중단")
    except Exception as e:
        print(f"\n💥 오류: {e}")
        import traceback
        traceback.print_exc()
    
    # CI 환경이 아닐 때만 입력 대기
    if not os.environ.get('CI'):
        input("\nEnter를 눌러 종료...")

if __name__ == "__main__":
    main()
