#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
라포랩스 현대카드 OAuth 자동화 (CI 환경 최적화 버전)
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

# 라이브러리 import
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.keys import Keys
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    import gspread
except ImportError as e:
    print(f"필수 라이브러리가 설치되어 있지 않습니다: {e}")
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
logger = logging.getLogger(__name__)

class HyundaiCardBot:
    """현대카드 OAuth 자동화 봇"""
    
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
        
        # 다운로드 경로
        self.download_path = os.path.join(os.path.expanduser("~"), "Downloads", "hyundai_auto")
        os.makedirs(self.download_path, exist_ok=True)
        
        # CI 환경 감지
        self.is_ci = os.environ.get('CI') == 'true' or os.environ.get('GITHUB_ACTIONS') == 'true'
        
        logger.info("🏢 라포랩스 현대카드 자동화 봇 시작")
        logger.info(f"환경: {'CI (GitHub Actions)' if self.is_ci else 'Local'}")
        logger.info(f"다운로드 경로: {self.download_path}")
        
    def authenticate(self):
        """OAuth 2.0 인증"""
        logger.info("🔐 OAuth 인증 중...")
        
        creds = None
        
        # 기존 토큰 확인
        if os.path.exists('token.pickle'):
            try:
                with open('token.pickle', 'rb') as token:
                    creds = pickle.load(token)
                logger.info("✅ 기존 토큰 로드 완료")
            except Exception as e:
                logger.warning(f"기존 토큰 로드 실패: {e}")
                if os.path.exists('token.pickle'):
                    os.remove('token.pickle')
        
        # 토큰 갱신 또는 새 인증
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    logger.info("🔄 토큰 갱신 중...")
                    creds.refresh(Request())
                    logger.info("✅ 토큰 갱신 완료")
                except Exception as e:
                    logger.warning(f"토큰 갱신 실패: {e}")
                    creds = None
            
            if not creds or not creds.valid:
                logger.info("새 OAuth 인증 시작...")
                
                # 클라이언트 자격증명 파일 찾기
                client_files = ['client_secret.json', 'credentials.json', 'oauth_credentials.json']
                client_file = None
                
                for file in client_files:
                    if os.path.exists(file):
                        client_file = file
                        logger.info(f"OAuth 파일 발견: {client_file}")
                        break
                
                if not client_file:
                    logger.error("OAuth 클라이언트 자격증명 파일을 찾을 수 없습니다!")
                    return None
                
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(client_file, self.SCOPES)
                    creds = flow.run_local_server(port=0)
                    logger.info("✅ 새 인증 완료")
                except Exception as e:
                    logger.error(f"OAuth 인증 실패: {e}")
                    return None
            
            # 토큰 저장
            try:
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)
                logger.info("토큰 저장 완료")
            except Exception as e:
                logger.error(f"토큰 저장 실패: {e}")
        
        logger.info("✅ OAuth 인증 완료")
        return creds
    
    def find_hyundai_email(self, gmail_service):
        """현대카드 이메일 찾기"""
        try:
            logger.info("📧 현대카드 이메일 검색 중...")
            
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
                    logger.info(f"  검색 쿼리: {query[:50]}...")
                    results = gmail_service.users().messages().list(
                        userId='me',
                        q=query,
                        maxResults=10
                    ).execute()
                    
                    messages = results.get('messages', [])
                    all_messages.extend(messages)
                    logger.info(f"    → {len(messages)}개 발견")
                    
                except Exception as e:
                    logger.warning(f"  검색 오류: {e}")
                    continue
            
            if not all_messages:
                logger.error("❌ 현대카드 이메일을 찾을 수 없습니다.")
                return None
            
            # 중복 제거
            unique_messages = {msg['id']: msg for msg in all_messages}
            latest_id = list(unique_messages.keys())[0]
            
            logger.info(f"✅ 이메일 발견: 총 {len(unique_messages)}개, 최신 선택")
            return latest_id
            
        except Exception as e:
            logger.error(f"❌ 이메일 검색 실패: {e}")
            return None
    
    def download_html_attachment(self, gmail_service, message_id):
        """HTML 첨부파일 다운로드"""
        try:
            logger.info("📥 HTML 첨부파일 다운로드 중...")
            
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
            
            logger.info(f"📧 제목: {subject}")
            logger.info(f"📧 발신자: {sender}")
            
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
                logger.error("❌ HTML 첨부파일을 찾을 수 없습니다.")
                return None
            
            filename = html_attachment.get('filename', 'secure_mail.html')
            attachment_id = html_attachment.get('body', {}).get('attachmentId')
            
            if not attachment_id:
                logger.error("❌ 첨부파일 ID를 찾을 수 없습니다.")
                return None
            
            logger.info(f"📄 HTML 파일: {filename}")
            
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
            logger.info(f"✅ HTML 다운로드 완료: {safe_filename} ({file_size} bytes)")
            
            return local_path
            
        except Exception as e:
            logger.error(f"❌ HTML 다운로드 실패: {e}")
            return None
    
    def process_secure_email(self, html_file):
        """보안메일 처리"""
        driver = None
        try:
            logger.info("🔐 보안메일 처리 시작...")
            
            # HTML 파일 유효성 확인
            file_size = os.path.getsize(html_file)
            logger.info(f"📄 파일 크기: {file_size} bytes")
            
            if file_size < 10000:
                logger.warning("⚠️ HTML 파일이 너무 작습니다 (손상 가능성)")
                with open(html_file, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read(500)
                    logger.info(f"파일 시작: {content[:200]}")
            
            # Chrome 설정
            chrome_options = Options()
            chrome_options.add_experimental_option("prefs", {
                "download.default_directory": self.download_path,
                "download.prompt_for_download": False,
                "profile.default_content_settings.popups": 0,
            })
            
            if self.is_ci:
                logger.info("💻 GitHub Actions 환경 감지 - Headless 모드 활성화")
                chrome_options.add_argument("--headless=new")
                chrome_options.add_argument("--no-sandbox")
                chrome_options.add_argument("--disable-dev-shm-usage")
                chrome_options.add_argument("--disable-gpu")
                chrome_options.add_argument("--disable-extensions")
                chrome_options.add_argument("--disable-sync")
                chrome_options.add_argument("--disable-plugins")
                chrome_options.add_argument("--disable-application-cache")
                chrome_options.add_argument("--no-first-run")
            else:
                chrome_options.add_argument("--window-size=1400,900")
            
            # 바이너리 경로 자동 감지
            chrome_bin = os.environ.get('CHROME_BIN')
            chromedriver_path = os.environ.get('CHROMEDRIVER_PATH')
            
            if chrome_bin and os.path.exists(chrome_bin):
                chrome_options.binary_location = chrome_bin
                logger.info(f"✅ Chrome 바이너리: {chrome_bin}")
            
            if chromedriver_path and os.path.exists(chromedriver_path):
                logger.info(f"✅ ChromeDriver: {chromedriver_path}")
                service = Service(chromedriver_path)
                driver = webdriver.Chrome(service=service, options=chrome_options)
            else:
                logger.info("✅ Selenium Manager로 자동 관리")
                driver = webdriver.Chrome(options=chrome_options)
            
            logger.info("✅ Chrome 브라우저 초기화 성공")
            
            # HTML 파일 열기
            file_url = f"file://{os.path.abspath(html_file)}"
            logger.info(f"📄 HTML 파일 로드: {os.path.basename(html_file)}")
            
            driver.get(file_url)
            time.sleep(5)
            
            # 페이지 HTML 확인
            page_html = driver.page_source
            logger.info(f"🔍 페이지 HTML 분석...")
            logger.info(f"페이지 HTML 일부: {page_html[:200]}")
            
            # 페이지의 모든 input 요소 개수
            all_inputs = driver.find_elements(By.TAG_NAME, "input")
            logger.info(f"전체 input 요소 개수: {len(all_inputs)}")
            
            if len(all_inputs) == 0:
                logger.warning("⚠️ 입력 필드가 없습니다. 페이지 새로고침 시도...")
                driver.refresh()
                time.sleep(5)
                all_inputs = driver.find_elements(By.TAG_NAME, "input")
                logger.info(f"새로고침 후 input 요소 개수: {len(all_inputs)}")
            
            # 인증번호 입력 필드 찾기
            logger.info("🔍 인증번호 입력 필드 찾기...")
            
            auth_input = None
            
            # 방법 1: p2_temp 클릭 후 p2에 입력
            try:
                logger.info("  시도 1: p2_temp → p2 방식")
                temp_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.NAME, "p2_temp"))
                )
                if temp_input.is_displayed():
                    temp_input.click()
                    time.sleep(2)
                    
                    try:
                        password_input = driver.find_element(By.NAME, "p2")
                        if password_input.is_displayed():
                            auth_input = password_input
                            logger.info("  ✅ p2 필드로 전환 성공")
                    except:
                        pass
            except:
                pass
            
            # 방법 2: p2 직접 접근
            if not auth_input:
                try:
                    logger.info("  시도 2: p2 직접 접근")
                    auth_input = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.NAME, "p2"))
                    )
                    if auth_input.is_displayed():
                        logger.info("  ✅ p2 필드 발견")
                except:
                    pass
            
            # 방법 3: CSS 선택자로 검색
            if not auth_input:
                selectors = [
                    ("input[type='password']", "password 타입"),
                    ("input[name='p2_temp']", "p2_temp"),
                    ("input[type='text']", "text 타입"),
                    ("input[placeholder*='번호']", "placeholder 번호"),
                    ("input[placeholder*='인증']", "placeholder 인증"),
                ]
                
                for selector, desc in selectors:
                    try:
                        logger.info(f"  시도: {desc}")
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                        for elem in elements:
                            if elem.is_displayed() and elem.is_enabled():
                                auth_input = elem
                                logger.info(f"  ✅ {desc} 필드 발견")
                                break
                        if auth_input:
                            break
                    except:
                        pass
            
            if not auth_input:
                logger.error("❌ 인증번호 입력 필드를 찾을 수 없습니다.")
                logger.error("페이지 소스를 저장합니다...")
                
                # 디버그 파일 저장
                debug_file = os.path.join(self.download_path, "debug_page_source.html")
                with open(debug_file, 'w', encoding='utf-8') as f:
                    f.write(page_html)
                logger.info(f"디버그 파일: {debug_file}")
                
                # 모든 input 요소 정보
                logger.info("페이지의 모든 input 요소:")
                for i, inp in enumerate(all_inputs[:10]):
                    name = inp.get_attribute('name')
                    type_attr = inp.get_attribute('type')
                    placeholder = inp.get_attribute('placeholder')
                    logger.info(f"  [{i}] name={name}, type={type_attr}, placeholder={placeholder}")
                
                return None
            
            # 인증번호 입력
            logger.info(f"✍️ 인증번호 입력: {self.AUTH_CODE}")
            auth_input.click()
            time.sleep(0.5)
            auth_input.clear()
            time.sleep(0.5)
            auth_input.send_keys(self.AUTH_CODE)
            time.sleep(1)
            
            logger.info("✅ 인증번호 입력 완료")
            
            # 폼 제출
            logger.info("📤 폼 제출...")
            auth_input.send_keys(Keys.RETURN)
            time.sleep(10)
            
            # ZIP 다운로드 대기
            logger.info("📦 ZIP 다운로드 대기...")
            
            start_time = time.time()
            max_wait = 120  # 2분
            
            while (time.time() - start_time) < max_wait:
                time.sleep(2)
                
                zip_files = list(Path(self.download_path).glob("*.zip"))
                downloading = list(Path(self.download_path).glob("*.crdownload"))
                
                if zip_files and not downloading:
                    latest_zip = max(zip_files, key=lambda x: x.stat().st_mtime)
                    if latest_zip.stat().st_size > 0:
                        logger.info(f"✅ ZIP 다운로드 완료: {latest_zip.name}")
                        return str(latest_zip)
                
                elapsed = int(time.time() - start_time)
                if elapsed % 20 == 0:
                    logger.info(f"  대기 중... {elapsed}초")
            
            logger.error("❌ ZIP 다운로드 타임아웃")
            return None
            
        except Exception as e:
            logger.error(f"❌ 보안메일 처리 실패: {e}")
            import traceback
            traceback.print_exc()
            return None
        finally:
            if driver:
                try:
                    driver.quit()
                    logger.info("🔒 브라우저 종료")
                except:
                    pass
    
    def extract_and_process_data(self, zip_file):
        """ZIP 압축해제 및 데이터 처리"""
        try:
            logger.info(f"📦 ZIP 파일 처리: {os.path.basename(zip_file)}")
            
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
                logger.info("✅ 압축 해제 완료")
            except Exception as e:
                logger.warning(f"압축 해제 실패, 대체 방법 시도: {e}")
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    for member in zip_ref.namelist():
                        try:
                            zip_ref.extract(member, extract_path)
                        except:
                            pass
                logger.info("✅ 압축 해제 완료 (대체 방법)")
            
            # 엑셀 파일 찾기
            excel_files = list(extract_path.rglob("*.xlsx")) + list(extract_path.rglob("*.xls"))
            
            if not excel_files:
                logger.error("❌ 엑셀 파일을 찾을 수 없습니다.")
                return None
            
            excel_file = excel_files[0]
            logger.info(f"📊 엑셀 파일: {excel_file.name}")
            
            # 엑셀 데이터 읽기
            xl = pd.ExcelFile(excel_file)
            sheet_names = xl.sheet_names
            logger.info(f"📋 시트 목록: {sheet_names}")
            
            # 두 번째 시트 우선
            sheet_name = sheet_names[1] if len(sheet_names) > 1 else sheet_names[0]
            logger.info(f"선택된 시트: {sheet_name}")
            
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            logger.info(f"✅ 데이터 읽기 완료: {len(df)}행 × {len(df.columns)}열")
            return df
            
        except Exception as e:
            logger.error(f"❌ ZIP 처리 실패: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def update_spreadsheet(self, gspread_client, data):
        """구글 스프레드시트 업데이트"""
        try:
            logger.info("📝 구글 스프레드시트 업데이트...")
            
            spreadsheet = gspread_client.open_by_key(self.SPREADSHEET_ID)
            
            try:
                worksheet = spreadsheet.worksheet(self.SHEET_NAME)
            except:
                worksheet = spreadsheet.add_worksheet(title=self.SHEET_NAME, rows=1000, cols=26)
                logger.info(f"새 워크시트 생성: {self.SHEET_NAME}")
            
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
            
            logger.info(f"✅ 스프레드시트 업데이트 완료: {len(data)}행")
            return True
            
        except Exception as e:
            logger.error(f"❌ 스프레드시트 업데이트 실패: {e}")
            return False
    
    def run(self):
        """전체 자동화 실행"""
        logger.info("🚀 현대카드 자동화 시작!")
        logger.info("="*60)
        
        start_time = time.time()
        
        try:
            # 1. OAuth 인증
            logger.info("\n1️⃣ OAuth 인증...")
            creds = self.authenticate()
            if not creds:
                return False
            
            # 2. Google 서비스 생성
            logger.info("\n2️⃣ Google 서비스 연결...")
            gmail_service = build('gmail', 'v1', credentials=creds)
            gspread_client = gspread.authorize(creds)
            
            # 3. 이메일 검색
            logger.info("\n3️⃣ 현대카드 이메일 검색...")
            message_id = self.find_hyundai_email(gmail_service)
            if not message_id:
                return False
            
            # 4. HTML 다운로드
            logger.info("\n4️⃣ HTML 첨부파일 다운로드...")
            html_file = self.download_html_attachment(gmail_service, message_id)
            if not html_file:
                return False
            
            # 5. 보안메일 처리
            logger.info("\n5️⃣ 보안메일 처리...")
            zip_file = self.process_secure_email(html_file)
            if not zip_file:
                return False
            
            # 6. 데이터 처리
            logger.info("\n6️⃣ 데이터 추출...")
            data = self.extract_and_process_data(zip_file)
            if data is None:
                return False
            
            # 7. 스프레드시트 업데이트
            logger.info("\n7️⃣ 스프레드시트 업데이트...")
            success = self.update_spreadsheet(gspread_client, data)
            
            if success:
                elapsed = int(time.time() - start_time)
                logger.info("\n" + "="*60)
                logger.info("🎉 자동화 완료!")
                logger.info(f"⏱️  소요시간: {elapsed}초")
                logger.info(f"📊 데이터: {len(data)}행 × {len(data.columns)}열")
                logger.info(f"🔗 링크: https://docs.google.com/spreadsheets/d/{self.SPREADSHEET_ID}")
                return True
            
            return False
            
        except Exception as e:
            logger.error(f"❌ 자동화 실패: {e}")
            import traceback
            traceback.print_exc()
            return False

def main():
    """메인 실행 함수"""
    logger.info("🏢 라포랱스 현대카드 OAuth 자동화")
    logger.info("📧 Gmail → HTML → 보안메일 → Google Sheets")
    logger.info("="*50)
    
    try:
        bot = HyundaiCardBot()
        success = bot.run()
        
        if success:
            logger.info("\n🎊 자동화 성공!")
        else:
            logger.error("\n😞 자동화 실패")
            logger.error("hyundai_automation.log 파일을 확인해주세요.")
    
    except KeyboardInterrupt:
        logger.info("\n⏹️ 사용자 중단")
    except Exception as e:
        logger.error(f"\n💥 오류: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
