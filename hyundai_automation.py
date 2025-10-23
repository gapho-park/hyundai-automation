#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
라포랩스 현대카드 OAuth 자동화 (GitHub Actions 버전 - 수정됨)
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
import shutil
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
    from selenium.webdriver.common.keys import Keys
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    import gspread
    from webdriver_manager.chrome import ChromeDriverManager
except ImportError as e:
    print(f"❌ 필수 라이브러리가 설치되어 있지 않습니다: {e}")
    print("\n다음 명령어로 설치해주세요:")
    print("pip install selenium pandas google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client gspread openpyxl xlrd webdriver-manager")
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
        
        # CI 환경 감지
        self.is_ci = os.environ.get('CI') == 'true' or os.environ.get('GITHUB_ACTIONS') == 'true'
        
        # 다운로드 경로 설정
        if self.is_ci:
            self.download_path = "/tmp/hyundai_auto"
        else:
            self.download_path = os.path.join(os.path.expanduser("~"), "Downloads", "hyundai_auto")
        
        os.makedirs(self.download_path, exist_ok=True)
        
        logger.info("🏢 라포랩스 현대카드 자동화 봇 시작")
        logger.info(f"환경: {'CI (GitHub Actions)' if self.is_ci else '로컬'}")
        logger.info(f"다운로드 경로: {self.download_path}")
        
    def init_chrome_driver(self):
        """Chrome WebDriver 초기화 (GitHub Actions 호환)"""
        try:
            logger.info("🤖 Chrome WebDriver 초기화 중...")
            
            chrome_options = Options()
            
            # 공통 설정
            chrome_options.add_experimental_option("prefs", {
                "download.default_directory": self.download_path,
                "download.prompt_for_download": False,
                "plugins.always_open_pdf_externally": True,
                "safebrowsing.enabled": False
            })
            
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-plugins")
            chrome_options.add_argument("--disable-infobars")
            
            # CI 환경 설정
            if self.is_ci:
                logger.info("💻 GitHub Actions 환경 감지 - Headless 모드 활성화")
                chrome_options.add_argument("--headless=new")
                chrome_options.add_argument("--disable-crash-reporter")
            else:
                logger.info("💻 로컬 환경 - GUI 모드")
            
            # Chrome 바이너리 경로 설정 (여러 가능성 확인)
            chrome_bin_candidates = [
                os.getenv('CHROME_BIN'),
                '/usr/bin/chromium-browser',
                '/usr/bin/google-chrome',
                '/usr/bin/chromium',
                '/snap/chromium/current/usr/bin/chromium'
            ]
            
            chrome_bin = None
            for candidate in chrome_bin_candidates:
                if candidate and os.path.exists(candidate):
                    chrome_bin = candidate
                    logger.info(f"✅ Chrome 바이너리 발견: {chrome_bin}")
                    break
            
            if chrome_bin:
                chrome_options.binary_location = chrome_bin
            else:
                logger.warning("⚠️ Chrome 바이너리를 찾을 수 없음, 기본 설정 사용")
            
            # ChromeDriver 설정
            try:
                chromedriver_path = os.getenv('CHROMEDRIVER_PATH')
                if chromedriver_path and os.path.exists(chromedriver_path):
                    service = Service(chromedriver_path)
                    logger.info(f"✅ ChromeDriver: {chromedriver_path}")
                else:
                    # webdriver-manager 사용 (자동 관리)
                    chromedriver_path = ChromeDriverManager().install()
                    service = Service(chromedriver_path)
                    logger.info(f"✅ ChromeDriver (webdriver-manager): {chromedriver_path}")
            except Exception as e:
                logger.warning(f"ChromeDriver 자동 설정 실패, 기본 설정 사용: {e}")
                service = Service()
            
            driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # 타임아웃 설정
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            
            logger.info("✅ Chrome WebDriver 초기화 성공")
            return driver
            
        except Exception as e:
            logger.error(f"❌ Chrome WebDriver 초기화 실패: {e}")
            raise
    
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
                creds = None
        
        # 토큰 갱신
        if creds and creds.expired and creds.refresh_token:
            try:
                logger.info("🔄 토큰 갱신 중...")
                creds.refresh(Request())
                logger.info("✅ 토큰 갱신 완료")
            except Exception as e:
                logger.warning(f"토큰 갱신 실패: {e}")
                creds = None
        
        # 새 인증 필요
        if not creds or not creds.valid:
            logger.info("새 OAuth 인증이 필요합니다...")
            
            # 클라이언트 자격증명 파일 찾기
            client_files = ['client_secret.json', 'credentials.json', 'oauth_credentials.json']
            client_file = None
            
            for file in client_files:
                if os.path.exists(file):
                    client_file = file
                    logger.info(f"✅ OAuth 파일 발견: {client_file}")
                    break
            
            if not client_file:
                logger.error("❌ OAuth 클라이언트 자격증명 파일을 찾을 수 없습니다!")
                logger.error("다음 파일 중 하나를 현재 폴더에 두세요:")
                for file in client_files:
                    logger.error(f"  - {file}")
                return None
            
            # CI 환경에서는 로컬 서버 사용 불가
            if self.is_ci:
                logger.error("❌ CI 환경에서 새 인증 불가")
                logger.error("GitHub Secrets에 token.pickle (base64)을 미리 저장해야 합니다.")
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
                logger.info("✅ 토큰 저장 완료")
            except Exception as e:
                logger.error(f"토큰 저장 실패: {e}")
        
        logger.info("✅ OAuth 인증 완료")
        return creds
    
    def find_hyundai_email(self, gmail_service):
        """현대카드 이메일 찾기 (재시도 로직 포함)"""
        try:
            logger.info("📧 현대카드 이메일 검색 중...")
            
            queries = [
                'from:"현대카드 MY COMPANY" subject:"라포랩스 보유내역" has:attachment newer_than:14d',
                'from:"현대카드 MY COMPANY" subject:"보유내역" has:attachment newer_than:14d',
                'from:"현대카드" subject:"보유내역" has:attachment newer_than:21d',
                'from:"MY COMPANY" has:attachment newer_than:21d',
                'from:"현대카드" has:attachment newer_than:30d'
            ]
            
            all_messages = []
            
            for query in queries:
                try:
                    logger.info(f"  검색 쿼리: {query[:40]}...")
                    results = gmail_service.users().messages().list(
                        userId='me',
                        q=query,
                        maxResults=10
                    ).execute()
                    
                    messages = results.get('messages', [])
                    all_messages.extend(messages)
                    logger.info(f"    → {len(messages)}개 발견")
                    
                    if messages:
                        break  # 첫 번째 쿼리에서 결과 있으면 종료
                        
                except Exception as e:
                    logger.warning(f"  검색 오류: {e}")
                    continue
            
            if not all_messages:
                logger.error("❌ 현대카드 이메일을 찾을 수 없습니다.")
                return None
            
            # 중복 제거 및 최신 선택
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
            
            message = gmail_service.users().messages().get(
                userId='me',
                id=message_id,
                format='full'
            ).execute()
            
            # 제목과 발신자 추출
            headers = message['payload'].get('headers', [])
            subject = next((h['value'] for h in headers if h['name'].lower() == 'subject'), 'Unknown')
            sender = next((h['value'] for h in headers if h['name'].lower() == 'from'), 'Unknown')
            
            logger.info(f"📧 제목: {subject}")
            logger.info(f"📧 발신자: {sender}")
            
            # HTML 첨부파일 재귀 검색
            def find_html_attachment(part):
                if part.get('filename') and part['filename'].lower().endswith(('.html', '.htm')):
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
        """보안메일 처리 (Selenium 사용)"""
        driver = None
        try:
            logger.info("🔐 보안메일 처리 시작...")
            
            driver = self.init_chrome_driver()
            wait = WebDriverWait(driver, 20)
            
            logger.info("✅ Chrome 브라우저 실행 완료")
            
            # HTML 파일 열기
            file_url = f"file://{os.path.abspath(html_file)}"
            logger.info(f"📄 HTML 파일 로드: {os.path.basename(html_file)}")
            driver.get(file_url)
            
            time.sleep(5)
            
            # 인증번호 입력 필드 찾기
            logger.info("🔍 인증번호 입력 필드 찾기...")
            
            input_field = None
            selectors = [
                "input[name='p2']",
                "input[name='p2_temp']",
                "input[type='password']",
                "input[type='text'][id*='auth']",
                "input[type='text']"
            ]
            
            for selector in selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        if element.is_displayed() and element.is_enabled():
                            input_field = element
                            logger.info(f"✅ 입력 필드 발견: {selector}")
                            break
                    if input_field:
                        break
                except:
                    continue
            
            if not input_field:
                logger.error("❌ 인증번호 입력 필드를 찾을 수 없습니다.")
                return None
            
            # 인증번호 입력
            try:
                input_field.click()
                time.sleep(1)
                input_field.clear()
                time.sleep(0.5)
                input_field.send_keys(self.AUTH_CODE)
                logger.info(f"✅ 인증번호 입력: {self.AUTH_CODE}")
                
                # 엔터키로 폼 제출
                input_field.send_keys(Keys.RETURN)
                logger.info("✅ 엔터키로 폼 제출")
                
            except Exception as e:
                logger.warning(f"입력 필드 조작 실패, JavaScript 시도: {e}")
                driver.execute_script(f"document.getElementsByName('p2')[0].value = '{self.AUTH_CODE}';")
                logger.info(f"✅ JavaScript로 인증번호 입력: {self.AUTH_CODE}")
                
                # 폼 제출
                try:
                    driver.execute_script("document.getElementById('decForm').submit();")
                except:
                    # 대체 제출 방법
                    driver.execute_script("document.querySelector('form').submit();")
                logger.info("✅ JavaScript로 폼 제출")
            
            logger.info("⏳ 폼 제출 후 페이지 로딩 대기...")
            time.sleep(10)
            
            # ZIP 다운로드 링크 클릭
            logger.info("📦 ZIP 다운로드 링크 찾기...")
            
            try:
                # 다운로드 링크 찾기
                download_link = None
                for selector in ["a[href*='.zip']", "a[download]", "a"]:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for element in elements:
                        try:
                            href = element.get_attribute('href') or ''
                            text = element.text.strip()
                            if '.zip' in href.lower() or '.zip' in text.lower() or '다운로드' in text:
                                download_link = element
                                break
                        except:
                            continue
                    if download_link:
                        break
                
                if download_link:
                    try:
                        download_link.click()
                    except:
                        driver.execute_script("arguments[0].click();", download_link)
                    logger.info("✅ ZIP 다운로드 시작")
                else:
                    logger.warning("⚠️ 다운로드 링크를 찾을 수 없음")
                    
            except Exception as e:
                logger.warning(f"다운로드 링크 클릭 실패: {e}")
            
            # ZIP 다운로드 대기
            logger.info("⏳ ZIP 파일 다운로드 대기...")
            
            start_time = time.time()
            while (time.time() - start_time) < 120:  # 2분 타임아웃
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
            
            logger.error("❌ ZIP 다운로드 타임아웃 (120초)")
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
            
            # 기존 폴더 정리
            if extract_path.exists():
                shutil.rmtree(extract_path)
            
            extract_path.mkdir()
            
            # 압축 해제
            try:
                with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                    zip_ref.extractall(extract_path)
                logger.info("✅ 압축 해제 완료")
            except Exception as e:
                logger.warning(f"표준 압축 해제 실패: {e}, 대체 방법 시도")
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
            try:
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
                logger.error(f"엑셀 읽기 실패: {e}")
                return None
            
        except Exception as e:
            logger.error(f"❌ ZIP 처리 실패: {e}")
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
                logger.info(f"✅ 새 워크시트 생성: {self.SHEET_NAME}")
            
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
    
    def run_automation(self):
        """전체 자동화 실행"""
        logger.info("🚀 현대카드 자동화 시작!")
        logger.info("="*60)
        
        start_time = time.time()
        
        try:
            # 1. OAuth 인증
            logger.info("\n1️⃣ OAuth 인증...")
            creds = self.authenticate()
            if not creds:
                logger.error("OAuth 인증 실패")
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
    logger.info("🏢 라포랩스 현대카드 OAuth 자동화")
    logger.info("📧 Gmail → HTML → 보안메일 → Google Sheets")
    logger.info("="*60)
    
    # CI 환경이면 자동 실행
    is_ci = os.environ.get('CI') == 'true' or os.environ.get('GITHUB_ACTIONS') == 'true'
    
    if is_ci:
        logger.info("\n🤖 CI 환경에서 자동 실행 중...")
    else:
        logger.info("\n📋 확인사항:")
        logger.info("✅ client_secret.json 파일")
        logger.info("✅ token.pickle (또는 새 인증 진행)")
        logger.info("✅ Chrome 브라우저")
        response = input("\n🚀 자동화를 시작하시겠습니까? (y/N): ").strip().lower()
        if response != 'y':
            logger.info("자동화를 취소했습니다.")
            return
    
    try:
        bot = HyundaiCardBot()
        success = bot.run_automation()
        
        if success:
            logger.info("\n🎊 자동화 성공!")
            sys.exit(0)
        else:
            logger.error("\n😞 자동화 실패")
            logger.error("hyundai_automation.log 파일을 확인해주세요.")
            sys.exit(1)
    
    except KeyboardInterrupt:
        logger.info("\n⏹️ 사용자 중단")
        sys.exit(0)
    except Exception as e:
        logger.error(f"\n💥 오류: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
