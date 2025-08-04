import os, time, re, requests, fitz
from openai import OpenAI
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchWindowException

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

class Browser:
    def __init__(self, headless=False):
        options = Options()
        if headless:
            options.add_argument("--headless")
        options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 300)

    def quit(self):
        self.driver.quit()

class OutlookBot(Browser):
    def login(self):
        self.driver.get("https://outlook.office.com/mail/")
        self.wait.until(EC.presence_of_element_located((By.NAME, "loginfmt"))).send_keys(os.getenv("EMAIL_ID"))
        for _ in range(5):
            try:
                self.wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                time.sleep(5)
                break
            except StaleElementReferenceException:
                time.sleep(1)
        self.wait.until(EC.presence_of_element_located((By.NAME, "passwd"))).send_keys(os.getenv("EMAIL_PASSWORD"))
        for _ in range(5):
            try:
                self.wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
                break
            except StaleElementReferenceException:
                time.sleep(1)
        print("🞡 '계속 로그인' 버튼 대기 중...")
        for attempt in range(12):
            try:
                stay_button = self.driver.find_element(By.ID, "idSIButton9")
                if stay_button.is_displayed() and stay_button.is_enabled():
                    stay_button.click()
                    print("'계속 로그인' 클릭 완료")
                    return
            except StaleElementReferenceException:
                print(f"버튼 요소 stale - 재시도 중 ({attempt + 1}/12)")
            except Exception as e:
                print(f"예외 발생: {type(e).__name__} - {e}")
            time.sleep(1)
        print("'계속 로그인' 클릭 실패: 요소를 찾을 수 없음")

    def get_code_from_email(self):
        print("📨 Outlook 받은편지함 접속 중...")
        self.driver.get("https://outlook.office.com/mail/inbox")
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='option']")))
            items = self.driver.find_elements(By.CSS_SELECTOR, "div[role='option']")
            if not items:
                print("메일 항목이 없습니다.")
                return None
            items[0].click()
            time.sleep(3)
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[id^='UniqueMessageBody']")))
            el = self.driver.find_element(By.CSS_SELECTOR, "div[id^='UniqueMessageBody']")
            text = el.text.strip()
            print(f"이메일 본문 내용: {text}")
            match = re.search(r"\b\d{6}\b", text)
            print(f"추출된 인증코드: {match.group() if match else 'None'}")
            return match.group() if match else None
        except Exception as e:
            print(f"본문 추출 실패: {type(e).__name__} - {e}")
            return None

class LMSBot(Browser):
    def __init__(self):
        super().__init__()
        self.outlook = OutlookBot()
        self.outlook.driver = self.driver
        self.outlook.wait = self.wait

    def login_and_authenticate(self):
        self.driver.get("https://uwlms.uhs.ac.kr/login.php")
        self.wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(os.getenv("LMS_ID"))
        self.driver.find_element(By.NAME, "password").send_keys(os.getenv("LMS_PASSWORD"))
        self.driver.find_element(By.NAME, "loginbutton").click()
        print("LMS 로그인 완료")
        self.wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'course_box')]//a[contains(@href, 'course/view.php')]")))
        course_link = self.driver.find_element(By.XPATH, "//div[text()='학부']/ancestor::div[contains(@class, 'course_box')]//a[contains(@href, 'course/view.php')]")
        self.driver.execute_script("""const banners = document.querySelectorAll('img[role=\"presentation\"], .modal, .popup');banners.forEach(el => el.style.display = 'none');""")
        self.driver.execute_script("arguments[0].scrollIntoView(true);", course_link)
        time.sleep(5)
        self.driver.execute_script("arguments[0].click();", course_link)
        print("강의 페이지 이동 완료")
        self.wait.until(EC.element_to_be_clickable((By.ID, "btn-emailAuth"))).click()
        self.wait.until(EC.element_to_be_clickable((By.ID, "btn-send-email"))).click()
        print("인증메일 발송 요청 완료")
        self.driver.execute_script("window.open('https://outlook.office.com/mail/');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        self.outlook.login()
        code = self.outlook.get_code_from_email()
        time.sleep(5)
        try:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
        except NoSuchWindowException:
            print("Outlook 탭 닫기 실패 또는 LMS 탭 전환 실패")
            return
        if code:
            print(f"인증코드: {code}")
            try:
                self.wait.until(EC.presence_of_element_located((By.ID, "auth-email-code"))).send_keys(code)
                self.driver.find_element(By.ID, "btn-auth-email-code").click()
                time.sleep(5)
                try:
                    self.wait.until(EC.invisibility_of_element_located((By.ID, "auth-email-code")))
                    print("인증 완료 (입력창 사라짐)")
                    self.driver.get("https://uwlms.uhs.ac.kr/")
                    time.sleep(5)
                    self.navigate_and_process_departments(from_authenticated_page=True)
                except:
                    print("인증 실패 (인증창 유지됨) → 스크린샷으로 확인 필요")
            except Exception as e:
                import traceback
                print("인증 입력 실패:", type(e).__name__)
                traceback.print_exc()
        else:
            print("인증코드 수신 실패")

    def navigate_and_process_departments(self, from_authenticated_page=False):
        if not from_authenticated_page:
            self.driver.get("https://uwlms.uhs.ac.kr/")
            time.sleep(5)
        course_links = self.driver.find_elements(By.XPATH, "//div[contains(@class, 'course_box')]//a[contains(@href, 'course/view.php')]")
        hrefs = [link.get_attribute("href") for link in course_links]
        print(f"감지된 전체 강의 수: {len(hrefs)}")
        for href in hrefs:
            try:
                self.driver.get(href)
                print(f"➡️ 학부 페이지 진입: {href}")
                if "auth" in self.driver.current_url or "email" in self.driver.current_url:
                    print("인증 페이지로 리디렉션됨 - 이전 인증이 유지되지 않음")
                    continue
                self.download_and_summarize_pdfs()
                self.driver.back()
                time.sleep(10)
            except Exception as e:
                print(f"❌ 학부 순회 중 오류 발생: {type(e).__name__} - {str(e)}")

    def download_and_summarize_pdfs(self):
        print("📥 PDF 링크 수집 시작...")
        pdf_xpaths = [
            "//li[contains(@class, 'activity') and contains(@class, 'ubfile')]//a[contains(@href, 'mod/ubfile/view.php')]",
            "//a[contains(@href, '/pluginfile.php')]",
            "//a[contains(@href, '/mod/resource/view.php')]",
            "//div[contains(@id, 'module-')]//a[contains(@href, 'pluginfile.php')]"
        ]
        links = []
        for xpath in pdf_xpaths:
            found = self.driver.find_elements(By.XPATH, xpath)
            if found:
                links = found
                print(f"XPath '{xpath}'에서 PDF 링크 {len(found)}개 발견")
                break
        if not links:
            print("PDF 링크를 찾을 수 없습니다.")
            return
        os.makedirs("downloads/원본", exist_ok=True)
        os.makedirs("downloads/요약본", exist_ok=True)
        try:
            professor_el = self.driver.find_element(By.CSS_SELECTOR, "h4.media-heading")
            professor = professor_el.text.strip()
        except:
            professor = "UnknownProfessor"
        cookies = {cookie['name']: cookie['value'] for cookie in self.driver.get_cookies()}
        for idx, link in enumerate(links):
            try:
                href = link.get_attribute("href")
                link_text = link.text.strip()
                subject_match = re.search(r"\[(.*?)\]", link_text)
                subject = subject_match.group(1) if subject_match else "UnknownSubject"
                if subject == "UnknownSubject":
                    subject = os.path.basename(href).split('?')[0].replace('.pdf', '')
                raw_filename = f"{subject}_{professor}교수님_{idx}.pdf".replace(" ", "_")
                pdf_path = os.path.join("downloads", "원본", raw_filename)
                headers = {"User-Agent": "Mozilla/5.0"}
                response = requests.get(href, headers=headers, cookies=cookies)
                if response.status_code == 200 and b'%PDF' in response.content[:1024]:
                    with open(pdf_path, "wb") as f:
                        f.write(response.content)
                    if not self.is_valid_pdf(pdf_path):
                        print(f"저장된 PDF가 되었으나 fitze에서 열 수 없음: {pdf_path}")
                        return
                    print(f"PDF 다운로드 완료: {pdf_path}")
                    summary = self.summarize_pdf(pdf_path)
                    summary_filename = os.path.join("downloads", "요약본", f"{subject}_{professor}교수님_요약.txt".replace(" ", "_"))
                    self.save_summary_as_txt(summary, summary_filename)
                    print(f"요약 저장 완료: {summary_filename}")
                else:
                    print(f"PDF 직접 다운로드 실패: {href} → 브라우저로 처리 시도")
                    self.driver.execute_script("window.open(arguments[0]);", href)
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    time.sleep(2)
                    try:
                        text = self.driver.find_element(By.TAG_NAME, "body").text
                        summary = self.summarize_text(text)
                        self.save_summary_as_txt(summary, f"downloads/요약본/{subject}_{professor}교슈님_요약.txt")
                        print("브라우저에서 텍스트 추출 완료")
                    except Exception as e:
                        print(f"브라우저 PDF 처리 실패: {str(e)}")
                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])
            except Exception as e:
                print(f"PDF 처리 중 오류: {type(e).__name__} - {str(e)}")

    def is_valid_pdf(self, file_path):
        try:
            with fitz.open(file_path) as doc:
                return doc.page_count > 0
        except Exception as e:
            print(f"PDF 유효성 검사 실패: {e}")
            return False

    def summarize_pdf(self, path):
        try:
            doc = fitz.open(path)
            text = "".join(page.get_text() for page in doc)
            doc.close()
            return self.summarize_text(text if text.strip() else "PDF 본문 없음")
        except Exception as e:
            print(f"PDF 파싱 실패: {type(e).__name__} - {str(e)}")
            return "PDF 파싱 중 오류 발생"

    def summarize_text(self, text):
        prompt = f"다음은 강의 자료의 내용입니다. 핵심 내용을 전체 요약해주세요:\n\n{text[:7000]}"
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
        )
        return response.choices[0].message.content

    def save_summary_as_txt(self, summary_text, filename):
        with open(filename, "w", encoding="utf-8") as f:
            f.write(summary_text)

if __name__ == "__main__":
    bot = LMSBot()
    try:
        bot.login_and_authenticate()
    finally:
        bot.quit()
