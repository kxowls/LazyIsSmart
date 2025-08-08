import customtkinter as ctk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException, NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
import threading
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from typing import List, Tuple

# --- Google Play Book Deletion Tool ---
class BookDeletionApp:
    # Constants
    BASE_URL = 'https://play.google.com/books/publish/u/0/?hl=ko'
    WAIT_TIME = 10
    PAGE_LOAD_WAIT = 5

    def __init__(self, master):
        self.master = master

        # Variables
        self.excel_file_path = None
        self.driver = None
        self.current_url = None
        self.total_processed = 0
        self.total_success = 0
        self.total_errors = 0

        self.create_widgets()

    def create_widgets(self):
        main_container = ctk.CTkFrame(self.master, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        self.create_header_section(main_container)
        self.create_file_section(main_container)
        self.create_driver_section(main_container)
        self.create_progress_section(main_container)
        self.create_log_section(main_container)
        self.create_button_section(main_container)

    def create_header_section(self, parent):
        header_frame = ctk.CTkFrame(parent, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))
        title_label = ctk.CTkLabel(header_frame, text="📚 전자책 삭제 프로그램", font=ctk.CTkFont(size=28, weight="bold"), text_color=("#1f538d", "#1f538d"))
        title_label.pack(pady=(0, 5))
        subtitle_label = ctk.CTkLabel(header_frame, text="구글 플레이 도서에서 특정 사용자의 전자책을 자동으로 삭제합니다", font=ctk.CTkFont(size=14), text_color=("#666666", "#cccccc"))
        subtitle_label.pack()

    def create_file_section(self, parent):
        file_frame = ctk.CTkFrame(parent, corner_radius=12)
        file_frame.pack(fill="x", pady=(0, 15))

        section_title = ctk.CTkLabel(file_frame, text="📁 파일 선택", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")

        button_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=15, pady=(10, 10))

        select_btn = ctk.CTkButton(button_frame, text="📄 엑셀 파일 선택", command=self.select_excel, font=ctk.CTkFont(size=14, weight="bold"), height=40, corner_radius=8)
        select_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        template_btn = ctk.CTkButton(button_frame, text="📋 템플릿 생성", command=self.create_template, font=ctk.CTkFont(size=14, weight="bold"), height=40, corner_radius=8, fg_color=("#28a745", "#28a745"), hover_color=("#218838", "#218838"))
        template_btn.pack(side="right", fill="x", expand=True, padx=(5, 0))

        self.excel_label = ctk.CTkLabel(file_frame, text="선택된 엑셀: 없음", font=ctk.CTkFont(size=12), text_color=("#666666", "#cccccc"))
        self.excel_label.pack(pady=(0, 15), padx=15, anchor="w")

    def create_driver_section(self, parent):
        driver_frame = ctk.CTkFrame(parent, corner_radius=12)
        driver_frame.pack(fill="x", pady=(0, 15))
        section_title = ctk.CTkLabel(driver_frame, text="🚀 크롬드라이버 설정", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        info_frame = ctk.CTkFrame(driver_frame, fg_color="transparent")
        info_frame.pack(fill="x", padx=15, pady=(0, 15))
        status_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
        status_frame.pack(fill="x")
        status_icon = ctk.CTkLabel(status_frame, text="✅", font=ctk.CTkFont(size=16))
        status_icon.pack(side="left", padx=(0, 8))
        status_text = ctk.CTkLabel(status_frame, text="자동으로 관리됩니다", font=ctk.CTkFont(size=14, weight="bold"), text_color=("#28a745", "#28a745"))
        status_text.pack(side="left")
        desc_text = ctk.CTkLabel(info_frame, text="webdriver-manager가 자동으로 최신 버전을 다운로드하고 관리합니다", font=ctk.CTkFont(size=12), text_color=("#666666", "#cccccc"))
        desc_text.pack(anchor="w")

    def create_progress_section(self, parent):
        progress_frame = ctk.CTkFrame(parent, corner_radius=12)
        progress_frame.pack(fill="x", pady=(0, 15))
        section_title = ctk.CTkLabel(progress_frame, text="📊 진행률", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        self.progress_var = ctk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(progress_frame, variable=self.progress_var, height=8, corner_radius=4)
        self.progress_bar.pack(fill="x", padx=15, pady=(0, 15))
        self.progress_bar.set(0)

    def create_log_section(self, parent):
        log_frame = ctk.CTkFrame(parent, corner_radius=12)
        log_frame.pack(fill="both", expand=True, pady=(0, 15))
        section_title = ctk.CTkLabel(log_frame, text="📝 작업 로그", font=ctk.CTkFont(size=18, weight="bold"), text_color=("#1f538d", "#1f538d"))
        section_title.pack(pady=(15, 10), padx=15, anchor="w")
        self.log_text = ctk.CTkTextbox(log_frame, font=ctk.CTkFont(size=12), corner_radius=8)
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))

    def create_button_section(self, parent):
        button_frame = ctk.CTkFrame(parent, fg_color="transparent")
        button_frame.pack(fill="x")
        self.start_button = ctk.CTkButton(button_frame, text="🚀 작업 시작", command=self.start_deletion_thread, font=ctk.CTkFont(size=16, weight="bold"), height=50, corner_radius=10, fg_color=("#007bff", "#007bff"), hover_color=("#0056b3", "#0056b3"))
        self.start_button.pack(side="left", fill="x", expand=True, padx=(0, 10))
        quit_button = ctk.CTkButton(button_frame, text="❌ 창 닫기", command=self.master.destroy, font=ctk.CTkFont(size=16, weight="bold"), height=50, corner_radius=10, fg_color=("#dc3545", "#dc3545"), hover_color=("#c82333", "#c82333"))
        quit_button.pack(side="right", fill="x", expand=True, padx=(10, 0))

    def log_message(self, message, level="INFO"):
        timestamp = time.strftime("%H:%M:%S")
        emoji_map = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌", "SUMMARY": "📊"}
        emoji = emoji_map.get(level, "ℹ️")
        log_entry = f"[{timestamp}] {emoji} {level}: {message}\n"
        self.log_text.insert("end", log_entry)
        self.log_text.see("end")
        self.master.update_idletasks()

    def handle_error(self, error, context=""):
        error_message = f"{context}: {str(error)}" if context else str(error)
        self.log_message(error_message, "ERROR")
        self.total_errors += 1

    def select_excel(self):
        file_path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel 파일", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                if len(df.columns) < 2:
                    messagebox.showerror("오류", "엑셀 파일에 최소 2개의 열(A, B)이 필요합니다.")
                    return
                self.excel_file_path = file_path
                filename = os.path.basename(file_path)
                self.excel_label.configure(text=f"선택된 엑셀: {filename}")
                self.log_message(f"엑셀 파일이 선택되었습니다: {filename}", "SUCCESS")
                self.log_message(f"총 {len(df)}개의 항목이 발견되었습니다", "INFO")
            except Exception as e:
                messagebox.showerror("오류", f"엑셀 파일을 읽는 중 오류가 발생했습니다:\n{str(e)}")

    def create_template(self):
        try:
            file_path = filedialog.asksaveasfilename(title="템플릿 파일 저장", defaultextension=".xlsx", filetypes=[("Excel 파일", "*.xlsx")])
            if file_path:
                template_data = {'구글플레이도서_URL': ['https://play.google.com/books/publish/u/0/book/123456789'], '삭제할_이메일': ['user1@example.com']}
                df = pd.DataFrame(template_data)
                df.to_excel(file_path, index=False)
                self.log_message(f"템플릿 파일이 생성되었습니다: {os.path.basename(file_path)}", "SUCCESS")
                messagebox.showinfo("완료", f"템플릿 파일이 생성되었습니다:\n{file_path}")
        except Exception as e:
            messagebox.showerror("오류", f"템플릿 생성 중 오류가 발생했습니다:\n{str(e)}")

    def check_email_exists(self, email):
        try:
            email_xpath = f"//h4[@class='mat-mdc-list-item-title mdc-list-item__primary-text' and contains(text(),'{email}')]"
            elements = self.driver.find_elements(By.XPATH, email_xpath)
            return len(elements) > 0
        except Exception as e:
            self.log_message(f"이메일 존재 확인 중 오류: {str(e)}", "ERROR")
            return False

    def click_delete_button(self, email):
        try:
            email_xpath = f"//h4[@class='mat-mdc-list-item-title mdc-list-item__primary-text' and contains(text(),'{email}')]"
            email_element = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, email_xpath)))
            item_element = email_element.find_element(By.XPATH, "./ancestor::mat-list-item")
            delete_button = item_element.find_element(By.XPATH, ".//button[@aria-label='삭제']")
            self.driver.execute_script("arguments[0].scrollIntoView(true);", delete_button)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", delete_button)
            return True
        except Exception as e:
            self.log_message(f"삭제 버튼 클릭 실패: {str(e)}", "ERROR")
            return False

    def click_confirm_button(self):
        try:
            confirm_selectors = ["//button[contains(text(), '삭제')]", "//button[contains(text(), '확인')]"]
            confirm_button = None
            for selector in confirm_selectors:
                try:
                    confirm_button = WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, selector)))
                    break
                except:
                    continue
            if confirm_button:
                self.driver.execute_script("arguments[0].click();", confirm_button)
                return True
            else:
                self.log_message(f"삭제 확인 버튼을 찾을 수 없습니다", "WARNING")
                return False
        except Exception as e:
            self.log_message(f"삭제 확인 버튼 클릭 실패: {str(e)}", "ERROR")
            return False

    def verify_deletion(self, email):
        try:
            self.driver.refresh()
            time.sleep(3)
            email_xpath = f"//h4[@class='mat-mdc-list-item-title mdc-list-item__primary-text' and contains(text(),'{email}')]"
            remaining_elements = self.driver.find_elements(By.XPATH, email_xpath)
            return len(remaining_elements) == 0
        except Exception as e:
            self.log_message(f"삭제 완료 확인 중 오류: {str(e)}", "WARNING")
            return False

    def start_deletion_thread(self):
        if not self.excel_file_path:
            messagebox.showerror("오류", "엑셀 파일을 선택해주세요.")
            return
        self.start_button.configure(state="disabled", text="🔄 처리 중...")
        threading.Thread(target=self.start_deletion_logic, daemon=True).start()

    def start_deletion_logic(self):
        try:
            start_time = time.time()
            self.log_message("작업을 시작합니다...", "INFO")
            self.log_message("크롬드라이버를 자동으로 설정합니다...", "INFO")

            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service)

            df = pd.read_excel(self.excel_file_path)
            book_urls = df.iloc[:, 0].dropna().tolist()
            emails = df.iloc[:, 1].dropna().tolist()

            self.driver.get(self.BASE_URL)
            self.log_message("로그인이 필요합니다. 로그인을 완료한 후 확인 버튼을 눌러주세요.", "WARNING")
            messagebox.showinfo("안내", "로그인을 완료한 후 확인을 눌러주세요.")

            url_groups = {}
            for url, email in zip(book_urls, emails):
                if url not in url_groups:
                    url_groups[url] = []
                url_groups[url].append(email)

            actual_total_items = len(book_urls)
            processed_count = 0

            for url, email_list in url_groups.items():
                if url != self.current_url:
                    self.log_message(f"페이지 로딩: {url}", "INFO")
                    self.driver.get(url)
                    WebDriverWait(self.driver, self.WAIT_TIME).until(lambda d: d.execute_script("return document.readyState") == "complete")
                    time.sleep(self.PAGE_LOAD_WAIT)
                    self.current_url = url

                for email in email_list:
                    processed_count += 1
                    progress = (processed_count / actual_total_items)
                    self.progress_var.set(progress)

                    self.log_message(f"[{processed_count}/{actual_total_items}] {email} 처리 중...", "INFO")

                    if self.check_email_exists(email):
                        if self.click_delete_button(email):
                            time.sleep(1)
                            if self.click_confirm_button():
                                time.sleep(3)
                                if self.verify_deletion(email):
                                    self.log_message(f"{email} 삭제 성공", "SUCCESS")
                                    self.total_success += 1
                                else:
                                    self.log_message(f"{email} 삭제 실패 (삭제 후에도 확인됨)", "ERROR")
                                    self.total_errors += 1
                            else:
                                self.log_message(f"{email} 삭제 확인 실패", "ERROR")
                                self.total_errors += 1
                        else:
                            self.log_message(f"{email} 삭제 버튼 클릭 실패", "ERROR")
                            self.total_errors += 1
                    else:
                        self.log_message(f"{email} 해당 이메일을 찾을 수 없습니다", "WARNING")
                    self.total_processed += 1
                    time.sleep(1)

            end_time = time.time()
            self.log_message(f"총 처리 시간: {end_time - start_time:.1f}초", "INFO")
            self.show_summary()

        except Exception as e:
            self.handle_error(e, "전체 작업 실패")
        finally:
            if self.driver:
                self.driver.quit()
            self.start_button.configure(state="normal", text="🚀 작업 시작")

    def show_summary(self):
        summary = f"📊 작업 완료!\n\n총 처리: {self.total_processed}개\n성공: {self.total_success}개\n실패: {self.total_errors}개"
        self.log_message(summary, "SUMMARY")
        messagebox.showinfo("작업 완료", summary)


# --- Hanbit Media Ebook Registration Tool ---

class EbookRegistrator:
    """Handles the Selenium automation logic for ebook registration."""
    def __init__(self, log_widget):
        self.driver = None
        self.wait = None
        self.log_widget = log_widget
        self.setup_driver()

    def log(self, message):
        """Logs a message to the GUI widget."""
        timestamp = time.strftime("%H:%M:%S")
        self.log_widget.insert("end", f"[{timestamp}] {message}\n")
        self.log_widget.see("end")

    def setup_driver(self):
        """Sets up the Chrome driver using webdriver-manager."""
        self.log("크롬 드라이버를 설정합니다...")
        try:
            chrome_options = Options()
            # chrome_options.add_argument("--headless") # Headless mode can be enabled here if needed
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 20)
            self.log("크롬 드라이버 설정 완료.")
        except Exception as e:
            self.log(f"드라이버 설정 오류: {e}")
            messagebox.showerror("드라이버 오류", f"크롬 드라이버를 설정하는 중 오류가 발생했습니다: {e}")
            raise

    def login(self):
        """Navigates to login page and waits for user to log in."""
        self.log("로그인 페이지로 이동합니다. 로그인 후 확인을 눌러주세요.")
        self.driver.get("https://www.hanbit.co.kr/login?redirect=https://www.hanbit.co.kr/hb_admin")
        messagebox.showinfo("로그인 필요", "로그인 페이지가 열렸습니다.\n로그인을 완료한 후 확인을 눌러주세요.")

        self.wait.until(lambda driver: "hb_admin" in driver.current_url)
        self.log("로그인 성공을 확인했습니다.")

    def click_register_button(self):
        """Clicks the last register button on the search list."""
        # Implementation from original script...
        pass # This will be filled in with the full logic

    def process_ebook_registration(self, user_name: str, email: str, book_name: str) -> bool:
        """Processes a single ebook registration."""
        try:
            self.log(f"처리 시작: {user_name} / {book_name}")
            self.driver.get("https://www.hanbit.co.kr/hb_admin/front.acaebookask.php")
            time.sleep(2)

            self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn_sch"))).click()
            time.sleep(2)
            self.driver.switch_to.window(self.driver.window_handles[-1])

            search_key = self.wait.until(EC.presence_of_element_located((By.ID, "search_key")))
            search_key.click()
            time.sleep(1)
            search_key.find_elements(By.TAG_NAME, "option")[0].click()

            keyword_input = self.wait.until(EC.presence_of_element_located((By.ID, "keyword")))
            keyword_input.send_keys(user_name + "\n")
            time.sleep(2)

            # Click register button in popup
            register_buttons = self.wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "btn_register")))
            if register_buttons:
                register_buttons[-1].click()

            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(2)

            email_input = self.wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='hao_gmail']")))
            email_input.clear()
            email_input.send_keys(email)
            time.sleep(1)

            checkbox = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='frm']/section/div/div/div[1]/div[2]/div[2]/div/input[2]")))
            checkbox.click()
            time.sleep(1)

            ebook_search_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn-primary') and contains(@class, 'mb10') and contains(., 'eBook 검색')]")))
            ebook_search_button.click()
            time.sleep(2)

            original_window = self.driver.current_window_handle
            for window_handle in self.driver.window_handles:
                if window_handle != original_window:
                    self.driver.switch_to.window(window_handle)
                    break

            try:
                self.wait.until(EC.presence_of_element_located((By.ID, "keyword")))
                ebook_input = self.wait.until(EC.presence_of_element_located((By.ID, "keyword")))
                ebook_input.clear()
                ebook_input.send_keys(book_name)
                time.sleep(1)
                ebook_input.send_keys("\n")
                time.sleep(2)

                try:
                    register_button = self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn_register")))
                    register_button.click()
                except TimeoutException:
                    self.log(f"오류: '{book_name}' 책 검색 후 등록 버튼을 찾을 수 없습니다.")
                    self.driver.switch_to.window(original_window)
                    return False

                self.driver.close()
                self.driver.switch_to.window(original_window)
                time.sleep(2)

                self.wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn_register"))).click()
                time.sleep(1)
                try:
                    alert = Alert(self.driver)
                    alert.accept()
                    self.log("팝업창 확인 완료.")
                except NoAlertPresentException:
                    pass
                self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body"))).send_keys(Keys.ENTER)
                time.sleep(2)

                self.log(f"성공: {user_name} - {book_name} - {email} 입력 완료!")
                return True

            except (NoSuchElementException, TimeoutException, NoAlertPresentException) as e:
                self.log(f"오류 발생: {e}")
                self.driver.close()
                self.driver.switch_to.window(original_window)
                return False

        except Exception as e:
            self.log(f"처리 중 심각한 오류 발생: {e}")
            return False

    def close(self):
        if self.driver:
            self.driver.quit()

class EbookRegistrationApp:
    def __init__(self, master):
        self.master = master
        self.excel_file_path = ""

        # Main container
        main_container = ctk.CTkFrame(self.master, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # --- UI Elements ---
        # File selection
        file_frame = ctk.CTkFrame(main_container)
        file_frame.pack(fill="x", pady=(0, 15))

        self.excel_label = ctk.CTkLabel(file_frame, text="엑셀 파일 경로:")
        self.excel_label.pack(side="left", padx=(15, 10), pady=15)

        self.excel_path_entry = ctk.CTkEntry(file_frame, width=300)
        self.excel_path_entry.pack(side="left", fill="x", expand=True, pady=15)

        self.excel_button = ctk.CTkButton(file_frame, text="찾아보기", command=self.select_excel_file, width=100)
        self.excel_button.pack(side="left", padx=(10, 15), pady=15)

        # Run button
        self.run_button = ctk.CTkButton(main_container, text="🚀 실행", command=self.run_script, height=40, font=ctk.CTkFont(size=16, weight="bold"))
        self.run_button.pack(fill="x", pady=(0, 15))

        # Log area
        self.log_text = ctk.CTkTextbox(main_container, height=200, font=ctk.CTkFont(size=12))
        self.log_text.pack(fill="both", expand=True)

    def select_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            self.excel_file_path = file_path
            self.excel_path_entry.delete(0, "end")
            self.excel_path_entry.insert(0, file_path)
            self.log_text.insert("end", f"엑셀 파일 선택됨: {os.path.basename(file_path)}\n")

    def run_script(self):
        if not self.excel_file_path:
            messagebox.showerror("오류", "엑셀 파일을 먼저 선택해주세요.")
            return

        self.run_button.configure(state="disabled", text="🔄 처리 중...")
        self.log_text.delete("1.0", "end")

        # Run the automation in a separate thread to keep the GUI responsive
        threading.Thread(target=self.run_script_threaded, daemon=True).start()

    def run_script_threaded(self):
        try:
            df = pd.read_excel(self.excel_file_path)
            names = df.iloc[:, 0].dropna().tolist()
            emails = df.iloc[:, 1].dropna().tolist()
            books = df.iloc[:, 2].dropna().tolist()

            if not names or not emails or not books:
                messagebox.showerror("오류", "엑셀 파일에 필요한 데이터(이름, 이메일, 책 제목)가 부족합니다.")
                self.run_button.configure(state="normal", text="🚀 실행")
                return

        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 처리 중 오류 발생: {e}")
            self.run_button.configure(state="normal", text="🚀 실행")
            return

        ebook_reg = None
        try:
            ebook_reg = EbookRegistrator(log_widget=self.log_text)
            ebook_reg.login()

            success_count = 0
            total_count = len(names)
            ebook_reg.log(f"총 {total_count}건의 등록을 시작합니다.")

            for i, (user_name, email, book_name) in enumerate(zip(names, emails, books)):
                ebook_reg.log(f"--- [{i+1}/{total_count}] 진행 중 ---")
                if ebook_reg.process_ebook_registration(user_name, email, book_name):
                    success_count += 1

            summary = f"총 {total_count}건 중 {success_count}건 처리 완료되었습니다."
            ebook_reg.log(f"--- 작업 완료 ---")
            ebook_reg.log(summary)
            messagebox.showinfo("완료", summary)

        except Exception as e:
            messagebox.showerror("오류", f"전체 작업 중 오류 발생: {e}")
            if ebook_reg:
                ebook_reg.log(f"치명적인 오류로 작업을 중단합니다: {e}")
        finally:
            if ebook_reg:
                ebook_reg.close()
            self.run_button.configure(state="normal", text="🚀 실행")


# --- Main Menu ---
class MainMenu:
    def __init__(self, root):
        self.root = root
        self.root.title("Ebook Tool Suite")
        self.root.geometry("500x200")

        # Prevent the main window from being closed while a tool is open
        self.toplevel_window = None

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        title_label = ctk.CTkLabel(main_frame, text="Ebook Tool Suite", font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=10)

        subtitle_label = ctk.CTkLabel(main_frame, text="실행할 작업을 선택하세요.", font=ctk.CTkFont(size=14))
        subtitle_label.pack(pady=(0, 20))

        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x", expand=True)

        google_button = ctk.CTkButton(button_frame, text="📚 Google Play 도서 삭제", command=self.open_google_tool, height=40)
        google_button.pack(side="left", fill="x", expand=True, padx=(0, 10))

        hanbit_button = ctk.CTkButton(button_frame, text="🚀 한빛미디어 eBook 등록", command=self.open_hanbit_tool, height=40)
        hanbit_button.pack(side="right", fill="x", expand=True, padx=(10, 0))

    def open_tool_window(self, tool_class, title, geometry):
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = ctk.CTkToplevel(self.root)
            self.toplevel_window.title(title)
            self.toplevel_window.geometry(geometry)
            tool_class(self.toplevel_window)
        else:
            self.toplevel_window.focus()

    def open_google_tool(self):
        self.open_tool_window(BookDeletionApp, "Google Play 도서 삭제", "900x700")

    def open_hanbit_tool(self):
        self.open_tool_window(EbookRegistrationApp, "한빛미디어 eBook 등록", "600x400")

if __name__ == "__main__":
    root = ctk.CTk()
    app = MainMenu(root)
    root.mainloop()
