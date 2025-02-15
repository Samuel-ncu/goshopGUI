import sys
import os
import re
import random

import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QListWidget, QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel, QMessageBox, QDialog, QHBoxLayout, QLineEdit, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from playwright.sync_api import sync_playwright
from PyQt5.QtWidgets import QFileDialog

# 新增使用者對話窗
class AddUserDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("新增使用者")
        self.resize(300, 100)
        layout = QVBoxLayout()
        self.label = QLabel("請輸入新使用者名稱：")
        layout.addWidget(self.label)
        self.user_edit = QLineEdit()
        layout.addWidget(self.user_edit)
        self.confirm_btn = QPushButton("確認新增")
        self.confirm_btn.clicked.connect(self.accept)
        layout.addWidget(self.confirm_btn)
        self.setLayout(layout)

    def get_username(self):
        return self.user_edit.text().strip()

# 更新產品URL對話窗
class UpdateProductURLDialog(QDialog):
    def __init__(self, products_file, parent=None):
        super().__init__(parent)
        self.setWindowTitle("更新產品URL")
        self.products_file = products_file
        self.df = pd.read_excel(products_file)
        self.current_index = 0
        self.initUI()
        self.load_current_record()

    def initUI(self):
        layout = QVBoxLayout()
        # 第一行：No. 與 產品名稱，以及 Copy 按鈕
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("No.:"))
        self.no_label = QLabel("")
        top_layout.addWidget(self.no_label)
        top_layout.addSpacing(20)
        top_layout.addWidget(QLabel("產品名稱:"))
        self.name_label = QLabel("")
        top_layout.addWidget(self.name_label)
        self.copy_btn = QPushButton("Copy")
        self.copy_btn.setToolTip("複製產品名稱")
        self.copy_btn.clicked.connect(self.copy_name)
        top_layout.addWidget(self.copy_btn)
        layout.addLayout(top_layout)
        # 第二行：URL 輸入框
        url_layout = QHBoxLayout()
        url_layout.addWidget(QLabel("URL:"))
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("請輸入百寶倉URL")
        url_layout.addWidget(self.url_edit)
        layout.addLayout(url_layout)
        # 第三行：功能按鈕
        btn_layout = QHBoxLayout()
        self.show_btn = QPushButton("顯示內容")
        self.show_btn.clicked.connect(self.show_url)
        btn_layout.addWidget(self.show_btn)
        self.save_btn = QPushButton("存入URL")
        self.save_btn.clicked.connect(self.save_url)
        btn_layout.addWidget(self.save_btn)
        self.prev_btn = QPushButton("上一筆")
        self.prev_btn.clicked.connect(self.load_prev)
        btn_layout.addWidget(self.prev_btn)
        self.next_btn = QPushButton("下一筆")
        self.next_btn.clicked.connect(self.load_next)
        btn_layout.addWidget(self.next_btn)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def load_current_record(self):
        if self.current_index < 0 or self.current_index >= len(self.df):
            return
        record = self.df.iloc[self.current_index]
        self.no_label.setText(str(record["#"]))
        self.name_label.setText(str(record["Name"]))
        url = record.get("url", "")
        if pd.isna(url) or url == "":
            self.url_edit.setText("請輸入百寶倉URL")
        else:
            self.url_edit.setText(str(url))
        self.prev_btn.setEnabled(self.current_index > 0)
        self.next_btn.setEnabled(self.current_index < len(self.df) - 1)

    def copy_name(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.name_label.text())
        print("產品名稱已複製到剪貼簿。")

    def show_url(self):
        url = self.url_edit.text().strip()
        if url and url != "請輸入百寶倉URL":
            self.show_btn.setEnabled(False)
            self.worker = ShowUrlWorker(url)
            self.worker.finished_signal.connect(lambda: self.show_btn.setEnabled(True))
            self.worker.start()
        else:
            QMessageBox.information(self, "提示", "請先輸入正確的 URL。", QMessageBox.Ok)

    def save_url(self):
        new_url = self.url_edit.text().strip()
        if not self.is_valid_url(new_url):
            QMessageBox.warning(self, "錯誤", "URL 格式不正確。")
            return
        self.df.at[self.current_index, "url"] = new_url
        try:
            self.df.to_excel(self.products_file, index=False)
            QMessageBox.information(self, "提示", "URL 已儲存。", QMessageBox.Ok)
            print("URL 已儲存到 products_list.xlsx。")
        except Exception as e:
            QMessageBox.warning(self, "錯誤", f"儲存 URL 時出錯：{e}", QMessageBox.Ok)

    def load_prev(self):
        if self.current_index > 0:
            self.current_index -= 1
            self.load_current_record()

    def load_next(self):
        if self.current_index < len(self.df) - 1:
            self.current_index += 1
            self.load_current_record()

    @staticmethod
    def is_valid_url(url):
        pattern = re.compile(r'^https?://(?:www\.)?\S+\.\S+')
        return bool(pattern.match(url))

# 顯示URL的工作線程
class ShowUrlWorker(QThread):
    finished_signal = pyqtSignal()

    def __init__(self, url, parent=None):
        super().__init__(parent)
        self.url = url

    def run(self):
        try:
            with sync_playwright() as p:

                browser = self.playwright.chromium.launch(channel="msedge", headless=False)
                page = browser.new_page()
                page.goto(self.url)
                page.wait_for_event("close")
                browser.close()
        except Exception as e:
            print(f"ShowUrlWorker error: {e}")
        finally:
            self.finished_signal.emit()

# 新增訂單號碼範圍對話框
class OrderRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("依訂單號碼擷取")
        self.resize(300, 150)
        layout = QVBoxLayout()

        self.start_order_label = QLabel("開始訂單號碼：")
        layout.addWidget(self.start_order_label)
        self.start_order_edit = QLineEdit()
        layout.addWidget(self.start_order_edit)

        self.end_order_label = QLabel("結束訂單號碼：")
        layout.addWidget(self.end_order_label)
        self.end_order_edit = QLineEdit()
        layout.addWidget(self.end_order_edit)

        self.start_scrape_btn = QPushButton("開始擷取")
        self.start_scrape_btn.clicked.connect(self.accept)
        layout.addWidget(self.start_scrape_btn)

        self.setLayout(layout)

    def get_order_range(self):
        start_order = self.start_order_edit.text().strip()
        end_order = self.end_order_edit.text().strip()
        return start_order, end_order

# 主應用程式
class OrderScraperApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Goshop 訂單與產品資料抓取工具")
        self.setGeometry(200, 200, 600, 600)
        layout = QVBoxLayout()

        # 初始化UI
        self.initUI(layout)
        self.setLayout(layout)

        # 初始化變數
        self.base_dir = os.getcwd()  # 當前工作目錄
        self.users_file = os.path.join(self.base_dir, "users.xlsx")  # 使用者列表文件
        self.current_user_dir = None  # 當前使用者的目錄
        self.playwright = None
        self.browser = None
        self.page = None

        # 檢查使用者文件是否存在
        if not os.path.exists(self.users_file):
            self.disable_buttons()
            self.log("尚未建立 users.xlsx，請先新增使用者。")
        else:
            self.load_users()

    def initUI(self, layout):
        # 資訊標籤
        self.info_label = QLabel(
            "【訂單資料】\n1. 點擊『啟動瀏覽器並登入』後，手動登入 Goshophsn。\n"
            "2. 登入完成後，返回此視窗點擊【抓取訂單】。\n"
            "   (抓取過程中若遇到 lastorder.txt 中的 Order Code (Delivery Status 為 pending)，則停止抓取。)\n\n"
            "【產品資料】\n點擊【更新產品資料】後，程式將至產品頁面抓取資料並存成 products_list.xlsx。\n"
            "【更新產品URL】則會讀取 products_list.xlsx 資料，讓您逐筆編輯 URL。\n\n"
            "【使用者管理】\n請先建立 users.xlsx 後，下拉式選單選擇使用者，\n"
            "否則其他功能按鈕將被禁用。"
        )
        self.info_label.setAlignment(Qt.AlignLeft)
        layout.addWidget(self.info_label)

        # 日誌顯示
        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        # 使用者管理
        self.user_combo = QComboBox()
        self.user_combo.currentIndexChanged.connect(self.change_base_dir)
        layout.addWidget(QLabel("請選擇使用者："))
        layout.addWidget(self.user_combo)
        self.add_user_btn = QPushButton("新增使用者")
        self.add_user_btn.clicked.connect(self.add_user)
        layout.addWidget(self.add_user_btn)

        # 功能按鈕
        self.open_browser_btn = QPushButton("啟動瀏覽器並登入")
        self.open_browser_btn.clicked.connect(self.open_browser)
        layout.addWidget(self.open_browser_btn)

        self.scrape_orders_btn = QPushButton("抓取訂單")
        self.scrape_orders_btn.clicked.connect(self.scrape_data)
        layout.addWidget(self.scrape_orders_btn)

        self.update_products_btn = QPushButton("更新產品資料")
        self.update_products_btn.clicked.connect(self.update_products_data)
        layout.addWidget(self.update_products_btn)

        self.update_product_url_btn = QPushButton("更新產品URL")
        self.update_product_url_btn.clicked.connect(self.update_product_url)
        layout.addWidget(self.update_product_url_btn)

        # 新增按鈕
        self.scrape_by_order_range_btn = QPushButton("依訂單號碼擷取")
        self.scrape_by_order_range_btn.clicked.connect(self.scrape_by_order_range)
        layout.addWidget(self.scrape_by_order_range_btn)

        # 新增 "選擇訂單並出貨" 按鈕
        self.select_order_btn = QPushButton("選擇訂單並出貨")
        self.select_order_btn.clicked.connect(self.select_and_ship_order)
        layout.addWidget(self.select_order_btn)

        # 銷售資訊顯示
        self.sales_info_label = QLabel("銷售總合：尚無資料", self)
        self.sales_info_label.setAlignment(Qt.AlignLeft)
        layout.addWidget(self.sales_info_label)

    def show_order_confirmation_dialog(self, df_orders):
        """顯示訂單確認對話框，讓使用者確認訂單內容"""
        if df_orders.empty:
            QMessageBox.warning(self, "錯誤", "無法顯示確認對話框，訂單資料為空。", QMessageBox.Ok)
            return

        print("✅ 進入 show_order_confirmation_dialog 方法")  # Debug

        # 計算總數量
        total_quantity = df_orders["Quantity"].sum()

        # 建立對話框
        dialog = QDialog(self)
        dialog.setWindowTitle("確認訂單內容")
        dialog.resize(400, 400)

        layout = QVBoxLayout()

        # 顯示訂單數
        order_count_label = QLabel(f"📦 訂單數: {len(df_orders)}")
        layout.addWidget(order_count_label)

        # 顯示總數量
        total_quantity_label = QLabel(f"總數量: {total_quantity}")
        layout.addWidget(total_quantity_label)

        # 建立列表元件，顯示所有訂單產品資訊
        order_list = QListWidget()
        for idx, row in enumerate(df_orders.iterrows(), start=1):
            order_info = row[1]  # row 是一個 tuple，第二個元素是 DataFrame 的行
            order_list.addItem(
                f"{idx:2}. {order_info['Product Name']} - {order_info['Attribute']} - 數量: {order_info['Quantity']}")
        layout.addWidget(order_list)

        # 確認進入出貨流程的按鈕
        confirm_button = QPushButton("確認後進入出貨流程")
        confirm_button.clicked.connect(lambda: self.start_shipping_process(df_orders, dialog))
        layout.addWidget(confirm_button)

        dialog.setLayout(layout)

        print("✅ 嘗試執行 dialog.exec_()")  # Debug
        dialog.setWindowModality(Qt.ApplicationModal)  # 確保對話框顯示於最上層
        dialog.exec_()  # 顯示對話框
        print("✅ 對話框關閉")  # Debug

    def log(self, message):
        self.log_text.append(message)
        self.log_text.ensureCursorVisible()
        print(message)

    def disable_buttons(self):
        self.scrape_orders_btn.setEnabled(False)
        self.update_products_btn.setEnabled(False)
        self.update_product_url_btn.setEnabled(False)
        self.scrape_by_order_range_btn.setEnabled(False)

    def load_users(self):
        if os.path.exists(self.users_file):
            try:
                df_users = pd.read_excel(self.users_file)
                self.user_combo.clear()
                for user in df_users["user"]:
                    self.user_combo.addItem(user)
                self.log("載入使用者資料成功。")
            except Exception as e:
                self.log(f"載入 users.xlsx 時出錯：{e}")
        else:
            self.log("未找到 users.xlsx，請建立此檔案。")

    def change_base_dir(self):
        user = self.user_combo.currentText()
        if user:
            self.current_user_dir = os.path.join(self.base_dir, user)
            if not os.path.exists(self.current_user_dir):
                os.makedirs(self.current_user_dir)
                self.log(f"建立新目錄：{self.current_user_dir}")
            self.log(f"已切換到使用者目錄：{self.current_user_dir}")
        else:
            self.log("未選擇使用者。")

    def process_shipping(self, file_path):
        # 讀取選擇的 Excel 訂單檔案，確認 '合併後資料' 頁面，並顯示確認對話框
        print(f"✅ 開始處理出貨流程，檔案路徑: {file_path}")  # Debug
        try:
            # 嘗試讀取 Excel 檔案的 "合併後資料" 工作表
            try:
                df_orders = pd.read_excel(file_path, sheet_name="合併後資料")
            except Exception as e:
                QMessageBox.critical(self, "錯誤",
                                     f"無法讀取 '合併後資料' 頁面，請確認檔案格式是否正確。\n錯誤訊息: {e}")
                return

            # 確保工作表有內容
            if df_orders.empty or "Product Name" not in df_orders.columns or "Attribute" not in df_orders.columns or "Quantity" not in df_orders.columns:
                QMessageBox.critical(self, "錯誤",
                                     "合併後資料缺少必要欄位 (Product Name, Attribute, Quantity)，請確認訂單檔案格式。")
                return

            # 開啟訂單確認對話框
            print("✅ 顯示訂單確認對話框")  # Debug
            self.show_order_confirmation_dialog(df_orders)

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"出貨流程發生錯誤: {e}")

    def start_shipping_process(self, df_orders, dialog):
        """按下確認後開始逐筆處理訂單出貨"""
        dialog.accept()  # 關閉訂單確認視窗
        QMessageBox.information(self, "開始出貨", "即將進入逐筆出貨流程，請稍候...")

        # 啟動瀏覽器並導航到登錄頁面
        try:
            self.log("正在啟動瀏覽器並導航到登錄頁面...")
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(channel="msedge", headless=False, args=["--disable-extensions"])
            context = self.browser.new_context()
            self.page = context.new_page()
            self.page.goto("https://baibaoshop.com/my-account/")
            self.page.mouse.move(random.randint(0, 1000), random.randint(0, 1000))
            # self.time.sleep(random.uniform(1, 3))  # 隨機暫停
            # self.page.click("body > div.wd-page-wrapper.website-wrapper > header > div > div.whb-row.whb-general-header.whb-not-sticky-row.whb-with-bg.whb-border-fullwidth.whb-color-light.whb-flex-equal-sides > div > div > div.whb-column.whb-col-right.whb-visible-lg > div.wd-header-my-account.wd-tools-element.wd-event-hover.wd-design-1.wd-account-style-icon.whb-vssfpylqqax9pvkfnxoz > a > span.wd-tools-icon")
            # self.page.click("body > div.wd-page-wrapper.website-wrapper > header > div > div.whb-row.whb-general-header.whb-not-sticky-row.whb-with-bg.whb-border-fullwidth.whb-color-light.whb-flex-equal-sides > div > div > div.whb-column.whb-col-left.whb-visible-lg > div.site-logo > a > imgselector")  # 替換為實際的選擇器
            # self.time.sleep(random.uniform(1, 3))  # 隨機暫停
            self.log("請在新開啟的瀏覽器中手動登入。")
            QMessageBox.information(self, "請手動登入", "請在瀏覽器中手動登入 Goshop，完成後點擊 '確定' 按鈕以繼續。")

            self.log("用戶已確認登入，繼續執行下一步...")
        except Exception as e:
            self.log(f"啟動瀏覽器時出錯：{e}")
            return

        # 遍歷所有訂單逐筆出貨
        print("✅ 開始逐筆出貨流程",df_orders)  # Debug

        for idx, row in df_orders.iterrows():
            product_name = row["Product Name"]
            attribute = row["Attribute"]
            quantity = row["Quantity"]
            print("Product Name",product_name)  # Debug
            link_url = row["LINK"]
            print("URL",link_url)  # Debug
            # 打開每個訂單的 URL
            try:
                self.log(f"正在打開訂單 URL: {link_url}")
                self.page.goto(link_url)
                # 在這裡可以加入與出貨 API 整合的邏輯
                self.log(f"正在出貨: {product_name} - {attribute} - 數量: {quantity}")
                QMessageBox.information(self, "出貨中",
                                        f"正在出貨\n\n產品: {product_name}\n規格: {attribute}\n數量: {quantity}")
            except Exception as e:
                self.log(f"打開訂單 URL 時出錯：{e}")

        # 出貨完成
        QMessageBox.information(self, "出貨完成", "所有訂單已成功完成出貨！")
        self.log("所有訂單已成功完成出貨！")

        # 關閉瀏覽器
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

    def select_and_ship_order(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(
            self, "選擇訂單檔案", self.current_user_dir, "Excel Files (goshop_orders_*.xlsx)", options=options
        )

        if file_path:
            self.log(f"已選擇訂單檔案: {file_path}")
            try:
                self.process_shipping(file_path)  # 呼叫出貨處理
            except Exception as e:
                self.log(f"讀取訂單檔案時出錯: {e}")
        else:
            self.log("未選擇任何檔案。")

    def add_user(self):
        dialog = AddUserDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            new_user = dialog.get_username()
            if new_user:
                if os.path.exists(self.users_file):
                    df_users = pd.read_excel(self.users_file)
                else:
                    df_users = pd.DataFrame(columns=["user"])
                if new_user in df_users["user"].values:
                    QMessageBox.information(self, "提示", "此使用者已存在。", QMessageBox.Ok)
                else:
                    df_users = pd.concat([df_users, pd.DataFrame({"user": [new_user]})], ignore_index=True)
                    df_users.to_excel(self.users_file, index=False)
                    self.log(f"使用者 {new_user} 已新增到 {self.users_file}。")
                    self.user_combo.addItem(new_user)
                    self.change_base_dir()
            else:
                QMessageBox.information(self, "提示", "請輸入使用者名稱。", QMessageBox.Ok)

    def open_browser(self):
        self.log("正在啟動瀏覽器...")
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(channel="msedge", headless=False)
            context = self.browser.new_context()
            self.page = context.new_page()
            self.page.goto("https://goshophsn.com/users/login")
            self.log("請在新開啟的瀏覽器中手動登入 Goshophsn。\n登入完成後，返回此視窗並點擊【抓取訂單】、"
                     "【更新產品資料】或【更新產品URL】。")
        except Exception as e:
            self.log(f"啟動瀏覽器時出錯：{e}")

    def scrape_data(self):
        user_prefix = self.user_combo.currentText()  # 獲取目前選擇的使用者名稱
        # 從使用者目錄下讀取產品目錄 (products_list.xlsx)
        products_file = os.path.join(self.current_user_dir, "products_list.xlsx")
        if not os.path.exists(products_file):
            QMessageBox.information(self, "提示", "請先建立產品目錄 (products_list.xlsx)", QMessageBox.Ok)
            return

        if not self.page:
            self.log("請先啟動瀏覽器並手動登入。")
            return

        stop_order_code = None
        lastorder_file = os.path.join(self.current_user_dir, "lastorder.txt")
        if os.path.exists(lastorder_file):
            try:
                with open(lastorder_file, "r") as f:
                    stop_order_code = f.read().strip()

                self.log(f"讀取到 lastorder.txt 的 Order Code：{stop_order_code}")
            except Exception as e:
                self.log(f"讀取 lastorder.txt 出錯：{e}")
        else:
            self.log("未找到 lastorder.txt，將分別存 Pending 與非 Pending 的訂單。")
            stop_order_code = None

        try:
            self.log("正在導航到訂單頁面...")
            self.page.goto("https://goshophsn.com/seller/orders")
            self.page.wait_for_load_state('networkidle')

            # 分別儲存 Pending 與非 Pending 的訂單
            pending_orders = []
            rest_orders = []
            stop_grabbing = False

            while not stop_grabbing:
                self.page.wait_for_selector("table tbody tr", timeout=10000)
                self.log("正在抓取當前分頁訂單資料...")
                table_rows = self.page.locator("table tbody tr")
                for i in range(table_rows.count()):
                    row_data = table_rows.nth(i).locator("td").all_inner_texts()
                    cleaned_row_data = [cell.strip() for cell in row_data]
                    status = cleaned_row_data[7].lower()  # Delivery Status 欄位
                    order_code = cleaned_row_data[1]
                    if stop_order_code and order_code == stop_order_code:
                        print(f"遇到訂單編號 {order_code}，停止抓取。")
                        stop_grabbing = True
                        break
                    # 若為 Pending 訂單
                    if status == "pending" :

                        try:
                            cleaned_row_data[4] = float(cleaned_row_data[4].replace("$", "").replace(",", ""))
                        except Exception:
                            cleaned_row_data[4] = 0.0
                        try:
                            cleaned_row_data[5] = float(cleaned_row_data[5].replace("$", "").replace(",", ""))
                        except Exception:
                            cleaned_row_data[5] = 0.0
                        try:
                            cleaned_row_data[6] = float(cleaned_row_data[6].replace("$", "").replace(",", ""))
                        except Exception:
                            cleaned_row_data[6] = 0.0
                        try:
                            cleaned_row_data[2] = int(cleaned_row_data[2])
                        except Exception:
                            cleaned_row_data[2] = 0
                        pending_orders.append(cleaned_row_data)
                    else:
                        # 非 Pending 訂單
                        try:
                            cleaned_row_data[4] = float(cleaned_row_data[4].replace("$", "").replace(",", ""))
                        except Exception:
                            cleaned_row_data[4] = 0.0
                        try:
                            cleaned_row_data[5] = float(cleaned_row_data[5].replace("$", "").replace(",", ""))
                        except Exception:
                            cleaned_row_data[5] = 0.0
                        try:
                            cleaned_row_data[6] = float(cleaned_row_data[6].replace("$", "").replace(",", ""))
                        except Exception:
                            cleaned_row_data[6] = 0.0
                        try:
                            cleaned_row_data[2] = int(cleaned_row_data[2])
                        except Exception:
                            cleaned_row_data[2] = 0
                        rest_orders.append(cleaned_row_data)
                if not stop_grabbing:
                    next_button = self.page.locator("a[aria-label='Next »']")
                    if next_button.is_visible():
                        next_button.click()
                        self.page.wait_for_load_state('networkidle')
                    else:
                        self.log("所有分頁抓取完畢。")
                        break
                else:
                    self.log("抓取已因遇到 lastorder.txt 指定的 Order Code 而停止。")
                    break

            columns = ["#", "Order Code", "Num. of Products", "Customer", "Amount", "Service charge",
                       "Final price", "Delivery Status", "Payment Status", "Product Info", "Options"]

            if os.path.exists(lastorder_file):
                # 若 lastorder.txt 存在，僅存 Pending 訂單
                df_pending = pd.DataFrame(pending_orders, columns=columns)
                file_path = os.path.join(self.base_dir, f"goshop_orders_{datetime.now().strftime('%Y%m%d')}.xlsx")

                df_pending.to_excel(file_path, index=False)
                self.log(f"訂單資料已存成 Excel 檔案：{file_path}")
                self.update_sales_file(df_pending)
            else:
                # 若 lastorder.txt 不存在，分別存 Pending 與 Rest 訂單
                df_pending = pd.DataFrame(pending_orders, columns=columns)
                df_rest = pd.DataFrame(rest_orders, columns=columns)
                file_path_pending = os.path.join(self.base_dir, f"goshop_orders_{datetime.now().strftime('%Y%m%d')}.xlsx")
                file_path_rest = os.path.join(self.base_dir, "rest-order.xlsx")
                df_pending.to_excel(file_path_pending, index=False)
                df_rest.to_excel(file_path_rest, index=False)
                self.log(f"訂單資料已分別存成 Excel 檔案：{file_path_pending} (Pending) 與 {file_path_rest} (Rest)")
                if not df_pending.empty:
                    first_order_code = df_pending["Order Code"].iloc[0].strip()
                    with open(lastorder_file, "w") as f:
                        f.write(first_order_code)
                    self.log(f"已建立 {lastorder_file}，內容為第一筆訂單的 Order Code：{first_order_code}")
                self.update_sales_file_split(df_pending, df_rest)
        except Exception as e:
            self.log(f"抓取資料時出錯：{e}")
        finally:
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()

    def scrape_by_order_range(self):
        if not self.current_user_dir:
            self.log("請先選擇使用者。")
            return

        dialog = OrderRangeDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            start_order, end_order = dialog.get_order_range()
            if start_order and end_order:
                self.scrape_data_by_order_range(start_order, end_order)
            else:
                QMessageBox.information(self, "提示", "請輸入開始和結束訂單號碼。", QMessageBox.Ok)

    def scrape_data_by_order_range(self, start_order, end_order):
        if not self.page:
            self.log("請先啟動瀏覽器並手動登入。")
            return

        try:
            self.log("正在導航到訂單頁面...")
            self.page.goto("https://goshophsn.com/seller/orders")
            self.page.wait_for_load_state('networkidle')

            all_data = []
            start_scraping = False  # 控制是否記錄資料的旗標
            found_end_order = False  # 新增標誌來確認是否已找到結束訂單

            while True:
                self.page.wait_for_selector("table tbody tr", timeout=10000)
                self.log("正在抓取當前分頁訂單資料...")
                table_rows = self.page.locator("table tbody tr")
                row_count = table_rows.count()

                for i in range(row_count):
                    row_data = table_rows.nth(i).locator("td").all_inner_texts()
                    cleaned_row_data = [cell.strip() for cell in row_data]
                    order_code = cleaned_row_data[1]  # 假設第 2 欄是訂單編號
                    print(f"當前處理訂單編號: {order_code}")

                    # 先檢查是否到達起始訂單
                    if order_code == start_order:
                        start_scraping = True
                        self.log("找到起始訂單，開始記錄資料...")

                    # 再檢查是否要結束
                    if order_code == end_order:
                        self.log("已找到結束訂單，停止記錄並退出...")
                        if start_scraping:  # 確保只有當在記錄狀態時才添加結束訂單
                            all_data.append(cleaned_row_data)
                            print("已記錄:", cleaned_row_data)
                        found_end_order = True  # 設置找到結束標誌
                        break  # 跳出內部迴圈

                    # 最後處理資料記錄
                    if start_scraping:
                        all_data.append(cleaned_row_data)
                        print("已記錄:", cleaned_row_data)

                # 內部迴圈結束後檢查是否找到結束訂單
                if found_end_order:
                    break  # 完全終止外部迴圈

                # 尚未找到結束訂單時檢查翻頁
                next_button = self.page.locator("a[aria-label='Next »']")
                if next_button.is_visible():
                    self.log("正在翻到下一頁...")
                    next_button.click()
                    self.page.wait_for_load_state('networkidle')
                else:
                    self.log("已遍歷所有分頁，但未找到結束訂單。")
                    break  # 沒有更多分頁可處理

            # 以下保存資料的邏輯保持不變
            if not all_data:
                self.log("未抓取到任何訂單資料。")
                return

            df_original = pd.DataFrame(all_data, columns=[
                "#", "Order Code", "Num. of Products", "Customer", "Amount", "Service charge",
                "Final price", "Delivery Status", "Payment Status", "Product Info", "Options"
            ])

            split_df, merged_df = self.split_and_merge_orders(df_original)

            file_path = os.path.join(self.current_user_dir, f"goshop_orders_{start_order}_to_{end_order}.xlsx")
            with pd.ExcelWriter(file_path) as writer:
                df_original.to_excel(writer, sheet_name="原始資料", index=False)
                split_df.to_excel(writer, sheet_name="拆分後資料", index=False)
                merged_df.to_excel(writer, sheet_name="合併後資料", index=False)

            self.log(f"訂單資料已存成 Excel 檔案：{file_path}")
            # self.update_sales_file(df_original)

        except Exception as e:
            self.log(f"抓取訂單資料時發生錯誤: {e}")

    def update_sales_file(self, df):
        try:
            today = datetime.now().strftime("%Y-%m-%d")
            total_amount = df["Amount"].sum()
            total_service_charge = df["Service charge"].sum()
            total_final_price = df["Final price"].sum()

            new_data = {
                "日期": [today],
                "Amount": [total_amount],
                "Service charge": [total_service_charge],
                "Final price": [total_final_price]
            }

            sales_file = os.path.join(self.current_user_dir, "sales.xlsx")
            if os.path.exists(sales_file):
                sales_df = pd.read_excel(sales_file)
                sales_df = pd.concat([sales_df, pd.DataFrame(new_data)], ignore_index=True)
            else:
                sales_df = pd.DataFrame(new_data)

            sales_df.to_excel(sales_file, index=False)
            self.log(f"已更新或建立 {sales_file} 檔案。")
            self.log(f"銷售總合 -> Amount: {total_amount:.2f}, Service charge: {total_service_charge:.2f}, Final price: {total_final_price:.2f}")
        except Exception as e:
            self.log(f"更新銷售檔案時出錯：{e}")

    def update_sales_file_split(self, df_pending, df_rest):
        try:
            today = datetime.now().strftime("%Y-%m-%d")

            total_amount_pending = df_pending["Amount"].sum()
            total_service_charge_pending = df_pending["Service charge"].sum()
            total_final_price_pending = df_pending["Final price"].sum()

            total_amount_rest = df_rest["Amount"].sum()
            total_service_charge_rest = df_rest["Service charge"].sum()
            total_final_price_rest = df_rest["Final price"].sum()

            new_data_pending = {
                "日期": [today],
                "Amount": [total_amount_pending],
                "Service charge": [total_service_charge_pending],
                "Final price": [total_final_price_pending]
            }

            new_data_rest = {
                "日期": [today],
                "Amount": [total_amount_rest],
                "Service charge": [total_service_charge_rest],
                "Final price": [total_final_price_rest]
            }

            sales_file_pending = os.path.join(self.current_user_dir, "sales_pending.xlsx")
            sales_file_rest = os.path.join(self.current_user_dir, "sales_rest.xlsx")

            if os.path.exists(sales_file_pending):
                sales_df_pending = pd.read_excel(sales_file_pending)
                sales_df_pending = pd.concat([sales_df_pending, pd.DataFrame(new_data_pending)], ignore_index=True)
            else:
                sales_df_pending = pd.DataFrame(new_data_pending)

            if os.path.exists(sales_file_rest):
                sales_df_rest = pd.read_excel(sales_file_rest)
                sales_df_rest = pd.concat([sales_df_rest, pd.DataFrame(new_data_rest)], ignore_index=True)
            else:
                sales_df_rest = pd.DataFrame(new_data_rest)

            sales_df_pending.to_excel(sales_file_pending, index=False)
            sales_df_rest.to_excel(sales_file_rest, index=False)

            self.log(f"已更新或建立 {sales_file_pending} 與 {sales_file_rest} 檔案。")
            self.log(f"銷售總合 (Pending) -> Amount: {total_amount_pending:.2f}, Service charge: {total_service_charge_pending:.2f}, Final price: {total_final_price_pending:.2f}")
            self.log(f"銷售總合 (Rest) -> Amount: {total_amount_rest:.2f}, Service charge: {total_service_charge_rest:.2f}, Final price: {total_final_price_rest:.2f}")
        except Exception as e:
            self.log(f"更新銷售檔案時出錯：{e}")

    def split_and_merge_orders(self, df):
        split_rows = []
        for idx, row in df.iterrows():
            product_info = row["Product Info"]
            lines = product_info.strip().split("\n")
            for line in lines:
                parts = [p.strip() for p in line.split("|")]
                if len(parts) >= 3:
                    product_name = parts[0]
                    attribute = parts[1].replace("；", "").strip()
                    quantity_str = parts[2].strip()
                    try:
                        quantity = int(quantity_str)
                    except ValueError:
                        print(f"警告：數量無法解析，忽略此產品。訂單編號：{row['Order Code']}，產品資訊：{line}")
                        continue
                    split_rows.append([row["Order Code"], product_name, attribute, quantity])
                else:
                    print(f"警告：無法解析產品資訊：{line}")
        split_df = pd.DataFrame(split_rows, columns=["Order Code", "Product Name", "Attribute", "Quantity"])
        merged_df = split_df.groupby(["Product Name", "Attribute"], as_index=False).agg({
            "Order Code": lambda x: ";".join(x),
            "Quantity": "sum"
        })
        return split_df, merged_df

    def update_products_data(self):
        if not self.current_user_dir:
            self.log("請先選擇使用者。")
            return

        products_file = os.path.join(self.current_user_dir, "products_list.xlsx")
        if not os.path.exists(products_file):
            QMessageBox.information(self, "提示", "產品目錄不存在，開始抓取產品資料...", QMessageBox.Ok)
            self.scrape_products_data()
            return

        try:
            df_products = pd.read_excel(products_file)
            if df_products.shape[1] < 10:
                df_products["url"] = ""
                df_products = df_products[["#", "Thumbnail Image", "Name", "Category", "Current Qty",
                                           "Base Price", "Published", "Examine Status", "Options", "url"]]
                df_products.to_excel(products_file, index=False)
                self.log(f"更新產品檔案欄位，補上 'url' 欄。")
            self.update_product_url()
        except Exception as e:
            self.log(f"更新產品資料時出錯：{e}")
            reply = QMessageBox.question(self, "錯誤", "更新產品資料時出錯，是否重新登入？", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.open_browser()
            else:
                self.log("取消重新登入。")

    def scrape_products_data(self):
        if not self.page:
            self.log("請先啟動瀏覽器並手動登入。")
            return

        try:
            self.log("正在導航到產品頁面...")
            self.page.goto("https://goshophsn.com/seller/products")
            self.page.wait_for_load_state('networkidle')

            all_data = []
            while True:
                self.page.wait_for_selector("table tbody tr", timeout=10000)
                self.log("正在抓取當前分頁產品資料...")
                table_rows = self.page.locator("table tbody tr")
                for i in range(table_rows.count()):
                    row_data = table_rows.nth(i).locator("td").all_inner_texts()
                    cleaned_row_data = [cell.strip() for cell in row_data]
                    all_data.append(cleaned_row_data)

                next_button = self.page.locator("a[aria-label='Next »']")
                if next_button.is_visible():
                    next_button.click()
                    self.page.wait_for_load_state('networkidle')
                else:
                    self.log("所有分頁抓取完畢。")
                    break

            if not all_data:
                self.log("未抓取到任何產品資料。")
                return

            df_products = pd.DataFrame(all_data, columns=[
                "#", "Thumbnail Image", "Name", "Category", "Current Qty",
                "Base Price", "Published", "Examine Status", "Options"
            ])
            df_products["url"] = ""

            products_file = os.path.join(self.current_user_dir, "products_list.xlsx")
            df_products.to_excel(products_file, index=False)
            self.log(f"產品資料已存成 Excel 檔案：{products_file}")

        except Exception as e:
            self.log(f"抓取產品資料時出錯：{e}")
        finally:
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()

    def update_product_url(self):
        if not self.current_user_dir:
            self.log("請先選擇使用者。")
            return

        products_file = os.path.join(self.current_user_dir, "products_list.xlsx")
        if not os.path.exists(products_file):
            QMessageBox.information(self, "提示", "請重建產品目錄", QMessageBox.Ok)
            return
        try:
            dialog = UpdateProductURLDialog(products_file, self)
            dialog.exec_()
        except Exception as e:
            self.log(f"更新產品URL時出錯：{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OrderScraperApp()
    window.show()
    sys.exit(app.exec_())
土