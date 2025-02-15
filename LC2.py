#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import random
import sys
import os
import time
import traceback
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel, QMessageBox, QDialog,
    QHBoxLayout, QLineEdit, QComboBox, QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from playwright.sync_api import sync_playwright


# ===============================
# 輔助對話框：新增使用者
# ===============================
class AddUserDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("新增使用者")
        self.resize(300, 100)
        layout = QVBoxLayout()
        self.username_edit = QLineEdit(self)
        self.username_edit.setPlaceholderText("請輸入使用者名稱")
        layout.addWidget(self.username_edit)
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("確定")
        cancel_btn = QPushButton("取消")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def get_username(self):
        return self.username_edit.text().strip()


# ===============================
# 輔助對話框：訂單範圍輸入
# ===============================
class OrderRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("依訂單號碼擷取")
        self.resize(300, 120)
        layout = QVBoxLayout()

        self.start_edit = QLineEdit(self)
        self.start_edit.setPlaceholderText("起始訂單號碼")
        layout.addWidget(QLabel("起始訂單號碼："))
        layout.addWidget(self.start_edit)

        self.end_edit = QLineEdit(self)
        self.end_edit.setPlaceholderText("結束訂單號碼")
        layout.addWidget(QLabel("結束訂單號碼："))
        layout.addWidget(self.end_edit)

        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("確定")
        cancel_btn = QPushButton("取消")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def get_order_range(self):
        return self.start_edit.text().strip(), self.end_edit.text().strip()


# ===============================
# 輔助對話框：更新產品 URL
# 將產品清單讀取進 QTableWidget 供使用者編輯 URL 欄位
# ===============================
class UpdateProductURLDialog(QDialog):
    def __init__(self, products_file, parent=None):
        super().__init__(parent)
        self.setWindowTitle("更新產品 URL")
        self.resize(800, 600)
        self.products_file = products_file
        self.df_products = pd.read_excel(products_file)
        # 如果沒有 url 欄位就新增
        if "url" not in self.df_products.columns:
            self.df_products["url"] = ""
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["產品名稱", "URL"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setRowCount(len(self.df_products))

        for row in range(len(self.df_products)):
            product_name = str(self.df_products.iloc[row].get("Name", ""))
            url = str(self.df_products.iloc[row].get("url", ""))
            name_item = QTableWidgetItem(product_name)
            name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.table.setItem(row, 0, name_item)
            url_item = QTableWidgetItem(url)
            url_item.setFlags(Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.table.setItem(row, 1, url_item)

        layout.addWidget(self.table)
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("儲存")
        cancel_btn = QPushButton("取消")
        save_btn.clicked.connect(self.save_data)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def save_data(self):
        # 將更新後的 URL 寫回 dataframe 並存檔
        urls = []
        for row in range(self.table.rowCount()):
            url_item = self.table.item(row, 1)
            urls.append(url_item.text() if url_item is not None else "")
        self.df_products["url"] = urls
        try:
            self.df_products.to_excel(self.products_file, index=False)
            QMessageBox.information(self, "提示", "產品 URL 已更新！")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"儲存產品 URL 時發生錯誤：{e}")


# ===============================
# 逐筆出貨對話框
# ===============================
class OrderProcessingDialog(QDialog):
    def __init__(self, df_orders, parent=None):
        super().__init__(parent)
        self.setWindowTitle("訂單處理")
        self.df_orders = df_orders
        self.current_index = 0
        # 透過 parent 取得瀏覽器 page（若有）
        self.page = parent.page if hasattr(parent, 'page') else None
        self.initUI()
        self.show_current_order()

    def initUI(self):
        layout = QVBoxLayout()

        self.order_label = QLabel("")
        layout.addWidget(self.order_label)

        self.next_button = QPushButton("下一筆")
        self.next_button.clicked.connect(self.process_next_order)
        layout.addWidget(self.next_button)

        self.setLayout(layout)

    def show_current_order(self):
        if self.current_index < 0 or self.current_index >= len(self.df_orders):
            return
        row = self.df_orders.iloc[self.current_index]
        product_name = row["Product Name"]
        attribute = row["Attribute"]
        quantity = row["Quantity"]
        link_url = row["URL"]

        # 嘗試打開每個訂單的 URL
        try:
            self.log(f"正在打開訂單 URL: {link_url}")
            if self.page:
                self.page.goto(link_url)
                time.sleep(random.uniform(1, 3))  # 隨機暫停
            self.order_label.setText(f"正在出貨: {product_name} - {attribute} - 數量: {quantity}")
        except Exception as e:
            self.log(f"打開訂單 URL 時出錯：{e}")

    def process_next_order(self):
        # 此處可加入出貨 API 整合邏輯
        self.log(f"正在出貨: {self.order_label.text()}")
        if self.current_index < len(self.df_orders) - 1:
            self.current_index += 1
            self.show_current_order()
        else:
            QMessageBox.information(self, "出貨完成", "所有訂單已成功完成出貨！")
            self.close()

    def log(self, message):
        print(message)


# ===============================
# 主應用程式
# ===============================
class OrderScraperApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Goshop 訂單與產品資料抓取工具")
        self.setGeometry(200, 200, 600, 600)
        layout = QVBoxLayout()
        self.initUI(layout)
        self.setLayout(layout)

        # 初始化變數
        self.base_dir = os.getcwd()  # 當前工作目錄（僅存放 users.xlsx）
        self.users_file = os.path.join(self.base_dir, "users.xlsx")  # 使用者列表文件
        self.current_user_dir = None  # 當前使用者的目錄
        self.playwright = None
        self.browser = None
        self.page = None
        self.df_orders = None  # 儲存訂單資料

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
        layout.addWidget(QLabel("請選擇使用者："))
        self.user_combo = QComboBox()
        self.user_combo.currentIndexChanged.connect(self.change_base_dir)
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

        # 新增按鈕：依訂單號碼擷取
        self.scrape_by_order_range_btn = QPushButton("依訂單號碼擷取")
        self.scrape_by_order_range_btn.clicked.connect(self.scrape_by_order_range)
        layout.addWidget(self.scrape_by_order_range_btn)

        # 新增按鈕：選擇訂單並出貨
        self.select_order_btn = QPushButton("選擇訂單並出貨")
        self.select_order_btn.clicked.connect(self.select_and_ship_order)
        layout.addWidget(self.select_order_btn)

        # 新增按鈕：逐筆下單（啟動逐筆出貨流程）
        self.process_orders_btn = QPushButton("逐筆下單")
        self.process_orders_btn.clicked.connect(self.start_order_processing)
        layout.addWidget(self.process_orders_btn)

        # 銷售資訊顯示
        self.sales_info_label = QLabel("銷售總合：尚無資料", self)
        self.sales_info_label.setAlignment(Qt.AlignLeft)
        layout.addWidget(self.sales_info_label)

    def disable_buttons(self):
        self.scrape_orders_btn.setEnabled(False)
        self.update_products_btn.setEnabled(False)
        self.update_product_url_btn.setEnabled(False)
        self.scrape_by_order_range_btn.setEnabled(False)
        self.select_order_btn.setEnabled(False)
        self.process_orders_btn.setEnabled(False)

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

    def log(self, message):
        # 同時列印到 console 與日誌文字編輯器
        print(message)
        self.log_text.append(message)

    # -------------------------------
    # 出貨處理流程（讀取訂單 Excel 後確認出貨）
    # -------------------------------
    def process_shipping(self, file_path):
        self.log(f"✅ 開始處理出貨流程，檔案路徑: {file_path}")
        try:
            # 嘗試讀取 Excel 檔案的 "合併後資料" 工作表
            try:
                df_orders = pd.read_excel(file_path, sheet_name="合併後資料")
            except Exception as e:
                QMessageBox.critical(self, "錯誤",
                                     f"無法讀取 '合併後資料' 頁面，請確認檔案格式是否正確。\n錯誤訊息: {e}")
                return

            # 確保工作表有必要欄位
            required_cols = ["Product Name", "Attribute", "Quantity", "URL"]
            if df_orders.empty or not all(col in df_orders.columns for col in required_cols):
                QMessageBox.critical(self, "錯誤",
                                     "合併後資料缺少必要欄位 (Product Name, Attribute, Quantity, URL)，請確認訂單檔案格式。")
                return

            self.show_order_confirmation_dialog(df_orders)

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"出貨流程發生錯誤: {traceback.format_exc()}")

    def show_order_confirmation_dialog(self, df_orders):
        reply = QMessageBox.question(self, "訂單確認", "是否開始出貨所有訂單？", QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.df_orders = df_orders
            self.start_shipping_process(df_orders)

    def start_shipping_process(self, df_orders):
        QMessageBox.information(self, "開始出貨", "即將進入逐筆出貨流程，請稍候...")
        # 啟動瀏覽器並導航到登錄頁面
        try:
            self.log("正在啟動瀏覽器並導航到登錄頁面...")
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(channel="msedge", headless=False)
            context = self.browser.new_context()
            self.page = context.new_page()
            self.page.goto("https://baibaoshop.com/")
            self.page.mouse.move(random.randint(0, 1000), random.randint(0, 1000))
            time.sleep(random.uniform(1, 3))
            # 請根據實際情況調整下方選擇器
            self.page.click(
                "body > div.wd-page-wrapper.website-wrapper > header > div > div.whb-row.whb-general-header.whb-not-sticky-row.whb-with-bg.whb-border-fullwidth.whb-color-light.whb-flex-equal-sides > div > div > div.whb-column.whb-col-left.whb-visible-lg > div.site-logo > a > img")
            time.sleep(random.uniform(1, 3))
            self.log("請在新開啟的瀏覽器中手動登入。")
        except Exception as e:
            self.log(f"啟動瀏覽器時出錯：{e}")
            return

        # 遍歷所有訂單逐筆出貨
        for idx, row in df_orders.iterrows():
            product_name = row["Product Name"]
            attribute = row["Attribute"]
            quantity = row["Quantity"]
            link_url = row["URL"]

            try:
                self.log(f"正在打開訂單 URL: {link_url}")
                self.page.goto(link_url)
                time.sleep(random.uniform(1, 3))
                self.log(f"正在出貨: {product_name} - {attribute} - 數量: {quantity}")
                QMessageBox.information(self, "出貨中",
                                        f"正在出貨\n\n產品: {product_name}\n規格: {attribute}\n數量: {quantity}")
            except Exception as e:
                self.log(f"打開訂單 URL 時出錯：{e}")

        QMessageBox.information(self, "出貨完成", "所有訂單已成功完成出貨！")
        self.log("所有訂單已成功完成出貨！")
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

    def select_and_ship_order(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(
            self, "選擇訂單檔案", self.current_user_dir if self.current_user_dir else self.base_dir,
            "Excel Files (goshop_orders_*.xlsx)", options=options
        )

        if file_path:
            self.log(f"已選擇訂單檔案: {file_path}")
            try:
                self.process_shipping(file_path)
            except Exception as e:
                self.log(f"讀取訂單檔案時出錯: {traceback.format_exc()}")
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
        if self.browser:
            self.log("瀏覽器已經啟動。")
            return
        try:
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(channel="msedge", headless=False)
            context = self.browser.new_context()
            self.page = context.new_page()
            self.page.goto("https://goshophsn.com/users/login")
            self.log(
                "請在新開啟的瀏覽器中手動登入 Goshophsn。\n登入完成後，返回此視窗並點擊【抓取訂單】、【更新產品資料】或【更新產品URL】。")
        except Exception as e:
            self.log(f"啟動瀏覽器時出錯：{e}")

    def scrape_data(self):
        # 請先確認已選擇使用者
        if not self.current_user_dir:
            self.log("請先選擇使用者。")
            return

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
                with open(lastorder_file, "r", encoding="utf-8") as f:
                    stop_order_code = f.read().strip()
                self.log(f"讀取到 lastorder.txt 的 Order Code：{stop_order_code}")
            except Exception as e:
                self.log(f"讀取 lastorder.txt 出錯：{e}")
        else:
            stop_order_code = None
            self.log("未找到 lastorder.txt，將分別存 Pending 與非 Pending 的訂單。")

        try:
            self.log("正在導航到訂單頁面...")
            self.page.goto("https://goshophsn.com/seller/orders")
            self.page.wait_for_load_state('networkidle')

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
                    order_code = cleaned_row_data[1] if len(cleaned_row_data) > 1 else ""
                    status = cleaned_row_data[7].lower() if len(cleaned_row_data) > 7 else ""
                    if stop_order_code and order_code == stop_order_code:
                        self.log(f"遇到訂單編號 {order_code}，停止抓取。")
                        stop_grabbing = True
                        break
                    # 處理數值欄位
                    for idx in [4, 5, 6]:
                        if len(cleaned_row_data) > idx:
                            try:
                                cleaned_row_data[idx] = float(cleaned_row_data[idx].replace("$", "").replace(",", ""))
                            except Exception:
                                cleaned_row_data[idx] = 0.0
                    try:
                        if len(cleaned_row_data) > 2:
                            cleaned_row_data[2] = int(cleaned_row_data[2])
                    except Exception:
                        cleaned_row_data[2] = 0
                    if status == "pending":
                        pending_orders.append(cleaned_row_data)
                    else:
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
                # 僅存 Pending 訂單
                df_pending = pd.DataFrame(pending_orders, columns=columns)
                file_path = os.path.join(self.current_user_dir,
                                         f"goshop_orders_{datetime.now().strftime('%Y%m%d')}.xlsx")
                df_pending.to_excel(file_path, index=False)
                self.log(f"訂單資料已存成 Excel 檔案：{file_path}")
                self.update_sales_file(df_pending)
            else:
                df_pending = pd.DataFrame(pending_orders, columns=columns)
                df_rest = pd.DataFrame(rest_orders, columns=columns)
                file_path_pending = os.path.join(self.current_user_dir,
                                                 f"goshop_orders_{datetime.now().strftime('%Y%m%d')}.xlsx")
                file_path_rest = os.path.join(self.current_user_dir, "rest-order.xlsx")
                df_pending.to_excel(file_path_pending, index=False)
                df_rest.to_excel(file_path_rest, index=False)
                self.log(f"訂單資料已分別存成 Excel 檔案：{file_path_pending} (Pending) 與 {file_path_rest} (Rest)")
                if not df_pending.empty:
                    first_order_code = str(df_pending["Order Code"].iloc[0]).strip()
                    with open(lastorder_file, "w", encoding="utf-8") as f:
                        f.write(first_order_code)
                    self.log(f"已建立 {lastorder_file}，內容為第一筆訂單的 Order Code：{first_order_code}")
                self.update_sales_file_split(df_pending, df_rest)
        except Exception as e:
            self.log(f"抓取資料時出錯：{traceback.format_exc()}")
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
            start_scraping = False
            found_end_order = False

            while True:
                self.page.wait_for_selector("table tbody tr", timeout=10000)
                self.log("正在抓取當前分頁訂單資料...")
                table_rows = self.page.locator("table tbody tr")
                row_count = table_rows.count()

                for i in range(row_count):
                    row_data = table_rows.nth(i).locator("td").all_inner_texts()
                    cleaned_row_data = [cell.strip() for cell in row_data]
                    order_code = cleaned_row_data[1] if len(cleaned_row_data) > 1 else ""
                    self.log(f"當前處理訂單編號: {order_code}")

                    if order_code == start_order:
                        start_scraping = True
                        self.log("找到起始訂單，開始記錄資料...")
                    if order_code == end_order:
                        self.log("已找到結束訂單，停止記錄並退出...")
                        if start_scraping:
                            all_data.append(cleaned_row_data)
                        found_end_order = True
                        break
                    if start_scraping:
                        all_data.append(cleaned_row_data)
                if found_end_order:
                    break
                next_button = self.page.locator("a[aria-label='Next »']")
                if next_button.is_visible():
                    self.log("正在翻到下一頁...")
                    next_button.click()
                    self.page.wait_for_load_state('networkidle')
                else:
                    self.log("已遍歷所有分頁，但未找到結束訂單。")
                    break

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
        except Exception as e:
            self.log(f"抓取訂單資料時發生錯誤: {traceback.format_exc()}")

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

            sales_file = os.path.join(self.current_user_dir, "sales.xlsx") if self.current_user_dir else "sales.xlsx"
            if os.path.exists(sales_file):
                sales_df = pd.read_excel(sales_file)
                sales_df = pd.concat([sales_df, pd.DataFrame(new_data)], ignore_index=True)
            else:
                sales_df = pd.DataFrame(new_data)

            sales_df.to_excel(sales_file, index=False)
            self.log(f"已更新或建立 {sales_file} 檔案。")
            self.log(
                f"銷售總合 -> Amount: {total_amount:.2f}, Service charge: {total_service_charge:.2f}, Final price: {total_final_price:.2f}")
        except Exception as e:
            self.log(f"更新銷售檔案時出錯：{traceback.format_exc()}")

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

            sales_file_pending = os.path.join(self.current_user_dir,
                                              "sales_pending.xlsx") if self.current_user_dir else "sales_pending.xlsx"
            sales_file_rest = os.path.join(self.current_user_dir,
                                           "sales_rest.xlsx") if self.current_user_dir else "sales_rest.xlsx"

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
            self.log(
                f"銷售總合 (Pending) -> Amount: {total_amount_pending:.2f}, Service charge: {total_service_charge_pending:.2f}, Final price: {total_final_price_pending:.2f}")
            self.log(
                f"銷售總合 (Rest) -> Amount: {total_amount_rest:.2f}, Service charge: {total_service_charge_rest:.2f}, Final price: {total_final_price_rest:.2f}")
        except Exception as e:
            self.log(f"更新銷售檔案時出錯：{traceback.format_exc()}")

    def split_and_merge_orders(self, df):
        split_rows = []
        if "Order Code" not in df.columns or "Product Info" not in df.columns:
            self.log("DataFrame 缺少必要欄位：Order Code 或 Product Info")
            return pd.DataFrame(), pd.DataFrame()
        for idx, row in df.iterrows():
            product_info = row["Product Info"]
            if not isinstance(product_info, str):
                continue
            lines = product_info.strip().split("\n")
            for line in lines:
                parts = [p.strip() for p in line.split("|")]
                if len(parts) >= 3:
                    product_name = parts[0]
                    attribute = parts[1].replace("；", "").strip()
                    quantity_str = parts[2]
                    try:
                        quantity = int(quantity_str)
                    except ValueError:
                        self.log(f"警告：數量無法解析，忽略此產品。訂單編號：{row['Order Code']}，產品資訊：{line}")
                        continue
                    split_rows.append([row["Order Code"], product_name, attribute, quantity])
                else:
                    self.log(f"警告：無法解析產品資訊：{line}")
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
            if df_products.shape[1] < 10 or "url" not in df_products.columns:
                df_products["url"] = ""
                cols = ["#", "Thumbnail Image", "Name", "Category", "Current Qty",
                        "Base Price", "Published", "Examine Status", "Options", "url"]
                df_products = df_products[cols]
                df_products.to_excel(products_file, index=False)
                self.log(f"更新產品檔案欄位，補上 'url' 欄。")
            self.update_product_url()
        except Exception as e:
            self.log(f"更新產品資料時出錯：{traceback.format_exc()}")
            reply = QMessageBox.question(self, "錯誤", "更新產品資料時出錯，是否重新登入？",
                                         QMessageBox.Yes | QMessageBox.No)
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
            self.log(f"抓取產品資料時出錯：{traceback.format_exc()}")
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
            self.log(f"更新產品URL時出錯：{traceback.format_exc()}")

    def start_order_processing(self):
        # 假設 self.df_orders 之前已經由其他流程取得
        if self.df_orders is None or self.df_orders.empty:
            QMessageBox.information(self, "提示", "請先取得訂單資料。", QMessageBox.Ok)
            return
        dialog = OrderProcessingDialog(self.df_orders, self)
        dialog.exec_()


# ===============================
# 主程式進入點
# ===============================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OrderScraperApp()
    window.show()
    sys.exit(app.exec_())
