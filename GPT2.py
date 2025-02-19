#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import random
import sys
import os
import time
import traceback
from tkinter.filedialog import dialogstates

import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel, QMessageBox, QDialog,
    QHBoxLayout, QLineEdit, QComboBox, QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from playwright.sync_api import sync_playwright
from PyQt5.QtGui import QClipboard

class DialogWindow(QWidget):
    def __init__(self):
        super().__init__()

    def show_dialog(self):
        """顯示 PyQt 對話框，等待使用者按下「確定」後才返回"""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("等待確認")
        msg_box.setText("請點擊確定以繼續執行 Playwright")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.setWindowFlag(Qt.WindowStaysOnTopHint)

        # 顯示對話框並阻塞程式直到按鈕被點擊
        msg_box.exec_()

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

        try:
            self.log(f"正在打開訂單 URL: {link_url}")
            if self.page:
                self.page.goto(link_url)
                # time.sleep(random.uniform(1, 3))
            self.order_label.setText(f"正在出貨: {product_name} - {attribute} - 數量: {quantity}")
        except Exception as e:
            self.log(f"打開訂單 URL 時出錯：{e}")
            QMessageBox.information(self, f"打開訂單 URL 時出錯：{e}")

    def process_next_order(self):
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
        self.base_dir = os.getcwd()  # users.xlsx 存放於此
        self.users_file = os.path.join(self.base_dir, "users.xlsx")
        self.current_user_dir = None  # 其他資料檔存放於各使用者目錄下
        self.playwright = None
        self.browser = None
        self.page = None
        self.df_orders = None  # 儲存訂單資料

        if not os.path.exists(self.users_file):
            self.disable_buttons()
            self.log("尚未建立 users.xlsx，請先新增使用者。")
        else:
            self.load_users()

    def initUI(self, layout):
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

        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

        layout.addWidget(QLabel("請選擇使用者："))
        self.user_combo = QComboBox()
        self.user_combo.currentIndexChanged.connect(self.change_base_dir)
        layout.addWidget(self.user_combo)
        self.add_user_btn = QPushButton("新增使用者")
        self.add_user_btn.clicked.connect(self.add_user)
        layout.addWidget(self.add_user_btn)

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

        self.scrape_by_order_range_btn = QPushButton("依訂單號碼擷取")
        self.scrape_by_order_range_btn.clicked.connect(self.scrape_by_order_range)
        layout.addWidget(self.scrape_by_order_range_btn)

        self.select_order_btn = QPushButton("選擇訂單並出貨")
        self.select_order_btn.clicked.connect(self.select_and_ship_order)
        layout.addWidget(self.select_order_btn)
        '''
        self.process_orders_btn = QPushButton("逐筆下單")
        self.process_orders_btn.clicked.connect(self.start_order_processing)
        layout.addWidget(self.process_orders_btn)
        '''
        # 完全關閉 Playwright 按鈕
        self.quit_button = QPushButton("完全關閉 Playwright")
        self.quit_button.clicked.connect(self.close_playwright)
        layout.addWidget(self.quit_button)

        self.sales_info_label = QLabel("銷售總合：尚無資料", self)
        self.sales_info_label.setAlignment(Qt.AlignLeft)
        layout.addWidget(self.sales_info_label)

    def close_playwright(self):
        """完全關閉 Playwright 並釋放所有資源"""
        self.log("🔴 正在完全關閉 Playwright...")

        try:
            # 關閉瀏覽器
            if self.browser:
                self.browser.close()
                self.browser = None
                self.log("✅ 瀏覽器已關閉")

            # 停止 Playwright
            if self.playwright:
                self.playwright.stop()
                self.playwright = None
                self.log("✅ Playwright 進程已完全停止")

        except Exception as e:
            self.log(f"❌ 退出 Playwright 時發生錯誤：{traceback.format_exc()}")

        QMessageBox.information(self, "Playwright 已關閉", "Playwright 已完全關閉，您可以重新啟動它。")


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
        print(message)
        self.log_text.append(message)

    # -------------------------------
    # 出貨處理流程
    # -------------------------------
    def process_shipping(self, file_path):
        self.log(f"✅ 開始處理出貨流程，檔案路徑: {file_path}")
        try:
            try:
                df_orders = pd.read_excel(file_path, sheet_name="合併後資料")
            except Exception as e:
                QMessageBox.critical(self, "錯誤",
                                     f"無法讀取 '合併後資料' 頁面，請確認檔案格式是否正確。\n錯誤訊息: {e}")
                return


            required_cols = ["Product Name", "Attribute", "Quantity","Product URL"]
            if df_orders.empty or not all(col in df_orders.columns for col in required_cols):
                QMessageBox.critical(self, "錯誤",
                                     "合併後資料缺少必要欄位 (Product Name, Attribute, Quantity, Product URL)，請確認訂單檔案格式。")
                return

            try:
                df_origin = pd.read_excel(file_path, sheet_name="原始資料")
            except Exception as e:
                QMessageBox.critical(self, "錯誤",
                                     f"無法讀取 '原始資料' 頁面，請確認檔案格式是否正確。\n錯誤訊息: {e}")
                return
            order_code_list = df_origin["Order Code"].tolist()
            print(order_code_list)

            self.show_order_confirmation_dialog(df_orders, order_code_list)

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"出貨流程發生錯誤: {traceback.format_exc()}")
    '''
    def show_order_confirmation_dialog(self, df_orders, order_code_list):
        first_order_code = order_code_list[0] if order_code_list else None
        last_order_code = order_code_list[-1] if order_code_list else None
        length_of_order_code_list = len(order_code_list)
        # 正確的 QMessageBox 語法
        reply = QMessageBox.question(
            self,
            "確認出貨",  # 訊息框標題
            f" {first_order_code} 到 {last_order_code} 共 {length_of_order_code_list} 筆訂單!!\n\n是否開始出貨？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.start_shipping_process(df_orders,[first_order_code,last_order_code,length_of_order_code_list])  # 直接傳遞 df_orders
    '''
    from PyQt5.QtWidgets import QMessageBox, QPushButton, QApplication
    from PyQt5.QtGui import QClipboard
    from PyQt5.QtWidgets import QMessageBox, QApplication

    def show_order_confirmation_dialog(self, df_orders, order_code_list):
        first_order_code = order_code_list[0] if order_code_list else "N/A"
        last_order_code = order_code_list[-1] if order_code_list else "N/A"
        length_of_order_code_list = len(order_code_list)
        user = self.user_combo.currentText()

        message = f"{user}\n訂單從 {first_order_code} 到 {last_order_code} 共 {length_of_order_code_list} 筆"

        # 顯示確認對話框
        reply = QMessageBox.question(
            self,
            "確認出貨",  # 對話框標題
            f"{message}，是否開始出貨所有訂單？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            # 複製訊息到剪貼簿
            clipboard = QApplication.clipboard()
            clipboard.setText(message)

            # 開始出貨流程
            self.start_shipping_process(df_orders, message)

    def start_shipping_process(self, df_orders, message):
        dialog = DialogWindow()
        QMessageBox.information(self, "開始出貨", "即將進入逐筆出貨流程，請稍候...")
        try:
            self.log("正在啟動瀏覽器並導航到登錄頁面...")
            self.playwright = sync_playwright().start()
            self.browser = self.playwright.chromium.launch(channel="msedge", headless=False)
            context = self.browser.new_context()
            self.page = context.new_page()
            self.page.goto("https://baibaoshop.com/")
            self.page.mouse.move(random.randint(0, 1000), random.randint(0, 1000))
            time.sleep(random.uniform(1, 3))
            self.page.click(
                "body > div.wd-page-wrapper.website-wrapper > header > div > div.whb-row.whb-general-header.whb-not-sticky-row.whb-with-bg.whb-border-fullwidth.whb-color-light.whb-flex-equal-sides > div > div > div.whb-column.whb-col-left.whb-visible-lg > div.site-logo > a > img")
            time.sleep(random.uniform(1, 3))
            self.log("請在新開啟的瀏覽器中手動登入。")
            print("等待使用者點擊 PyQt 對話框...")

            # 顯示 PyQt5 對話框
            dialog.show_dialog()

            print("使用者已確認，繼續執行 Playwright")
        except Exception as e:
            self.log(f"啟動瀏覽器時出錯：{e}")
            return
        sub_total = 0
        for idx, row in df_orders.iterrows():
            product_name = row["Product Name"]
            attribute = row["Attribute"]
            quantity = row["Quantity"]
            link_url = row["Product URL"]
            sub_total += quantity
            total_quantity = df_orders["Quantity"].sum()
            try:
                self.log(f"正在打開訂單 URL: {link_url}")
                self.page.goto(link_url)
                time.sleep(random.uniform(1, 3))
                self.log(f"正在出貨: {idx + 1}. {product_name} - {attribute} - 數量: {quantity}")
                '''
                QMessageBox.information(self, "出貨中",
                                      f"正在出貨\n\n產品: {product_name}\n規格: {attribute}\n數量: {quantity}")
                '''
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("出貨中")
                msg_box.setText(f"第 {idx + 1}筆.共{len(df_orders)}筆 總計{total_quantity}件中第{sub_total}件\n\n產品: {product_name}\n規格: {attribute}\n數量: {quantity}")
                msg_box.addButton("下一筆", QMessageBox.AcceptRole)
                msg_box.exec_()
            except Exception as e:
                self.log(f"打開訂單 URL 時出錯：{e}")

        # QMessageBox.information(self, "出貨完成", "所有訂單已成功完成出貨！")
        # 建立訊息框
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("確認出貨")
        msg_box.setText(message)

        # 添加按鈕
        yes_button = msg_box.addButton("是", QMessageBox.YesRole)
        # no_button = msg_box.addButton("否", QMessageBox.NoRole)
        copy_button = QPushButton("Copy")

        # 設定 Copy 按鈕點擊事件
        def copy_to_clipboard():
            clipboard = QApplication.clipboard()
            clipboard.setText(message)

        copy_button.clicked.connect(copy_to_clipboard)

        # 將 Copy 按鈕加入訊息框
        msg_box.addButton(copy_button, QMessageBox.ActionRole)

        # 顯示對話框並獲取結果
        msg_box.exec_()
        self.log("所有訂單已成功完成出貨！")
        """
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()
        """

    def select_and_ship_order(self):
        if not self.current_user_dir:
            self.log("請先選擇使用者。")
            return
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(
            self, "選擇訂單檔案", self.current_user_dir, "Excel Files (goshop_orders_*.xlsx)", options=options
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

    # -------------------------------
    # 訂單資料抓取功能（含 split_and_merge_orders 處理）
    # -------------------------------
    def scrape_data(self):
        if not self.current_user_dir:
            self.log("請先選擇使用者。")
            return

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
                        self.log("正在翻到下一頁...")
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
                df_pending = pd.DataFrame(pending_orders, columns=columns)
                # 呼叫 split_and_merge_orders
                print("split_and_merge_orders", df_pending)
                user = self.user_combo.currentText()
                split_df, merged_df = self.split_and_merge_orders(df_pending)
                file_path = os.path.join(self.current_user_dir,
                                         f"goshop_orders_{datetime.now().strftime('%Y%m%d')}_{user}.xlsx")
                with pd.ExcelWriter(file_path) as writer:
                    df_pending.to_excel(writer, sheet_name="原始資料", index=False)
                    split_df.to_excel(writer, sheet_name="拆分後資料", index=False)
                    merged_df.to_excel(writer, sheet_name="合併後資料", index=False)
                self.log(f"訂單資料已存成 Excel 檔案：{file_path}")
                if not df_pending.empty:
                    first_order_code = str(df_pending["Order Code"].iloc[0]).strip()
                    with open(lastorder_file, "w", encoding="utf-8") as f:
                        f.write(first_order_code)
                    self.log(f"已建立 {lastorder_file}，內容為第一筆訂單的 Order Code：{first_order_code}")
                self.update_sales_file(df_pending)
            else:
                df_pending = pd.DataFrame(pending_orders, columns=columns)
                df_rest = pd.DataFrame(rest_orders, columns=columns)
                split_df, merged_df = self.split_and_merge_orders(df_pending)
                user = self.user_combo.currentText()
                file_path_pending = os.path.join(self.current_user_dir,
                                                 f"goshop_orders_{datetime.now().strftime('%Y%m%d')}_{user}.xlsx")
                file_path_rest = os.path.join(self.current_user_dir, "rest-order_{user}.xlsx")
                with pd.ExcelWriter(file_path_pending) as writer:
                    df_pending.to_excel(writer, sheet_name="原始資料", index=False)
                    split_df.to_excel(writer, sheet_name="拆分後資料", index=False)
                    merged_df.to_excel(writer, sheet_name="合併後資料", index=False)
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
        '''    
        finally:
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        '''
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
            user = self.user_combo.currentText()
            file_path = os.path.join(self.current_user_dir, f"goshop_orders_{start_order}_to_{end_order}_{user}.xlsx")
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

            sales_file = os.path.join(self.current_user_dir, "sales.xlsx")
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
            self.log(
                f"銷售總合 (Pending) -> Amount: {total_amount_pending:.2f}, Service charge: {total_service_charge_pending:.2f}, Final price: {total_final_price_pending:.2f}")
            self.log(
                f"銷售總合 (Rest) -> Amount: {total_amount_rest:.2f}, Service charge: {total_service_charge_rest:.2f}, Final price: {total_final_price_rest:.2f}")
        except Exception as e:
            self.log(f"更新銷售檔案時出錯：{traceback.format_exc()}")

    def split_and_merge_orders(self, df):
        self.log("開始執行 split_and_merge_orders()")
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
                    quantity_str = parts[2].strip()
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
        # 新增 "Product URL" 欄位：從 products_list.xlsx 中比對 Name 欄位
        try:
            products_file = os.path.join(self.current_user_dir, "products_list.xlsx")
            if os.path.exists(products_file):
                df_products = pd.read_excel(products_file)
                if "Name" in df_products.columns and "url" in df_products.columns:
                    def get_product_url(product_name):
                        # 比對時使用 strip() 來避免前後空白影響比對
                        match = df_products[df_products["Name"].str.strip().eq(product_name.strip())]
                        if not match.empty:
                            return match.iloc[0]["url"]
                        else:
                            return ""
                    merged_df["Product URL"] = merged_df["Product Name"].apply(get_product_url)
                else:
                    self.log("產品目錄中缺少必要欄位：Name 或 url")
                    merged_df["Product URL"] = ""
            else:
                self.log("未找到產品目錄 products_list.xlsx")
                merged_df["Product URL"] = ""
        except Exception as e:
            self.log(f"加入 Product URL 時出錯：{traceback.format_exc()}")
            merged_df["Product URL"] = ""
        return split_df, merged_df
    """
    def split_and_merge_orders(self, df):
        self.log("開始執行 split_and_merge_orders()")
        print("split_and_merge_orders程式開始", df)
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
                    quantity_str = parts[2].strip()
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
    """
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
            self.update_orders_url()
        except Exception as e:
            self.log(f"更新產品資料時出錯：{traceback.format_exc()}")
            reply = QMessageBox.question(self, "錯誤", "更新產品資料時出錯，是否重新登入？",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.open_browser()
            else:
                self.log("取消重新登入。")

    # -------------------------------
    # 更新訂單檔中「合併後資料」的 Product URL 功能
    # -------------------------------
    def update_product_url(self):
        """
        此函數比對訂單檔（合併後資料工作表）中之 Product Name 與產品目錄（products_list.xlsx）中之 Name，
        若相符則於該筆資料新增一個「Product URL」欄，填入產品目錄中對應的 url 資料，
        更新後存回原訂單檔（以覆蓋「合併後資料」工作表）。
        """
        if not self.current_user_dir:
            self.log("請先選擇使用者。")
            return

        # 讓使用者選取要更新的訂單檔案（需包含「合併後資料」工作表）
        order_file, _ = QFileDialog.getOpenFileName(
            self, "選擇訂單檔案（含合併後資料）", self.current_user_dir, "Excel Files (*.xlsx)"
        )
        if not order_file:
            self.log("未選擇訂單檔案。")
            return

        try:
            df_order = pd.read_excel(order_file, sheet_name="合併後資料")
            self.log("已讀取訂單檔案中『合併後資料』工作表。")
        except Exception as e:
            self.log(f"讀取『合併後資料』工作表時出錯：{traceback.format_exc()}")
            return

        products_file = os.path.join(self.current_user_dir, "products_list.xlsx")
        if not os.path.exists(products_file):
            self.log("產品目錄不存在，請先更新產品資料。")
            return

        try:
            df_products = pd.read_excel(products_file)
            self.log("已讀取產品目錄資料。")
        except Exception as e:
            self.log(f"讀取產品目錄資料時出錯：{traceback.format_exc()}")
            return

        # 建立產品名稱與對應 url 的字典（假設產品目錄中欄位名稱分別為 Name 與 url）
        product_url_map = dict(zip(df_products["Name"], df_products["url"]))

        if "Product Name" not in df_order.columns:
            self.log("訂單檔『合併後資料』中無 Product Name 欄位。")
            return

        # 新增欄位「Product URL」，以對應產品目錄中相同產品名稱之 url，若無則空白
        df_order["Product URL"] = df_order["Product Name"].apply(lambda name: product_url_map.get(name, ""))

        try:
            # 利用 openpyxl 模組覆蓋更新「合併後資料」工作表
            with pd.ExcelWriter(order_file, engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as writer:
                df_order.to_excel(writer, sheet_name="合併後資料", index=False)
            self.log("更新訂單檔『合併後資料』中 Product URL 成功！")
            QMessageBox.information(self, "更新完成", "已成功更新訂單檔中『合併後資料』的 Product URL。")
        except Exception as e:
            self.log(f"寫入更新後資料時出錯：{traceback.format_exc()}")

    def start_order_processing(self):
        if self.df_orders is None or self.df_orders.empty:
            QMessageBox.information(self, "提示", "請先取得訂單資料。", QMessageBox.Ok)
            return
        dialog = OrderProcessingDialog(self.df_orders, self)
        dialog.exec_()

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
        '''
        finally:
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        '''
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
        if self.df_orders is None or self.df_orders.empty:
            QMessageBox.information(self, "提示", "請先取得訂單資料。", QMessageBox.Ok)
            return
        dialog = OrderProcessingDialog(self.df_orders, self)
        dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)  # 確保 QApplication 正確初始化
    window = OrderScraperApp()
    window.show()


    # 設定所有視窗置頂
    def set_all_windows_on_top(app):
        """讓所有 Qt UI 視窗保持在最上層"""
        for widget in app.topLevelWidgets():
            widget.setWindowFlags(widget.windowFlags() | Qt.WindowStaysOnTopHint)
            widget.show()  # 重新顯示以應用新設定


    set_all_windows_on_top(app)  # 呼叫函式，讓所有 UI 都保持最前

    sys.exit(app.exec_())  # 進入應用程式主迴圈