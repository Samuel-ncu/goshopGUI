from playwright.sync_api import sync_playwright
import pandas as pd

def scrape_orders_to_excel():
    with sync_playwright() as p:
        # 開啟 Microsoft Edge 瀏覽器
        browser = p.chromium.launch(channel="msedge", headless=False)
        context = browser.new_context()
        page = context.new_page()

        # 導向到登入頁面
        page.goto("https://goshophsn.com/")

        # 等待使用者手動完成登入
        print("請在瀏覽器中手動完成登入，完成後按下 'Enter' 繼續...")
        input("等待使用者登入後按下 'Enter'：")

        # 模擬點擊 "Orders" 連結
        orders_link = page.locator("text='Orders'")
        orders_link.click()
        page.wait_for_load_state('networkidle')

        # 儲存所有表格資料
        all_data = []

        # 逐頁抓取資料
        while True:
            try:
                # 等待表格加載
                page.wait_for_selector("table tbody tr", timeout=10000)
                print("正在抓取當前頁面資料...")

                # 抓取當前頁面表格數據
                table_rows = page.locator("table tbody tr")
                for i in range(table_rows.count()):
                    row_data = table_rows.nth(i).locator("td").all_inner_texts()
                    all_data.append(row_data)

                # 檢查是否有下一頁按鈕
                next_button = page.locator('a[aria-label="Next »"]')
                if next_button.is_visible():
                    next_button.click()
                    page.wait_for_load_state('networkidle')
                else:
                    print("所有分頁資料已抓取完成。")
                    break

            except Exception as e:
                print(f"抓取過程中出錯: {e}")
                break

        # 將資料存成 Excel
        if all_data:
            columns = ["#", "Order Code", "Num. of Products", "Customer", "Amount", "Service charge", "Final price", "Delivery Status", "Payment Status", "Product Info", "Options"]
            df = pd.DataFrame(all_data, columns=columns)
            file_path = "goshop_orders.xlsx"
            df.to_excel(file_path, index=False)
            print(f"所有訂單資料已存成 Excel 檔案：{file_path}")
        else:
            print("未抓取到任何資料。")
        input()
        browser.close()

# 執行程式
scrape_orders_to_excel()
