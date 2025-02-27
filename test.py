from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QPushButton, QLabel,
                             QScrollArea, QWidget, QApplication, QHBoxLayout)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QClipboard


class ShippingDialog(QDialog):
    def __init__(self, message, order_details, parent=None):
        super().__init__(parent)
        self.setWindowTitle("確認出貨")
        self.setMinimumSize(600, 400)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)

        # 主佈局
        layout = QVBoxLayout(self)

        # 提示訊息
        lbl_message = QLabel("確認要出貨嗎？")
        layout.addWidget(lbl_message)

        # 滾動區域（保持原有代碼）
        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        order_details_label = QLabel(f"訂單詳細資料:\n{order_details}")
        order_details_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        scroll_layout.addWidget(order_details_label)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)

        # 按鈕區域
        btn_layout = QHBoxLayout()

        # Copy 按鈕
        self.copy_btn = QPushButton("Copy")
        self.copy_btn.clicked.connect(self.on_copy)
        btn_layout.addWidget(self.copy_btn)

        # OK 按鈕
        self.ok_btn = QPushButton("OK")
        self.ok_btn.clicked.connect(self.accept)  # 觸發 accept() 關閉對話框
        btn_layout.addWidget(self.ok_btn)

        layout.addLayout(btn_layout)

    def on_copy(self):
        """複製內容到剪貼簿且不關閉對話框"""
        clipboard = QApplication.clipboard()
        clipboard.setText("需要複製的內容")  # 替換為實際內容
        # 可選：添加複製成功提示
        self.copy_btn.setText("已複製！")
        QApplication.processEvents()  # 立即更新按鈕文字


# 使用方式示例
if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)

    # 模擬數據
    message = "出貨確認訊息"
    order_details = "訂單編號: 12345\n商品: 範例商品\n數量: 2"

    dialog = ShippingDialog(message, order_details)
    if dialog.exec_() == QDialog.Accepted:
        print("用戶點擊了OK")
    else:
        print("對話框關閉")

    sys.exit(app.exec_())