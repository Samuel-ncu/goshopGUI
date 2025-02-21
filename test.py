import sys
from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QTextEdit, QPushButton, QDialogButtonBox

class ScrollableMessageBox(QDialog):
    def __init__(self, message, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Scrollable Message Box")

        layout = QVBoxLayout(self)

        # 創建一個可滾動的文本編輯器
        self.textEdit = QTextEdit(self)
        self.textEdit.setReadOnly(True)
        self.textEdit.setPlainText(message)
        layout.addWidget(self.textEdit)

        # 創建按鈕框
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok, self)
        self.buttons.accepted.connect(self.accept)
        layout.addWidget(self.buttons)

        self.setLayout(layout)

if __name__ == "__main__":
    app = QApplication(sys.argv)

    message = "這是一個很長的消息，需要滾動才能看到全部內容。\n" * 20
    dialog = ScrollableMessageBox(message)
    dialog.exec_()

    sys.exit(app.exec_())
