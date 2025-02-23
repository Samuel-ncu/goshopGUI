from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit
from PyQt5.QtGui import QClipboard
import sys
class ClipboardExample(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.input_field = QLineEdit(self)
        layout.addWidget(self.input_field)

        self.copy_button = QPushButton("Copy to Clipboard", self)
        self.copy_button.clicked.connect(self.copy_to_clipboard)
        layout.addWidget(self.copy_button)

        self.setLayout(layout)
        self.setWindowTitle("Clipboard Example")
        self.show()

    def copy_to_clipboard(self):
        clipboard = QApplication.clipboard()
        clipboard.setText(self.input_field.text())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ClipboardExample()
    sys.exit(app.exec_())