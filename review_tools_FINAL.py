import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTabWidget, QTableWidget, QTableWidgetItem
from PyQt5.QtGui import QPainter, QColor, QBrush, QFont, QPen
from PyQt5.QtCore import Qt, QRect


class StepProgressBar(QWidget):
    def __init__(self, steps=5):
        super().__init__()
        self.steps = steps
        self.current_step = 0
        self.setMinimumHeight(40)
        self.setMinimumWidth(300)
        self.setMouseTracking(True)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        width = self.width()
        height = self.height()
        step_width = width // self.steps
        radius = 8  # Radius for rounded corners

        for i in range(self.steps):
            if i < self.current_step:
                brush_color = QColor(100, 149, 237)  # Cornflower Blue for completed steps
            else:
                brush_color = QColor(211, 211, 211)  # Light Grey for remaining steps

            pen = QPen(Qt.NoPen)
            painter.setPen(pen)
            painter.setBrush(QBrush(brush_color))

            # Draw rounded rectangle for steps
            painter.drawRoundedRect(QRect(i * step_width, 0, step_width - 5, height), radius, radius)

            # Draw text
            painter.setPen(QColor(255, 255, 255))  # White for text
            font = QFont("Arial", 12, QFont.Bold)
            painter.setFont(font)
            painter.drawText(QRect(i * step_width, 0, step_width - 5, height), Qt.AlignCenter, f"{i + 1}")

    def mousePressEvent(self, event):
        width = self.width()
        step_width = width // self.steps
        clicked_step = event.x() // step_width
        if clicked_step < self.steps:
            self.current_step = clicked_step + 1
            self.update()
            self.parent().tab_switched(clicked_step)

    def advance_step(self):
        if self.current_step < self.steps:
            self.current_step += 1
            self.update()
            self.parent().tab_switched(self.current_step - 1)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('美化的步骤条和Tab页示例')
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        self.step_bar = StepProgressBar(steps=5)
        self.step_bar.setParent(self)
        layout.addWidget(self.step_bar)

        self.tabs = QTabWidget()
        # self.tabs.setStyleSheet("""
        #     QTabWidget::pane {
        #         border: 1px solid #cccccc;
        #         border-radius: 5px;
        #         padding: 10px;
        #     }
        #     QTabBar::tab {
        #         background: #e0e0e0;
        #         border: 1px solid #cccccc;
        #         border-radius: 5px;
        #         padding: 10px;
        #         min-width: 80px;
        #     }
        #     QTabBar::tab:selected {
        #         background: #ffffff;
        #         border-color: #3399ff;
        #     }
        # """)

        for i in range(5):
            tab = QWidget()
            tab_layout = QVBoxLayout()
            table = QTableWidget(5, 3)
            table.setHorizontalHeaderLabels(['Column 1', 'Column 2', 'Column 3'])
            for row in range(5):
                for col in range(3):
                    table.setItem(row, col, QTableWidgetItem(f"Step {i + 1} - Cell ({row + 1},{col + 1})"))
            tab_layout.addWidget(table)
            tab.setLayout(tab_layout)
            self.tabs.addTab(tab, f"Step {i + 1}")

        layout.addWidget(self.tabs)

        self.button = QPushButton('下一步')
        # self.button.setStyleSheet("""
        #     QPushButton {
        #         background-color: #3399ff;
        #         border: none;
        #         color: white;
        #         padding: 10px 20px;
        #         text-align: center;
        #         font-size: 16px;
        #         border-radius: 5px;
        #         margin: 10px 0;
        #     }
        #     QPushButton:hover {
        #         background-color: #0073e6;
        #     }
        # """)
        self.button.clicked.connect(self.on_button_clicked)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def on_button_clicked(self):
        self.step_bar.advance_step()

    def tab_switched(self, index):
        self.tabs.setCurrentIndex(index)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Windows')  # Windows , windowsvista , Fusion
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
