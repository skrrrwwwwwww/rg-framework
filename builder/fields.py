from PyQt6.QtWidgets import QTextEdit


class TeacherField(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPlaceholderText("доц. Федулов Я.С,\nасс. Маракулина Ю.Д")
        self.setMaximumHeight(80)
        self.setTabChangesFocus(True)

    def mousePressEvent(self, event):
        self.selectAll()
        super().mousePressEvent(event)