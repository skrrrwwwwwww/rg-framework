from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import *


class SectionEditorDialog(QDialog):
    def __init__(self, model, section_name, parent_name, full_name, parent=None):
        super().__init__(parent)
        self.model = model
        self.section_name = section_name
        self.parent_name = parent_name

        self.setWindowTitle(f"📝 Редактирование: {full_name}")
        self.setModal(True)
        self.setMinimumSize(600, 700)

        layout = QVBoxLayout(self)

        # Заголовок
        label = QLabel(f"<h3>«{full_name}»</h3>")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)

        # Текстовое поле
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Введите текст раздела...")
        self.text_edit.setPlainText(self.model.get_section_text(section_name, parent_name))
        layout.addWidget(self.text_edit)

        # Панель инструментов
        toolbar = self.create_toolbar()
        layout.addWidget(toolbar)

        # Кнопки сохранения
        btn_layout = QHBoxLayout()

        self.save_btn = QPushButton("💾 Сохранить")
        self.save_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        self.save_btn.clicked.connect(self.save)

        self.save_close_btn = QPushButton("💾 Сохранить и закрыть")
        self.save_close_btn.setStyleSheet("background-color: #2196F3; color: white;")
        self.save_close_btn.clicked.connect(self.save_and_close)

        self.close_btn = QPushButton("❌ Закрыть")
        self.close_btn.clicked.connect(self.reject)

        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.save_close_btn)
        btn_layout.addWidget(self.close_btn)

        layout.addLayout(btn_layout)

    def create_toolbar(self):
        toolbar = QGroupBox("Вставка элементов")
        layout = QHBoxLayout(toolbar)

        btn_text = QPushButton("📝 Текст")
        btn_text.clicked.connect(self.insert_text)

        btn_table = QPushButton("📊 Таблица")
        btn_table.clicked.connect(self.insert_table)

        btn_formula = QPushButton("🧮 Формула")
        btn_formula.clicked.connect(self.insert_formula)

        btn_image = QPushButton("🖼️ Скриншот")
        btn_image.clicked.connect(self.insert_image)

        btn_list = QPushButton("📋 Список")
        btn_list.clicked.connect(self.insert_list)

        layout.addWidget(btn_text)
        layout.addWidget(btn_table)
        layout.addWidget(btn_formula)
        layout.addWidget(btn_image)
        layout.addWidget(btn_list)

        return toolbar

    def insert_text(self):
        text, ok = QInputDialog.getMultiLineText(self, "Введите текст", "Текст:")
        if ok and text:
            self.text_edit.insertPlainText(f"\n{text}\n")

    def insert_table(self):
        cols, ok1 = QInputDialog.getInt(self, "Таблица", "Колонок:", 3, 1, 10)
        if not ok1:
            return
        rows, ok2 = QInputDialog.getInt(self, "Таблица", "Строк:", 3, 1, 50)
        if not ok2:
            return

        table_text = "\n"
        for i in range(rows):
            row = "| " + " | ".join([f"[{i+1},{j+1}]" for j in range(cols)]) + " |\n"
            table_text += row
        table_text += "\n"
        self.text_edit.insertPlainText(table_text)

    def insert_formula(self):
        formula, ok1 = QInputDialog.getText(self, "Формула", "Формула:", text="y = f(x)")
        if not ok1:
            return
        explanation, ok2 = QInputDialog.getText(self, "Пояснение", "Пояснение (Enter - пропустить):")
        text = f"\n[ФОРМУЛА] {formula}\n"
        if explanation:
            text += f"[ПОЯСНЕНИЕ] {explanation}\n"
        self.text_edit.insertPlainText(text)

    def insert_image(self):
        caption, ok1 = QInputDialog.getText(self, "Скриншот", "Подпись:", text="Результаты")
        if not ok1:
            return
        height, ok2 = QInputDialog.getInt(self, "Скриншот", "Высота (строк):", 5, 1, 20)
        self.text_edit.insertPlainText(f"\n[СКРИНШОТ: {caption}] (высота: {height} строк)\n")

    def insert_list(self):
        items = []
        while True:
            item, ok = QInputDialog.getText(self, "Элемент списка",
                                            "Введите элемент (Enter - пустая строка для завершения):")
            if not ok or not item.strip():
                break
            items.append(item.strip())
        if items:
            list_text = "\n"
            for i, item in enumerate(items, 1):
                list_text += f"{i}. {item}\n"
            list_text += "\n"
            self.text_edit.insertPlainText(list_text)

    def save(self):
        content = self.text_edit.toPlainText()
        self.model.set_section_text(self.section_name, content, self.parent_name)
        QMessageBox.information(self, "Сохранено", "Раздел сохранен")

    def save_and_close(self):
        self.save()
        self.accept()