from PyQt6.QtWidgets import *
from datetime import datetime


class MainTab(QWidget):
    def __init__(self):
        super().__init__()
        self.layout = QFormLayout(self)
        self.layout.setSpacing(15)

        # Дисциплина
        self.subject = QLineEdit()
        self.subject.setPlaceholderText("Технология объектного программирования")
        self.subject.setToolTip("Ctrl+↑/↓ - навигация, Tab - автозаполнение")
        self.subject.mousePressEvent = lambda e: self.subject.selectAll()
        self.layout.addRow("Дисциплина:", self.subject)

        # Тип работы
        self.work_type = QComboBox()
        self.work_type.addItems([
            "Лабораторная работа",
            "Курсовая работа",
            "Расчетно-графическая работа",
            "Отчет по практике",
            "Контрольная работа",
            "Реферат"
        ])
        self.work_type.setCurrentIndex(0)
        self.layout.addRow("Тип работы:", self.work_type)

        # Номер работы
        work_number_layout = QHBoxLayout()
        self.work_number = QSpinBox()
        self.work_number.setRange(1, 99)
        self.work_number.setValue(1)
        self.work_number.setPrefix("№")

        self.no_work_number = QCheckBox("Без номера")
        self.no_work_number.toggled.connect(self.toggle_work_number)

        work_number_layout.addWidget(self.work_number)
        work_number_layout.addWidget(self.no_work_number)
        self.layout.addRow("Номер работы:", work_number_layout)

        # Диапазон лабораторных
        lab_range_group = QGroupBox("Диапазон лабораторных работ")
        lab_range_layout = QHBoxLayout()

        self.lab_from = QSpinBox()
        self.lab_from.setRange(1, 99)
        self.lab_from.setValue(1)
        self.lab_from.setPrefix("с ")

        self.lab_to = QSpinBox()
        self.lab_to.setRange(1, 99)
        self.lab_to.setValue(8)
        self.lab_to.setPrefix("по ")

        self.create_all_btn = QPushButton("📦 Создать все")
        self.create_all_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 8px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)
        # Сигнал будет подключаться позже, из главного окна

        lab_range_layout.addWidget(self.lab_from)
        lab_range_layout.addWidget(self.lab_to)
        lab_range_layout.addWidget(self.create_all_btn)
        lab_range_group.setLayout(lab_range_layout)
        self.layout.addRow(lab_range_group)

        # Студент
        self.student = QLineEdit()
        self.student.setPlaceholderText("Шикарев Иван А")
        self.student.mousePressEvent = lambda e: self.student.selectAll()
        self.layout.addRow("Студент:", self.student)

        # Группа
        self.group = QLineEdit()
        self.group.setPlaceholderText("ПО1-23")
        self.group.mousePressEvent = lambda e: self.group.selectAll()
        self.layout.addRow("Группа:", self.group)

        # Преподаватель
        self.teacher = QTextEdit()
        self.teacher.setPlaceholderText("доц. Федулов Я.С,\nасс. Маракулина Ю.Д")
        self.teacher.setMaximumHeight(80)
        self.teacher.setTabChangesFocus(True)
        self.teacher.mousePressEvent = lambda e: self.teacher.selectAll()
        self.layout.addRow("Преподаватель:", self.teacher)

        # Вариант
        variant_layout = QHBoxLayout()
        self.variant = QSpinBox()
        self.variant.setRange(1, 99)
        self.variant.setValue(17)
        self.no_variant = QCheckBox("Без варианта")
        variant_layout.addWidget(self.variant)
        variant_layout.addWidget(self.no_variant)
        self.layout.addRow("Вариант:", variant_layout)

        # Год
        self.year = QSpinBox()
        self.year.setRange(2000, 2100)
        self.year.setValue(datetime.now().year)
        self.layout.addRow("Год:", self.year)

        # Оглавление
        self.include_toc = QCheckBox("Добавить оглавление")
        self.include_toc.setChecked(True)
        self.layout.addRow("", self.include_toc)

    def toggle_work_number(self, checked):
        self.work_number.setEnabled(not checked)

    def get_data(self):
        """Возвращает словарь со всеми полями"""
        return {
            "subject": self.subject.text().strip(),
            "work_type": self.work_type.currentText().strip(),
            "work_number": self.work_number.value(),
            "no_work_number": self.no_work_number.isChecked(),
            "lab_from": self.lab_from.value(),
            "lab_to": self.lab_to.value(),
            "student": self.student.text().strip(),
            "group": self.group.text().strip(),
            "teacher": self.teacher.toPlainText().strip(),
            "variant": None if self.no_variant.isChecked() else self.variant.value(),
            "year": self.year.value(),
            "include_toc": self.include_toc.isChecked()
        }

    def set_defaults(self):
        """Заполняет поля значениями по умолчанию"""
        self.subject.setText("Технология объектного программирования")
        self.work_type.setCurrentIndex(0)
        self.work_number.setValue(1)
        self.student.setText("Шикарев Иван А")
        self.group.setText("ПО1-23")
        self.teacher.setPlainText("доц. Федулов Я.С,\nасс. Маракулина Ю.Д")
        self.variant.setValue(17)
        self.no_variant.setChecked(False)
        self.include_toc.setChecked(True)

    def set_focus(self):
        self.subject.setFocus()