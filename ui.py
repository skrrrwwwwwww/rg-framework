import sys
from datetime import datetime
from pathlib import Path

from PyQt6.QtCore import *
from PyQt6.QtWidgets import *

from blocks import *
from builder import DocBuilder


class ReportGeneratorUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Генератор отчетов v2.0")
        self.setGeometry(100, 100, 800, 600)

        self.sections_content = {}  # {название_раздела: текст}

        # Центральный виджет
        central = QWidget()
        self.setCentralWidget(central)

        # Основной layout
        layout = QVBoxLayout(central)

        # Заголовок
        title = QLabel("🔧 КОНСТРУКТОР ОТЧЕТОВ")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; margin: 20px;")
        layout.addWidget(title)

        # Табы для разных разделов
        tabs = QTabWidget()
        layout.addWidget(tabs)

        # Вкладка "Основное"
        main_tab = self.create_main_tab()
        tabs.addTab(main_tab, "📋 Основное")

        # Вкладка "Разделы"
        sections_tab = self.create_sections_tab()
        tabs.addTab(sections_tab, "📑 Разделы")

        # Кнопка генерации
        self.generate_btn = QPushButton("🚀 СГЕНЕРИРОВАТЬ ОТЧЕТ")
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 16px;
                padding: 15px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        self.generate_btn.clicked.connect(self.generate_report)
        layout.addWidget(self.generate_btn)

        # Статус бар
        self.statusBar().showMessage("Готов к работе")

        # Устанавливаем фильтры событий для автозаполнения
        self.setup_event_filters()

    def setup_event_filters(self):
        """Устанавливает фильтры событий для полей ввода"""
        fields = [
            self.subject, self.work_type, self.work_number,
            self.student, self.group, self.teacher
        ]
        for field in fields:
            field.installEventFilter(self)

    def create_main_tab(self):
        widget = QWidget()
        layout = QFormLayout(widget)
        layout.setSpacing(15)

        # Поля ввода
        self.subject = QLineEdit()
        self.subject.setPlaceholderText("Технология объектного программирования")
        self.subject.mousePressEvent = lambda e: self.subject.selectAll()
        layout.addRow("Дисциплина:", self.subject)

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
        layout.addRow("Тип работы:", self.work_type)

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
        layout.addRow("Номер работы:", work_number_layout)

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
        self.create_all_btn.clicked.connect(self.create_all_labs)

        lab_range_layout.addWidget(self.lab_from)
        lab_range_layout.addWidget(self.lab_to)
        lab_range_layout.addWidget(self.create_all_btn)
        lab_range_group.setLayout(lab_range_layout)
        layout.addRow(lab_range_group)

        self.student = QLineEdit()
        self.student.setPlaceholderText("Шикарев Иван А")
        self.student.mousePressEvent = lambda e: self.student.selectAll()
        layout.addRow("Студент:", self.student)

        self.group = QLineEdit()
        self.group.setPlaceholderText("ПО1-23")
        self.group.mousePressEvent = lambda e: self.group.selectAll()
        layout.addRow("Группа:", self.group)

        self.teacher = QTextEdit()
        self.teacher.setPlaceholderText("доц. Федулов Я.С,\nасс. Маракулина Ю.Д")
        self.teacher.setMaximumHeight(80)
        self.teacher.setTabChangesFocus(True)
        self.teacher.mousePressEvent = lambda e: self.teacher.selectAll()
        layout.addRow("Преподаватель:", self.teacher)

        # Вариант
        variant_layout = QHBoxLayout()
        self.variant = QSpinBox()
        self.variant.setRange(1, 99)
        self.variant.setValue(17)
        self.no_variant = QCheckBox("Без варианта")
        variant_layout.addWidget(self.variant)
        variant_layout.addWidget(self.no_variant)
        layout.addRow("Вариант:", variant_layout)

        # Год
        self.year = QSpinBox()
        self.year.setRange(2000, 2100)
        self.year.setValue(datetime.now().year)
        layout.addRow("Год:", self.year)

        # Оглавление
        self.include_toc = QCheckBox("Добавить оглавление")
        self.include_toc.setChecked(True)
        layout.addRow("", self.include_toc)

        return widget

    def eventFilter(self, obj, event):
        """Автозаполнение по Tab"""
        if event.type() == QEvent.Type.KeyPress and event.key() == Qt.Key.Key_Tab:
            # Если поле пустое - заполняем
            if obj == self.subject and not self.subject.text():
                self.subject.setText("Технология объектного программирования")
                self.subject.selectAll()
                return True
            elif obj == self.work_type and not self.work_type.currentText():
                self.work_type.setCurrentIndex(0)
                return True
            elif obj == self.work_number and self.work_number.value() == 1:
                pass
            elif obj == self.student and not self.student.text():
                self.student.setText("Шикарев Иван А")
                self.student.selectAll()
                return True
            elif obj == self.group and not self.group.text():
                self.group.setText("ПО1-23")
                self.group.selectAll()
                return True
            elif obj == self.teacher and not self.teacher.toPlainText():  # ← ИСПРАВЛЕНО!
                self.teacher.setPlainText("доц. Федулов А.С,\nасс. Маракулина Ю.Д")
                self.teacher.selectAll()
                return True

        return super().eventFilter(obj, event)

    def toggle_work_number(self, checked):
        """Включает/выключает поле номера работы"""
        self.work_number.setEnabled(not checked)

    def create_sections_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Список разделов
        self.sections_list = QTreeWidget()  # ← Меняем на QTreeWidget!
        self.sections_list.setMinimumHeight(300)
        self.sections_list.setHeaderLabel("Разделы отчета")
        self.sections_list.itemDoubleClicked.connect(self.open_section_window)
        layout.addWidget(self.sections_list)

        # Кнопки управления разделами
        btn_layout = QHBoxLayout()

        add_section = QPushButton("➕ Добавить раздел")
        add_section.clicked.connect(self.add_section)
        btn_layout.addWidget(add_section)

        add_subsection = QPushButton("↳ Добавить подраздел")
        add_subsection.clicked.connect(self.add_subsection)
        btn_layout.addWidget(add_subsection)

        remove_section = QPushButton("➖ Удалить")
        remove_section.clicked.connect(self.remove_section)
        btn_layout.addWidget(remove_section)

        edit_section = QPushButton("✏️ Редактировать")
        edit_section.clicked.connect(self.open_section_window)
        btn_layout.addWidget(edit_section)

        layout.addLayout(btn_layout)

        return widget

    def open_section_window(self):
        """Открывает окно для работы с разделом"""
        current = self.sections_list.currentItem()
        if not current:
            QMessageBox.warning(self, "Ошибка", "Выберите раздел!")
            return

        # Определяем тип элемента и его название
        item_type = current.data(0, Qt.ItemDataRole.UserRole)
        section_name = current.text(0).strip()

        if item_type == "subsection":
            section_name = section_name.replace("    ", "")
            parent_item = current.parent()
            parent_name = parent_item.text(0)
            full_name = f"{parent_name} → {section_name}"
        else:
            parent_name = None
            full_name = section_name

        # Создаем окно раздела
        self.section_window = QDialog(self)
        self.section_window.setWindowTitle(f"📝 Редактирование: {full_name}")
        self.section_window.setModal(True)
        self.section_window.setMinimumWidth(600)
        self.section_window.setMinimumHeight(700)

        layout = QVBoxLayout(self.section_window)

        # Заголовок
        label = QLabel(f"<h3>«{full_name}»</h3>")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)

        # Текст раздела
        self.section_text = QTextEdit()
        self.section_text.setPlaceholderText("Введите текст раздела...")

        # Загружаем сохраненное содержимое
        if parent_name:
            # Это подраздел
            for sub in self.sections_content.get(parent_name, {}).get("subsections", []):
                if sub["name"] == section_name:
                    self.section_text.setPlainText(sub.get("text", ""))
                    break
        else:
            # Это главный раздел
            if section_name in self.sections_content:
                self.section_text.setPlainText(self.sections_content[section_name].get("text", ""))

        layout.addWidget(self.section_text)

        # Панель инструментов (такая же как была)
        toolbar = self.create_content_toolbar()
        layout.addWidget(toolbar)

        # Кнопки сохранения
        btn_layout = QHBoxLayout()

        btn_save = QPushButton("💾 Сохранить")
        btn_save.setStyleSheet("background-color: #4CAF50; color: white;")
        btn_save.clicked.connect(lambda: self.save_section_content(
            section_name, parent_name))

        btn_save_close = QPushButton("💾 Сохранить и закрыть")
        btn_save_close.setStyleSheet("background-color: #2196F3; color: white;")
        btn_save_close.clicked.connect(lambda: self.save_and_close_section(
            section_name, parent_name))

        btn_close = QPushButton("❌ Закрыть")
        btn_close.clicked.connect(self.section_window.reject)

        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_save_close)
        btn_layout.addWidget(btn_close)

        layout.addLayout(btn_layout)

        self.section_window.exec()

    def save_section_content(self, section_name, parent_name=None):
        """Сохраняет содержимое раздела"""
        content = self.section_text.toPlainText()

        if parent_name:
            # Это подраздел
            if parent_name not in self.sections_content:
                self.sections_content[parent_name] = {"text": "", "subsections": []}

            for sub in self.sections_content[parent_name]["subsections"]:
                if sub["name"] == section_name:
                    sub["text"] = content
                    break
        else:
            # Это главный раздел
            if section_name not in self.sections_content:
                self.sections_content[section_name] = {"text": "", "subsections": []}
            self.sections_content[section_name]["text"] = content

        QMessageBox.information(self, "Сохранено", f"Раздел сохранен")

    def save_and_close_section(self, section_name, parent_name=None):
        """Сохраняет и закрывает окно"""
        self.save_section_content(section_name, parent_name)
        self.section_window.accept()

    def add_section(self):
        """Добавляет главный раздел"""
        section, ok = QInputDialog.getText(self, "Новый раздел", "Введите название раздела:")
        if ok and section:
            item = QTreeWidgetItem(self.sections_list)
            item.setText(0, section)
            item.setData(0, Qt.ItemDataRole.UserRole, "section")  # тип: раздел
            self.sections_content[section] = {"text": "", "subsections": []}

    def add_subsection(self):
        """Добавляет подраздел к выбранному разделу"""
        current = self.sections_list.currentItem()
        if not current:
            QMessageBox.warning(self, "Ошибка", "Выберите раздел для добавления подраздела!")
            return

        # Находим родительский раздел
        parent_item = current
        while parent_item.parent():
            parent_item = parent_item.parent()
        parent_section = parent_item.text(0)

        subsection, ok = QInputDialog.getText(self, "Новый подраздел", "Введите название подраздела:")
        if ok and subsection:
            # Создаем элемент подраздела
            child = QTreeWidgetItem(current)
            child.setText(0, f"    {subsection}")
            child.setData(0, Qt.ItemDataRole.UserRole, "subsection")

            # Добавляем в структуру данных
            if parent_section not in self.sections_content:
                self.sections_content[parent_section] = {"text": "", "subsections": []}
            self.sections_content[parent_section]["subsections"].append({
                "name": subsection,
                "text": "",
                "parent": parent_section
            })

    def remove_section(self):
        """Удаляет выбранный элемент"""
        current = self.sections_list.currentItem()
        if not current:
            return

        section_name = current.text(0).strip()
        parent = current.parent()

        if parent:
            # Это подраздел
            parent_section = parent.text(0)
            subsection_name = section_name.replace("    ", "")
            # Удаляем из структуры данных
            if parent_section in self.sections_content:
                self.sections_content[parent_section]["subsections"] = [
                    s for s in self.sections_content[parent_section]["subsections"]
                    if s["name"] != subsection_name
                ]
        else:
            # Это главный раздел
            if section_name in self.sections_content:
                del self.sections_content[section_name]

        # Удаляем из дерева
        self.sections_list.takeTopLevelItem(self.sections_list.indexOfTopLevelItem(current))

    def fill_section_content(self, doc, section_name):
        """Интерактивное наполнение раздела контентом"""

        # Создаем диалог
        dialog = QDialog(self)
        dialog.setWindowTitle(f"📝 Наполнение: {section_name}")
        dialog.setModal(True)
        dialog.setMinimumWidth(450)
        dialog.setMinimumHeight(500)

        layout = QVBoxLayout(dialog)
        layout.setSpacing(10)

        # Заголовок (используем обычный QLabel без HTML)
        label = QLabel(f"Добавление контента в раздел:\n«{section_name}»")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                background-color: #f0f0f0;
                border-radius: 5px;
            }
        """)
        layout.addWidget(label)

        # Кнопки контента
        btn_text = QPushButton("📝 Добавить текст")
        btn_text.setMinimumHeight(45)
        btn_text.clicked.connect(lambda: self.add_text(doc))
        layout.addWidget(btn_text)

        btn_table = QPushButton("📊 Добавить таблицу")
        btn_table.setMinimumHeight(45)
        btn_table.clicked.connect(lambda: self.add_table(doc))
        layout.addWidget(btn_table)

        btn_formula = QPushButton("🧮 Добавить формулу")
        btn_formula.setMinimumHeight(45)
        btn_formula.clicked.connect(lambda: self.add_formula(doc))
        layout.addWidget(btn_formula)

        btn_image = QPushButton("🖼️ Место для скриншота")
        btn_image.setMinimumHeight(45)
        btn_image.clicked.connect(lambda: self.add_image(doc))
        layout.addWidget(btn_image)

        btn_list = QPushButton("📋 Добавить список")
        btn_list.setMinimumHeight(45)
        btn_list.clicked.connect(lambda: self.add_list(doc))
        layout.addWidget(btn_list)

        # Растягивающееся пространство
        layout.addStretch()

        # Разделитель
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(line)

        # Кнопка завершения
        btn_done = QPushButton("✅ Закончить раздел")
        btn_done.setMinimumHeight(45)
        btn_done.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        btn_done.clicked.connect(dialog.accept)
        layout.addWidget(btn_done)

        # Показываем диалог
        dialog.exec()

    def auto_fill_empty_fields(self):
        """Заполняет пустые поля значениями по умолчанию"""
        filled = False

        if not self.subject.text():
            self.subject.setText("Технология объектного программирования")
            filled = True

        if not self.work_type.currentText():
            self.work_type.setCurrentIndex(0)
            filled = True

        if not self.student.text():
            self.student.setText("Шикарев Иван А")
            filled = True

        if not self.group.text():
            self.group.setText("ПО1-23")
            filled = True

        if not self.teacher.toPlainText():
            self.teacher.setPlainText("доц. Федулов Я.С,\nасс. Маракулина Ю.Д")
            filled = True

        if not self.no_variant.isChecked() and self.variant.value() == 0:
            self.variant.setValue(17)
            filled = True

        if filled:
            QMessageBox.information(self, "Автозаполнение",
                                    "Пустые поля заполнены значениями по умолчанию")

        return filled

    def generate_report(self):
        try:
            self.statusBar().showMessage("Генерация...")
            self.generate_btn.setEnabled(False)

            self.auto_fill_empty_fields()

            # Сбор данных
            subject = self.subject.text().strip()
            if not subject:
                QMessageBox.warning(self, "Ошибка", "Введите дисциплину!")
                return

            work_type = self.work_type.currentText().strip()
            if not self.no_work_number.isChecked() and work_type == "Лабораторная работа":
                work_type = f"{work_type} №{self.work_number.value()}"

            student = self.student.text().strip() or "Шикарев Иван А"
            group = self.group.text().strip() or "ПО1-23"
            teacher = self.teacher.toPlainText().strip()
            variant = None if self.no_variant.isChecked() else self.variant.value()
            year = self.year.value()

            # Создание документа
            doc = DocBuilder(
                subject=subject,
                work_type=work_type,
                student=student,
                group=group,
                teacher=teacher,
                variant=variant,
                year=year,
            )

            # Оглавление
            if self.include_toc.isChecked():
                doc.add(TableOfContents())

            # В цикле добавления разделов:
            for i in range(self.sections_list.topLevelItemCount()):
                section_item = self.sections_list.topLevelItem(i)
                section_name = section_item.text(0)

                # Добавляем главный раздел
                doc.add(Heading(section_name, level=1))

                # Добавляем текст главного раздела
                if section_name in self.sections_content:
                    content = self.sections_content[section_name].get("text", "")
                    if content:
                        self.parse_and_add_content(doc, content)

                # Добавляем подразделы
                for j in range(section_item.childCount()):
                    child = section_item.child(j)
                    subsection_name = child.text(0).replace("    ", "")

                    doc.add(Heading(subsection_name, level=2))

                    # Добавляем текст подраздела
                    for sub in self.sections_content.get(section_name, {}).get("subsections", []):
                        if sub["name"] == subsection_name:
                            self.parse_and_add_content(doc, sub.get("text", ""))
                            break

            # Номера страниц
            doc.add_page_numbers_bottom(hide_on_first_page=True)

            # Сохранение
            path = doc.save()

            QMessageBox.information(self, "Успех!", f"✅ Отчет создан!\n\n📄 {path}")

            if QMessageBox.question(self, "Открыть?", "Открыть документ?") == QMessageBox.StandardButton.Yes:
                import os
                os.startfile(path)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"❌ {str(e)}")
        finally:
            self.statusBar().showMessage("Готово")
            self.generate_btn.setEnabled(True)

    def parse_and_add_content(self, doc, content):
        """Разбирает текст и добавляет соответствующие блоки"""
        lines = content.split('\n')
        for line in lines:
            if line.startswith('[ФОРМУЛА]'):
                formula = line.replace('[ФОРМУЛА]', '').strip()
                doc.add(Formula(formula))
            elif line.startswith('[ПОЯСНЕНИЕ]'):
                # Пояснение уже добавлено с формулой
                pass
            elif line.startswith('[СКРИНШОТ:'):
                # TODO: парсить скриншоты
                doc.add(ImagePlaceholder("Скриншот", height_lines=5))
            elif line.strip() and not line.startswith('|'):
                # Обычный текст
                doc.add(Paragraph(line.strip()))
                
    def insert_text(self, section_name):
        """Вставляет текст"""
        text, ok = QInputDialog.getMultiLineText(self, "Введите текст", "Текст:")
        if ok and text:
            cursor = self.section_text.textCursor()
            cursor.insertText(f"\n{text}\n")
            self.save_section_content(section_name)  # Автосохранение

    def insert_table(self, section_name):
        """Вставляет таблицу"""
        cols, ok1 = QInputDialog.getInt(self, "Таблица", "Колонок:", 3, 1, 10)
        if not ok1:
            return
        rows, ok2 = QInputDialog.getInt(self, "Таблица", "Строк:", 3, 1, 50)
        if not ok2:
            return

        # Создаем таблицу в текстовом виде
        table_text = "\n"
        for i in range(rows):
            row = "| " + " | ".join([f"[{i + 1},{j + 1}]" for j in range(cols)]) + " |\n"
            table_text += row
        table_text += "\n"

        cursor = self.section_text.textCursor()
        cursor.insertText(table_text)
        self.save_section_content(section_name)

    def insert_formula(self, section_name):
        """Вставляет формулу"""
        formula, ok1 = QInputDialog.getText(self, "Формула", "Формула:", text="y = f(x)")
        if not ok1:
            return
        explanation, ok2 = QInputDialog.getText(self, "Пояснение", "Пояснение (Enter - пропустить):")

        formula_text = f"\n[ФОРМУЛА] {formula}\n"
        if explanation:
            formula_text += f"[ПОЯСНЕНИЕ] {explanation}\n"

        cursor = self.section_text.textCursor()
        cursor.insertText(formula_text)
        self.save_section_content(section_name)

    def insert_image(self, section_name):
        """Вставляет место для скриншота"""
        caption, ok1 = QInputDialog.getText(self, "Скриншот", "Подпись:", text="Результаты")
        if not ok1:
            return
        height, ok2 = QInputDialog.getInt(self, "Скриншот", "Высота (строк):", 5, 1, 20)

        image_text = f"\n[СКРИНШОТ: {caption}] (высота: {height} строк)\n"

        cursor = self.section_text.textCursor()
        cursor.insertText(image_text)
        self.save_section_content(section_name)

    def insert_list(self, section_name):
        """Вставляет список"""
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

            cursor = self.section_text.textCursor()
            cursor.insertText(list_text)
            self.save_section_content(section_name)

    def save_section_to_doc(self, section_name):
        """Сохраняет содержимое раздела в документ"""
        # TODO: сохранять текст в doc
        QMessageBox.information(self, "Сохранено", f"Раздел «{section_name}» сохранен")
        self.section_window.accept()

    def add_text(self, doc):
        """Добавить текст"""
        text, ok = QInputDialog.getMultiLineText(
            self, "Введите текст",
            "Текст абзаца:",
            "Здесь будет ваш текст..."
        )
        if ok and text:
            doc.add(Paragraph(text))

    def add_table(self, doc):
        """Добавить таблицу"""
        cols, ok1 = QInputDialog.getInt(
            self, "Таблица", "Количество колонок:", 3, 1, 10)
        if not ok1:
            return

        rows, ok2 = QInputDialog.getInt(
            self, "Таблица", "Количество строк данных:", 3, 1, 50)
        if not ok2:
            return

        title, ok3 = QInputDialog.getText(
            self, "Таблица", "Название таблицы:",
            text="Результаты измерений")

        # Создаем диалог для ввода заголовков
        headers = []
        for i in range(cols):
            header, ok = QInputDialog.getText(
                self, f"Колонка {i + 1}",
                f"Название колонки {i + 1}:",
                text=f"Параметр {i + 1}")
            headers.append(header if ok else f"Колонка {i + 1}")

        # Формируем таблицу
        table_data = [headers]
        for i in range(rows):
            row = []
            for j in range(cols):
                row.append(f"[{headers[j]}_{i + 1}]")
            table_data.append(row)

        doc.add(Table(title if title else "Результаты", table_data))

    def add_formula(self, doc):
        """Добавить формулу"""
        formula, ok1 = QInputDialog.getText(
            self, "Формула",
            "Введите формулу (например y = ax + b):",
            text="y = f(x)")

        if not ok1:
            return

        explanation, ok2 = QInputDialog.getText(
            self, "Пояснение",
            "Введите пояснение (Enter - пропустить):")

        doc.add(Formula(formula, explanation=explanation if ok2 else None))

    def add_image(self, doc):
        """Добавить место для скриншота"""
        caption, ok1 = QInputDialog.getText(
            self, "Скриншот",
            "Подпись к рисунку:",
            text="Результаты выполнения")

        height, ok2 = QInputDialog.getInt(
            self, "Скриншот",
            "Высота (строк):", 5, 1, 20)

        if ok1:
            doc.add(ImagePlaceholder(caption, height_lines=height))

    def add_list(self, doc):
        """Добавить нумерованный список"""
        items = []

        dialog = QDialog(self)
        dialog.setWindowTitle("Введите элементы списка")
        layout = QVBoxLayout(dialog)

        list_widget = QListWidget()
        layout.addWidget(list_widget)

        input_field = QLineEdit()
        input_field.setPlaceholderText("Введите элемент и нажмите Добавить")
        layout.addWidget(input_field)

        btn_add = QPushButton("➕ Добавить")
        btn_add.clicked.connect(lambda: self.add_list_item(list_widget, input_field))

        btn_remove = QPushButton("➖ Удалить")
        btn_remove.clicked.connect(lambda: self.remove_list_item(list_widget))

        btn_done = QPushButton("✅ Готово")
        btn_done.clicked.connect(dialog.accept)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_remove)
        btn_layout.addWidget(btn_done)

        layout.addLayout(btn_layout)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            items = [list_widget.item(i).text()
                     for i in range(list_widget.count())]
            if items:
                doc.add(NumberedList(items))

    def add_list_item(self, list_widget, input_field):
        text = input_field.text().strip()
        if text:
            list_widget.addItem(text)
            input_field.clear()

    def remove_list_item(self, list_widget):
        current = list_widget.currentRow()
        if current >= 0:
            list_widget.takeItem(current)

    def create_all_labs(self):
        """Создает все лабораторные работы из диапазона"""
        try:
            from_num = self.lab_from.value()
            to_num = self.lab_to.value()

            if from_num > to_num:
                QMessageBox.warning(self, "Ошибка", "Начальный номер больше конечного!")
                return

            reply = QMessageBox.question(
                self, "Подтверждение",
                f"Создать {to_num - from_num + 1} лабораторных работ?\nс {from_num} по {to_num}",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply != QMessageBox.StandardButton.Yes:
                return

            self.statusBar().showMessage(f"Создание работ {from_num}-{to_num}...")
            self.generate_btn.setEnabled(False)

            created_files = []

            for lab_num in range(from_num, to_num + 1):
                try:
                    subject = self.subject.text().strip() or "Технология объектного программирования"
                    work_type = f"Лабораторная работа №{lab_num}"
                    student = self.student.text().strip() or "Шикарев Иван А"
                    group = self.group.text().strip() or "ПО1-23"
                    teacher = self.teacher.text().strip()
                    variant = None if self.no_variant.isChecked() else self.variant.value()
                    year = self.year.value()

                    doc = DocBuilder(
                        subject=subject,
                        work_type=work_type,
                        student=student,
                        group=group,
                        teacher=teacher,
                        variant=variant,
                        year=year,
                    )

                    if self.include_toc.isChecked():
                        doc.add(TableOfContents())

                    doc.add(Heading("Введение", level=1))
                    doc.add(Paragraph(f"Отчет по {work_type.lower()}"))

                    doc.add(Heading("Теоретические сведения", level=1))
                    doc.add(Heading("Основные положения", level=2))
                    doc.add(NumberedList(["Пункт 1", "Пункт 2", "Пункт 3"]))

                    doc.add(Heading("Ход работы", level=1))
                    doc.add(Paragraph("Описание выполнения..."))

                    doc.add(PageBreak())
                    doc.add(Heading("Заключение", level=1))
                    doc.add(Paragraph("Выводы..."))

                    doc.add_page_numbers_bottom(hide_on_first_page=True)

                    path = doc.save()
                    created_files.append(path)

                except Exception as e:
                    QMessageBox.warning(self, "Ошибка", f"Не удалось создать работу №{lab_num}: {str(e)}")

            if created_files:
                QMessageBox.information(
                    self, "Готово!",
                    f"✅ Создано {len(created_files)} файлов:\n" +
                    "\n".join([f"📄 {Path(f).name}" for f in created_files])
                )

                reply = QMessageBox.question(
                    self, "Открыть папку?",
                    "Открыть папку с созданными файлами?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if reply == QMessageBox.StandardButton.Yes:
                    import os
                    os.startfile("build")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"❌ {str(e)}")
        finally:
            self.statusBar().showMessage("Готово")
            self.generate_btn.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ReportGeneratorUI()
    window.show()
    sys.exit(app.exec())