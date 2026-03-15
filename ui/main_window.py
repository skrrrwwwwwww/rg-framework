from pathlib import Path

from PyQt6.QtCore import *
from PyQt6.QtGui import QShortcut, QKeySequence, QAction
from PyQt6.QtWidgets import *

from core.blocks import *
from core.builder import DocBuilder
from ui.dialogs import SectionEditorDialog
from ui.main_tab import MainTab
from models.section_model import SectionModel
from ui.sections_tab import SectionsTab


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Генератор отчетов Word")
        self.setGeometry(100, 100, 800, 600)

        # Модель данных
        self.section_model = SectionModel()

        # Центральный виджет
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Заголовок
        title = QLabel("🔧 КОНСТРУКТОР ОТЧЕТОВ")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 24px; font-weight: bold; margin: 20px;")
        layout.addWidget(title)

        # Табы
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Вкладка "Основное"
        self.main_tab = MainTab()
        self.tabs.addTab(self.main_tab, "📋 Основное")

        # Вкладка "Разделы"
        self.sections_tab = SectionsTab(self.section_model, self)
        self.tabs.addTab(self.sections_tab, "📑 Разделы")

        # Подключаем кнопку "Создать все" из main_tab
        self.main_tab.create_all_btn.clicked.connect(self.create_all_labs)

        # Подключаем редактирование раздела
        # (в sections_tab мы будем вызывать метод родителя)

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

        # Фильтры событий для автозаполнения
        self.setup_event_filters()

        # Меню и хоткеи
        self.create_menu()
        self.create_shortcuts()

        # Фокус
        self.main_tab.set_focus()

        self.statusBar().showMessage("Ctrl+S - сгенерировать | F5 - все лабы | Ctrl+N - сброс")

    # --- Методы, которые остаются в главном окне ---

    def setup_event_filters(self):
        """Устанавливает фильтры событий для полей ввода"""
        fields = [
            self.main_tab.subject, self.main_tab.work_type, self.main_tab.work_number,
            self.main_tab.student, self.main_tab.group, self.main_tab.teacher
        ]
        for field in fields:
            field.installEventFilter(self)

    def eventFilter(self, obj, event):
        """Автозаполнение по Tab"""
        if event.type() == QEvent.Type.KeyPress and event.key() == Qt.Key.Key_Tab:
            if obj == self.main_tab.subject and not self.main_tab.subject.text():
                self.main_tab.subject.setText("Технология объектного программирования")
                self.main_tab.subject.selectAll()
                return True
            elif obj == self.main_tab.work_type and not self.main_tab.work_type.currentText():
                self.main_tab.work_type.setCurrentIndex(0)
                return True
            elif obj == self.main_tab.work_number and self.main_tab.work_number.value() == 1:
                pass
            elif obj == self.main_tab.student and not self.main_tab.student.text():
                self.main_tab.student.setText("Шикарев Иван А")
                self.main_tab.student.selectAll()
                return True
            elif obj == self.main_tab.group and not self.main_tab.group.text():
                self.main_tab.group.setText("ПО1-23")
                self.main_tab.group.selectAll()
                return True
            elif obj == self.main_tab.teacher and not self.main_tab.teacher.toPlainText():
                self.main_tab.teacher.setPlainText("доц. Федулов А.С,\nасс. Маракулина Ю.Д")
                self.main_tab.teacher.selectAll()
                return True
        return super().eventFilter(obj, event)

    def create_shortcuts(self):
        # Ctrl+N - новый отчет (сброс)
        new_shortcut = QShortcut(QKeySequence("Ctrl+N"), self)
        new_shortcut.activated.connect(self.reset_form)

        # Ctrl+S - сохранить/сгенерировать
        save_shortcut = QShortcut(QKeySequence("Ctrl+S"), self)
        save_shortcut.activated.connect(self.generate_report)

        # Ctrl+Q - выход
        quit_shortcut = QShortcut(QKeySequence("Ctrl+Q"), self)
        quit_shortcut.activated.connect(self.close)

        # F5 - создать все лабы
        create_all_shortcut = QShortcut(QKeySequence("F5"), self)
        create_all_shortcut.activated.connect(self.create_all_labs)

        # Ctrl+T - добавить раздел
        add_section_shortcut = QShortcut(QKeySequence("Ctrl+T"), self)
        add_section_shortcut.activated.connect(self.sections_tab.add_section)

        # Ctrl+R - удалить раздел
        remove_section_shortcut = QShortcut(QKeySequence("Ctrl+R"), self)
        remove_section_shortcut.activated.connect(self.sections_tab.remove_item)

        # Ctrl+E - редактировать раздел
        edit_section_shortcut = QShortcut(QKeySequence("Ctrl+E"), self)
        edit_section_shortcut.activated.connect(self.open_current_section)

        # Alt+1 - перейти на вкладку "Основное"
        tab1_shortcut = QShortcut(QKeySequence("Alt+1"), self)
        tab1_shortcut.activated.connect(lambda: self.tabs.setCurrentIndex(0))

        # Alt+2 - перейти на вкладку "Разделы"
        tab2_shortcut = QShortcut(QKeySequence("Alt+2"), self)
        tab2_shortcut.activated.connect(lambda: self.tabs.setCurrentIndex(1))

    def open_current_section(self):
        """Открывает редактор для выбранного в данный момент раздела"""
        current = self.sections_tab.tree.currentItem()
        if current:
            self.open_section_editor(current)

    def open_section_editor(self, item):
        """Открывает диалог редактирования раздела"""
        item_type = item.data(0, Qt.ItemDataRole.UserRole)
        name = item.text(0).strip()
        if item_type == "subsection":
            name = name.replace("    ", "")
            parent_item = item.parent()
            parent_name = parent_item.text(0)
            full_name = f"{parent_name} → {name}"
        else:
            parent_name = None
            full_name = name

        dialog = SectionEditorDialog(self.section_model, name, parent_name, full_name, self)
        dialog.exec()
        self.sections_tab.refresh_tree()

    def create_menu(self):
        menubar = self.menuBar()

        # Меню "Файл"
        file_menu = menubar.addMenu("📁 Файл")

        new_action = QAction("🆕 Новый", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.reset_form)
        file_menu.addAction(new_action)

        save_action = QAction("💾 Сгенерировать", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.generate_report)
        file_menu.addAction(save_action)

        file_menu.addSeparator()

        quit_action = QAction("🚪 Выход", self)
        quit_action.setShortcut("Ctrl+Q")
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        # Меню "Разделы"
        sections_menu = menubar.addMenu("📚 Разделы")

        add_section_action = QAction("➕ Добавить раздел", self)
        add_section_action.setShortcut("Ctrl+T")
        add_section_action.triggered.connect(self.sections_tab.add_section)
        sections_menu.addAction(add_section_action)

        edit_section_action = QAction("✏️ Редактировать", self)
        edit_section_action.setShortcut("Ctrl+E")
        edit_section_action.triggered.connect(self.open_current_section)
        sections_menu.addAction(edit_section_action)

        remove_section_action = QAction("➖ Удалить", self)
        remove_section_action.setShortcut("Ctrl+R")
        remove_section_action.triggered.connect(self.sections_tab.remove_item)
        sections_menu.addAction(remove_section_action)

        # Меню "Вид"
        view_menu = menubar.addMenu("👁️ Вид")

        main_tab_action = QAction("📋 Основное", self)
        main_tab_action.setShortcut("Alt+1")
        main_tab_action.triggered.connect(lambda: self.tabs.setCurrentIndex(0))
        view_menu.addAction(main_tab_action)

        sections_tab_action = QAction("📑 Разделы", self)
        sections_tab_action.setShortcut("Alt+2")
        sections_tab_action.triggered.connect(lambda: self.tabs.setCurrentIndex(1))
        view_menu.addAction(sections_tab_action)

        # Меню "Справка"
        help_menu = menubar.addMenu("❓ Справка")

        shortcuts_action = QAction("⌨️ Горячие клавиши", self)
        shortcuts_action.triggered.connect(self.show_shortcuts)
        help_menu.addAction(shortcuts_action)

        # Добавляем действия в окно для глобального перехвата
        self.addAction(new_action)
        self.addAction(save_action)
        self.addAction(quit_action)
        self.addAction(add_section_action)
        self.addAction(edit_section_action)
        self.addAction(remove_section_action)
        self.addAction(main_tab_action)
        self.addAction(sections_tab_action)

    def show_shortcuts(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Горячие клавиши")
        msg.setIcon(QMessageBox.Icon.Information)
        text = """
        <h3>⌨️ Горячие клавиши</h3>

        <b>Навигация:</b><br>
        • Tab — автозаполнение/следующее поле<br>
        • Ctrl+↑/↓ — навигация по полям<br>
        • Alt+1 — вкладка "Основное"<br>
        • Alt+2 — вкладка "Разделы"<br>
        <br>

        <b>Действия:</b><br>
        • Ctrl+S — сгенерировать отчет<br>
        • Ctrl+N — новый отчет (сброс)<br>
        • Ctrl+Q — выход<br>
        <br>

        <b>Разделы:</b><br>
        • Ctrl+T — добавить раздел<br>
        • Ctrl+E — редактировать раздел<br>
        • Ctrl+R — удалить раздел<br>
        • F5 — создать все лабы<br>
        <br>

        <b>В окне раздела:</b><br>
        • Ctrl+Enter — сохранить и закрыть<br>
        • Esc — закрыть без сохранения
        """
        msg.setText(text)
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.exec()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Up and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.focusPreviousChild()
        elif event.key() == Qt.Key.Key_Down and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.focusNextChild()
        elif event.key() == Qt.Key.Key_Return and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.generate_report()
        else:
            super().keyPressEvent(event)

    def reset_form(self):
        reply = QMessageBox.question(self, "Сброс",
                                     "Сбросить все поля к значениям по умолчанию?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.main_tab.set_defaults()
            self.section_model.sections.clear()
            self.sections_tab.refresh_tree()
            QMessageBox.information(self, "Готово", "Форма сброшена")

    def auto_fill_empty_fields(self):
        """Заполняет пустые поля значениями по умолчанию"""
        filled = False
        data = self.main_tab.get_data()
        if not data["subject"]:
            self.main_tab.subject.setText("Технология объектного программирования")
            filled = True
        if not data["work_type"]:
            self.main_tab.work_type.setCurrentIndex(0)
            filled = True
        if not data["student"]:
            self.main_tab.student.setText("Шикарев Иван А")
            filled = True
        if not data["group"]:
            self.main_tab.group.setText("ПО1-23")
            filled = True
        if not data["teacher"]:
            self.main_tab.teacher.setPlainText("доц. Федулов Я.С,\nасс. Маракулина Ю.Д")
            filled = True
        if not data["variant"] and self.main_tab.variant.value() == 0:
            self.main_tab.variant.setValue(17)
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
            data = self.main_tab.get_data()

            subject = data["subject"]
            if not subject:
                QMessageBox.warning(self, "Ошибка", "Введите дисциплину!")
                return

            work_type = data["work_type"]
            if not data["no_work_number"] and work_type == "Лабораторная работа":
                work_type = f"{work_type} №{data['work_number']}"

            student = data["student"] or "Шикарев Иван А"
            group = data["group"] or "ПО1-23"
            teacher = data["teacher"]
            variant = data["variant"]
            year = data["year"]

            doc = DocBuilder(
                subject=subject,
                work_type=work_type,
                student=student,
                group=group,
                teacher=teacher,
                variant=variant,
                year=year,
            )

            if data["include_toc"]:
                doc.add(TableOfContents())

            # Добавляем разделы из модели
            for section_name, section_data in self.section_model.get_all_sections().items():
                doc.add(Heading(section_name, level=1))
                if section_data["text"]:
                    self.parse_and_add_content(doc, section_data["text"])

                for sub in section_data["subsections"]:
                    doc.add(Heading(sub["name"], level=2))
                    if sub["text"]:
                        self.parse_and_add_content(doc, sub["text"])

            doc.add_page_numbers_bottom(hide_on_first_page=True)
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
        lines = content.split('\n')
        for line in lines:
            if line.startswith('[ФОРМУЛА]'):
                formula = line.replace('[ФОРМУЛА]', '').strip()
                doc.add(Formula(formula))
            elif line.startswith('[ПОЯСНЕНИЕ]'):
                pass
            elif line.startswith('[СКРИНШОТ:'):
                doc.add(ImagePlaceholder("Скриншот", height_lines=5))
            elif line.strip() and not line.startswith('|'):
                doc.add(Paragraph(line.strip()))

    def create_all_labs(self):
        try:
            from_num = self.main_tab.lab_from.value()
            to_num = self.main_tab.lab_to.value()

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
            base_data = self.main_tab.get_data()
            subject = base_data["subject"] or "Технология объектного программирования"
            student = base_data["student"] or "Шикарев Иван А"
            group = base_data["group"] or "ПО1-23"
            teacher = base_data["teacher"]
            variant = base_data["variant"]
            year = base_data["year"]

            for lab_num in range(from_num, to_num + 1):
                try:
                    work_type = f"Лабораторная работа №{lab_num}"
                    doc = DocBuilder(
                        subject=subject,
                        work_type=work_type,
                        student=student,
                        group=group,
                        teacher=teacher,
                        variant=variant,
                        year=year,
                    )
                    if base_data["include_toc"]:
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
                reply = QMessageBox.question(self, "Открыть папку?", "Открыть папку с созданными файлами?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.Yes:
                    import os
                    os.startfile("build")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"❌ {str(e)}")
        finally:
            self.statusBar().showMessage("Готово")
            self.generate_btn.setEnabled(True)