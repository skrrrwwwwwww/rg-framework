from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import *
from ui.dialogs import SectionEditorDialog

class SectionsTab(QWidget):
    def __init__(self, model, main_window):
        super().__init__()
        self.model = model
        self.main_window = main_window
        self.layout = QVBoxLayout(self)

        # Дерево разделов
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Разделы отчета")
        self.tree.setMinimumHeight(300)
        self.layout.addWidget(self.tree)

        # Кнопки управления
        btn_layout = QHBoxLayout()

        self.add_section_btn = QPushButton("➕ Добавить раздел")
        self.add_section_btn.clicked.connect(self.add_section)

        self.add_subsection_btn = QPushButton("↳ Добавить подраздел")
        self.add_subsection_btn.clicked.connect(self.add_subsection)

        self.remove_btn = QPushButton("➖ Удалить")
        self.remove_btn.clicked.connect(self.remove_item)

        self.edit_btn = QPushButton("✏️ Редактировать")
        self.edit_btn.clicked.connect(self.open_editor)

        btn_layout.addWidget(self.add_section_btn)
        btn_layout.addWidget(self.add_subsection_btn)
        btn_layout.addWidget(self.remove_btn)
        btn_layout.addWidget(self.edit_btn)

        self.layout.addLayout(btn_layout)

        self.refresh_tree()

    def refresh_tree(self):
        """Обновляет дерево из model"""
        self.tree.clear()
        for section_name, section_data in self.model.get_all_sections().items():
            section_item = QTreeWidgetItem(self.tree)
            section_item.setText(0, section_name)
            section_item.setData(0, Qt.ItemDataRole.UserRole, "section")

            for sub in section_data.get("subsections", []):
                sub_item = QTreeWidgetItem(section_item)
                sub_item.setText(0, f"    {sub['name']}")
                sub_item.setData(0, Qt.ItemDataRole.UserRole, "subsection")

        self.tree.expandAll()

    def add_section(self):
        name, ok = QInputDialog.getText(self, "Новый раздел", "Введите название раздела:")
        if ok and name:
            self.model.add_section(name)
            self.refresh_tree()

    def add_subsection(self):
        current = self.tree.currentItem()
        if not current:
            QMessageBox.warning(self, "Ошибка", "Выберите родительский раздел!")
            return

        # Определяем родительский раздел
        parent_item = current
        while parent_item.parent():
            parent_item = parent_item.parent()
        parent_name = parent_item.text(0)

        sub_name, ok = QInputDialog.getText(self, "Новый подраздел", "Введите название подраздела:")
        if ok and sub_name:
            self.model.add_subsection(parent_name, sub_name)
            self.refresh_tree()

    def remove_item(self):
        current = self.tree.currentItem()
        if not current:
            return

        item_type = current.data(0, Qt.ItemDataRole.UserRole)
        name = current.text(0).strip()
        if item_type == "subsection":
            name = name.replace("    ", "")
            parent = current.parent()
            parent_name = parent.text(0)
            self.model.remove_subsection(parent_name, name)
        else:
            self.model.remove_section(name)

        self.refresh_tree()

    def open_editor(self):
        current = self.tree.currentItem()
        if not current:
            QMessageBox.warning(self, "Ошибка", "Выберите раздел!")
            return
        # Сигнал будет обрабатываться в главном окне
        # Здесь просто испускаем сигнал или вызываем метод родителя
        # Для простоты передадим в главное окно через сигнал
        self.main_window.open_section_editor(current)