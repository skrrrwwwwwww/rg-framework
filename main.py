from builder import DocBuilder
from blocks import (
    TableOfContents,
    Heading,
    Paragraph,
    NumberedList,
    Table,
    ImagePlaceholder,
    Formula,
    PageBreak,
)
from datetime import datetime
import os

def ask_nonempty(prompt):
    """Запрашивает непустой ввод от пользователя"""
    while True:
        value = input(prompt).strip()
        if value:
            return value
        print("Поле обязательно для заполнения")


def ask_optional_int(prompt, default=None):
    """Запрашивает опциональное целое число"""
    value = input(prompt).strip()
    if not value:
        return default
    try:
        return int(value)
    except ValueError:
        print("Введите целое число или оставьте пустым")
        return ask_optional_int(prompt, default)


def ask_teachers():
    """Запрашивает список преподавателей"""
    teachers = []
    print("\nВведите преподавателей (пустая строка завершает ввод)")

    while True:
        t = input("Преподаватель: ").strip()
        if not t:
            if not teachers:
                print("Нужно ввести хотя бы одного преподавателя")
                continue
            break
        teachers.append(t)

    return "\n".join(teachers)


def ask_yes_no(prompt, default=True):
    """Запрашивает ответ да/нет"""
    default_str = " (Y/n)" if default else " (y/N)"
    while True:
        value = input(prompt + default_str + ": ").strip().lower()
        if not value:
            return default
        if value in ['y', 'yes', 'да', 'д']:
            return True
        if value in ['n', 'no', 'нет', 'н']:
            return False
        print("Пожалуйста, ответьте y или n")


def ask_sections():
    """Пользователь вводит разделы отчета"""
    sections = []
    print("\n📑 Введите названия разделов отчета")
    print("   (пустая строка - закончить ввод)\n")

    while True:
        section = input(f"Раздел {len(sections) + 1}: ").strip()
        if not section:
            if len(sections) == 0:
                print("❌ Нужно добавить хотя бы один раздел!")
                continue
            break
        sections.append(section)

    return sections


def fill_section_interactively(doc, section_name):
    """Интерактивное наполнение раздела"""
    print(f"\n📝 Наполнение раздела: {section_name}")

    while True:
        print("\nЧто добавить в раздел?")
        print("1. Подраздел")
        print("2. Текст")
        print("3. Нумерованный список")
        print("4. Таблицу")
        print("5. Место для скриншота")
        print("6. Формулу")
        print("7. Разрыв страницы")
        print("0. Закончить раздел")

        choice = input("Выберите (0-7): ").strip()

        if choice == "0":
            break
        elif choice == "1":
            level = input("Уровень подраздела (2-5): ").strip()
            try:
                level = int(level)
                if level < 2 or level > 5:
                    level = 2
            except:
                level = 2
            title = input("Название подраздела: ").strip()
            if title:
                doc.add(Heading(title, level=level))
        elif choice == "2":
            text = input("Текст: ").strip()
            if text:
                doc.add(Paragraph(text))
        elif choice == "3":
            items = []
            print("Введите элементы списка (пустая строка - конец):")
            while True:
                item = input("- ").strip()
                if not item:
                    break
                items.append(item)
            if items:
                doc.add(NumberedList(items))
        elif choice == "4":
            cols = input("Количество колонок (2-5): ").strip()
            try:
                cols = int(cols)
                if cols < 2 or cols > 5:
                    cols = 3
            except:
                cols = 3

            rows = input("Количество строк данных: ").strip()
            try:
                rows = int(rows)
                if rows < 1:
                    rows = 3
            except:
                rows = 3

            # Создаем заголовки
            headers = []
            for i in range(cols):
                header = input(f"Название колонки {i + 1}: ").strip()
                headers.append(header if header else f"Колонка {i + 1}")

            # Создаем данные
            table_data = [headers]
            for i in range(rows):
                row = []
                for j in range(cols):
                    row.append(f"[{headers[j]}_{i + 1}]")
                table_data.append(row)

            title = input("Название таблицы: ").strip()
            doc.add(Table(title if title else "Результаты", table_data))
        elif choice == "5":
            caption = input("Подпись к рисунку: ").strip()
            height = input("Высота (строк, Enter=5): ").strip()
            try:
                height = int(height)
            except:
                height = 5
            doc.add(ImagePlaceholder(caption if caption else "Скриншот", height))
        elif choice == "6":
            formula = input("Формула (например y = ax + b): ").strip()
            explanation = input("Пояснение к формуле (Enter - пропустить): ").strip()
            doc.add(Formula(
                formula if formula else "y = f(x)",
                explanation=explanation if explanation else None
            ))
        elif choice == "7":
            doc.add(PageBreak())
        else:
            print("Неверный выбор, попробуйте снова")


def main():
    """Основная функция создания документа"""
    print("=" * 60)
    print("🔧 КОНСТРУКТОР ОТЧЕТОВ 🔧")
    print("=" * 60)

    # Сбор информации
    print("\n📋 ОСНОВНАЯ ИНФОРМАЦИЯ")
    subject = ask_nonempty("Дисциплина: ")
    work_type = ask_nonempty("Тип работы (например Лабораторная работа 1): ")
    student = ask_nonempty("Студент (Фамилия Имя О): ")
    group = ask_nonempty("Группа: ")
    teacher = ask_teachers()
    variant = ask_optional_int("Вариант (Enter = 17): ", default=17)
    year = ask_optional_int(f"Год (Enter = {datetime.now().year}): ",
                            default=datetime.now().year)

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
    print("\n📑 ОГЛАВЛЕНИЕ")
    include_toc = ask_yes_no("Добавить оглавление", default=True)
    if include_toc:
        doc.add(TableOfContents())

    # Ввод разделов
    print("\n📄 СТРУКТУРА ОТЧЕТА")
    sections = ask_sections()

    # Наполнение разделов
    for i, section in enumerate(sections, 1):
        print(f"\n{'=' * 60}")
        print(f"📌 РАЗДЕЛ {i}/{len(sections)}: {section}")
        print(f"{'=' * 60}")

        doc.add(Heading(section, level=1))
        fill_section_interactively(doc, section)

    # Добавление номеров страниц
    doc.add_page_numbers_bottom(hide_on_first_page=True)

    # Сохранение документа
    print("\n" + "=" * 60)
    print("💾 СОХРАНЕНИЕ ДОКУМЕНТА...")
    print("=" * 60)

    try:
        path = doc.save()
        print(f"\n✅ Документ успешно создан!")
        print(f"📄 Путь: {path}")

        # Открыть документ
        if ask_yes_no("\n📂 Открыть созданный документ?"):
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # Linux/Mac
                os.system(f'xdg-open "{path}"')

    except Exception as e:
        print(f"\n❌ Ошибка при создании документа: {e}")
        return

    print("\n✨ Работа программы завершена ✨")


if __name__ == "__main__":
    main()