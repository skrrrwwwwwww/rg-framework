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


def main():
    """Основная функция создания документа"""
    print("=" * 50)
    print("СОЗДАНИЕ ОТЧЕТА ПО ЛАБОРАТОРНОЙ РАБОТЕ")
    print("=" * 50)

    # Сбор информации
    subject = ask_nonempty("Дисциплина: ")
    work_type = ask_nonempty("Тип работы (например ЛР 1): ")
    student = ask_nonempty("Студент: ")
    group = ask_nonempty("Группа: ")
    teacher = ask_teachers()
    variant = ask_optional_int("Вариант (Enter = без варианта): ")
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

    # Добавление содержания
    doc.add(TableOfContents())

    # Введение
    doc.add(Heading("Введение", level=1))
    doc.add(Paragraph(
        f"В настоящем отчете рассматривается выполнение {work_type.lower()} "
        f"по дисциплине «{subject}». Ниже приведены теоретические "
        f"сведения, ход выполнения и результаты работы."
    ))

    # Теоретические сведения
    doc.add(Heading("Теоретические сведения", level=1))

    if ask_yes_no("\nДобавить раздел 'Основные положения'?"):
        doc.add(Heading("Основные положения", level=2))
        points = []
        print("Введите основные положения (пустая строка завершает ввод):")
        while True:
            point = input("- ").strip()
            if not point:
                break
            points.append(point)
        if points:
            doc.add(NumberedList(points))

    doc.add(Heading("Расчетные зависимости", level=2))
    doc.add(Paragraph("Основные расчетные зависимости:"))
    doc.add(Formula(
        "y = f(x_1, x_2, ..., x_n)",
        number=1,
        explanation="где y — результат; x_i — входные параметры; f — функция преобразования."
    ))

    # Ход работы
    doc.add(Heading("Ход работы", level=1))

    if ask_yes_no("Добавить таблицу с результатами?"):
        doc.add(Paragraph("Результаты измерений представлены в таблице 1."))
        # Создаем базовую таблицу
        table_data = [
            ["№", "Параметр", "Значение", "Погрешность"],
            ["1", "Измерение 1", "[вставить]", "[вставить]"],
            ["2", "Измерение 2", "[вставить]", "[вставить]"],
            ["3", "Измерение 3", "[вставить]", "[вставить]"],
        ]
        doc.add(Table("Результаты измерений", table_data))

    if ask_yes_no("Добавить место для скриншота?"):
        doc.add(Paragraph("На рисунке 1 представлен скриншот результатов выполнения."))
        doc.add(ImagePlaceholder("Скриншот результатов", height_lines=5))

    # Заключение
    doc.add(PageBreak())
    doc.add(Heading("Заключение", level=1))
    doc.add(Paragraph(
        "В результате выполнения лабораторной работы были получены необходимые "
        "результаты. Все поставленные задачи выполнены, цели достигнуты."
    ))

    # Добавление номеров страниц
    doc.add_page_numbers_bottom(hide_on_first_page=True)

    # Сохранение документа
    print("\n" + "=" * 50)
    print("СОЗДАНИЕ ДОКУМЕНТА...")
    print("=" * 50)

    try:
        path = doc.save()
        print(f"\n✅ Документ успешно создан!")
        print(f"📄 Путь: {path}")

        # Пытаемся открыть документ
        if ask_yes_no("\nОткрыть созданный документ?"):
            if os.name == 'nt':  # Windows
                os.startfile(path)
            elif os.name == 'posix':  # Linux/Mac
                os.system(f'xdg-open "{path}"')

    except Exception as e:
        print(f"\n❌ Ошибка при создании документа: {e}")
        return

    print("\nРабота программы завершена")


if __name__ == "__main__":
    main()