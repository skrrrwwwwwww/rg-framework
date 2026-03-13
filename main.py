from builder import DocBuilder
from blocks import Heading, Paragraph, Table, PageBreak

doc = DocBuilder(
    subject="Технология объектного программирования",
    work_type="Лабораторная работа №1",
    student="Шикарев И.А.",
    group="ПО1-23",
    teacher="Федулов Я.А.",
    variant=17,
    year=2025,
)

doc.add(Heading("Введение", level=1))
doc.add(Paragraph("Это текст введения."))

doc.add(Heading("Результаты", level=1))
doc.add(Table("Результаты измерений", [
    ["Параметр", "Значение"],
    ["A", 10],
    ["B", 20],
]))

doc.add(PageBreak())
doc.add(Heading("Заключение", level=1))
doc.add(Paragraph("Работа выполнена успешно."))

path = doc.save()
print("Создан файл:", path)