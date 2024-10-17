import openpyxl

# Загружаем файл Excel
wb = openpyxl.load_workbook('data.xlsx')
sheet = wb['1 курс']  # Выбираем лист

# Получаем ячейку
arr = []
for i in range(1, 20):
    item = sheet[i][14]
    arr.append(item)  #[f"E{i}"])
    cell = sheet[i][14]
    print(item, item.fill.fgColor.rgb)
    print(f"\tЗначение ячейки: {cell.value}")
    print(f"\tЦвет заливки: {cell.fill.fgColor.rgb}")
    # print(f"\tЦвет шрифта: {cell.font.color.rgb}")
    print(f"\tШрифт: {cell.font.name}")
    print(f"\tРазмер шрифта: {cell.font.size}")
    print(f"\tЖирный шрифт: {cell.font.bold}")
    print(f"\tКурсивный шрифт: {cell.font.italic}")
    print(f"\tПодчеркнутый шрифт: {cell.font.underline}")
    print(f"\tВыравнивание: {cell.alignment.horizontal}")
    print(f"\tОбтекание текста: {cell.alignment.wrapText}")
    # print(f"\tВысота строки: {cell.row_height}")
    # print(f"\tШирина столбца: {cell.column_width}")
    print(f"\tКомментарий: {cell.comment}")




print(sheet[1][1])
cell = sheet['E8']  # Измените 'A1' на нужную ячейку
cell1 = sheet['E9']
cell2 = sheet['E10']
cell3 = sheet['E11']

# Получаем цвет ячейки
color = cell.fill.fgColor.rgb
color1 = cell1.fill.fgColor.rgb
color2 = cell2.fill.fgColor.rgb
color3 = cell3.fill.fgColor.rgb
print(color3, color1, color2, color3)

# Преобразуем цвет в формат RGB
color_rgb = int(color, 16)
color_rgb1 = int(color1, 16)
color_rgb2 = int(color2, 16)
color_rgb3 = int(color3, 16)

# Выводим цвет
print(f'Цвет ячейки: {color_rgb}')
print(f'Цвет ячейки: {color_rgb1}')
print(f'Цвет ячейки: {color_rgb2}')
print(f'Цвет ячейки: {color_rgb3}')
