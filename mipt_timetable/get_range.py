import openpyxl

# Загружаем файл Excel
wb = openpyxl.load_workbook('data.xlsx')
sheet = wb['1 курс']

# Диапазон ячеек
my_string = "A1:B4"

# Разделяем строку на начальную и конечную ячейки
start_cell, end_cell = my_string.split(':')

# Получаем координаты начальной и конечной ячеек
start_row = int(start_cell[1:])
start_col = openpyxl.utils.column_index_from_string(start_cell[:1])
end_row = int(end_cell[1:])
end_col = openpyxl.utils.column_index_from_string(end_cell[:1])

# Создаем список для хранения значений ячеек
values = []

# Перебираем ячейки в диапазоне
for row in range(start_row, end_row + 1):
    for col in range(start_col, end_col + 1):
        cell = sheet.cell(row=row, column=col)
        values.append(cell.value)

# Выводим значения ячеек
print(values)
