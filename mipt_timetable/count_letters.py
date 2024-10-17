def count_letters(text):
  """Подсчитывает количество букв в строке.

  Args:
    text: Строка для подсчета букв.

  Returns:
    Количество букв в строке.
  """
  count = 0
  for char in text:
    if char.isalpha():
      count += 1
  return count

# Пример использования
my_string = "Привет, мир!"
letter_count = count_letters(my_string)
print(f"Количество букв в строке '{my_string}': {letter_count}")
