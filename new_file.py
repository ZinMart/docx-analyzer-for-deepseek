#!/usr/bin/env python3
"""
Безопасное создание файлов с UTF-8 кодировкой
Использование: python new_file.py имя_файла текст
"""

import sys


def create_file(filename, content):
    """Создать файл с UTF-8 кодировкой"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✅ Файл создан: {filename}")
        return True
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return False


def main():
    # Проверяем количество аргументов
    if len(sys.argv) < 3:
        print("Использование: python new_file.py имя_файла текст")
        print("Пример: python new_file.py test.txt 'Привет мир'")
        sys.exit(1)

    # Имя файла - первый аргумент
    filename = sys.argv[1]

    # Текст - все остальные аргументы объединяем
    # sys.argv[2:] - все аргументы со второго до конца
    text_parts = sys.argv[2:]
    content = ' '.join(text_parts)

    # Создаем файл
    if create_file(filename, content):
        sys.exit(0)  # Успех
    else:
        sys.exit(1)  # Ошибка


if __name__ == "__main__":
    main()