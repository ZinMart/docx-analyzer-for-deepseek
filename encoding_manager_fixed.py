#!/usr/bin/env python3
"""
МЕНЕДЖЕР КОДИРОВОК - гарантированная работа с UTF-8
"""
import os
import sys


class EncodingManager:
    @staticmethod
    def create_file(filepath, content):
        """Создать файл с гарантированной UTF-8 кодировкой"""
        try:
            # Создаем директорию если нужно
            dir_path = os.path.dirname(filepath)
            if dir_path and dir_path != '':
                os.makedirs(dir_path, exist_ok=True)

            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f'✅ Создан: {filepath}')
            return True
        except Exception as e:
            print(f'❌ Ошибка: {e}')
            return False

    @staticmethod
    def read_file(filepath):
        """Прочитать файл с правильной кодировкой"""
        if not os.path.exists(filepath):
            print(f'❌ Файл не найден: {filepath}')
            return None

        # Пробуем разные кодировки
        encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'cp866']

        for encoding in encodings:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    content = f.read()
                print(f'✅ Прочитан: {filepath} (кодировка: {encoding})')
                return content
            except UnicodeDecodeError:
                continue
            except Exception as e:
                print(f'❌ Ошибка чтения: {e}')
                return None

        print(f'❌ Не удалось определить кодировку: {filepath}')
        return None

    @staticmethod
    def check_file(filepath):
        """Проверить кодировку файла"""
        if not os.path.exists(filepath):
            print(f'❌ Файл не найден: {filepath}')
            return None

        content = EncodingManager.read_file(filepath)
        if content:
            # Простая проверка на русский текст
            russian = 'абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ'
            has_russian = any(ch in content for ch in russian)

            if has_russian:
                print(f'   Содержит русский текст: ДА')
            else:
                print(f'   Содержит русский текст: НЕТ')

            print(f'   Длина: {len(content)} символов')
            return True

        return False

    @staticmethod
    def fix_file(filepath):
        """Исправить кодировку файла (привести к UTF-8)"""
        if not os.path.exists(filepath):
            print(f'❌ Файл не найден: {filepath}')
            return False

        content = EncodingManager.read_file(filepath)
        if not content:
            return False

        try:
            # Сохраняем как чистый UTF-8
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f'✅ Исправлен: {filepath} → UTF-8')
            return True
        except Exception as e:
            print(f'❌ Ошибка записи: {e}')
            return False


def main():
    if len(sys.argv) < 2:
        print('Использование:')
        print('  python encoding_manager.py create <файл> "текст"')
        print('  python encoding_manager.py read <файл>')
        print('  python encoding_manager.py check <файл>')
        print('  python encoding_manager.py fix <файл>')
        return

    manager = EncodingManager()
    cmd = sys.argv[1]

    if cmd == 'create' and len(sys.argv) >= 4:
        # Объединяем все аргументы после имени файла
        text = ' '.join(sys.argv[3:])
        manager.create_file(sys.argv[2], text)

    elif cmd == 'read' and len(sys.argv) == 3:
        content = manager.read_file(sys.argv[2])
        if content:
            print('\n=== СОДЕРЖИМОЕ ФАЙЛА ===')
            print(content)
            print('=' * 30)

    elif cmd == 'check' and len(sys.argv) == 3:
        manager.check_file(sys.argv[2])

    elif cmd == 'fix' and len(sys.argv) == 3:
        manager.fix_file(sys.argv[2])

    else:
        print('❌ Неверные аргументы')
        print('Примеры:')
        print('  python encoding_manager.py create test.txt "Привет мир"')
        print('  python encoding_manager.py read test.txt')
        print('  python encoding_manager.py check test.txt')
        print('  python encoding_manager.py fix test.txt')


if __name__ == '__main__':
    main()