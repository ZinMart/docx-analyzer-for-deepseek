#!/usr/bin/env python3
"""Безопасное создание файлов"""
import sys

if len(sys.argv) != 3:
    print("Использование: python new_file.py ИМЯ_ФАЙЛА ТЕКСТ")
    sys.exit(1)

try:
    with open(sys.argv[1], 'w', encoding='utf-8') as f:
        f.write(sys.argv[2])
    print(f"Файл создан: {sys.argv[1]}")
except Exception as e:
    print(f"Ошибка: {e}")