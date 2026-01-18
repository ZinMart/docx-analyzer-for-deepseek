#!/usr/bin/env python3
"""Проверка кодировки файлов"""
import os

print("Проверка файлов...")

for root, dirs, files in os.walk('.'):
    if '.git' in root or '.venv' in root:
        continue
    for file in files:
        if file.endswith(('.py', '.md', '.txt', '.json')):
            path = os.path.join(root, file)
            try:
                with open(path, 'rb') as f:
                    data = f.read()
                if b'\x00' in data:
                    print(f"ПРОБЛЕМА: {path}: NULL байты")
                else:
                    print(f"OK: {path}")
            except:
                pass

print("Проверка завершена")