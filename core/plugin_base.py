# core/plugin_base.py
"""Базовый класс для всех плагинов"""


class DocumentPlugin:
    """Простейший плагин для анализа документов"""

    def __init__(self):
        self.name = "Базовый плагин"
        self.version = "1.0"
        self.supported_extensions = []  # Например: ['.docx', '.pdf']

    def can_handle(self, file_path):
        """Может ли этот плагин обработать файл?"""
        # Проверяем расширение файла
        import os
        file_ext = os.path.splitext(file_path)[1].lower()
        return file_ext in self.supported_extensions

    def analyze(self, file_path):
        """Проанализировать файл - БАЗОВЫЙ МЕТОД"""
        # Этот метод будут переопределять конкретные плагины
        return {
            "status": "not_implemented",
            "message": "Этот плагин не умеет анализировать файлы"
        }