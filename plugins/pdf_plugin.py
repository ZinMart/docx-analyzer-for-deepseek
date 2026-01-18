# plugins/pdf_plugin.py
"""Простейший плагин для анализа PDF файлов"""

import os
import PyPDF2
from core.plugin_base import DocumentPlugin


class PDFPlugin(DocumentPlugin):
    """Плагин для работы с PDF файлами"""

    def __init__(self):
        super().__init__()
        self.name = "PDF Анализатор"
        self.version = "1.0"
        self.supported_extensions = ['.pdf']

    def analyze(self, file_path):
        """Анализировать PDF файл"""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)

                # Собираем статистику
                stats = {
                    'file_name': os.path.basename(file_path),
                    'pages': len(pdf_reader.pages),
                    'author': pdf_reader.metadata.get('/Author', 'Не указан') if pdf_reader.metadata else 'Не указан',
                    'title': pdf_reader.metadata.get('/Title',
                                                     'Без названия') if pdf_reader.metadata else 'Без названия',
                    'encrypted': pdf_reader.is_encrypted
                }

                # Извлекаем текст с первых 3 страниц
                text_parts = []
                for i, page in enumerate(pdf_reader.pages[:3]):
                    text = page.extract_text()
                    if text.strip():
                        text_parts.append(f"--- Страница {i + 1} ---\n{text}")

                text_sample = "\n\n".join(text_parts)[:1000]

                return {
                    "status": "success",
                    "stats": stats,
                    "text_sample": text_sample
                }

        except Exception as e:
            return {
                "status": "error",
                "message": f"Ошибка при анализе PDF: {str(e)}"
            }