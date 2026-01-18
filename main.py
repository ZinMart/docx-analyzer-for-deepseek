"""
Главный файл программы DOCX Analyzer for DeepSeek
Создает графический интерфейс и запускает приложение
"""

import sys
import os
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton,
                             QVBoxLayout, QWidget, QFileDialog, QLabel,
                             QMessageBox)
from PyQt5.QtCore import Qt
from docx_analyzer import DocxAnalyzer


class MainWindow(QMainWindow):
    """Главное окно программы"""

    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.current_file = None

    def setup_ui(self):
        """Настройка графического интерфейса"""
        # Настройки окна
        self.setWindowTitle("DOCX Analyzer for DeepSeek")
        self.setGeometry(100, 100, 600, 400)  # x, y, width, height

        # Создаем виджеты (элементы интерфейса)
        self.title_label = QLabel("DOCX Анализатор для DeepSeek")
        self.title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.info_label = QLabel("Выберите DOCX файл для анализа")
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.file_label = QLabel("Файл не выбран")
        self.file_label.setStyleSheet("color: gray;")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Кнопки
        self.btn_select_file = QPushButton("📁 Выбрать DOCX файл")
        self.btn_select_file.clicked.connect(self.select_file)

        self.btn_select_folder = QPushButton("📂 Выбрать папку для сохранения")
        self.btn_select_folder.clicked.connect(self.select_folder)
        self.btn_select_folder.setEnabled(False)  # Пока неактивна

        self.btn_analyze = QPushButton("🔍 Анализировать файл")
        self.btn_analyze.clicked.connect(self.analyze_file)
        self.btn_analyze.setEnabled(False)  # Пока неактивна

        # Размещаем виджеты в layout (компоновка)
        layout = QVBoxLayout()
        layout.addWidget(self.title_label)
        layout.addWidget(self.info_label)
        layout.addSpacing(20)
        layout.addWidget(self.file_label)
        layout.addSpacing(20)
        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_select_folder)
        layout.addWidget(self.btn_analyze)
        layout.addStretch()  # Добавляем растягивающееся пространство

        # Создаем контейнер и устанавливаем layout
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file(self):
        """Обработчик кнопки выбора файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите DOCX файл",
            "",  # Начальная директория (пустая = последняя использованная)
            "Word Documents (*.docx *.doc);;All Files (*.*)"
        )

        if file_path:
            self.current_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.setText(f"✅ Выбран файл: {filename}")
            self.file_label.setStyleSheet("color: green;")
            self.btn_select_folder.setEnabled(True)
            self.btn_analyze.setEnabled(True)

            # Показываем информацию о файле
            QMessageBox.information(
                self,
                "Файл выбран",
                f"Файл '{filename}' готов к анализу.\n"
                f"Теперь выберите папку для сохранения результатов."
            )

    def select_folder(self):
        """Обработчик кнопки выбора папки"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения результатов"
        )

        if folder_path:
            QMessageBox.information(
                self,
                "Папка выбрана",
                f"Результаты будут сохранены в:\n{folder_path}"
            )

    def analyze_file(self):
        """Обработчик кнопки анализа файла - с выбором плагина"""
        if self.current_file:
            try:
                # Пробуем разные плагины по очереди
                plugins_to_try = []

                # 1. Пробуем DOCX плагин
                try:
                    from plugins.docx_plugin import DocxPlugin
                    plugins_to_try.append(DocxPlugin())
                except ImportError:
                    pass

                # 2. Пробуем PDF плагин
                try:
                    from plugins.pdf_plugin import PDFPlugin
                    plugins_to_try.append(PDFPlugin())
                except ImportError:
                    pass

                # 3. Ищем подходящий плагин
                suitable_plugin = None
                for plugin in plugins_to_try:
                    if plugin.can_handle(self.current_file):
                        suitable_plugin = plugin
                        break

                if suitable_plugin:
                    result = suitable_plugin.analyze(self.current_file)

                    if result["status"] == "success":
                        stats = result["stats"]
                        text = result["text_sample"]

                        message = f"📄 Файл: {stats['file_name']}\n"

                        if 'author' in stats:
                            message += f"👤 Автор: {stats['author']}\n"
                        if 'pages' in stats:
                            message += f"📄 Страниц: {stats['pages']}\n"
                        elif 'paragraphs' in stats:
                            message += f"📝 Абзацев: {stats['paragraphs']}\n"

                        message += f"\n📊 СТАТИСТИКА:\n"
                        for key, value in stats.items():
                            if key not in ['file_name', 'text_sample']:
                                message += f"• {key}: {value}\n"

                        message += f"\n📝 ТЕКСТ (первые 1000 символов):\n"
                        message += f"{text}..."

                        QMessageBox.information(self, "Результаты анализа", message)
                    else:
                        QMessageBox.critical(self, "Ошибка", result["message"])
                else:
                    QMessageBox.warning(self, "Не поддерживается",
                                        f"Формат файла не поддерживается\n\nПоддерживаемые форматы:\n• DOCX/DOC\n• PDF")

            except Exception as e:
                QMessageBox.critical(self, "Ошибка",
                                     f"Не удалось прочитать файл:\n{str(e)}")

def main():
    """Точка входа в программу"""
    app = QApplication(sys.argv)

    # Настройка стиля приложения
    app.setStyle('Fusion')

    # Создание и отображение главного окна
    window = MainWindow()
    window.show()

    # Запуск основного цикла приложения
    sys.exit(app.exec())


if __name__ == "__main__":
    main()