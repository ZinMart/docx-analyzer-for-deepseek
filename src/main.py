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
        """Обработчик кнопки анализа файла"""
        if self.current_file:
            try:
                # 1. Создаем анализатор
                analyzer = DocxAnalyzer(self.current_file)

                # 2. Получаем базовую информацию
                basic_info = analyzer.get_basic_info()

                # 3. Анализируем содержимое
                stats = analyzer.analyze()

                # 4. Извлекаем текст (первые 500 символов)
                text_sample = analyzer.extract_text()[:500]

                # 5. Формируем сообщение
                message = f"📄 Файл: {basic_info['filename']}\n"
                message += f"👤 Автор: {basic_info['author']}\n"
                message += f"📅 Создан: {basic_info['created']}\n\n"
                message += f"📊 СТАТИСТИКА:\n"
                message += f"• Абзацев: {stats['total_paragraphs']}\n"
                message += f"• Таблиц: {stats['tables']}\n"
                message += f"• Изображений: {stats['images']}\n"
                message += f"• Формул: {stats['formulas']}\n\n"
                message += f"📝 ТЕКСТ (первые 500 символов):\n"
                message += f"{text_sample}..."

                # 6. Показываем результаты
                QMessageBox.information(
                    self,
                    "Результаты анализа",
                    message
                )

            except Exception as e:
                # 7. Если ошибка - показываем детали
                import traceback
                error_details = traceback.format_exc()
                print("ОШИБКА АНАЛИЗА:", error_details)

                QMessageBox.critical(
                    self,
                    "Ошибка анализа",
                    f"Файл открыт, но анализ не удался:\n{str(e)}"
                )
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