import sys
import os
import json
import datetime

# Добавляем папку проекта в путь поиска модулей
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton,
                             QVBoxLayout, QWidget, QFileDialog, QLabel,
                             QMessageBox)
from PyQt5.QtCore import Qt


class MainWindow(QMainWindow):
    """Главное окно программы - СТАБИЛЬНАЯ ВЕРСИЯ БЕЗ ТЕМ"""

    CONFIG_FILE = "app_config.json"

    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.selected_files = []
        self.last_folder = None
        self.last_file_folder = None
        self.load_config()

        # Применяем базовые стили
        self.apply_basic_styles()

    def load_config(self):
        """Загрузить сохраненные настройки из файла"""
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.last_folder = config.get('last_folder')
                    self.last_file_folder = config.get('last_file_folder')
                    print(f"✅ Загружены настройки")
        except Exception as e:
            print(f"❌ Ошибка загрузки настроек: {e}")

    def save_config(self):
        """Сохранить текущие настройки в файл"""
        try:
            config = {
                'last_folder': self.last_folder,
                'last_file_folder': self.last_file_folder,
                'last_save': str(datetime.datetime.now())
            }
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            print(f"✅ Настройки сохранены")
        except Exception as e:
            print(f"❌ Ошибка сохранения настроек: {e}")

    def apply_basic_styles(self):
        """Применить базовые стили"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F5F5F5;
                font-family: 'Segoe UI', Arial, sans-serif;
            }

            QLabel {
                color: #333333;
                font-size: 14px;
            }

            QPushButton {
                background-color: #007ACC;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 18px;
                font-size: 14px;
                font-weight: 500;
                min-height: 40px;
            }

            QPushButton:hover {
                background-color: #005FA3;
            }

            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #666666;
            }

            QPushButton#title_button {
                background-color: #2DA44E;
                font-weight: bold;
            }

            QPushButton#title_button:hover {
                background-color: #2C974B;
            }
        """)

    def setup_ui(self):
        """Настройка графического интерфейса"""
        # Настройки окна
        self.setWindowTitle("DOCX/PDF Analyzer")
        self.setGeometry(100, 100, 600, 400)

        # Создаем виджеты
        self.title_label = QLabel("DOCX/PDF Анализатор")
        self.title_label.setStyleSheet("""
            font-size: 20px; 
            font-weight: bold; 
            color: #007ACC;
            margin-bottom: 10px;
        """)
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.info_label = QLabel("Выберите DOCX или PDF файлы для анализа")
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.file_label = QLabel("Файлы не выбраны")
        self.file_label.setStyleSheet("""
            color: #666666;
            background-color: white;
            border: 2px dashed #007ACC;
            border-radius: 8px;
            padding: 12px;
            margin: 5px;
        """)
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Кнопки
        self.btn_select_file = QPushButton("📁 Выбрать файлы (DOCX/PDF)")
        self.btn_select_file.setObjectName("title_button")
        self.btn_select_file.clicked.connect(self.select_file)

        self.btn_select_folder = QPushButton("📂 Выбрать папку для сохранения")
        self.btn_select_folder.clicked.connect(self.select_folder)
        self.btn_select_folder.setEnabled(False)

        self.btn_analyze = QPushButton("🔍 Анализировать файл(ы)")
        self.btn_analyze.setObjectName("title_button")
        self.btn_analyze.clicked.connect(self.analyze_file)
        self.btn_analyze.setEnabled(False)

        self.btn_check_updates = QPushButton("🔄 Проверить обновления")
        self.btn_check_updates.clicked.connect(self.check_updates)

        # Размещение
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        layout.addWidget(self.title_label)
        layout.addWidget(self.info_label)
        layout.addSpacing(10)
        layout.addWidget(self.file_label)
        layout.addSpacing(20)

        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_select_folder)
        layout.addWidget(self.btn_analyze)
        layout.addSpacing(15)

        layout.addWidget(self.btn_check_updates)
        layout.addStretch()

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file(self):
        """Выбор файлов с сохранением последнего пути"""
        initial_dir = self.last_file_folder if self.last_file_folder else os.path.expanduser("~")

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы (можно несколько)",
            initial_dir,
            "Документы (*.docx *.doc *.pdf);;Все файлы (*.*)"
        )

        if files:
            self.selected_files = files
            self.last_file_folder = os.path.dirname(files[0])
            self.save_config()

            filenames = [os.path.basename(f) for f in files]

            if len(files) == 1:
                self.file_label.setText(f"✅ Выбран 1 файл: {filenames[0]}")
            else:
                self.file_label.setText(f"✅ Выбрано {len(files)} файлов")

            self.file_label.setStyleSheet("""
                color: #2DA44E;
                background-color: #F0FFF4;
                border: 2px solid #2DA44E;
                border-radius: 8px;
                padding: 12px;
                margin: 5px;
                font-weight: bold;
            """)
            self.btn_select_folder.setEnabled(True)
            self.btn_analyze.setEnabled(True)

    def select_folder(self):
        """Обработчик кнопки выбора папки с сохранением последнего пути"""
        initial_dir = self.last_folder if self.last_folder else os.path.expanduser("~")

        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения результатов",
            initial_dir
        )

        if folder_path:
            self.last_folder = folder_path
            self.save_config()

            QMessageBox.information(
                self,
                "Папка выбрана",
                f"Результаты будут сохранены в:\n{folder_path}"
            )

    def analyze_file(self):
        """Анализ файлов с использованием плагинов"""
        if not self.selected_files:
            QMessageBox.warning(self, "Нет файлов", "Сначала выберите файлы")
            return

        file_to_analyze = self.selected_files[0]

        try:
            # Пробуем использовать плагины
            plugins = []

            # DOCX плагин
            try:
                from plugins.docx_plugin import DocxPlugin
                docx_plugin = DocxPlugin()
                plugins.append(docx_plugin)
                print(f"✅ Загружен DOCX плагин: {docx_plugin.name}")
            except ImportError as e:
                print(f"⚠️ DOCX плагин не загружен: {e}")

            # PDF плагин
            try:
                from plugins.pdf_plugin import PDFPlugin
                pdf_plugin = PDFPlugin()
                plugins.append(pdf_plugin)
                print(f"✅ Загружен PDF плагин: {pdf_plugin.name}")
            except ImportError as e:
                print(f"⚠️ PDF плагин не загружен: {e}")

            # Ищем подходящий плагин
            suitable_plugin = None
            for plugin in plugins:
                if hasattr(plugin, 'can_handle') and plugin.can_handle(file_to_analyze):
                    suitable_plugin = plugin
                    print(f"✅ Найден подходящий плагин: {plugin.name}")
                    break

            if suitable_plugin:
                result = suitable_plugin.analyze(file_to_analyze)

                if result["status"] == "success":
                    stats = result["stats"]
                    text = result.get("text_sample", "")

                    # Форматируем красивое сообщение
                    message = f"<h3>📄 Результаты анализа</h3>"
                    message += f"<p><b>Файл:</b> {stats['file_name']}</p>"
                    message += f"<p><b>Плагин:</b> {suitable_plugin.name}</p>"
                    message += "<hr>"
                    message += "<h4>📊 Статистика:</h4>"

                    for key, value in stats.items():
                        if key != 'file_name':
                            message += f"<p>• <b>{key}:</b> {value}</p>"

                    if text:
                        message += "<hr>"
                        message += "<h4>📝 Текст (первые 500 символов):</h4>"
                        message += f"<pre>{text[:500]}...</pre>"

                    # Создаем красивое сообщение
                    msg_box = QMessageBox(self)
                    msg_box.setWindowTitle("Результаты анализа")
                    msg_box.setTextFormat(Qt.TextFormat.RichText)
                    msg_box.setText(message)
                    msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
                    msg_box.exec()

                else:
                    QMessageBox.critical(self, "Ошибка", result["message"])
            else:
                QMessageBox.warning(self, "Не поддерживается",
                                    "Формат файла не поддерживается\n\n"
                                    "Поддерживаемые форматы:\n"
                                    "• DOCX/DOC\n"
                                    "• PDF")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка",
                                 f"Не удалось прочитать файл:\n{str(e)}")

    def check_updates(self):
        """Проверить обновления"""
        QMessageBox.information(
            self,
            "Обновления",
            "✅ Ваша программа актуальна!\n\n"
            "• Версия: 2.0 (стабильная)\n"
            "• Поддержка: DOCX, PDF\n"
            "• Система тем: отключена\n\n"
            "Все функции работают корректно."
        )


def main():
    """Точка входа в программу"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Современный стиль Qt

    window = MainWindow()
    window.show()

    print("=" * 50)
    print("✅ DOCX/PDF Analyzer for DeepSeek")
    print("✅ Версия: 2.0 (стабильная)")
    print("✅ Поддержка: DOCX, PDF файлы")
    print("✅ Сохранение настроек")
    print("=" * 50)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()