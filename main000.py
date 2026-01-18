import sys
import os
import json
import datetime

# Добавляем папку проекта в путь поиска модулей
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton,
                             QVBoxLayout, QWidget, QFileDialog, QLabel,
                             QMessageBox)
from PyQt5.QtCore import Qt


class MainWindow(QMainWindow):
    """Главное окно программы"""

    CONFIG_FILE = "app_config.json"

    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.selected_files = []
        self.last_folder = None
        self.last_file_folder = None
        self.load_config()

    def load_config(self):
        """Загрузить сохраненные настройки из файла"""
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.last_folder = config.get('last_folder')
                    self.last_file_folder = config.get('last_file_folder')
                    print(f"✅ Загружены настройки: {config}")
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
            print(f"✅ Настройки сохранены: {config}")
        except Exception as e:
            print(f"❌ Ошибка сохранения настроек: {e}")

    def setup_ui(self):
        """Настройка графического интерфейса"""
        # Настройки окна
        self.setWindowTitle("DOCX Analyzer for DeepSeek")
        self.setGeometry(100, 100, 600, 400)

        # Создаем виджеты
        self.title_label = QLabel("DOCX Анализатор для DeepSeek")
        self.title_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.info_label = QLabel("Выберите DOCX или PDF файлы для анализа")
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.file_label = QLabel("Файлы не выбраны")
        self.file_label.setStyleSheet("color: gray;")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Кнопки
        self.btn_select_file = QPushButton("📁 Выбрать файлы (DOCX/PDF)")
        self.btn_select_file.clicked.connect(self.select_file)

        self.btn_select_folder = QPushButton("📂 Выбрать папку для сохранения")
        self.btn_select_folder.clicked.connect(self.select_folder)
        self.btn_select_folder.setEnabled(False)

        self.btn_analyze = QPushButton("🔍 Анализировать файл(ы)")
        self.btn_analyze.clicked.connect(self.analyze_file)
        self.btn_analyze.setEnabled(False)

        self.btn_check_updates = QPushButton("🔄 Проверить обновления")
        self.btn_check_updates.clicked.connect(self.check_updates)  # ← ИСПРАВЛЕНО

        # Размещение
        layout = QVBoxLayout()
        layout.addWidget(self.title_label)
        layout.addWidget(self.info_label)
        layout.addSpacing(20)
        layout.addWidget(self.file_label)
        layout.addSpacing(20)
        layout.addWidget(self.btn_select_file)
        layout.addWidget(self.btn_select_folder)
        layout.addWidget(self.btn_analyze)
        layout.addWidget(self.btn_check_updates)
        layout.addStretch()

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file(self):
        """Выбор файлов с сохранением последнего пути"""
        # Начальная папка - либо последняя выбранная, либо домашняя
        initial_dir = self.last_file_folder if self.last_file_folder else os.path.expanduser("~")

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы (можно несколько)",
            initial_dir,
            "Документы (*.docx *.doc *.pdf);;All Files (*.*)"
        )

        if files:
            self.selected_files = files

            # Сохраняем папку последнего файла для следующего выбора
            self.last_file_folder = os.path.dirname(files[0])
            self.save_config()  # Сохраняем настройки

            filenames = [os.path.basename(f) for f in files]

            if len(files) == 1:
                self.file_label.setText(f"✅ Выбран 1 файл: {filenames[0]}")
            else:
                self.file_label.setText(f"✅ Выбрано {len(files)} файлов")

            self.file_label.setStyleSheet("color: green;")
            self.btn_select_folder.setEnabled(True)
            self.btn_analyze.setEnabled(True)

    def select_folder(self):
        """Обработчик кнопки выбора папки с сохранением последнего пути"""
        # Начальная папка - либо последняя выбранная, либо домашняя
        initial_dir = self.last_folder if self.last_folder else os.path.expanduser("~")

        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения результатов",
            initial_dir
        )

        if folder_path:
            # Сохраняем путь для следующего раза
            self.last_folder = folder_path
            self.save_config()  # Сохраняем настройки

            QMessageBox.information(
                self,
                "Папка выбрана",
                f"Результаты будут сохранены в:\n{folder_path}"
            )

    def analyze_file(self):
        """Обработчик кнопки анализа файла - с выбором плагина"""
        # Проверяем есть ли выбранные файлы
        if not self.selected_files:
            QMessageBox.warning(self, "Нет файлов", "Сначала выберите файлы")
            return

        # Берем первый файл из списка для анализа
        file_to_analyze = self.selected_files[0]

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
                if plugin.can_handle(file_to_analyze):
                    suitable_plugin = plugin
                    break

            if suitable_plugin:
                result = suitable_plugin.analyze(file_to_analyze)

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

    def check_updates(self):
        """Проверить обновления"""
        try:
            from update_client import SimpleUpdateClient

            client = SimpleUpdateClient()
            updates = client.check_updates()

            if not updates:
                QMessageBox.information(self, "Обновления",
                                        "✅ Все обновления установлены!\n\n"
                                        "Ваша программа актуальна.")
            else:
                message = f"📦 Доступно {len(updates)} обновлений:\n\n"
                for update in updates:
                    message += f"• {update['name']} (v{update['version']})\n"

                message += "\nНажмите 'ОК' чтобы обновиться."

                reply = QMessageBox.question(self, "Обновления доступны",
                                             message, QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)

                if reply == QMessageBox.StandardButton.Ok:
                    self.perform_update(updates)

        except ImportError:
            QMessageBox.warning(self, "Обновления",
                                "Модуль обновлений не установлен")

    def perform_update(self, updates):
        """Выполнить обновление"""
        QMessageBox.information(self, "Обновление",
                                "Обновление будет выполнено в фоновом режиме.\n"
                                "Программа продолжит работу.\n\n"
                                "После загрузки обновлений потребуется перезапуск.")


def main():
    """Точка входа в программу"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()

    print("✅ Программа запущена успешно!")
    print("✅ Окно должно быть открыто")
    print("✅ Проверьте панель задач Windows")

    sys.exit(app.exec())


if __name__ == "__main__":
    main()