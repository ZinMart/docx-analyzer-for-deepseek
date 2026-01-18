"""
Простейший клиент для проверки обновлений ВСЕГО
"""

import json
import os
from pathlib import Path


class SimpleUpdateClient:
    """Клиент проверки обновлений"""

    def __init__(self):
        self.current_versions = {
            "core": "1.0.0",
            "docx_plugin": "1.0.0",
            "pdf_plugin": "1.0.0"
        }

    def check_updates(self):
        """Проверить все доступные обновления"""
        print("🔍 Проверяю обновления...")

        try:
            # Пока читаем из локального файла (позже - с сервера)
            updates_file = Path("update_server/all_updates.json")

            if not updates_file.exists():
                print("❌ Файл обновлений не найден")
                return []

            with open(updates_file, 'r', encoding='utf-8') as f:
                all_updates = json.load(f)

            available_updates = []

            # Проверяем обновления ядра
            for core_update in all_updates.get("core_updates", []):
                if self._is_newer_version(core_update["version"], self.current_versions["core"]):
                    available_updates.append({
                        "type": "core",
                        "name": "Ядро программы",
                        "version": core_update["version"],
                        "description": core_update["description"],
                        "size": core_update["size_kb"]
                    })

            # Проверяем обновления плагинов
            for plugin_update in all_updates.get("plugin_updates", []):
                plugin_name = plugin_update["name"]
                current_ver = self.current_versions.get(plugin_name, "0.0.0")

                if self._is_newer_version(plugin_update["version"], current_ver):
                    available_updates.append({
                        "type": "plugin",
                        "name": f"Плагин: {plugin_name}",
                        "version": plugin_update["version"],
                        "description": plugin_update["description"],
                        "size": plugin_update["size_kb"]
                    })

            return available_updates

        except Exception as e:
            print(f"❌ Ошибка при проверке обновлений: {e}")
            return []

    def _is_newer_version(self, new_version, current_version):
        """Сравнить версии (упрощенно)"""
        # Простое сравнение строк
        return new_version != current_version

    def show_updates(self):
        """Показать доступные обновления"""
        updates = self.check_updates()

        if not updates:
            print("✅ Все обновления установлены!")
            return

        print(f"\n📦 Доступно {len(updates)} обновлений:")
        print("-" * 50)

        for i, update in enumerate(updates, 1):
            print(f"{i}. [{update['type'].upper()}] {update['name']}")
            print(f"   Версия: {update['version']}")
            print(f"   Описание: {update['description']}")
            print(f"   Размер: {update['size']} КБ")
            print()


if __name__ == "__main__":
    client = SimpleUpdateClient()
    client.show_updates()