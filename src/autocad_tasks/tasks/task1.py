"""Task 1: Выставить рамки для печати"""

import json
import os
import time
from typing import Dict, List, Callable
from autocad_tasks.utils.autocad_utils import AutoCADHelper


class Task1:
    """Задача 1: Выставить рамки для печати (номер листа на задний план, рамку на передний)"""

    def __init__(self, log_callback: Callable = None,
                 load_previous: bool = False):
        self.results_file = "task1_results.json"
        self.log_callback = log_callback
        self.report_text = ""

        # Загружаем предыдущие результаты только если явно запрошено
        if load_previous:
            self.last_results = self.load_last_results()
        else:
            self.last_results = {"success": [], "errors": [], "report": "",
                                 "timestamp": None}

        try:
            self.autocad = AutoCADHelper()
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"❌ Ошибка подключения: {e}")
            raise
        self.results: Dict[str, List[str]] = {"success": [], "errors": []}

    def log(self, message: str):
        """Вывод сообщения через callback или в консоль"""
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def load_last_results(self) -> Dict:
        """Загрузить результаты последнего выполнения"""
        if os.path.exists(self.results_file):
            try:
                with open(self.results_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {"success": [], "errors": [], "report": "", "timestamp": None}

    def save_results(self):
        """Сохранить результаты выполнения"""
        self.last_results = {
            "success": self.results["success"],
            "errors": self.results["errors"],
            "report": self.report_text,
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        try:
            with open(self.results_file, 'w', encoding='utf-8') as f:
                json.dump(self.last_results, f, ensure_ascii=False, indent=2)
        except:
            pass

    def get_status(self) -> Dict:
        """Получить статус последнего выполнения"""
        return self.last_results

    def run(self) -> Dict[str, List[str]]:
        """Выполнение задачи"""
        self.log("\n" + "=" * 60)
        self.log("📐 ЗАДАЧА 1: Выставить рамки для печати")
        self.log("=" * 60)
        self.log("\n📌 Описание:")
        self.log("  • Номер листа (1.06, 3.51) → на задний план")
        self.log("  • Рамка чертежа (8, 5) → на передний план")
        self.log("-" * 60)

        self.log("\n🔄 Запуск обработки...")
        self.log("⏳ Пожалуйста, подождите, идет обработка листов...")
        self.log("")

        success, success_layouts, error_layouts, report_text = self.autocad.run_lisp_script(
            self.log)
        self.report_text = report_text

        # Формируем результаты
        if success:
            self.results["success"] = success_layouts if success_layouts else [
                "Все листы обработаны"]
            self.results["errors"] = error_layouts

            self.log("\n" + "=" * 60)
            if error_layouts:
                self.log("⚠️ РЕЗУЛЬТАТ: Задача выполнена с ошибками!")
            else:
                self.log("✅ РЕЗУЛЬТАТ: Задача выполнена успешно!")
            self.log("=" * 60)

            # Выводим отчет в лог
            if report_text:
                self.log("\n📋 ПОДРОБНЫЙ ОТЧЕТ:")
                self.log("-" * 60)
                for line in report_text.split('\n'):
                    if line.strip():
                        self.log(f"  {line}")
                self.log("-" * 60)
        else:
            self.results["success"] = []
            self.results["errors"] = ["Ошибка выполнения LISP скрипта"]
            self.log("\n" + "=" * 60)
            self.log("❌ РЕЗУЛЬТАТ: Ошибка при выполнении задачи!")
            self.log("=" * 60)

        self.save_results()

        self.log("\n" + "-" * 60)
        self.log(f"📊 Статистика:")
        self.log(f"   ✅ Успешно: {len(self.results['success'])} листов")
        self.log(f"   ❌ Ошибок: {len(self.results['errors'])} листов")
        self.log("-" * 60)

        if self.results['errors']:
            self.log("\n❌ Список листов с ошибками:")
            for error in self.results['errors']:
                self.log(f"   • {error}")

        return self.results