"""Task 2: Setup print settings"""

import json
import os
import time
from typing import Dict, List, Callable, Tuple
from autocad_tasks.utils.autocad_utils import AutoCADHelper


class Task2:
    """Задача 2: Установка переменных для печати"""

    def __init__(self, log_callback: Callable = None, base_path: str = None, load_previous: bool = False):
        self.results_file = "task2_results.json"
        self.log_callback = log_callback
        self.report_text = ""
        self.base_path = base_path

        if load_previous:
            self.last_results = self.load_last_results()
        else:
            self.last_results = {"success": [], "errors": [], "report": "", "timestamp": None}

        try:
            self.autocad = AutoCADHelper()
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"❌ Ошибка подключения: {e}")
            raise
        self.results: Dict[str, List[str]] = {"success": [], "errors": []}

    def log(self, message: str):
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)

    def load_last_results(self) -> Dict:
        if os.path.exists(self.results_file):
            try:
                with open(self.results_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {"success": [], "errors": [], "report": "", "timestamp": None}

    def save_results(self):
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
        return self.last_results

    def run(self) -> Dict[str, List[str]]:
        self.log("\n" + "=" * 60)
        self.log("⚙️ ЗАДАЧА 2: Установка переменных для печати")
        self.log("=" * 60)
        self.log("\n📌 Описание:")
        self.log("  • PDFFRAME = 2 (отображать рамки PDF)")
        self.log("  • IMAGEFRAME = 2 (отображать рамки изображений)")
        self.log("  • Контуры маскировки: отображать, но не печатать")
        self.log("  • Сохранить чертеж после изменений")
        self.log("-" * 60)

        self.log("\n🔄 Запуск настройки...")
        self.log("⏳ Пожалуйста, подождите...")
        self.log("")

        success, _, _, report_text = self.autocad.run_lisp_script("setup_print.lsp", "SetupPrint", self.log)
        self.report_text = report_text

        if success:
            self.results["success"] = ["Задача выполнена успешно"]
            self.results["errors"] = []

            self.log("\n" + "=" * 60)
            self.log("✅ РЕЗУЛЬТАТ: Задача выполнена успешно!")
            self.log("=" * 60)

            if report_text:
                self.log("\n📋 ПОДРОБНЫЙ ОТЧЕТ:")
                self.log("-" * 60)
                for line in report_text.split('\n'):
                    if line.strip():
                        self.log(f"  {line}")
                self.log("-" * 60)
            else:
                self.log("\n💡 Подробности смотрите в командной строке AutoCAD (нажмите F2)")
        else:
            self.results["success"] = []
            self.results["errors"] = ["Ошибка выполнения LISP скрипта"]
            self.log("\n" + "=" * 60)
            self.log("❌ РЕЗУЛЬТАТ: Ошибка при выполнении задачи!")
            self.log("=" * 60)

        self.save_results()

        self.log("\n" + "-" * 60)
        self.log(f"📊 Статистика:")
        self.log(f"   ✅ Успешно: {len(self.results['success'])}")
        self.log(f"   ❌ Ошибок: {len(self.results['errors'])}")
        self.log("-" * 60)

        return self.results