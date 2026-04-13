"""Task 5: Audit DWG - check and fix errors"""

import json
import os
import time
from typing import Dict, List, Callable, Tuple
from autocad_tasks.utils.autocad_utils import AutoCADHelper


class Task5:
    """Задача 5: Проверка и исправление ошибок чертежа (AUDIT)"""

    def __init__(self, log_callback: Callable = None,
                 load_previous: bool = False):
        self.results_file = "task5_results.json"
        self.log_callback = log_callback
        self.report_text = ""

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
        self.log("🔍 ЗАДАЧА 5: Проверка и исправление ошибок чертежа")
        self.log("=" * 60)
        self.log("\n📌 Описание:")
        self.log("  • Применение команды AUDIT")
        self.log("  • Переход на лист Model")
        self.log("  • Автоматическое исправление найденных ошибок")
        self.log("  • Возврат на исходный лист")
        self.log("  • Сохранение чертежа")
        self.log("-" * 60)

        self.log("\n🔄 Запуск проверки...")
        self.log("⏳ Пожалуйста, подождите...")
        self.log("")

        # Запускаем LISP скрипт
        success, report_text = self.run_lisp_script()
        self.report_text = report_text

        # Формируем результаты
        if success:
            self.results["success"] = ["Проверка выполнена успешно"]
            self.results["errors"] = []

            self.log("\n" + "=" * 60)
            self.log("✅ РЕЗУЛЬТАТ: Задача выполнена успешно!")
            self.log("=" * 60)
            self.log("\n📊 РЕЗУЛЬТАТ ПРОВЕРКИ:")
            self.log("-" * 60)
            self.log("  ✅ Выполнена проверка чертежа (AUDIT)")
            self.log("  🛠️ Найденные ошибки исправлены автоматически")
            self.log("-" * 60)
            self.log("\n💡 Детальный отчет об ошибках")
            self.log("   смотрите в командной строке AutoCAD (нажмите F2)")
        else:
            self.results["success"] = []
            self.results["errors"] = ["Ошибка выполнения AUDIT"]
            self.log("\n" + "=" * 60)
            self.log("❌ РЕЗУЛЬТАТ: Ошибка при выполнении задачи!")
            self.log("=" * 60)

        self.save_results()

        self.log("\n" + "-" * 60)
        self.log(f"📊 Статистика выполнения:")
        self.log(f"   ✅ Успешно: {len(self.results['success'])}")
        self.log(f"   ❌ Ошибок: {len(self.results['errors'])}")
        self.log("-" * 60)

        return self.results

    def run_lisp_script(self) -> Tuple[bool, str]:
        """Запустить LISP скрипт для AUDIT"""
        try:
            lisp_path = r"D:\Pycharm Projects\AutoCAD\RunPlot\audit_dwg.lsp"

            if not os.path.exists(lisp_path):
                self.log(f"  ❌ LISP файл не найден: {lisp_path}")
                return False, ""

            self.log("  🚀 Запуск AUDIT...")

            # Загружаем и запускаем LISP
            lisp_path_fixed = lisp_path.replace('\\', '/')
            cmd = f'(load "{lisp_path_fixed}")\nAUDITDWG\n'
            self.autocad.acad.ActiveDocument.SendCommand(cmd)

            # Ждем выполнения
            time.sleep(8)

            self.log("  ✅ AUDIT выполнен")

            # Сохраняем чертеж
            self.log("  💾 Сохранение чертежа...")
            self.autocad.acad.ActiveDocument.SendCommand("_.QSAVE\n")
            time.sleep(2)
            self.log("  ✅ Чертеж сохранен")

            return True, "AUDIT выполнен. Чертеж сохранен."

        except Exception as e:
            self.log(f"  ❌ Ошибка: {e}")
            return False, str(e)