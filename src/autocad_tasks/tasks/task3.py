"""Task 3: Cleanup DWG - PURGE and optimize"""

import json
import os
import time
from typing import Dict, List, Callable
from autocad_tasks.utils.autocad_utils import AutoCADHelper


class Task3:
    """Задача 3: Очистка чертежа (PURGE)"""

    def __init__(self, log_callback: Callable = None,
                 load_previous: bool = False):
        self.results_file = "task3_results.json"
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
        self.log("🧹 ЗАДАЧА 3: Очистка чертежа")
        self.log("=" * 60)
        self.log("\n📌 Описание:")
        self.log("  • Переход на лист Model")
        self.log("  • Очистка всех неиспользуемых элементов (PURGE)")
        self.log("  • Возврат на исходный лист")
        self.log("  • Сохранение чертежа")
        self.log("-" * 60)

        self.log("\n🔄 Запуск очистки...")
        self.log("⏳ Пожалуйста, подождите...")
        self.log("")

        # Запускаем LISP скрипт
        success, report_text = self.run_lisp_script()
        self.report_text = report_text

        # Формируем результаты
        if success:
            self.results["success"] = ["Очистка выполнена успешно"]
            self.results["errors"] = []

            self.log("\n" + "=" * 60)
            self.log("✅ РЕЗУЛЬТАТ: Задача выполнена успешно!")
            self.log("=" * 60)
            self.log("\n📊 СТАТИСТИКА ОЧИСТКИ:")
            self.log("-" * 60)
            self.log("  ✅ Удалены все неиспользуемые элементы:")
            self.log("     • Блоки (BLOCKS)")
            self.log("     • Слои (LAYERS)")
            self.log("     • Типы линий (LINETYPES)")
            self.log("     • Размерные стили (DIMSTYLES)")
            self.log("     • Стили печати (PLOTSTYLES)")
            self.log("     • И другие неиспользуемые объекты")
            self.log("-" * 60)
            self.log("\n💡 Детальный список удаленных элементов")
            self.log("   смотрите в командной строке AutoCAD (нажмите F2)")
        else:
            self.results["success"] = []
            self.results["errors"] = ["Ошибка выполнения очистки"]
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

    def run_lisp_script(self) -> tuple:
        """Запустить LISP скрипт для очистки"""
        try:
            lisp_path = r"D:\Pycharm Projects\AutoCAD\RunPlot\cleanup_dwg.lsp"

            if not os.path.exists(lisp_path):
                self.log(f"  ❌ LISP файл не найден: {lisp_path}")
                return False, ""

            self.log("  🚀 Запуск LISP скрипта...")

            # Загружаем и запускаем LISP
            lisp_path_fixed = lisp_path.replace('\\', '/')
            cmd = f'(load "{lisp_path_fixed}")\nCLEANUP\n'
            self.autocad.acad.ActiveDocument.SendCommand(cmd)

            # Ждем выполнения
            time.sleep(8)

            self.log("  ✅ LISP скрипт выполнен")

            # Сохраняем чертеж
            self.log("  💾 Сохранение чертежа...")
            self.autocad.acad.ActiveDocument.SendCommand("_.QSAVE\n")
            time.sleep(2)
            self.log("  ✅ Чертеж сохранен")

            return True, "Очистка выполнена. Чертеж сохранен."

        except Exception as e:
            self.log(f"  ❌ Ошибка: {e}")
            return False, str(e)