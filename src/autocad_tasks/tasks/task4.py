"""Task 4: Remove duplicate objects - OVERKILL"""

import json
import os
import time
from typing import Dict, List, Callable, Tuple
from autocad_tasks.utils.autocad_utils import AutoCADHelper


class Task4:
    """Задача 4: Удаление дублирующихся объектов (OVERKILL)"""

    def __init__(self, log_callback: Callable = None,
                 load_previous: bool = False):
        self.results_file = "task4_results.json"
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
        self.log("🗑️ ЗАДАЧА 4: Удаление дублирующихся объектов")
        self.log("=" * 60)
        self.log("\n📌 Описание:")
        self.log("  • Применение команды OVERKILL")
        self.log("  • Переход на лист Model")
        self.log("  • Выделение всех объектов")
        self.log("  • Удаление дублирующихся и перекрывающихся объектов")
        self.log("  • Возврат на исходный лист")
        self.log("  • Сохранение чертежа")
        self.log("-" * 60)

        self.log("\n🔄 Запуск OVERKILL...")
        self.log("⏳ Пожалуйста, подождите...")
        self.log("")

        # Запускаем LISP скрипт
        success, report_text = self.run_lisp_script()
        self.report_text = report_text

        # Формируем результаты
        if success:
            self.results["success"] = ["OVERKILL выполнен успешно"]
            self.results["errors"] = []

            self.log("\n" + "=" * 60)
            self.log("✅ РЕЗУЛЬТАТ: Задача выполнена успешно!")
            self.log("=" * 60)
            self.log("\n📊 РЕЗУЛЬТАТ ОЧИСТКИ:")
            self.log("-" * 60)
            self.log("  ✅ Выполнена очистка чертежа от дублирующихся объектов")
            self.log("  🧹 Удалены перекрывающиеся и повторяющиеся элементы")
            self.log("  📏 Параметры OVERKILL:")
            self.log("      • Допуск: 0.001")
            self.log("      • Объединение частично перекрывающихся линий")
            self.log("      • Объединение линий конец-в-конец")
            self.log("      • Оптимизация сегментов полилиний")
            self.log("-" * 60)
            self.log("\n💡 Детальная статистика удаленных объектов")
            self.log("   смотрите в командной строке AutoCAD (нажмите F2)")
        else:
            self.results["success"] = []
            self.results["errors"] = ["Ошибка выполнения OVERKILL"]
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
        """Запустить LISP скрипт для OVERKILL"""
        try:
            lisp_path = r"D:\Pycharm Projects\AutoCAD\RunPlot\overkill_dwg.lsp"

            if not os.path.exists(lisp_path):
                self.log(f"  ❌ LISP файл не найден: {lisp_path}")
                return False, ""

            self.log("  🚀 Запуск LISP скрипта...")

            # Загружаем и запускаем LISP
            lisp_path_fixed = lisp_path.replace('\\', '/')
            cmd = f'(load "{lisp_path_fixed}")\nOVERKILLCLEAN\n'
            self.autocad.acad.ActiveDocument.SendCommand(cmd)

            # Ждем выполнения (OVERKILL может быть долгим)
            time.sleep(15)

            self.log("  ✅ LISP скрипт выполнен")

            # Сохраняем чертеж
            self.log("  💾 Сохранение чертежа...")
            self.autocad.acad.ActiveDocument.SendCommand("_.QSAVE\n")
            time.sleep(2)
            self.log("  ✅ Чертеж сохранен")

            return True, "OVERKILL выполнен. Чертеж сохранен."

        except Exception as e:
            self.log(f"  ❌ Ошибка: {e}")
            return False, str(e)