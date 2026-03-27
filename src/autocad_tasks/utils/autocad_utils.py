"""AutoCAD helper utilities - LISP with result file"""

import time
import pythoncom
import os
from typing import List, Tuple, Callable
import comtypes.client


class AutoCADHelper:
    """Класс-помощник для работы с AutoCAD"""

    def __init__(self):
        """Подключение к AutoCAD"""
        pythoncom.CoInitialize()

        try:
            print("Подключение к AutoCAD...")
            self.acad = comtypes.client.GetActiveObject(
                "AutoCAD.Application.24")
            self.acad.Visible = True
            time.sleep(1)

            print(f"✓ Подключено к AutoCAD")

        except Exception as e:
            raise ConnectionError(f"Не удалось подключиться: {e}")

    def run_lisp_script(self, log_callback: Callable = None) -> Tuple[
        bool, List[str], List[str], str]:
        """
        Загрузить и выполнить LISP скрипт, вернуть результаты из файла
        """
        success_layouts = []
        error_layouts = []
        report_text = ""

        try:
            lisp_path = r"D:\Pycharm Projects\AutoCAD\RunPlot\process_layouts.lsp"
            results_file = os.path.join(os.environ.get('TEMP', 'C:\\Temp'),
                                        "acad_results.txt")

            # Удаляем старый файл результатов, если существует
            if os.path.exists(results_file):
                try:
                    os.remove(results_file)
                except:
                    pass

            if not os.path.exists(lisp_path):
                if log_callback:
                    log_callback(f"  ❌ LISP файл не найден: {lisp_path}")
                return False, [], [], ""

            if log_callback:
                log_callback("  🚀 Запуск LISP скрипта...")
                log_callback("  ⏳ Обработка листов...")

            # Загружаем и запускаем LISP
            lisp_path_fixed = lisp_path.replace('\\', '/')
            cmd = f'(load "{lisp_path_fixed}")\nPPGO\n'
            self.acad.ActiveDocument.SendCommand(cmd)

            # Ждем выполнения (увеличиваем время для обработки всех листов)
            time.sleep(15)

            # Читаем файл результатов с правильной кодировкой
            if os.path.exists(results_file):
                # Пробуем разные кодировки
                encodings = ['cp1251', 'utf-8', 'cp866', 'latin-1']

                for encoding in encodings:
                    try:
                        with open(results_file, 'r', encoding=encoding) as f:
                            report_text = f.read()
                        if log_callback:
                            log_callback(
                                f"  ✅ Файл прочитан (кодировка: {encoding})")
                        break
                    except UnicodeDecodeError:
                        continue
                    except Exception as e:
                        if log_callback:
                            log_callback(
                                f"  ⚠️ Ошибка чтения ({encoding}): {e}")
                        continue

                if not report_text:
                    # Если ничего не прочиталось, читаем в бинарном режиме
                    with open(results_file, 'rb') as f:
                        raw = f.read()
                        report_text = raw.decode('cp1251', errors='replace')
                    if log_callback:
                        log_callback("  ✅ Файл прочитан (бинарный режим)")

                # Парсим файл для получения списков
                in_success = False
                in_errors = False

                for line in report_text.split('\n'):
                    line = line.strip()
                    if line == "=== УСПЕШНЫЕ ЛИСТЫ ===":
                        in_success = True
                        in_errors = False
                    elif line == "=== ЛИСТЫ С ОШИБКАМИ ===":
                        in_success = False
                        in_errors = True
                    elif in_success and line.startswith('•'):
                        success_layouts.append(line.replace('•', '').strip())
                    elif in_errors and line.startswith('•'):
                        error_layouts.append(line.replace('•', '').strip())

                if log_callback:
                    if success_layouts:
                        log_callback(
                            f"  📊 Найдено успешных листов: {len(success_layouts)}")
                    if error_layouts:
                        log_callback(
                            f"  ⚠️ Найдено листов с ошибками: {len(error_layouts)}")
            else:
                if log_callback:
                    log_callback("  ⚠️ Файл с результатами не найден")

            return True, success_layouts, error_layouts, report_text

        except Exception as e:
            if log_callback:
                log_callback(f"  ❌ Ошибка: {e}")
            return False, [], [], ""

    def __del__(self):
        """Очистка COM"""
        try:
            pythoncom.CoUninitialize()
        except:
            pass