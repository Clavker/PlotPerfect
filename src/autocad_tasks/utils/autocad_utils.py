"""AutoCAD helper utilities - LISP with result file"""

import time
import pythoncom
import os
import sys
from typing import List, Tuple, Callable
import comtypes.client


def get_base_path():
    """Получить базовый путь для файлов (работает и для exe и для скрипта)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class AutoCADHelper:
    """Класс-помощник для работы с AutoCAD"""

    def __init__(self):
        """Подключение к AutoCAD"""
        pythoncom.CoInitialize()
        self.base_path = get_base_path()

        try:
            print("Подключение к AutoCAD...")
            self.acad = comtypes.client.GetActiveObject("AutoCAD.Application.24")
            self.acad.Visible = True
            time.sleep(1)

            print(f"✓ Подключено к AutoCAD")

        except Exception as e:
            raise ConnectionError(f"Не удалось подключиться: {e}")

    def run_lisp_script(self, lisp_filename: str, command: str, log_callback: Callable = None) -> Tuple[bool, List[str], List[str], str]:
        """
        Загрузить и выполнить LISP скрипт, вернуть результаты из файла
        """
        success_layouts = []
        error_layouts = []
        report_text = ""

        try:
            lisp_path = os.path.join(self.base_path, lisp_filename)
            results_file = os.path.join(os.environ.get('TEMP', 'C:\\Temp'), "acad_results.txt")

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
                log_callback(f"  🚀 Запуск LISP скрипта {lisp_filename}...")

            lisp_path_fixed = lisp_path.replace('\\', '/')
            cmd = f'(load "{lisp_path_fixed}")\n{command}\n'
            self.acad.ActiveDocument.SendCommand(cmd)

            # Ждем выполнения
            time.sleep(15)

            # Пытаемся прочитать файл результатов ТОЛЬКО если он существует
            if os.path.exists(results_file):
                encodings = ['cp1251', 'utf-8', 'cp866', 'latin-1']
                for encoding in encodings:
                    try:
                        with open(results_file, 'r', encoding=encoding) as f:
                            report_text = f.read()
                        break
                    except:
                        continue

                if not report_text:
                    with open(results_file, 'rb') as f:
                        raw = f.read()
                        report_text = raw.decode('cp1251', errors='replace')

                # Парсим результаты ТОЛЬКО если есть информация о листах
                if "УСПЕШНЫЕ ЛИСТЫ" in report_text or "Успешные листы" in report_text:
                    in_success = False
                    in_errors = False

                    for line in report_text.split('\n'):
                        line = line.strip()
                        if "УСПЕШНЫЕ ЛИСТЫ" in line or "Успешные листы" in line:
                            in_success = True
                            in_errors = False
                        elif "ЛИСТЫ С ОШИБКАМИ" in line or "Ошибки" in line:
                            in_success = False
                            in_errors = True
                        elif line.startswith('•') or line.startswith('-'):
                            item = line.replace('•', '').replace('-', '').strip()
                            if item and item != "Model" and "---" not in item:
                                if in_success:
                                    success_layouts.append(item)
                                elif in_errors:
                                    error_layouts.append(item)

                if log_callback:
                    if success_layouts:
                        log_callback(f"  📊 Найдено успешных листов: {len(success_layouts)}")
                    if error_layouts:
                        log_callback(f"  ⚠️ Найдено листов с ошибками: {len(error_layouts)}")

            # Для задач, которые не создают файл результатов, просто выводим сообщение
            if not report_text:
                if log_callback:
                    log_callback(f"  ✅ Команда выполнена (подробности в AutoCAD, F2)")

            # Всегда возвращаем success=True, если команда выполнилась
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