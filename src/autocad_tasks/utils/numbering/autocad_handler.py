"""Модуль для работы с AutoCAD COM объектами."""

import pythoncom
import win32com.client
from typing import List, Optional, Tuple, Any


class AutoCADHandler:
    """Класс для управления подключением к AutoCAD и выбором объектов."""

    def __init__(self):
        """Инициализация обработчика AutoCAD."""
        self.acad: Optional[Any] = None
        self.text_objects: List[Any] = []

    def connect(self) -> bool:
        """Подключение к запущенному экземпляру AutoCAD."""
        try:
            # Инициализируем COM для текущего потока
            pythoncom.CoInitialize()
            self.acad = win32com.client.GetActiveObject("AutoCAD.Application")
            print(f"Подключено к AutoCAD версии: {self.acad.Version}")
            return True
        except Exception as e:
            print(f"Ошибка подключения к AutoCAD: {e}")
            return False

    def disconnect(self) -> None:
        """Отключение от AutoCAD."""
        try:
            if self.acad:
                self.acad = None
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"Ошибка при отключении: {e}")

    def select_text_objects(self) -> List[Any]:
        """Выбор текстовых объектов в AutoCAD."""
        if not self.acad:
            if not self.connect():
                return []

        try:
            doc = self.acad.ActiveDocument

            # Удаляем временную выборку, если она существует
            try:
                selection = doc.SelectionSets.Item("TempSelectionNumbering")
                selection.Delete()
            except:
                pass

            # Создаем новую выборку
            selection = doc.SelectionSets.Add("TempSelectionNumbering")

            # Запрашиваем выбор объектов на экране
            selection.SelectOnScreen()

            # Собираем текстовые объекты
            self.text_objects = []
            for obj in selection:
                obj_type = str(obj.ObjectName).lower()
                if 'text' in obj_type or 'mtext' in obj_type or 'attribute' in obj_type:
                    self.text_objects.append(obj)

            selection.Delete()
            return self.text_objects

        except Exception as e:
            print(f"Ошибка при выборе объектов: {e}")
            # Если произошла ошибка, возможно COM упал
            self.acad = None
            return []

    def get_text_coordinates(self, obj: Any) -> Tuple[float, float]:
        """Получение координат текстового объекта."""
        try:
            if hasattr(obj, 'InsertionPoint'):
                point = obj.InsertionPoint
            elif hasattr(obj, 'TextPosition'):
                point = obj.TextPosition
            else:
                return (0.0, 0.0)
            return (float(point[0]), float(point[1]))
        except:
            return (0.0, 0.0)

    def update_text(self, obj: Any, new_text: str) -> bool:
        """Обновление текста объекта."""
        try:
            if hasattr(obj, 'TextString'):
                obj.TextString = new_text
            elif hasattr(obj, 'textString'):
                obj.textString = new_text
            return True
        except Exception as e:
            print(f"Ошибка обновления текста: {e}")
            return False

    def update_display(self) -> None:
        """Обновление отображения в AutoCAD."""
        try:
            if self.acad:
                self.acad.Update()
                self.acad.ZoomExtents()
        except:
            pass