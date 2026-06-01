"""Модуль для управления нумерацией текстовых объектов."""

from typing import List, Tuple, Any, Callable, Optional
from .autocad_handler import AutoCADHandler


class NumberingEngine:
    """Движок для автоматической нумерации."""

    def __init__(self, handler: AutoCADHandler):
        """Инициализация движка нумерации."""
        self.handler = handler
        self.text_objects: List[Any] = []

    def set_objects(self, objects: List[Any]) -> None:
        """Установка списка объектов для обработки."""
        self.text_objects = objects

    def sort_objects(self, sort_direction: str = "left_to_right") -> List[
        Tuple[Any, float, float]]:
        """
        Сортировка объектов по координатам.

        Направления:
        - left_to_right: слева направо (по X возрастанию)
        - right_to_left: справа налево (по X убыванию)
        - bottom_to_top: снизу вверх (по Y возрастанию)
        - top_to_bottom: сверху вниз (по Y убыванию)
        """
        objects_with_coords = []
        for obj in self.text_objects:
            x, y = self.handler.get_text_coordinates(obj)
            objects_with_coords.append((obj, x, y))

        if sort_direction == "left_to_right":
            objects_with_coords.sort(
                key=lambda item: item[1])  # по X возрастанию
        elif sort_direction == "right_to_left":
            objects_with_coords.sort(
                key=lambda item: -item[1])  # по X убыванию
        elif sort_direction == "bottom_to_top":
            objects_with_coords.sort(
                key=lambda item: item[2])  # по Y возрастанию
        elif sort_direction == "top_to_bottom":
            objects_with_coords.sort(
                key=lambda item: -item[2])  # по Y убыванию
        else:
            objects_with_coords.sort(
                key=lambda item: item[1])  # по умолчанию слева направо

        return objects_with_coords

    @staticmethod
    def format_number(number: int, format_type: str = "1,2,3") -> str:
        """Форматирование номера."""
        if format_type == "01,02,03":
            return f"{number:02d}"
        elif format_type == "001,002,003":
            return f"{number:03d}"
        else:  # "1,2,3"
            return str(number)

    def apply_numbering(
            self,
            prefix: str = "",
            start_number: int = 1,
            suffix: str = "",
            sort_direction: str = "left_to_right",
            format_type: str = "1,2,3",
            progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> int:
        """Применение нумерации к объектам."""
        if not self.text_objects:
            return 0

        # Сортируем объекты
        sorted_objects = self.sort_objects(sort_direction)

        # Применяем нумерацию
        success_count = 0
        total = len(sorted_objects)

        for i, (obj, x, y) in enumerate(sorted_objects):
            current_number = start_number + i
            formatted_number = self.format_number(current_number, format_type)
            new_text = f"{prefix}{formatted_number}{suffix}"

            if self.handler.update_text(obj, new_text):
                success_count += 1

            if progress_callback:
                progress_callback(i + 1, total)

        # Обновляем отображение
        self.handler.update_display()

        return success_count

    def preview_numbering(
            self,
            prefix: str = "",
            start_number: int = 1,
            suffix: str = "",
            sort_direction: str = "left_to_right",
            format_type: str = "1,2,3"
    ) -> List[Tuple[str, Tuple[float, float]]]:
        """Предпросмотр нумерации без применения."""
        if not self.text_objects:
            return []

        sorted_objects = self.sort_objects(sort_direction)
        preview = []

        for i, (obj, x, y) in enumerate(sorted_objects):
            current_number = start_number + i
            formatted_number = self.format_number(current_number, format_type)
            new_text = f"{prefix}{formatted_number}{suffix}"
            preview.append((new_text, (x, y)))

        return preview