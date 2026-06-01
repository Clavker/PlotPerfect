"""Модуль для графического интерфейса автонумерации."""

import tkinter as tk
from tkinter import ttk, messagebox
from .autocad_handler import AutoCADHandler
from .numbering_engine import NumberingEngine


class NumberingWindow:
    """Окно автонумерации (вызывается как подпрограмма)."""

    def __init__(self, parent=None):
        """Инициализация окна."""
        self.parent = parent
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("Автонумерация в AutoCAD")
        self.window.geometry("600x700")
        self.window.resizable(True, True)

        self.handler = AutoCADHandler()
        self.engine = NumberingEngine(self.handler)
        self.setup_ui()

        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        """Обработка закрытия окна."""
        self.handler.disconnect()
        self.window.destroy()

    def run(self):
        """Запуск окна."""
        self.window.mainloop()

    def setup_ui(self):
        """Настройка пользовательского интерфейса."""
        # Основной контейнер
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = ttk.Label(
            main_frame,
            text="Автонумерация текстовых объектов AutoCAD",
            font=("Arial", 12, "bold")
        )
        title_label.pack(pady=10)

        # Кнопка выбора объектов
        self.select_btn = ttk.Button(
            main_frame,
            text="1. Выбрать текстовые объекты в AutoCAD",
            command=self.select_objects
        )
        self.select_btn.pack(pady=5, fill=tk.X)

        # Статус выбора
        self.status_label = ttk.Label(
            main_frame,
            text="Статус: Не выбрано",
            foreground="red"
        )
        self.status_label.pack(pady=5)

        # Разделитель
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)

        # Настройки нумерации
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки нумерации",
                                        padding="10")
        settings_frame.pack(fill=tk.X, pady=5)

        # Префикс
        ttk.Label(settings_frame, text="Префикс:").grid(row=0, column=0,
                                                        sticky=tk.W, pady=2)
        self.prefix_var = tk.StringVar(value="№")
        self.prefix_entry = ttk.Entry(settings_frame,
                                      textvariable=self.prefix_var, width=25)
        self.prefix_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2,
                               padx=5)

        # Начальное число
        ttk.Label(settings_frame, text="Начальное число:").grid(row=1,
                                                                column=0,
                                                                sticky=tk.W,
                                                                pady=2)
        self.start_var = tk.StringVar(value="1")
        self.start_entry = ttk.Entry(settings_frame,
                                     textvariable=self.start_var, width=25)
        self.start_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2,
                              padx=5)

        # Суффикс
        ttk.Label(settings_frame, text="Суффикс:").grid(row=2, column=0,
                                                        sticky=tk.W, pady=2)
        self.suffix_var = tk.StringVar(value="")
        self.suffix_entry = ttk.Entry(settings_frame,
                                      textvariable=self.suffix_var, width=25)
        self.suffix_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2,
                               padx=5)

        # Формат номера
        ttk.Label(settings_frame, text="Формат номера:").grid(row=3, column=0,
                                                              sticky=tk.W,
                                                              pady=2)
        self.format_var = tk.StringVar(value="1,2,3")
        format_combo = ttk.Combobox(
            settings_frame,
            textvariable=self.format_var,
            values=["1,2,3", "01,02,03", "001,002,003"],
            width=23
        )
        format_combo.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=2, padx=5)

        settings_frame.columnconfigure(1, weight=1)

        # Порядок сортировки
        sort_frame = ttk.LabelFrame(main_frame, text="Порядок сортировки",
                                    padding="10")
        sort_frame.pack(fill=tk.X, pady=5)

        self.sort_var = tk.StringVar(value="left_to_right")

        ttk.Radiobutton(sort_frame, text="Слева направо (по X ↑)",
                        variable=self.sort_var, value="left_to_right").pack(
            anchor=tk.W)
        ttk.Radiobutton(sort_frame, text="Справа налево (по X ↓)",
                        variable=self.sort_var, value="right_to_left").pack(
            anchor=tk.W)
        ttk.Radiobutton(sort_frame, text="Снизу вверх (по Y ↑)",
                        variable=self.sort_var, value="bottom_to_top").pack(
            anchor=tk.W)
        ttk.Radiobutton(sort_frame, text="Сверху вниз (по Y ↓)",
                        variable=self.sort_var, value="top_to_bottom").pack(
            anchor=tk.W)

        # Прогресс-бар
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=5)

        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)

        self.preview_btn = ttk.Button(
            button_frame,
            text="Предпросмотр",
            command=self.preview_numbering,
            state='disabled'
        )
        self.preview_btn.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.apply_btn = ttk.Button(
            button_frame,
            text="2. Применить нумерацию",
            command=self.apply_numbering,
            state='disabled'
        )
        self.apply_btn.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)

        # Текстовое поле для предпросмотра
        preview_frame = ttk.LabelFrame(main_frame, text="Предпросмотр",
                                       padding="5")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.preview_text = tk.Text(preview_frame, height=12, wrap=tk.WORD)
        self.preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL,
                                  command=self.preview_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.preview_text.configure(yscrollcommand=scrollbar.set)

        # Инструкция
        self.create_instruction(main_frame)

    def create_instruction(self, parent):
        """Создание инструкции."""
        instruction = """Инструкция:
1. Убедитесь, что AutoCAD запущен
2. Нажмите кнопку "Выбрать текстовые объекты"
3. Выделите нужные текстовые объекты в AutoCAD
4. Настройте параметры нумерации
5. Нажмите "Предпросмотр" для проверки
6. Нажмите "Применить нумерацию" для выполнения"""

        instruction_frame = ttk.LabelFrame(parent, text="Инструкция",
                                           padding="5")
        instruction_frame.pack(fill=tk.X, pady=10)

        instruction_label = ttk.Label(
            instruction_frame,
            text=instruction,
            foreground="gray",
            justify=tk.LEFT
        )
        instruction_label.pack(anchor=tk.W)

    def select_objects(self):
        """Выбор текстовых объектов."""
        try:
            objects = self.handler.select_text_objects()
            if objects:
                self.engine.set_objects(objects)
                self.status_label.config(
                    text=f"Статус: Выбрано {len(objects)} объектов",
                    foreground="green"
                )
                self.preview_btn.config(state='normal')
                self.apply_btn.config(state='normal')
                messagebox.showinfo(
                    "Успех",
                    f"Выбрано {len(objects)} текстовых объектов"
                )
                self.preview_numbering()
            else:
                self.status_label.config(
                    text="Статус: Объекты не выбраны",
                    foreground="red"
                )
                self.preview_btn.config(state='disabled')
                self.apply_btn.config(state='disabled')
                messagebox.showwarning(
                    "Предупреждение",
                    "Не выбрано ни одного текстового объекта!"
                )
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при выборе объектов: {e}")

    def preview_numbering(self):
        """Предпросмотр нумерации."""
        try:
            start_number = int(self.start_var.get())
        except ValueError:
            messagebox.showerror("Ошибка",
                                 "Начальное число должно быть целым числом!")
            return

        prefix = self.prefix_var.get()
        suffix = self.suffix_var.get()
        sort_direction = self.sort_var.get()
        format_type = self.format_var.get()

        preview = self.engine.preview_numbering(
            prefix=prefix,
            start_number=start_number,
            suffix=suffix,
            sort_direction=sort_direction,
            format_type=format_type
        )

        self.preview_text.delete(1.0, tk.END)

        if preview:
            self.preview_text.insert(tk.END, "Результат нумерации:\n")
            self.preview_text.insert(tk.END, "-" * 60 + "\n")

            direction_names = {
                "left_to_right": "Слева направо",
                "right_to_left": "Справа налево",
                "bottom_to_top": "Снизу вверх",
                "top_to_bottom": "Сверху вниз"
            }
            dir_name = direction_names.get(sort_direction, sort_direction)
            self.preview_text.insert(tk.END, f"Сортировка: {dir_name}\n")
            self.preview_text.insert(tk.END, "-" * 60 + "\n")

            for i, (text, (x, y)) in enumerate(preview, 1):
                self.preview_text.insert(
                    tk.END,
                    f"{i:3}. Текст: '{text}' | X={x:8.1f} | Y={y:8.1f}\n"
                )
        else:
            self.preview_text.insert(tk.END, "Нет объектов для предпросмотра")

    def apply_numbering(self):
        """Применение нумерации."""
        if not self.engine.text_objects:
            messagebox.showwarning(
                "Предупреждение",
                "Сначала выберите текстовые объекты!"
            )
            return

        try:
            start_number = int(self.start_var.get())
        except ValueError:
            messagebox.showerror("Ошибка",
                                 "Начальное число должно быть целым числом!")
            return

        prefix = self.prefix_var.get()
        suffix = self.suffix_var.get()
        sort_direction = self.sort_var.get()
        format_type = self.format_var.get()

        self.progress['value'] = 0

        def update_progress(current, total):
            self.progress['value'] = (current / total) * 100
            self.window.update_idletasks()

        success_count = self.engine.apply_numbering(
            prefix=prefix,
            start_number=start_number,
            suffix=suffix,
            sort_direction=sort_direction,
            format_type=format_type,
            progress_callback=update_progress
        )

        self.progress['value'] = 100

        if success_count > 0:
            self.status_label.config(
                text=f"Статус: Обновлено {success_count} объектов",
                foreground="blue"
            )
            messagebox.showinfo(
                "Успех",
                f"Нумерация успешно применена!\n"
                f"Обновлено {success_count} из {len(self.engine.text_objects)} объектов."
            )
            self.preview_numbering()
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось обновить ни одного объекта!")