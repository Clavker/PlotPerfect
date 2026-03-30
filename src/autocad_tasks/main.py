"""Main GUI application for AutoCAD Tasks"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from autocad_tasks.tasks.task1 import Task1
from autocad_tasks.tasks.task2 import Task2
from autocad_tasks.tasks.task3 import Task3


class AutoCADTasksApp:
    """Главное окно приложения"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("AutoCAD Task Manager")
        self.root.geometry("800x800")  # Увеличим высоту для третьей задачи

        self.tasks_status = {}
        self.task_instances = {}
        self.task_frames = {}
        self.task_results = {}
        self.task_completed = {}

        self.create_widgets()

    def create_widgets(self):
        # Верхняя панель с кнопкой Auto
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        self.auto_btn = ttk.Button(
            top_frame, text="🚀 AUTO", command=self.run_all_tasks, width=15
        )
        self.auto_btn.pack(side=tk.LEFT, padx=5)

        info_label = ttk.Label(
            top_frame,
            text="Выберите задачи для автоматического выполнения",
            foreground="gray",
        )
        info_label.pack(side=tk.LEFT, padx=10)

        # Статус оптимизации документа
        self.optimization_status = ttk.Label(
            top_frame,
            text="❌ Документ не оптимизирован для печати",
            font=("Arial", 9, "bold"),
            foreground="red"
        )
        self.optimization_status.pack(side=tk.RIGHT, padx=10)

        ttk.Separator(self.root, orient="horizontal").pack(fill=tk.X, padx=10,
                                                           pady=5)

        tasks_label = ttk.Label(
            self.root, text="📋 Доступные задачи:", font=("Arial", 10, "bold")
        )
        tasks_label.pack(anchor=tk.W, padx=10, pady=(10, 5))

        self.tasks_frame = ttk.Frame(self.root)
        self.tasks_frame.pack(fill=tk.X, padx=10, pady=5)

        # Добавление задачи 1
        self.add_task_widget(
            "Задача 1: Выставить рамки для печати",
            self.run_task1,
            "Отправляет номер листа (1.06, 3.51) на задний план, рамку чертежа (8, 5) на передний",
        )

        # Добавление задачи 2
        self.add_task_widget(
            "Задача 2: Установка переменных для печати",
            self.run_task2,
            "Устанавливает PDFFRAME=2, IMAGEFRAME=2, контуры маскировки без печати",
        )

        # Добавление задачи 3
        # Добавление задачи 3
        self.add_task_widget(
            "Задача 3: Очистка чертежа от неиспользуемых объектов",
            self.run_task3,
            "Применение команды PURGE - удаление всех неиспользуемых элементов (блоки, слои, стили и т.д.)",
        )

        ttk.Separator(self.root, orient="horizontal").pack(fill=tk.X, padx=10,
                                                           pady=5)

        # Панель для лога с кнопкой очистки
        log_panel = ttk.Frame(self.root)
        log_panel.pack(fill=tk.X, padx=10, pady=(10, 5))

        log_label = ttk.Label(
            log_panel, text="📝 Лог выполнения:", font=("Arial", 9, "bold")
        )
        log_label.pack(side=tk.LEFT)

        # Кнопка очистки лога
        clear_btn = ttk.Button(
            log_panel,
            text="🗑️ Очистить лог",
            command=self.clear_log,
            width=14
        )
        clear_btn.pack(side=tk.RIGHT, padx=(0, 0))

        # Создаем текстовое поле с прокруткой
        text_frame = ttk.Frame(self.root)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(
            text_frame,
            height=15,
            width=85,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def add_task_widget(self, task_name: str, task_function,
                        description: str = ""):
        """Добавляет виджет для задачи с отображением статуса"""
        frame = ttk.Frame(self.tasks_frame)
        frame.pack(fill=tk.X, pady=5)
        self.task_frames[task_name] = frame

        # Галочка для Auto режима
        var = tk.BooleanVar(value=True)
        self.tasks_status[task_name] = var

        chk = ttk.Checkbutton(frame, variable=var)
        chk.pack(side=tk.LEFT, padx=5)

        # Название задачи
        label = ttk.Label(frame, text=task_name, font=("Arial", 9, "bold"))
        label.pack(side=tk.LEFT, padx=5)

        # Кнопка запуска
        run_btn = ttk.Button(
            frame,
            text="▶ Запустить",
            command=lambda: self.run_single_task(task_function, task_name),
            width=12
        )
        run_btn.pack(side=tk.RIGHT, padx=5)

        # Метка статуса
        status_label = ttk.Label(frame, text="⚪ Не выполнялась",
                                 foreground="gray", width=15)
        status_label.pack(side=tk.RIGHT, padx=10)

        self.task_frames[task_name + "_status"] = status_label
        self.task_completed[task_name] = False

        # Описание
        if description:
            desc_label = ttk.Label(
                self.tasks_frame,
                text=f"  📌 {description}",
                foreground="blue",
                font=("Arial", 8),
                wraplength=700
            )
            desc_label.pack(anchor=tk.W, padx=(30, 10), pady=(0, 5))

    def update_task_status(self, task_name: str, result):
        """Обновить статус задачи в интерфейсе"""
        status_key = task_name + "_status"
        if status_key in self.task_frames:
            if result.get("errors") and len(result.get("errors")) > 0:
                self.task_frames[status_key].config(
                    text="❌ Ошибка", foreground="red"
                )
                self.task_completed[task_name] = False
            elif result.get("success") and len(result.get("success")) > 0:
                self.task_frames[status_key].config(
                    text="✅ Выполнена", foreground="green"
                )
                self.task_completed[task_name] = True
            else:
                self.task_frames[status_key].config(
                    text="⚠️ Нет данных", foreground="orange"
                )
                self.task_completed[task_name] = False

        self.update_optimization_status()

    def update_optimization_status(self):
        """Обновить общий статус оптимизации документа"""
        all_completed = all(self.task_completed.values())

        if all_completed and self.task_completed:
            self.optimization_status.config(
                text="✅ Документ AutoCAD полностью оптимизирован для печати",
                foreground="green"
            )
        else:
            any_completed = any(self.task_completed.values())
            if any_completed:
                completed_count = sum(
                    1 for v in self.task_completed.values() if v)
                total_count = len(self.task_completed)
                self.optimization_status.config(
                    text=f"⚠️ Оптимизировано {completed_count}/{total_count} задач",
                    foreground="orange"
                )
            else:
                self.optimization_status.config(
                    text="❌ Документ не оптимизирован для печати",
                    foreground="red"
                )

    def reset_task_statuses(self, selected_tasks=None):
        """Сбросить статусы только для выбранных задач"""
        tasks_to_reset = selected_tasks if selected_tasks else self.task_completed.keys()

        for task_name in tasks_to_reset:
            if task_name in self.task_completed:
                status_key = task_name + "_status"
                if status_key in self.task_frames:
                    if not self.task_completed[task_name]:
                        self.task_frames[status_key].config(
                            text="⚪ Не выполнялась", foreground="gray"
                        )

    def clear_log(self):
        """Очистить лог выполнения"""
        self.log_text.delete(1.0, tk.END)
        self.log("📝 Лог очищен")
        self.log("=" * 60)

    def log(self, message: str):
        """Вывод сообщения в лог с автоскроллом"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def run_single_task(self, task_func, task_name: str):
        """Запуск одной задачи"""
        self.reset_task_statuses([task_name])

        self.log(f"\n{'█' * 60}")
        self.log(f"▶ ЗАПУСК: {task_name}")
        self.log(f"{'█' * 60}")

        if task_name == "Задача 1: Выставить рамки для печати":
            task = Task1(log_callback=self.log)
            result = task.run()
            self.task_results[task_name] = result
            self.update_task_status(task_name, result)
            self.show_task_report(task_name, result)
        elif task_name == "Задача 2: Установка переменных для печати":
            task = Task2(log_callback=self.log)
            result = task.run()
            self.task_results[task_name] = result
            self.update_task_status(task_name, result)
            self.show_task_report(task_name, result)
        elif task_name == "Задача 3: Очистка чертежа":
            task = Task3(log_callback=self.log)
            result = task.run()
            self.task_results[task_name] = result
            self.update_task_status(task_name, result)
            self.show_task_report(task_name, result)

    def run_all_tasks(self):
        """Запуск всех отмеченных задач"""
        selected_tasks = [name for name, var in self.tasks_status.items() if
                          var.get()]

        if not selected_tasks:
            messagebox.showwarning(
                "Нет задач", "Не выбрано ни одной задачи для выполнения"
            )
            return

        self.reset_task_statuses(selected_tasks)

        self.log("\n" + "█" * 60)
        self.log("🚀 ЗАПУСК ВСЕХ ЗАДАЧ")
        self.log("█" * 60)

        for task_name in selected_tasks:
            if task_name == "Задача 1: Выставить рамки для печати":
                task = Task1(log_callback=self.log)
                result = task.run()
                self.task_results[task_name] = result
                self.update_task_status(task_name, result)
            elif task_name == "Задача 2: Установка переменных для печати":
                task = Task2(log_callback=self.log)
                result = task.run()
                self.task_results[task_name] = result
                self.update_task_status(task_name, result)
            elif task_name == "Задача 3: Очистка чертежа":
                task = Task3(log_callback=self.log)
                result = task.run()
                self.task_results[task_name] = result
                self.update_task_status(task_name, result)

        self.show_final_report()

    def run_task1(self):
        """Выполнение задачи 1"""
        task = Task1(log_callback=self.log)
        result = task.run()
        self.task_results["Задача 1: Выставить рамки для печати"] = result
        self.update_task_status("Задача 1: Выставить рамки для печати", result)
        return result

    def run_task2(self):
        """Выполнение задачи 2"""
        task = Task2(log_callback=self.log)
        result = task.run()
        self.task_results["Задача 2: Установка переменных для печати"] = result
        self.update_task_status("Задача 2: Установка переменных для печати",
                                result)
        return result

    def run_task3(self):
        """Выполнение задачи 3"""
        task = Task3(log_callback=self.log)
        result = task.run()
        self.task_results["Задача 3: Очистка чертежа"] = result
        self.update_task_status("Задача 3: Очистка чертежа", result)
        return result

    def show_task_report(self, task_name: str, result):
        """Показать отчет по задаче"""
        if result["errors"]:
            messagebox.showwarning(
                "Завершено с ошибками",
                f"{task_name}\n\n✅ Успешно: {len(result['success'])}\n❌ Ошибки: {len(result['errors'])}\n\n📋 Подробности в логе",
            )
        else:
            messagebox.showinfo(
                "Успешно",
                f"{task_name}\n\n✅ Задача выполнена успешно!\n\nПодробности в логе"
            )

    def show_final_report(self):
        """Показать итоговый отчет"""
        messagebox.showinfo(
            "Итоговый отчет",
            "✅ Выполнение всех задач завершено!\n📋 Подробности смотрите в логе."
        )


def main():
    root = tk.Tk()
    app = AutoCADTasksApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()