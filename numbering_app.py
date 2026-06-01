"""Автонумерация текстовых объектов в AutoCAD (comtypes версия)"""

import tkinter as tk
from tkinter import ttk, messagebox
import pythoncom
import comtypes.client
import threading
from typing import List, Tuple


class AutoCADHandler:
    def __init__(self):
        self.acad = None
        self.doc = None

    def connect(self) -> bool:
        try:
            pythoncom.CoInitialize()
            progids = ["AutoCAD.Application.24", "AutoCAD.Application"]
            for progid in progids:
                try:
                    self.acad = comtypes.client.GetActiveObject(progid)
                    break
                except:
                    continue
            if self.acad:
                self.acad.Visible = True
                self.doc = self.acad.ActiveDocument
                return True
            return False
        except Exception as e:
            print(f"Ошибка подключения: {e}")
            return False

    def select_text_objects(self) -> List:
        try:
            # Удаляем старую выборку
            try:
                sel = self.doc.SelectionSets.Item("TempNumbering")
                sel.Delete()
            except:
                pass

            sel = self.doc.SelectionSets.Add("TempNumbering")
            sel.SelectOnScreen()

            objects = []
            for i in range(sel.Count):
                obj = sel.Item(i)
                obj_name = str(obj.ObjectName).lower()
                if 'text' in obj_name or 'mtext' in obj_name:
                    # Сохраняем ObjectID для последующего получения объекта
                    objects.append(obj.ObjectID)

            sel.Delete()
            return objects
        except Exception as e:
            print(f"Ошибка выбора: {e}")
            return []

    def get_object_by_id(self, obj_id):
        """Получить объект по ObjectID"""
        try:
            return self.doc.ObjectIDToObject(obj_id)
        except:
            return None

    def get_coords(self, obj) -> Tuple[float, float]:
        try:
            if hasattr(obj, 'InsertionPoint'):
                p = obj.InsertionPoint
                return (float(p[0]), float(p[1]))
        except:
            pass
        return (0.0, 0.0)

    def disconnect(self):
        try:
            self.acad = None
            pythoncom.CoUninitialize()
        except:
            pass


class NumberingApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Автонумерация AutoCAD")
        self.root.geometry("580x650")
        self.root.resizable(False, False)

        self.handler = AutoCADHandler()
        self.text_ids = []  # Храним ObjectID

        self.setup_ui()

    def setup_ui(self):
        title = tk.Label(self.root,
                         text="Автонумерация текстовых объектов AutoCAD",
                         font=("Arial", 12, "bold"))
        title.pack(pady=10)

        self.btn_select = tk.Button(self.root,
                                    text="1. Выбрать текстовые объекты в AutoCAD",
                                    command=self.select_objects,
                                    bg="lightblue", height=2)
        self.btn_select.pack(pady=5, padx=20, fill="x")

        self.status_label = tk.Label(self.root, text="Статус: Не выбрано",
                                     fg="red")
        self.status_label.pack(pady=5)

        settings_frame = tk.LabelFrame(self.root, text="Настройки нумерации",
                                       padx=10, pady=10)
        settings_frame.pack(pady=10, padx=20, fill="x")

        # Префикс
        tk.Label(settings_frame, text="Префикс:").grid(row=0, column=0,
                                                       sticky="w", pady=5)
        self.prefix_entry = tk.Entry(settings_frame, width=20)
        self.prefix_entry.grid(row=0, column=1, padx=10, pady=5)
        self.prefix_entry.insert(0, "№")

        # Начальное число
        tk.Label(settings_frame, text="Начальное число:").grid(row=1, column=0,
                                                               sticky="w",
                                                               pady=5)
        self.start_entry = tk.Entry(settings_frame, width=20)
        self.start_entry.grid(row=1, column=1, padx=10, pady=5)
        self.start_entry.insert(0, "1")

        # Суффикс
        tk.Label(settings_frame, text="Суффикс:").grid(row=2, column=0,
                                                       sticky="w", pady=5)
        self.suffix_entry = tk.Entry(settings_frame, width=20)
        self.suffix_entry.grid(row=2, column=1, padx=10, pady=5)

        # Формат
        tk.Label(settings_frame, text="Формат номера:").grid(row=3, column=0,
                                                             sticky="w",
                                                             pady=5)
        self.format_var = tk.StringVar(value="1,2,3")
        format_combo = ttk.Combobox(settings_frame,
                                    textvariable=self.format_var,
                                    values=["1,2,3", "01,02,03",
                                            "001,002,003"], width=18)
        format_combo.grid(row=3, column=1, padx=10, pady=5)

        # Сортировка
        tk.Label(settings_frame, text="Сортировка:").grid(row=4, column=0,
                                                          sticky="w", pady=5)
        self.sort_var = tk.StringVar(value="left_to_right")
        sort_frame = tk.Frame(settings_frame)
        sort_frame.grid(row=4, column=1, sticky="w")
        tk.Radiobutton(sort_frame, text="Слева направо",
                       variable=self.sort_var,
                       value="left_to_right").pack(anchor="w")
        tk.Radiobutton(sort_frame, text="Справа налево",
                       variable=self.sort_var,
                       value="right_to_left").pack(anchor="w")
        tk.Radiobutton(sort_frame, text="Снизу вверх", variable=self.sort_var,
                       value="bottom_to_top").pack(anchor="w")
        tk.Radiobutton(sort_frame, text="Сверху вниз", variable=self.sort_var,
                       value="top_to_bottom").pack(anchor="w")

        self.btn_apply = tk.Button(self.root, text="2. Применить нумерацию",
                                   command=self.apply_numbering,
                                   state="disabled",
                                   bg="lightgreen", height=2)
        self.btn_apply.pack(pady=10, padx=20, fill="x")

        inst_frame = tk.LabelFrame(self.root, text="Инструкция", padx=10,
                                   pady=5)
        inst_frame.pack(pady=10, padx=20, fill="x")
        instruction = """1. Запустите AutoCAD
2. Нажмите "Выбрать текстовые объекты"
3. Выделите нужные тексты в AutoCAD
4. Настройте параметры
5. Нажмите "Применить нумерацию"""
        tk.Label(inst_frame, text=instruction, justify="left",
                 fg="gray").pack()

    def select_objects(self):
        self.btn_select.config(state='disabled')
        self.status_label.config(text="Статус: Подключение к AutoCAD...",
                                 fg="orange")
        self.root.update()

        def do_select():
            try:
                if not self.handler.connect():
                    self.root.after(0, lambda: self.select_error(
                        "Не удалось подключиться к AutoCAD"))
                    return

                self.root.after(0, lambda: self.status_label.config(
                    text="Статус: Выберите объекты в AutoCAD...", fg="orange"))

                obj_ids = self.handler.select_text_objects()

                if obj_ids:
                    self.text_ids = obj_ids
                    self.root.after(0,
                                    lambda: self.select_success(len(obj_ids)))
                else:
                    self.root.after(0, lambda: self.select_error(
                        "Не выбрано ни одного текстового объекта"))
            except Exception as e:
                self.root.after(0, lambda: self.select_error(str(e)))

        threading.Thread(target=do_select, daemon=True).start()

    def select_success(self, count):
        self.status_label.config(text=f"Статус: Выбрано {count} объектов",
                                 fg="green")
        self.btn_apply.config(state='normal')
        self.btn_select.config(state='normal')
        messagebox.showinfo("Успех", f"Выбрано {count} текстовых объектов")

    def select_error(self, msg):
        self.status_label.config(text="Статус: Ошибка", fg="red")
        self.btn_select.config(state='normal')
        messagebox.showerror("Ошибка", msg)

    def apply_numbering(self):
        if not self.text_ids:
            messagebox.showwarning("Ошибка", "Сначала выберите объекты!")
            return

        try:
            start_num = int(self.start_entry.get())
        except ValueError:
            messagebox.showerror("Ошибка",
                                 "Начальное число должно быть целым!")
            return

        prefix = self.prefix_entry.get()
        suffix = self.suffix_entry.get()
        sort_dir = self.sort_var.get()
        format_type = self.format_var.get()

        # Подключаемся к AutoCAD для обновления
        if not self.handler.connect():
            messagebox.showerror("Ошибка",
                                 "Не удалось подключиться к AutoCAD para обновления")
            return

        # Получаем объекты по ObjectID
        objects = []
        for obj_id in self.text_ids:
            obj = self.handler.get_object_by_id(obj_id)
            if obj:
                objects.append(obj)

        if not objects:
            messagebox.showerror("Ошибка",
                                 "Не удалось найти выбранные объекты в чертеже")
            self.handler.disconnect()
            return

        # Сортируем
        objects_with_coords = []
        for obj in objects:
            try:
                x, y = self.handler.get_coords(obj)
                objects_with_coords.append((obj, x, y))
            except:
                pass

        if sort_dir == "left_to_right":
            objects_with_coords.sort(key=lambda item: item[1])
        elif sort_dir == "right_to_left":
            objects_with_coords.sort(key=lambda item: -item[1])
        elif sort_dir == "bottom_to_top":
            objects_with_coords.sort(key=lambda item: item[2])
        elif sort_dir == "top_to_bottom":
            objects_with_coords.sort(key=lambda item: -item[2])

        # Применяем нумерацию
        success_count = 0
        for i, (obj, x, y) in enumerate(objects_with_coords):
            try:
                num = start_num + i
                if format_type == "01,02,03":
                    num_str = f"{num:02d}"
                elif format_type == "001,002,003":
                    num_str = f"{num:03d}"
                else:
                    num_str = str(num)
                new_text = f"{prefix}{num_str}{suffix}"

                obj.TextString = new_text
                success_count += 1

            except Exception as e:
                print(f"Ошибка обновления: {e}")

        # Отключаемся
        self.handler.disconnect()

        if success_count > 0:
            messagebox.showinfo("Успех",
                                f"Обновлено {success_count} из {len(objects)} объектов")
            self.status_label.config(
                text=f"Статус: Обновлено {success_count} объектов", fg="blue")
        else:
            messagebox.showerror("Ошибка",
                                 "Не удалось обновить ни одного объекта!")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = NumberingApp()
    app.run()