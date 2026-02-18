# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
import os
import json

DATABASE_FILE = "production_database.xlsx"


def initialize_database():
    if not os.path.exists(DATABASE_FILE):
        wb = Workbook()
        materials_sheet = wb.active
        materials_sheet.title = "Materials"
        materials_sheet.append([
            "ID", "Марка", "Толщина", "Длина", "Ширина",
            "Количество штук", "Общая площадь", "Зарезервировано", "Доступно", "Дата добавления"
        ])
        orders_sheet = wb.create_sheet("Orders")
        orders_sheet.append(["ID заказа", "Название заказа", "Заказчик", "Дата создания", "Статус", "Примечания"])
        order_details_sheet = wb.create_sheet("OrderDetails")
        order_details_sheet.append(["ID", "ID заказа", "Название детали", "Количество", "Порезано", "Погнуто"])
        reservations_sheet = wb.create_sheet("Reservations")
        reservations_sheet.append(["ID резерва", "ID заказа", "ID детали", "Название детали", "ID материала", "Марка", "Толщина", "Длина", "Ширина", "Зарезервировано штук", "Списано", "Остаток к списанию", "Дата резерва"])
        writeoffs_sheet = wb.create_sheet("WriteOffs")
        writeoffs_sheet.append([
            "ID списания", "ID резерва", "ID заказа", "ID материала", "Марка", "Толщина", "Длина", "Ширина",
            "Количество", "Дата списания", "Комментарий"
        ])
        wb.save(DATABASE_FILE)
        print(f"База данных '{DATABASE_FILE}' создана!")


def load_data(sheet_name):
    try:
        df = pd.read_excel(DATABASE_FILE, sheet_name=sheet_name, engine='openpyxl')
        if df.empty:
            return df
        df = df.fillna("")
        return df
    except Exception as e:
        print(f"Ошибка загрузки данных из {sheet_name}: {e}")
        return pd.DataFrame()


def save_data(sheet_name, dataframe):
    try:
        book = load_workbook(DATABASE_FILE)
        if sheet_name in book.sheetnames:
            del book[sheet_name]
        sheet = book.create_sheet(sheet_name)
        for col_num, column_title in enumerate(dataframe.columns, 1):
            sheet.cell(row=1, column=col_num).value = str(column_title)
        for row_num, row_data in enumerate(dataframe.values, 2):
            for col_num, cell_value in enumerate(row_data, 1):
                sheet.cell(row=row_num, column=col_num).value = cell_value
        book.save(DATABASE_FILE)
        book.close()
    except Exception as e:
        print(f"Ошибка сохранения данных в {sheet_name}: {e}")
        messagebox.showerror("Ошибка", f"Невозможно сохранить данные: {e}")


class ProductionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Система учета производства")
        self.root.geometry("1400x800")
        self.root.configure(bg='#f0f0f0')

        # Инициализация переменных toggles
        self.materials_toggles = {}
        self.orders_toggles = {}
        self.reservations_toggles = {}
        self.balance_toggles = {}
        self.writeoffs_toggles = {}

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.materials_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.materials_frame, text='Материалы на складе')
        self.setup_materials_tab()

        self.orders_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.orders_frame, text='Заказы')
        self.setup_orders_tab()

        self.reservations_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.reservations_frame, text='Резервирование')
        self.setup_reservations_tab()

        self.writeoffs_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.writeoffs_frame, text='Списание материалов')
        self.setup_writeoffs_tab()

        self.balance_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.balance_frame, text='Баланс материалов')
        self.setup_balance_tab()

        # Загрузка настроек и обработчик закрытия
        self.load_toggle_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_filter_panel(self, parent_frame, tree_widget, columns_to_filter, refresh_callback):
        """Создание панели фильтрации для любой таблицы"""
        filter_frame = tk.LabelFrame(parent_frame, text="🔍 Фильтры", bg='#e8f4f8', font=("Arial", 10, "bold"))
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        filter_entries = {}
        row = 0
        col = 0
        max_cols = 4

        for column_name in columns_to_filter:
            filter_container = tk.Frame(filter_frame, bg='#e8f4f8')
            filter_container.grid(row=row, column=col, padx=5, pady=3, sticky='w')

            tk.Label(filter_container, text=f"{column_name}:", bg='#e8f4f8', font=("Arial", 9)).pack(side=tk.LEFT)

            entry = tk.Entry(filter_container, width=15, font=("Arial", 9))
            entry.pack(side=tk.LEFT, padx=5)

            filter_entries[column_name] = entry

            entry.bind('<KeyRelease>', lambda e, tree=tree_widget, filters=filter_entries, cb=refresh_callback:
            self.apply_filters(tree, filters, cb))

            col += 1
            if col >= max_cols:
                col = 0
                row += 1

        buttons_container = tk.Frame(filter_frame, bg='#e8f4f8')
        buttons_container.grid(row=row + 1, column=0, columnspan=max_cols, pady=5)

        tk.Button(buttons_container, text="🗑️ Очистить фильтры", bg='#95a5a6', fg='white',
                  font=("Arial", 9),
                  command=lambda: self.clear_filters(filter_entries, tree_widget, refresh_callback)).pack(side=tk.LEFT,
                                                                                                          padx=5)

        tk.Button(buttons_container, text="🔄 Обновить", bg='#3498db', fg='white',
                  font=("Arial", 9), command=refresh_callback).pack(side=tk.LEFT, padx=5)

        return filter_entries

    def apply_filters(self, tree, filter_entries, refresh_callback):
        """Применить фильтры к таблице"""
        active_filters = {}
        for col_name, entry in filter_entries.items():
            filter_text = entry.get().strip().lower()
            if filter_text:
                active_filters[col_name] = filter_text

        if not active_filters:
            refresh_callback()
            return

        all_items = []
        for item in tree.get_children():
            all_items.append(tree.item(item)['values'])

        for item in tree.get_children():
            tree.delete(item)

        columns = tree['columns']
        for item_values in all_items:
            match = True
            for col_name, filter_text in active_filters.items():
                try:
                    col_index = columns.index(col_name)
                    cell_value = str(item_values[col_index]).lower()
                    if filter_text not in cell_value:
                        match = False
                        break
                except (ValueError, IndexError):
                    continue

            if match:
                tree.insert("", "end", values=item_values)

    def clear_filters(self, filter_entries, tree, refresh_callback):
        """Очистить все фильтры"""
        for entry in filter_entries.values():
            entry.delete(0, tk.END)
        refresh_callback()

        def auto_resize_columns(self, tree):
            """Автоматическая подгонка ширины колонок по содержимому"""
            for col in tree["columns"]:
                max_width = len(col) * 10

                tree.column(col, width=max_width)
                tree.update_idletasks()

                for item in tree.get_children():
                    try:
                        col_index = tree["columns"].index(col)
                        cell_value = str(tree.item(item)['values'][col_index])
                        cell_width = len(cell_value) * 8 + 20
                        if cell_width > max_width:
                            max_width = cell_width
                    except:
                        pass

                max_width = min(max_width, 400)
                max_width = max(max_width, 80)

                tree.column(col, width=max_width)

                max_width = min(max_width, 400)
                max_width = max(max_width, 80)

                tree.column(col, width=max_width)

    def create_visibility_toggles(self, parent_frame, tree_widget, toggle_options, refresh_callback):
        """Создание переключателей видимости для таблицы"""
        toggles_frame = tk.Frame(parent_frame, bg='#fff9e6')
        toggles_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(toggles_frame, text="👁️ Отображение:", bg='#fff9e6', font=("Arial", 10, "bold")).pack(side=tk.LEFT,
                                                                                                       padx=5)

        toggle_vars = {}

        for option_key, option_text in toggle_options.items():
            var = tk.BooleanVar(value=True)
            toggle_vars[option_key] = var

            cb = tk.Checkbutton(
                toggles_frame,
                text=option_text,
                variable=var,
                bg='#fff9e6',
                font=("Arial", 9),
                command=refresh_callback
            )
            cb.pack(side=tk.LEFT, padx=10)

        return toggle_vars

    def auto_resize_columns(self, tree):
        """Автоматическая подгонка ширины колонок"""
        for col in tree["columns"]:
            max_width = 100
            for item in tree.get_children():
                try:
                    col_index = tree["columns"].index(col)
                    cell_value = str(tree.item(item)['values'][col_index])
                    cell_width = len(cell_value) * 8 + 20
                    if cell_width > max_width:
                        max_width = cell_width
                except:
                    pass
            max_width = min(max_width, 400)
            max_width = max(max_width, 80)
            tree.column(col, width=max_width)

    def save_toggle_settings(self):
        """Сохранить настройки переключателей"""
        settings = {}

        if hasattr(self, 'materials_toggles'):
            settings['materials'] = {k: v.get() for k, v in self.materials_toggles.items()}

        if hasattr(self, 'orders_toggles'):
            settings['orders'] = {k: v.get() for k, v in self.orders_toggles.items()}

        if hasattr(self, 'reservations_toggles'):
            settings['reservations'] = {k: v.get() for k, v in self.reservations_toggles.items()}

        if hasattr(self, 'balance_toggles'):
            settings['balance'] = {k: v.get() for k, v in self.balance_toggles.items()}

        if hasattr(self, 'writeoffs_toggles'):
            settings['writeoffs'] = {k: v.get() for k, v in self.writeoffs_toggles.items()}

        try:
            with open('toggle_settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except:
            pass

    def load_toggle_settings(self):
        """Загрузить настройки переключателей"""
        try:
            with open('toggle_settings.json', 'r', encoding='utf-8') as f:
                settings = json.load(f)

            if 'materials' in settings and hasattr(self, 'materials_toggles'):
                for k, v in settings['materials'].items():
                    if k in self.materials_toggles:
                        self.materials_toggles[k].set(v)

            if 'orders' in settings and hasattr(self, 'orders_toggles'):
                for k, v in settings['orders'].items():
                    if k in self.orders_toggles:
                        self.orders_toggles[k].set(v)

            if 'reservations' in settings and hasattr(self, 'reservations_toggles'):
                for k, v in settings['reservations'].items():
                    if k in self.reservations_toggles:
                        self.reservations_toggles[k].set(v)

            if 'balance' in settings and hasattr(self, 'balance_toggles'):
                for k, v in settings['balance'].items():
                    if k in self.balance_toggles:
                        self.balance_toggles[k].set(v)

            if 'writeoffs' in settings and hasattr(self, 'writeoffs_toggles'):
                for k, v in settings['writeoffs'].items():
                    if k in self.writeoffs_toggles:
                        self.writeoffs_toggles[k].set(v)

            self.refresh_materials()
            self.refresh_orders()
            self.refresh_reservations()
            self.refresh_balance()
            if hasattr(self, 'refresh_writeoffs'):
                self.refresh_writeoffs()
        except:
            pass

    def on_closing(self):
        """Обработчик закрытия окна"""
        self.save_toggle_settings()
        self.root.destroy()

    def setup_materials_tab(self):
        header = tk.Label(self.materials_frame, text="Учет листового проката на складе",
                          font=("Arial", 16, "bold"), bg='white', fg='#2c3e50')
        header.pack(pady=10)
        tree_frame = tk.Frame(self.materials_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        self.materials_tree = ttk.Treeview(tree_frame,
                                           columns=("ID", "Марка", "Толщина", "Длина", "Ширина", "Кол-во шт", "Площадь",
                                                    "Резерв", "Доступно", "Дата"),
                                           show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.materials_tree.yview)
        scroll_x.config(command=self.materials_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        columns_config = {"ID": 50, "Марка": 100, "Толщина": 80, "Длина": 80, "Ширина": 80,
                          "Кол-во шт": 80, "Площадь": 100, "Резерв": 80, "Доступно": 80, "Дата": 100}
        for col, width in columns_config.items():
            self.materials_tree.heading(col, text=col)
            self.materials_tree.column(col, width=width, anchor=tk.CENTER)
        self.materials_tree.pack(fill=tk.BOTH, expand=True)

        # Панель фильтрации
        self.materials_filters = self.create_filter_panel(
            self.materials_frame,
            self.materials_tree,
            ["ID", "Марка", "Толщина", "Длина", "Ширина", "Кол-во шт", "Резерв", "Доступно"],
            self.refresh_materials
        )

        # Переключатели видимости
        self.materials_toggles = self.create_visibility_toggles(
            self.materials_frame,
            self.materials_tree,
            {
                'show_zero_stock': '📦 Показать с нулевым остатком',
                'show_zero_available': '✅ Показать с нулём доступных'
            },
            self.refresh_materials
        )

        buttons_frame = tk.Frame(self.materials_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        btn_style = {"font": ("Arial", 10), "width": 15, "height": 2}
        tk.Button(buttons_frame, text="Добавить", bg='#27ae60', fg='white', command=self.add_material,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Импорт из Excel", bg='#9b59b6', fg='white', command=self.import_materials,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Скачать шаблон", bg='#3498db', fg='white', command=self.download_template,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Редактировать", bg='#f39c12', fg='white', command=self.edit_material,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Удалить", bg='#e74c3c', fg='white', command=self.delete_material,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Обновить", bg='#95a5a6', fg='white', command=self.refresh_materials,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        self.refresh_materials()

    def refresh_materials(self):
        for i in self.materials_tree.get_children():
            self.materials_tree.delete(i)
        df = load_data("Materials")
        if not df.empty:
            show_zero_stock = True
            show_zero_available = True

            if hasattr(self, 'materials_toggles') and self.materials_toggles:
                show_zero_stock = self.materials_toggles.get('show_zero_stock', tk.BooleanVar(value=True)).get()
                show_zero_available = self.materials_toggles.get('show_zero_available', tk.BooleanVar(value=True)).get()

            for index, row in df.iterrows():
                qty = int(row["Количество штук"])
                available = int(row["Доступно"])

                if not show_zero_stock and qty == 0:
                    continue
                if not show_zero_available and available == 0:
                    continue

                values = [row["ID"], row["Марка"], row["Толщина"], row["Длина"], row["Ширина"],
                          row["Количество штук"], row["Общая площадь"], row["Зарезервировано"],
                          row["Доступно"], row["Дата добавления"]]
                self.materials_tree.insert("", "end", values=values)

        self.auto_resize_columns(self.materials_tree)

    def download_template(self):
        file_path = filedialog.asksaveasfilename(title="Сохранить шаблон", defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                 initialfile="template_materials.xlsx")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Материалы"
            ws.append(["Марка", "Толщина", "Длина", "Ширина", "Количество штук"])
            examples = [["09Г2С", 10, 6000, 1500, 5], ["Ст3", 12, 6000, 1500, 3], ["40Х", 8, 3000, 1250, 10]]
            for example in examples:
                ws.append(example)
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws.column_dimensions[column].width = max_length + 2
            wb.save(file_path)
            messagebox.showinfo("Успех", f"Шаблон сохранен в:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать шаблон: {e}")

    def import_materials(self):
        file_path = filedialog.askopenfilename(title="Выберите файл Excel с материалами",
                                               filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not file_path:
            return
        try:
            import_df = pd.read_excel(file_path, engine='openpyxl')
            required_columns = ["Марка", "Толщина", "Длина", "Ширина", "Количество штук"]
            missing_columns = [col for col in required_columns if col not in import_df.columns]
            if missing_columns:
                messagebox.showerror("Ошибка", f"В файле отсутствуют колонки:\n{', '.join(missing_columns)}")
                return
            materials_df = load_data("Materials")
            current_max_id = 0 if materials_df.empty else int(materials_df["ID"].max())
            imported_count = 0
            errors = []
            for idx, row in import_df.iterrows():
                try:
                    if pd.isna(row["Марка"]) or row["Марка"] == "":
                        continue
                    marka = str(row["Марка"]).strip()
                    thickness = float(row["Толщина"])
                    length = float(row["Длина"])
                    width = float(row["Ширина"])
                    quantity = int(row["Количество штук"])
                    duplicate = materials_df[(materials_df["Марка"] == marka) & (materials_df["Толщина"] == thickness) &
                                             (materials_df["Длина"] == length) & (materials_df["Ширина"] == width)]
                    if not duplicate.empty:
                        material_id = duplicate.iloc[0]["ID"]
                        old_qty = int(duplicate.iloc[0]["Количество штук"])
                        new_qty = old_qty + quantity
                        reserved = int(duplicate.iloc[0]["Зарезервировано"])
                        area = (length * width * new_qty) / 1000000
                        materials_df.loc[materials_df["ID"] == material_id, "Количество штук"] = new_qty
                        materials_df.loc[materials_df["ID"] == material_id, "Общая площадь"] = round(area, 2)
                        materials_df.loc[materials_df["ID"] == material_id, "Доступно"] = new_qty - reserved
                    else:
                        current_max_id += 1
                        area = (length * width * quantity) / 1000000
                        new_row = pd.DataFrame([{"ID": current_max_id, "Марка": marka, "Толщина": thickness,
                                                 "Длина": length, "Ширина": width, "Количество штук": quantity,
                                                 "Общая площадь": round(area, 2), "Зарезервировано": 0,
                                                 "Доступно": quantity,
                                                 "Дата добавления": datetime.now().strftime("%Y-%m-%d")}])
                        materials_df = pd.concat([materials_df, new_row], ignore_index=True)
                    imported_count += 1
                except Exception as e:
                    errors.append(f"Строка {idx + 2}: {str(e)}")
            save_data("Materials", materials_df)
            self.refresh_materials()
            self.refresh_balance()
            result_msg = f"Успешно импортировано: {imported_count} материалов"
            if errors:
                result_msg += f"\n\nОшибки:\n" + "\n".join(errors[:10])
            messagebox.showinfo("Результат импорта", result_msg)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать данные:\n{e}")

    def add_material(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Добавить материал")
        add_window.geometry("450x500")
        add_window.configure(bg='#ecf0f1')
        tk.Label(add_window, text="Добавление листового проката", font=("Arial", 12, "bold"), bg='#ecf0f1').pack(
            pady=10)
        fields = [("Марка стали:", "marka"), ("Толщина (мм):", "thickness"), ("Длина (мм):", "length"),
                  ("Ширина (мм):", "width"), ("Количество штук:", "quantity")]
        entries = {}
        for label_text, key in fields:
            frame = tk.Frame(add_window, bg='#ecf0f1')
            frame.pack(fill=tk.X, padx=20, pady=5)
            tk.Label(frame, text=label_text, width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
            entry = tk.Entry(frame, font=("Arial", 10))
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            entries[key] = entry

        def save_material():
            try:
                marka = entries["marka"].get().strip()
                thickness = float(entries["thickness"].get().strip())
                length = float(entries["length"].get().strip())
                width = float(entries["width"].get().strip())
                quantity = int(entries["quantity"].get().strip())
                if not marka:
                    messagebox.showwarning("Предупреждение", "Заполните марку стали!")
                    return
                area = (length * width * quantity) / 1000000
                df = load_data("Materials")
                new_id = 1 if df.empty else int(df["ID"].max()) + 1
                new_row = pd.DataFrame(
                    [{"ID": new_id, "Марка": marka, "Толщина": thickness, "Длина": length, "Ширина": width,
                      "Количество штук": quantity, "Общая площадь": round(area, 2), "Зарезервировано": 0,
                      "Доступно": quantity, "Дата добавления": datetime.now().strftime("%Y-%m-%d")}])
                df = pd.concat([df, new_row], ignore_index=True)
                save_data("Materials", df)
                self.refresh_materials()
                self.refresh_balance()
                add_window.destroy()
                messagebox.showinfo("Успех", "Материал успешно добавлен!")
            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось добавить материал: {e}")

        tk.Button(add_window, text="Сохранить", bg='#27ae60', fg='white', font=("Arial", 12, "bold"),
                  command=save_material).pack(pady=20)

    def edit_material(self):
        selected = self.materials_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите материал для редактирования")
            return
        item_id = self.materials_tree.item(selected)["values"][0]
        df = load_data("Materials")
        row = df[df["ID"] == item_id].iloc[0]
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактировать материал")
        edit_window.geometry("450x500")
        edit_window.configure(bg='#ecf0f1')
        tk.Label(edit_window, text="Редактирование материала", font=("Arial", 12, "bold"), bg='#ecf0f1').pack(pady=10)
        fields = [("Марка стали:", "Марка"), ("Толщина (мм):", "Толщина"), ("Длина (мм):", "Длина"),
                  ("Ширина (мм):", "Ширина"), ("Количество штук:", "Количество штук")]
        entries = {}
        for label_text, key in fields:
            frame = tk.Frame(edit_window, bg='#ecf0f1')
            frame.pack(fill=tk.X, padx=20, pady=5)
            tk.Label(frame, text=label_text, width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
            entry = tk.Entry(frame, font=("Arial", 10))
            entry.insert(0, str(row[key]))
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            entries[key] = entry

        def save_changes():
            try:
                thickness = float(entries["Толщина"].get())
                length = float(entries["Длина"].get())
                width = float(entries["Ширина"].get())
                quantity = int(entries["Количество штук"].get())
                reserved = int(row["Зарезервировано"])
                area = (length * width * quantity) / 1000000
                df.loc[df["ID"] == item_id, "Марка"] = entries["Марка"].get()
                df.loc[df["ID"] == item_id, "Толщина"] = thickness
                df.loc[df["ID"] == item_id, "Длина"] = length
                df.loc[df["ID"] == item_id, "Ширина"] = width
                df.loc[df["ID"] == item_id, "Количество штук"] = quantity
                df.loc[df["ID"] == item_id, "Общая площадь"] = round(area, 2)
                df.loc[df["ID"] == item_id, "Доступно"] = quantity - reserved
                save_data("Materials", df)
                self.refresh_materials()
                self.refresh_balance()
                edit_window.destroy()
                messagebox.showinfo("Успех", "Материал успешно обновлен!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить материал: {e}")

        tk.Button(edit_window, text="Сохранить", bg='#3498db', fg='white', font=("Arial", 12, "bold"),
                  command=save_changes).pack(pady=20)

    def delete_material(self):
        selected = self.materials_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите материалы для удаления")
            return
        count = len(selected)
        if messagebox.askyesno("Подтверждение", f"Удалить выбранные материалы ({count} шт)?"):
            df = load_data("Materials")
            for item in selected:
                item_id = self.materials_tree.item(item)["values"][0]
                df = df[df["ID"] != item_id]
            save_data("Materials", df)
            self.refresh_materials()
            self.refresh_balance()  # <-- ЭТА СТРОКА ДОЛЖНА БЫТЬ!
            messagebox.showinfo("Успех", f"Удалено материалов: {count}")

    def setup_orders_tab(self):
        header = tk.Label(self.orders_frame, text="Управление заказами", font=("Arial", 16, "bold"), bg='white',
                          fg='#2c3e50')
        header.pack(pady=10)
        orders_label = tk.Label(self.orders_frame, text="Список заказов", font=("Arial", 12, "bold"), bg='white')
        orders_label.pack(pady=5)
        tree_frame = tk.Frame(self.orders_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        self.orders_tree = ttk.Treeview(tree_frame,
                                        columns=("ID", "Название", "Заказчик", "Дата", "Статус", "Примечания"),
                                        show="headings", yscrollcommand=scroll_y.set, height=8)
        scroll_y.config(command=self.orders_tree.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        columns_config = {"ID": 80, "Название": 200, "Заказчик": 150, "Дата": 100, "Статус": 100, "Примечания": 200}
        for col, width in columns_config.items():
            self.orders_tree.heading(col, text=col)
            self.orders_tree.column(col, width=width, anchor=tk.CENTER)
        self.orders_tree.pack(fill=tk.BOTH, expand=True)
        self.orders_tree.bind('<<TreeviewSelect>>', self.on_order_select)

        # Панель фильтрации заказов
        self.orders_filters = self.create_filter_panel(
            self.orders_frame,
            self.orders_tree,
            ["ID", "Название", "Заказчик", "Статус"],
            self.refresh_orders
        )

        # Переключатели видимости заказов
        self.orders_toggles = self.create_visibility_toggles(
            self.orders_frame,
            self.orders_tree,
            {
                'show_completed': '✅ Показать завершённые',
                'show_cancelled': '❌ Показать отменённые'
            },
            self.refresh_orders
        )

        buttons_frame = tk.Frame(self.orders_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=5)
        btn_style = {"font": ("Arial", 10), "width": 15, "height": 2}
        tk.Button(buttons_frame, text="Добавить заказ", bg='#27ae60', fg='white', command=self.add_order,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Импорт из Excel", bg='#9b59b6', fg='white', command=self.import_orders,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Скачать шаблон", bg='#3498db', fg='white', command=self.download_orders_template,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Редактировать", bg='#f39c12', fg='white', command=self.edit_order,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Удалить заказ", bg='#e74c3c', fg='white', command=self.delete_order,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Обновить", bg='#95a5a6', fg='white', command=self.refresh_orders,
                  **btn_style).pack(side=tk.LEFT, padx=5)

        details_label = tk.Label(self.orders_frame, text="Детали выбранного заказа", font=("Arial", 12, "bold"),
                                 bg='white')
        details_label.pack(pady=5)
        details_tree_frame = tk.Frame(self.orders_frame, bg='white')
        details_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y2 = tk.Scrollbar(details_tree_frame, orient=tk.VERTICAL)
        self.order_details_tree = ttk.Treeview(details_tree_frame,
                                               columns=("ID", "ID заказа", "Название детали", "Количество", "Порезано",
                                                        "Погнуто"),
                                               )
        scroll_y2.config(command=self.order_details_tree.yview)
        scroll_y2.pack(side=tk.RIGHT, fill=tk.Y)
        for col in ["ID", "ID заказа", "Название детали", "Количество", "Порезано", "Погнуто"]:
            self.order_details_tree.heading(col, text=col)
            self.order_details_tree.column(col, width=150, anchor=tk.CENTER)
        self.order_details_tree.pack(fill=tk.BOTH, expand=True)

        # Панель фильтрации деталей
        self.order_details_filters = self.create_filter_panel(
            self.orders_frame,
            self.order_details_tree,
            ["Название детали", "Количество"],
            self.refresh_order_details
        )

        details_buttons_frame = tk.Frame(self.orders_frame, bg='white')
        details_buttons_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(details_buttons_frame, text="Добавить деталь", bg='#27ae60', fg='white',
                  command=self.add_order_detail, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(details_buttons_frame, text="Редактировать деталь", bg='#f39c12', fg='white',
                  command=self.edit_order_detail, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(details_buttons_frame, text="Удалить деталь", bg='#e74c3c', fg='white',
                  command=self.delete_order_detail, **btn_style).pack(side=tk.LEFT, padx=5)
        self.refresh_orders()

    def on_order_select(self, event):
        self.refresh_order_details()

    def refresh_orders(self):
        for i in self.orders_tree.get_children():
            self.orders_tree.delete(i)
        df = load_data("Orders")
        if not df.empty:
            show_completed = True
            show_cancelled = True

            if hasattr(self, 'orders_toggles') and self.orders_toggles:
                show_completed = self.orders_toggles.get('show_completed', tk.BooleanVar(value=True)).get()
                show_cancelled = self.orders_toggles.get('show_cancelled', tk.BooleanVar(value=True)).get()

            for index, row in df.iterrows():
                status = row["Статус"]

                if not show_completed and status == "Завершен":
                    continue
                if not show_cancelled and status == "Отменен":
                    continue

                values = [row["ID заказа"], row["Название заказа"], row["Заказчик"],
                          row["Дата создания"], row["Статус"], row["Примечания"]]
                self.orders_tree.insert("", "end", values=values)
                self.auto_resize_columns(self.orders_tree)

    def refresh_order_details(self):
        for i in self.order_details_tree.get_children():
            self.order_details_tree.delete(i)

        selected = self.orders_tree.selection()

        # ЗАЩИТА: Если ничего не выбрано или выбрано несколько - выходим
        if not selected or len(selected) != 1:
            return

        try:
            order_id = self.orders_tree.item(selected[0])["values"][0]
        except (IndexError, KeyError, tk.TclError):
            # Если не удалось получить ID - выходим
            return

        df = load_data("OrderDetails")

        if not df.empty:
            # Настраиваем теги для цветовой индикации
            self.order_details_tree.tag_configure('completed', background='#c8e6c9')  # Зеленый - завершено
            self.order_details_tree.tag_configure('in_progress', background='#fff9c4')  # Желтый - в процессе
            self.order_details_tree.tag_configure('not_started', background='#ffcccc')  # Красный - не начато

            order_details = df[df["ID заказа"] == order_id]
            for index, row in order_details.iterrows():
                # Безопасное получение значений с обработкой пустых строк
                cut_raw = row.get("Порезано", 0) if "Порезано" in row else 0
                bent_raw = row.get("Погнуто", 0) if "Погнуто" in row else 0

                # Преобразуем в int с защитой от пустых строк
                try:
                    cut = int(cut_raw) if cut_raw != '' and pd.notna(cut_raw) else 0
                except (ValueError, TypeError):
                    cut = 0

                try:
                    bent = int(bent_raw) if bent_raw != '' and pd.notna(bent_raw) else 0
                except (ValueError, TypeError):
                    bent = 0

                qty = int(row["Количество"])

                values = (row["ID"], row["ID заказа"], row["Название детали"], qty, cut, bent)

                # Определяем статус выполнения
                if bent == qty and qty > 0:
                    tag = 'completed'  # Все детали погнуты = готово
                elif cut > 0 or bent > 0:
                    tag = 'in_progress'  # Что-то порезано или погнуто = в работе
                else:
                    tag = 'not_started'  # Ничего не сделано

                self.order_details_tree.insert("", "end", values=values, tags=(tag,))

            self.auto_resize_columns(self.order_details_tree)

    def download_orders_template(self):
        file_path = filedialog.asksaveasfilename(title="Сохранить шаблон", defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                                                 initialfile="template_orders.xlsx")
        if not file_path:
            return
        try:
            wb = Workbook()
            ws_orders = wb.active
            ws_orders.title = "Заказы"
            headers_orders = ["Название заказа", "Заказчик", "Статус", "Примечания"]
            ws_orders.append(headers_orders)
            examples_orders = [
                ["Заказ №1 - Металлоконструкции", "ООО Стройтех", "Новый", "Срочный заказ"],
                ["Заказ №2 - Лестница", "ИП Иванов", "В работе", ""],
                ["Заказ №3 - Ограждение", "ООО Метпром", "Новый", "Требуется предоплата"]
            ]
            for example in examples_orders:
                ws_orders.append(example)
            for col in ws_orders.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws_orders.column_dimensions[column].width = max_length + 2
            ws_details = wb.create_sheet("Детали")
            headers_details = ["Название заказа", "Название детали", "Количество"]
            ws_details.append(headers_details)
            examples_details = [
                ["Заказ №1 - Металлоконструкции", "Балка двутавровая 20", 15],
                ["Заказ №1 - Металлоконструкции", "Швеллер 16", 8],
                ["Заказ №2 - Лестница", "Ступень 300x250", 12],
                ["Заказ №2 - Лестница", "Поручень", 2],
                ["Заказ №3 - Ограждение", "Стойка 50x50", 20]
            ]
            for example in examples_details:
                ws_details.append(example)
            for col in ws_details.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                ws_details.column_dimensions[column].width = max_length + 2
            wb.save(file_path)
            messagebox.showinfo("Успех",
                                f"Шаблон сохранен в:\n{file_path}\n\n📋 ИНСТРУКЦИЯ:\n\nЛист 'Заказы':\n• Название заказа - уникальное имя\n• Заказчик - обязательно\n• Статус: Новый, В работе, Завершен, Отменен\n• Примечания - опционально\n\nЛист 'Детали':\n• Название заказа - должно совпадать с листом 'Заказы'\n• Название детали - обязательно\n• Количество - число")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать шаблон: {e}")

    def import_orders(self):
        file_path = filedialog.askopenfilename(title="Выберите файл Excel с заказами",
                                               filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if not file_path:
            return
        try:
            try:
                orders_import_df = pd.read_excel(file_path, sheet_name="Заказы", engine='openpyxl')
            except:
                messagebox.showerror("Ошибка", "В файле отсутствует лист 'Заказы'!\n\nИспользуйте шаблон.")
                return
            try:
                details_import_df = pd.read_excel(file_path, sheet_name="Детали", engine='openpyxl')
                has_details = True
            except:
                details_import_df = pd.DataFrame()
                has_details = False
            required_columns_orders = ["Название заказа", "Заказчик"]
            missing_columns = [col for col in required_columns_orders if col not in orders_import_df.columns]
            if missing_columns:
                messagebox.showerror("Ошибка",
                                     f"В листе 'Заказы' отсутствуют колонки:\n{', '.join(missing_columns)}\n\nИспользуйте кнопку 'Скачать шаблон'.")
                return
            if has_details and not details_import_df.empty:
                required_columns_details = ["Название заказа", "Название детали", "Количество"]
                missing_details = [col for col in required_columns_details if col not in details_import_df.columns]
                if missing_details:
                    messagebox.showwarning("Предупреждение",
                                           f"В листе 'Детали' отсутствуют колонки:\n{', '.join(missing_details)}\n\nДетали не будут импортированы.")
                    has_details = False
            orders_df = load_data("Orders")
            current_max_order_id = 1000 if orders_df.empty else int(orders_df["ID заказа"].max())
            order_details_df = load_data("OrderDetails")
            current_max_detail_id = 0 if order_details_df.empty else int(order_details_df["ID"].max())
            imported_orders = 0
            imported_details = 0
            errors = []
            valid_statuses = ["Новый", "В работе", "Завершен", "Отменен"]
            order_name_to_id = {}
            for idx, row in orders_import_df.iterrows():
                try:
                    if pd.isna(row["Название заказа"]) or str(row["Название заказа"]).strip() == "":
                        continue
                    if pd.isna(row["Заказчик"]) or str(row["Заказчик"]).strip() == "":
                        errors.append(f"Заказы, строка {idx + 2}: Отсутствует заказчик")
                        continue
                    order_name = str(row["Название заказа"]).strip()
                    customer = str(row["Заказчик"]).strip()
                    status = "Новый"
                    if "Статус" in orders_import_df.columns and not pd.isna(row["Статус"]):
                        status_input = str(row["Статус"]).strip()
                        if status_input in valid_statuses:
                            status = status_input
                        else:
                            errors.append(
                                f"Заказы, строка {idx + 2}: Неверный статус '{status_input}', установлен 'Новый'")
                    notes = ""
                    if "Примечания" in orders_import_df.columns and not pd.isna(row["Примечания"]):
                        notes = str(row["Примечания"]).strip()
                    current_max_order_id += 1
                    new_order_id = current_max_order_id
                    new_row = pd.DataFrame([{
                        "ID заказа": new_order_id,
                        "Название заказа": order_name,
                        "Заказчик": customer,
                        "Дата создания": datetime.now().strftime("%Y-%m-%d"),
                        "Статус": status,
                        "Примечания": notes
                    }])
                    orders_df = pd.concat([orders_df, new_row], ignore_index=True)
                    imported_orders += 1
                    order_name_to_id[order_name] = new_order_id
                except Exception as e:
                    errors.append(f"Заказы, строка {idx + 2}: {str(e)}")
            if has_details and not details_import_df.empty:
                for idx, row in details_import_df.iterrows():
                    try:
                        if pd.isna(row["Название заказа"]) or str(row["Название заказа"]).strip() == "":
                            continue
                        order_name = str(row["Название заказа"]).strip()
                        if order_name not in order_name_to_id:
                            errors.append(f"Детали, строка {idx + 2}: Заказ '{order_name}' не найден в листе 'Заказы'")
                            continue
                        if pd.isna(row["Название детали"]) or str(row["Название детали"]).strip() == "":
                            errors.append(f"Детали, строка {idx + 2}: Отсутствует название детали")
                            continue
                        detail_name = str(row["Название детали"]).strip()
                        if pd.isna(row["Количество"]):
                            errors.append(
                                f"Детали, строка {idx + 2}: Отсутствует количество для детали '{detail_name}'")
                            continue
                        try:
                            quantity = float(row["Количество"])
                            quantity = int(quantity)
                            if quantity <= 0:
                                errors.append(
                                    f"Детали, строка {idx + 2}: Количество должно быть больше нуля для детали '{detail_name}'")
                                continue
                        except (ValueError, TypeError):
                            errors.append(
                                f"Детали, строка {idx + 2}: Неверное количество '{row['Количество']}' для детали '{detail_name}'")
                            continue
                        current_max_detail_id += 1
                        order_id = order_name_to_id[order_name]
                        new_detail = pd.DataFrame([{
                            "ID": current_max_detail_id,
                            "ID заказа": order_id,
                            "Название детали": detail_name,
                            "Количество": quantity
                        }])
                        order_details_df = pd.concat([order_details_df, new_detail], ignore_index=True)
                        imported_details += 1
                    except Exception as e:
                        errors.append(f"Детали, строка {idx + 2}: {str(e)}")
            save_data("Orders", orders_df)
            if imported_details > 0:
                save_data("OrderDetails", order_details_df)
            self.refresh_orders()
            result_msg = f"✅ Успешно импортировано:\n• Заказов: {imported_orders}\n• Деталей: {imported_details}"
            if errors:
                result_msg += f"\n\n⚠ Ошибки ({len(errors)}):\n" + "\n".join(errors[:15])
                if len(errors) > 15:
                    result_msg += f"\n... и еще {len(errors) - 15} ошибок"
            messagebox.showinfo("Результат импорта", result_msg)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать данные:\n{e}")

    def add_order(self):
        add_window = tk.Toplevel(self.root)
        add_window.title("Добавить заказ")
        add_window.geometry("450x450")
        add_window.configure(bg='#ecf0f1')
        tk.Label(add_window, text="Создание нового заказа", font=("Arial", 12, "bold"), bg='#ecf0f1').pack(pady=10)
        fields = [("Название заказа:", "name"), ("Заказчик:", "customer"), ("Примечания:", "notes")]
        entries = {}
        for label_text, key in fields:
            frame = tk.Frame(add_window, bg='#ecf0f1')
            frame.pack(fill=tk.X, padx=20, pady=5)
            tk.Label(frame, text=label_text, width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
            entry = tk.Entry(frame, font=("Arial", 10))
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            entries[key] = entry
        status_frame = tk.Frame(add_window, bg='#ecf0f1')
        status_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(status_frame, text="Статус:", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)
        status_var = tk.StringVar(value="Новый")
        status_combo = ttk.Combobox(status_frame, textvariable=status_var,
                                    values=["Новый", "В работе", "Завершен", "Отменен"],
                                    font=("Arial", 10), state="readonly")
        status_combo.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        def save_order():
            try:
                name = entries["name"].get().strip()
                customer = entries["customer"].get().strip()
                if not name or not customer:
                    messagebox.showwarning("Предупреждение", "Заполните название и заказчика!")
                    return
                df = load_data("Orders")
                new_id = 1001 if df.empty else int(df["ID заказа"].max()) + 1
                new_row = pd.DataFrame([{"ID заказа": new_id, "Название заказа": name, "Заказчик": customer,
                                         "Дата создания": datetime.now().strftime("%Y-%m-%d"),
                                         "Статус": status_var.get(), "Примечания": entries["notes"].get()}])
                df = pd.concat([df, new_row], ignore_index=True)
                save_data("Orders", df)
                self.refresh_orders()
                add_window.destroy()
                messagebox.showinfo("Успех", f"Заказ #{new_id} успешно создан!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось создать заказ: {e}")

        tk.Button(add_window, text="Создать заказ", bg='#27ae60', fg='white', font=("Arial", 12, "bold"),
                  command=save_order).pack(pady=20)

    def edit_order(self):
        selected = self.orders_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите заказ для редактирования")
            return
        item_id = self.orders_tree.item(selected)["values"][0]
        df = load_data("Orders")
        row = df[df["ID заказа"] == item_id].iloc[0]
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактировать заказ")
        edit_window.geometry("450x450")
        edit_window.configure(bg='#ecf0f1')
        tk.Label(edit_window, text=f"Редактирование заказа #{item_id}", font=("Arial", 12, "bold"), bg='#ecf0f1').pack(
            pady=10)
        fields = [("Название заказа:", "Название заказа"), ("Заказчик:", "Заказчик"), ("Примечания:", "Примечания")]
        entries = {}
        for label_text, key in fields:
            frame = tk.Frame(edit_window, bg='#ecf0f1')
            frame.pack(fill=tk.X, padx=20, pady=5)
            tk.Label(frame, text=label_text, width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
            entry = tk.Entry(frame, font=("Arial", 10))
            entry.insert(0, str(row[key]))
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            entries[key] = entry
        status_frame = tk.Frame(edit_window, bg='#ecf0f1')
        status_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(status_frame, text="Статус:", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)
        status_var = tk.StringVar(value=row["Статус"])
        status_combo = ttk.Combobox(status_frame, textvariable=status_var,
                                    values=["Новый", "В работе", "Завершен", "Отменен"],
                                    font=("Arial", 10), state="readonly")
        status_combo.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        def save_changes():
            try:
                df.loc[df["ID заказа"] == item_id, "Название заказа"] = entries["Название заказа"].get()
                df.loc[df["ID заказа"] == item_id, "Заказчик"] = entries["Заказчик"].get()
                df.loc[df["ID заказа"] == item_id, "Статус"] = status_var.get()
                df.loc[df["ID заказа"] == item_id, "Примечания"] = entries["Примечания"].get()
                save_data("Orders", df)
                self.refresh_orders()
                edit_window.destroy()
                messagebox.showinfo("Успех", "Заказ успешно обновлен!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить заказ: {e}")

        tk.Button(edit_window, text="Сохранить", bg='#3498db', fg='white', font=("Arial", 12, "bold"),
                  command=save_changes).pack(pady=20)

    def delete_order(self):
        selected = self.orders_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите заказы для удаления")
            return
        count = len(selected)
        if messagebox.askyesno("Подтверждение", f"Удалить выбранные заказы ({count} шт)?"):
            df = load_data("Orders")
            details_df = load_data("OrderDetails")
            for item in selected:
                item_id = self.orders_tree.item(item)["values"][0]
                df = df[df["ID заказа"] != item_id]
                if not details_df.empty:
                    details_df = details_df[details_df["ID заказа"] != item_id]
            save_data("Orders", df)
            if not details_df.empty or len(selected) > 0:
                save_data("OrderDetails", details_df)
            self.refresh_orders()
            self.refresh_order_details()
            messagebox.showinfo("Успех", f"Удалено заказов: {count}")

    def add_order_detail(self):
        selected = self.orders_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Сначала выберите заказ!")
            return
        order_id = self.orders_tree.item(selected)["values"][0]
        add_window = tk.Toplevel(self.root)
        add_window.title("Добавить деталь")
        add_window.geometry("400x300")
        add_window.configure(bg='#ecf0f1')
        tk.Label(add_window, text=f"Добавление детали к заказу #{order_id}", font=("Arial", 12, "bold"),
                 bg='#ecf0f1').pack(pady=10)
        name_frame = tk.Frame(add_window, bg='#ecf0f1')
        name_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(name_frame, text="Название детали:", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)
        name_entry = tk.Entry(name_frame, font=("Arial", 10))
        name_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
        qty_frame = tk.Frame(add_window, bg='#ecf0f1')
        qty_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(qty_frame, text="Количество:", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        def save_detail():
            try:
                detail_name = name_entry.get().strip()
                quantity = int(qty_entry.get().strip())
                if not detail_name:
                    messagebox.showwarning("Предупреждение", "Введите название детали!")
                    return
                df = load_data("OrderDetails")
                new_id = 1 if df.empty else int(df["ID"].max()) + 1
                new_row = pd.DataFrame(
                    [{"ID": new_id, "ID заказа": order_id, "Название детали": detail_name,
                      "Количество": quantity, "Порезано": 0, "Погнуто": 0}])
                df = pd.concat([df, new_row], ignore_index=True)
                save_data("OrderDetails", df)
                self.refresh_order_details()
                add_window.destroy()
                messagebox.showinfo("Успех", "Деталь добавлена!")
            except ValueError:
                messagebox.showerror("Ошибка", "Количество должно быть числом!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось добавить деталь: {e}")

        tk.Button(add_window, text="Добавить", bg='#27ae60', fg='white', font=("Arial", 12, "bold"),
                  command=save_detail).pack(pady=20)

    def delete_order_detail(self):
        selected = self.order_details_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите детали для удаления")
            return
        count = len(selected)
        if messagebox.askyesno("Подтверждение", f"Удалить выбранные детали ({count} шт)?"):
            df = load_data("OrderDetails")
            for item in selected:
                detail_id = self.order_details_tree.item(item)["values"][0]
                df = df[df["ID"] != detail_id]
            save_data("OrderDetails", df)
            self.refresh_order_details()
            messagebox.showinfo("Успех", f"Удалено деталей: {count}")

            def edit_order_detail(self):
                """Редактирование детали заказа с учетом этапов производства"""
                selected = self.order_details_tree.selection()
                if not selected:
                    messagebox.showwarning("Предупреждение", "Выберите деталь для редактирования")
                    return

                detail_id = self.order_details_tree.item(selected)["values"][0]
                df = load_data("OrderDetails")
                row = df[df["ID"] == detail_id].iloc[0]

                edit_window = tk.Toplevel(self.root)
                edit_window.title("Редактировать деталь")
                edit_window.geometry("450x550")
                edit_window.configure(bg='#ecf0f1')

                tk.Label(edit_window, text=f"Редактирование детали #{detail_id}",
                         font=("Arial", 12, "bold"), bg='#ecf0f1').pack(pady=10)

                # Название детали
                name_frame = tk.Frame(edit_window, bg='#ecf0f1')
                name_frame.pack(fill=tk.X, padx=20, pady=5)
                tk.Label(name_frame, text="Название детали:", width=20, anchor='w',
                         bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
                name_entry = tk.Entry(name_frame, font=("Arial", 10))
                name_entry.insert(0, str(row["Название детали"]))
                name_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

                # Общее количество
                qty_frame = tk.Frame(edit_window, bg='#ecf0f1')
                qty_frame.pack(fill=tk.X, padx=20, pady=5)
                tk.Label(qty_frame, text="📋 Общее количество:", width=20, anchor='w',
                         bg='#ecf0f1', font=("Arial", 10, "bold")).pack(side=tk.LEFT)
                qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
                qty_entry.insert(0, str(int(row["Количество"])))
                qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

                # Разделитель для этапов производства
                tk.Label(edit_window, text="━" * 50, bg='#ecf0f1', fg='#95a5a6').pack(pady=10)
                tk.Label(edit_window, text="Этапы производства", font=("Arial", 11, "bold"),
                         bg='#ecf0f1', fg='#2980b9').pack(pady=5)

                # Порезано (этап 1)
                cut_frame = tk.Frame(edit_window, bg='#ecf0f1')
                cut_frame.pack(fill=tk.X, padx=20, pady=5)
                tk.Label(cut_frame, text="✂️ Порезано:", width=20, anchor='w',
                         bg='#ecf0f1', font=("Arial", 10, "bold"), fg='#27ae60').pack(side=tk.LEFT)
                cut_entry = tk.Entry(cut_frame, font=("Arial", 10))
                cut_raw = row.get("Порезано", 0) if "Порезано" in row else 0
                try:
                    cut_value = int(cut_raw) if cut_raw != '' and pd.notna(cut_raw) else 0
                except (ValueError, TypeError):
                    cut_value = 0
                cut_entry.insert(0, str(cut_value))
                cut_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

                # Погнуто (этап 2)
                bent_frame = tk.Frame(edit_window, bg='#ecf0f1')
                bent_frame.pack(fill=tk.X, padx=20, pady=5)
                tk.Label(bent_frame, text="🔧 Погнуто:", width=20, anchor='w',
                         bg='#ecf0f1', font=("Arial", 10, "bold"), fg='#f39c12').pack(side=tk.LEFT)
                bent_entry = tk.Entry(bent_frame, font=("Arial", 10))
                bent_raw = row.get("Погнуто", 0) if "Погнуто" in row else 0
                try:
                    bent_value = int(bent_raw) if bent_raw != '' and pd.notna(bent_raw) else 0
                except (ValueError, TypeError):
                    bent_value = 0
                bent_entry.insert(0, str(bent_value))
                bent_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

                # Информация
                info_frame = tk.Frame(edit_window, bg='#d1ecf1', relief=tk.RIDGE, borderwidth=2)
                info_frame.pack(fill=tk.X, padx=20, pady=10)
                tk.Label(info_frame, text="ℹ️ Информация о производстве:", font=("Arial", 9, "bold"),
                         bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=5, pady=2)
                tk.Label(info_frame, text="• Общее количество - всего деталей в заказе",
                         font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
                tk.Label(info_frame, text="• Порезано - количество заготовок после резки металла",
                         font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
                tk.Label(info_frame, text="• Погнуто - количество деталей после гибки",
                         font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
                tk.Label(info_frame, text="• Корректировка значений производится вручную",
                         font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)

                def save_changes():
                    try:
                        new_name = name_entry.get().strip()
                        new_qty = int(qty_entry.get().strip())
                        new_cut = int(cut_entry.get().strip())
                        new_bent = int(bent_entry.get().strip())

                        if not new_name:
                            messagebox.showwarning("Предупреждение", "Введите название детали!")
                            return

                        if new_qty < 0 or new_cut < 0 or new_bent < 0:
                            messagebox.showerror("Ошибка", "Значения не могут быть отрицательными!")
                            return

                        if new_cut > new_qty:
                            if not messagebox.askyesno("Предупреждение",
                                                       f"Порезано ({new_cut}) больше общего количества ({new_qty}).\n"
                                                       "Возможно, есть излишки заготовок.\n\nПродолжить?"):
                                return

                        if new_bent > new_cut:
                            if not messagebox.askyesno("Предупреждение",
                                                       f"Погнуто ({new_bent}) больше порезанных ({new_cut}).\n"
                                                       "Проверьте правильность данных.\n\nПродолжить?"):
                                return

                        # Обновляем данные
                        df.loc[df["ID"] == detail_id, "Название детали"] = new_name
                        df.loc[df["ID"] == detail_id, "Количество"] = new_qty
                        df.loc[df["ID"] == detail_id, "Порезано"] = new_cut
                        df.loc[df["ID"] == detail_id, "Погнуто"] = new_bent

                        save_data("OrderDetails", df)
                        self.refresh_order_details()
                        edit_window.destroy()

                        # Расчет остатков
                        to_cut = new_qty - new_cut
                        to_bend = new_cut - new_bent

                        messagebox.showinfo("Успех",
                                            f"✅ Деталь обновлена!\n\n"
                                            f"📋 Общее количество: {new_qty}\n"
                                            f"✂️ Порезано: {new_cut} (осталось порезать: {to_cut})\n"
                                            f"🔧 Погнуто: {new_bent} (осталось погнуть: {to_bend})")

                    except ValueError:
                        messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
                    except Exception as e:
                        messagebox.showerror("Ошибка", f"Не удалось обновить деталь: {e}")

                tk.Button(edit_window, text="💾 Сохранить изменения", bg='#3498db', fg='white',
                          font=("Arial", 12, "bold"), command=save_changes).pack(pady=15)

    def delete_order_detail(self):
        selected = self.order_details_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите детали для удаления")
            return
        count = len(selected)
        if messagebox.askyesno("Подтверждение", f"Удалить выбранные детали ({count} шт)?"):
            df = load_data("OrderDetails")
            for item in selected:
                detail_id = self.order_details_tree.item(item)["values"][0]
                df = df[df["ID"] != detail_id]
            save_data("OrderDetails", df)
            self.refresh_order_details()
            messagebox.showinfo("Успех", f"Удалено деталей: {count}")

    def edit_order_detail(self):
        """Редактирование детали заказа с учетом этапов производства"""
        selected = self.order_details_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите деталь для редактирования")
            return

        detail_id = self.order_details_tree.item(selected)["values"][0]
        df = load_data("OrderDetails")
        row = df[df["ID"] == detail_id].iloc[0]

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактировать деталь")
        edit_window.geometry("450x550")
        edit_window.configure(bg='#ecf0f1')

        tk.Label(edit_window, text=f"Редактирование детали #{detail_id}",
                 font=("Arial", 12, "bold"), bg='#ecf0f1').pack(pady=10)

        # Название детали
        name_frame = tk.Frame(edit_window, bg='#ecf0f1')
        name_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(name_frame, text="Название детали:", width=20, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
        name_entry = tk.Entry(name_frame, font=("Arial", 10))
        name_entry.insert(0, str(row["Название детали"]))
        name_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Общее количество
        qty_frame = tk.Frame(edit_window, bg='#ecf0f1')
        qty_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(qty_frame, text="📋 Общее количество:", width=20, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.insert(0, str(int(row["Количество"])))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Разделитель для этапов производства
        tk.Label(edit_window, text="━" * 50, bg='#ecf0f1', fg='#95a5a6').pack(pady=10)
        tk.Label(edit_window, text="Этапы производства", font=("Arial", 11, "bold"),
                 bg='#ecf0f1', fg='#2980b9').pack(pady=5)

        # Порезано (этап 1)
        cut_frame = tk.Frame(edit_window, bg='#ecf0f1')
        cut_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(cut_frame, text="✂️ Порезано:", width=20, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10, "bold"), fg='#27ae60').pack(side=tk.LEFT)
        cut_entry = tk.Entry(cut_frame, font=("Arial", 10))
        cut_value = row.get("Порезано", 0) if "Порезано" in row and pd.notna(row["Порезано"]) else 0
        cut_entry.insert(0, str(int(cut_value)))
        cut_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Погнуто (этап 2)
        bent_frame = tk.Frame(edit_window, bg='#ecf0f1')
        bent_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(bent_frame, text="🔧 Погнуто:", width=20, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10, "bold"), fg='#f39c12').pack(side=tk.LEFT)
        bent_entry = tk.Entry(bent_frame, font=("Arial", 10))
        bent_value = row.get("Погнуто", 0) if "Погнуто" in row and pd.notna(row["Погнуто"]) else 0
        bent_entry.insert(0, str(int(bent_value)))
        bent_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Информация
        info_frame = tk.Frame(edit_window, bg='#d1ecf1', relief=tk.RIDGE, borderwidth=2)
        info_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(info_frame, text="ℹ️ Информация о производстве:", font=("Arial", 9, "bold"),
                 bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=5, pady=2)
        tk.Label(info_frame, text="• Общее количество - всего деталей в заказе",
                 font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
        tk.Label(info_frame, text="• Порезано - количество заготовок после резки металла",
                 font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
        tk.Label(info_frame, text="• Погнуто - количество деталей после гибки",
                 font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
        tk.Label(info_frame, text="• Корректировка значений производится вручную",
                 font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)

        def save_changes():
            try:
                new_name = name_entry.get().strip()
                new_qty = int(qty_entry.get().strip())
                new_cut = int(cut_entry.get().strip())
                new_bent = int(bent_entry.get().strip())

                if not new_name:
                    messagebox.showwarning("Предупреждение", "Введите название детали!")
                    return

                if new_qty < 0 or new_cut < 0 or new_bent < 0:
                    messagebox.showerror("Ошибка", "Значения не могут быть отрицательными!")
                    return

                if new_cut > new_qty:
                    if not messagebox.askyesno("Предупреждение",
                                               f"Порезано ({new_cut}) больше общего количества ({new_qty}).\n"
                                               "Возможно, есть излишки заготовок.\n\nПродолжить?"):
                        return

                if new_bent > new_cut:
                    if not messagebox.askyesno("Предупреждение",
                                               f"Погнуто ({new_bent}) больше порезанных ({new_cut}).\n"
                                               "Проверьте правильность данных.\n\nПродолжить?"):
                        return

                # Обновляем данные
                df.loc[df["ID"] == detail_id, "Название детали"] = new_name
                df.loc[df["ID"] == detail_id, "Количество"] = new_qty
                df.loc[df["ID"] == detail_id, "Порезано"] = new_cut
                df.loc[df["ID"] == detail_id, "Погнуто"] = new_bent

                save_data("OrderDetails", df)
                self.refresh_order_details()
                edit_window.destroy()

                # Расчет остатков
                to_cut = new_qty - new_cut
                to_bend = new_cut - new_bent

                messagebox.showinfo("Успех",
                                    f"✅ Деталь обновлена!\n\n"
                                    f"📋 Общее количество: {new_qty}\n"
                                    f"✂️ Порезано: {new_cut} (осталось порезать: {to_cut})\n"
                                    f"🔧 Погнуто: {new_bent} (осталось погнуть: {to_bend})")

            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить деталь: {e}")

        tk.Button(edit_window, text="💾 Сохранить изменения", bg='#3498db', fg='white',
                  font=("Arial", 12, "bold"), command=save_changes).pack(pady=15)

    def setup_reservations_tab(self):
        header = tk.Label(self.reservations_frame, text="Резервирование материалов", font=("Arial", 16, "bold"),
                          bg='white', fg='#2c3e50')
        header.pack(pady=10)
        tree_frame = tk.Frame(self.reservations_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        self.reservations_tree = ttk.Treeview(tree_frame,
                                              columns=("ID", "Заказчик | Заказ", "Деталь", "Материал", "Марка",
                                                       "Толщина",
                                                       "Размер", "Резерв", "Списано", "Остаток", "Дата"),
                                              show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.reservations_tree.yview)
        scroll_x.config(command=self.reservations_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        columns_widths = {
            "ID": 60,
            "Заказчик | Заказ": 250,
            "Деталь": 150,
            "Материал": 80,
            "Марка": 100,
            "Толщина": 80,
            "Размер": 120,
            "Резерв": 80,
            "Списано": 80,
            "Остаток": 80,
            "Дата": 100
        }

        for col in self.reservations_tree["columns"]:
            self.reservations_tree.heading(col, text=col)
            width = columns_widths.get(col, 110)
            self.reservations_tree.column(col, width=width, anchor=tk.CENTER)
        self.reservations_tree.pack(fill=tk.BOTH, expand=True)

        # Панель фильтрации
        self.reservations_filters = self.create_filter_panel(
            self.reservations_frame,
            self.reservations_tree,
            ["ID", "Заказчик | Заказ", "Деталь", "Марка", "Толщина", "Резерв", "Списано", "Остаток"],
            self.refresh_reservations
        )

        # Переключатели видимости
        self.reservations_toggles = self.create_visibility_toggles(
            self.reservations_frame,
            self.reservations_tree,
            {
                'show_fully_written_off': '📝 Показать полностью списанные'
            },
            self.refresh_reservations
        )

        buttons_frame = tk.Frame(self.reservations_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        btn_style = {"font": ("Arial", 10), "width": 18, "height": 2}
        tk.Button(buttons_frame, text="Зарезервировать", bg='#27ae60', fg='white', command=self.add_reservation,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Удалить резерв", bg='#e74c3c', fg='white', command=self.delete_reservation,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Редактировать", bg='#f39c12', fg='white', command=self.edit_reservation,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Обновить", bg='#95a5a6', fg='white', command=self.refresh_reservations,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Задание на лазер", bg='#e67e22', fg='white', command=self.export_laser_task,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        self.refresh_reservations()

    def refresh_reservations(self):
        for i in self.reservations_tree.get_children():
            self.reservations_tree.delete(i)

        reservations_df = load_data("Reservations")
        orders_df = load_data("Orders")

        if not reservations_df.empty:
            show_fully_written_off = True

            if hasattr(self, 'reservations_toggles') and self.reservations_toggles:
                show_fully_written_off = self.reservations_toggles.get('show_fully_written_off',
                                                                       tk.BooleanVar(value=True)).get()

            for index, row in reservations_df.iterrows():
                remainder = int(row["Остаток к списанию"])
                if not show_fully_written_off and remainder == 0:
                    continue

                # Получаем информацию о заказе
                order_id = int(row["ID заказа"])
                order_display = f"#{order_id}"

                if not orders_df.empty:
                    order_row = orders_df[orders_df["ID заказа"] == order_id]
                    if not order_row.empty:
                        customer = order_row.iloc[0]["Заказчик"]
                        order_name = order_row.iloc[0]["Название заказа"]
                        order_display = f"{customer} | {order_name}"

                size_str = f"{row['Ширина']}x{row['Длина']}"
                detail_name = row.get("Название детали", "Не указана") if "Название детали" in row else "Не указана"

                values = [
                    row["ID резерва"],
                    order_display,  # Вместо ID заказа показываем "Заказчик | Название"
                    detail_name,
                    row["ID материала"],
                    row["Марка"],
                    row["Толщина"],
                    size_str,
                    row["Зарезервировано штук"],
                    row["Списано"],
                    row["Остаток к списанию"],
                    row["Дата резерва"]
                ]

                self.reservations_tree.insert("", "end", values=values)

            self.auto_resize_columns(self.reservations_tree)

    def add_reservation(self):
        orders_df = load_data("Orders")
        if orders_df.empty:
            messagebox.showwarning("Предупреждение", "Сначала создайте заказы!")
            return

        add_window = tk.Toplevel(self.root)
        add_window.title("Создать резерв")
        add_window.geometry("550x850")
        add_window.configure(bg='#ecf0f1')
        tk.Label(add_window, text="Резервирование материала под заказ", font=("Arial", 12, "bold"), bg='#ecf0f1').pack(
            pady=10)

        # ЗАКАЗ С ПОИСКОМ
        order_frame = tk.Frame(add_window, bg='#ecf0f1')
        order_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(order_frame, text="Заказ:", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)

        all_order_options = [
            f"ID:{int(row['ID заказа'])} | {row['Заказчик']} | {row['Название заказа']}"
            for _, row in orders_df.iterrows()
        ]

        order_search_var = tk.StringVar()
        order_search_entry = tk.Entry(order_frame, textvariable=order_search_var, font=("Arial", 10), width=35)
        order_search_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        order_results_frame = tk.Frame(add_window, bg='#ecf0f1')
        order_results_frame.pack(fill=tk.X, padx=20, pady=5)

        order_scroll = tk.Scrollbar(order_results_frame, orient=tk.VERTICAL)
        order_listbox = tk.Listbox(order_results_frame, height=3, font=("Arial", 9),
                                   yscrollcommand=order_scroll.set)
        order_scroll.config(command=order_listbox.yview)
        order_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        order_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for option in all_order_options:
            order_listbox.insert(tk.END, option)

        selected_order = {"value": None}

        def on_order_search(*args):
            search_text = order_search_var.get().lower()
            order_listbox.delete(0, tk.END)
            for option in all_order_options:
                if search_text in option.lower():
                    order_listbox.insert(tk.END, option)

        def on_select_order(event):
            try:
                selection = order_listbox.get(order_listbox.curselection())
                selected_order["value"] = selection
                order_search_var.set(selection)
                update_details_list()
            except:
                pass

        order_search_var.trace('w', on_order_search)
        order_listbox.bind('<<ListboxSelect>>', on_select_order)
        order_listbox.bind('<Double-Button-1>', on_select_order)

        # ДЕТАЛЬ ЗАКАЗА
        detail_frame = tk.Frame(add_window, bg='#ecf0f1')
        detail_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(detail_frame, text="Деталь заказа:", width=20, anchor='w', bg='#ecf0f1',
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)

        detail_var = tk.StringVar()
        detail_combo = ttk.Combobox(detail_frame, textvariable=detail_var, font=("Arial", 10), state="readonly",
                                    width=35)
        detail_combo.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        selected_detail = {"id": None, "name": None}

        def update_details_list():
            detail_combo['values'] = []
            detail_var.set("")
            selected_detail["id"] = None
            selected_detail["name"] = None

            if not selected_order["value"]:
                return

            try:
                order_id = int(selected_order["value"].split(" - ")[0])
                order_details_df = load_data("OrderDetails")

                if not order_details_df.empty:
                    details = order_details_df[order_details_df["ID заказа"] == order_id]

                    if not details.empty:
                        detail_options = ["[Без привязки к детали]"]
                        detail_options.extend([f"ID:{int(row['ID'])} - {row['Название детали']}"
                                               for _, row in details.iterrows()])
                        detail_combo['values'] = detail_options
                        detail_combo.current(0)
                    else:
                        detail_combo['values'] = ["[Нет деталей у заказа]"]
                        detail_combo.current(0)
                else:
                    detail_combo['values'] = ["[Нет деталей у заказа]"]
                    detail_combo.current(0)
            except:
                pass

        def on_detail_select(event):
            value = detail_var.get()
            if value and value.startswith("ID:"):
                try:
                    selected_detail["id"] = int(value.split("ID:")[1].split(" - ")[0])
                    selected_detail["name"] = value.split(" - ")[1]
                except:
                    selected_detail["id"] = None
                    selected_detail["name"] = None
            else:
                selected_detail["id"] = None
                selected_detail["name"] = None

        detail_combo.bind('<<ComboboxSelected>>', on_detail_select)

        # МАТЕРИАЛ С ПОИСКОМ
        material_frame = tk.Frame(add_window, bg='#ecf0f1')
        material_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(material_frame, text="Материал:", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)

        materials_df = load_data("Materials")
        all_material_options = ["[Добавить вручную]"]
        if not materials_df.empty:
            all_material_options.extend([
                                            f"{int(row['ID'])} - {row['Марка']} {row['Толщина']}мм {row['Ширина']}x{row['Длина']} (доступно: {int(row['Доступно'])} шт)"
                                            for _, row in materials_df.iterrows()])

        search_container = tk.Frame(material_frame, bg='#ecf0f1')
        search_container.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        material_search_var = tk.StringVar()
        material_search_entry = tk.Entry(search_container, textvariable=material_search_var, font=("Arial", 10))
        material_search_entry.pack(fill=tk.X)

        selected_reserve = {"value": None}

        search_results_frame = tk.Frame(add_window, bg='#ecf0f1')
        search_results_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        scroll_results = tk.Scrollbar(search_results_frame, orient=tk.VERTICAL)
        results_listbox = tk.Listbox(search_results_frame, height=5, font=("Arial", 9),
                                     yscrollcommand=scroll_results.set)
        scroll_results.config(command=results_listbox.yview)
        scroll_results.pack(side=tk.RIGHT, fill=tk.Y)
        results_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for option in all_material_options:
            results_listbox.insert(tk.END, option)

        selected_material = {"value": None}

        def on_search_change(*args):
            search_text = material_search_var.get().lower()
            results_listbox.delete(0, tk.END)

            for option in all_material_options:
                if search_text in option.lower():
                    results_listbox.insert(tk.END, option)

        def on_select_material(event):
            try:
                selection = results_listbox.get(results_listbox.curselection())
                selected_material["value"] = selection
                material_search_var.set(selection)
            except:
                pass

        material_search_var.trace('w', on_search_change)
        results_listbox.bind('<<ListboxSelect>>', on_select_material)
        results_listbox.bind('<Double-Button-1>', on_select_material)

        # ПАРАМЕТРЫ МАТЕРИАЛА (ручной ввод)
        manual_frame = tk.LabelFrame(add_window, text="Параметры материала (для ручного ввода)", bg='#ecf0f1',
                                     font=("Arial", 10, "bold"))
        manual_frame.pack(fill=tk.X, padx=20, pady=10)
        manual_entries = {}
        manual_fields = [("Марка стали:", "marka"), ("Толщина (мм):", "thickness"), ("Длина (мм):", "length"),
                         ("Ширина (мм):", "width")]
        for label_text, key in manual_fields:
            frame = tk.Frame(manual_frame, bg='#ecf0f1')
            frame.pack(fill=tk.X, padx=10, pady=3)
            tk.Label(frame, text=label_text, width=18, anchor='w', bg='#ecf0f1', font=("Arial", 9)).pack(side=tk.LEFT)
            entry = tk.Entry(frame, font=("Arial", 9))
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            manual_entries[key] = entry

        # КОЛИЧЕСТВО
        qty_frame = tk.Frame(add_window, bg='#ecf0f1')
        qty_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(qty_frame, text="Количество (шт):", width=20, anchor='w', bg='#ecf0f1',
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        def save_reservation():
            try:
                order_value = selected_order["value"] or order_search_var.get()
                if not order_value:
                    messagebox.showwarning("Предупреждение", "Выберите заказ!")
                    return

                material_value = selected_material["value"] or material_search_var.get()
                if not material_value:
                    messagebox.showwarning("Предупреждение", "Выберите материал!")
                    return

                # Парсим ID из формата "ID:1001 | Заказчик | Название"
                order_id = int(order_value.split("ID:")[1].split(" | ")[0])
                quantity = int(qty_entry.get())

                # Получаем ID и название детали
                detail_id = selected_detail["id"] if selected_detail["id"] else -1
                detail_name = selected_detail["name"] if selected_detail["name"] else "Не указана"

                if material_value == "[Добавить вручную]":
                    marka = manual_entries["marka"].get().strip()
                    thickness = float(manual_entries["thickness"].get().strip())
                    length = float(manual_entries["length"].get().strip())
                    width = float(manual_entries["width"].get().strip())
                    if not marka:
                        messagebox.showwarning("Предупреждение", "Заполните марку стали!")
                        return
                    material_id = -1
                else:
                    material_id = int(material_value.split(" - ")[0])
                    material_row = materials_df[materials_df["ID"] == material_id].iloc[0]
                    marka = material_row["Марка"]
                    thickness = material_row["Толщина"]
                    length = material_row["Длина"]
                    width = material_row["Ширина"]

                reservations_df = load_data("Reservations")
                new_id = 1 if reservations_df.empty else int(reservations_df["ID резерва"].max()) + 1

                new_row = pd.DataFrame([{
                    "ID резерва": new_id,
                    "ID заказа": order_id,
                    "ID детали": detail_id,
                    "Название детали": detail_name,
                    "ID материала": material_id,
                    "Марка": marka,
                    "Толщина": thickness,
                    "Длина": length,
                    "Ширина": width,
                    "Зарезервировано штук": quantity,
                    "Списано": 0,
                    "Остаток к списанию": quantity,
                    "Дата резерва": datetime.now().strftime("%Y-%m-%d")
                }])

                reservations_df = pd.concat([reservations_df, new_row], ignore_index=True)
                save_data("Reservations", reservations_df)

                if material_id != -1:
                    materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = int(
                        material_row["Зарезервировано"]) + quantity
                    materials_df.loc[materials_df["ID"] == material_id, "Доступно"] = int(
                        material_row["Доступно"]) - quantity
                    save_data("Materials", materials_df)
                    self.refresh_materials()

                self.refresh_reservations()
                self.refresh_balance()
                add_window.destroy()

                detail_info = f"\nДеталь: {detail_name}" if detail_name != "Не указана" else ""
                messagebox.showinfo("Успех", f"Резерв #{new_id} успешно создан!{detail_info}")

            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось создать резерв: {e}")

        tk.Button(add_window, text="Зарезервировать", bg='#27ae60', fg='white', font=("Arial", 12, "bold"),
                  command=save_reservation).pack(pady=15)

    def delete_reservation(self):
        selected = self.reservations_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите резервы для удаления")
            return
        count = len(selected)
        if messagebox.askyesno("Подтверждение",
                               f"Удалить выбранные резервы ({count} шт)?\n\nМатериалы вернутся на склад!"):
            reservations_df = load_data("Reservations")
            materials_df = load_data("Materials")
            for item in selected:
                reserve_id = self.reservations_tree.item(item)["values"][0]
                reserve_row = reservations_df[reservations_df["ID резерва"] == reserve_id].iloc[0]
                material_id = reserve_row["ID материала"]
                if material_id != -1:
                    quantity_to_return = int(reserve_row["Остаток к списанию"])
                    if not materials_df[materials_df["ID"] == material_id].empty:
                        mat_row = materials_df[materials_df["ID"] == material_id].iloc[0]
                        materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = int(
                            mat_row["Зарезервировано"]) - quantity_to_return
                        materials_df.loc[materials_df["ID"] == material_id, "Доступно"] = int(
                            mat_row["Доступно"]) + quantity_to_return
                reservations_df = reservations_df[reservations_df["ID резерва"] != reserve_id]
            save_data("Reservations", reservations_df)
            save_data("Materials", materials_df)
            self.refresh_materials()
            self.refresh_reservations()
            self.refresh_balance()
            messagebox.showinfo("Успех", f"Удалено резервов: {count}")

    def edit_reservation(self):
        """Редактирование резервирования"""
        selected = self.reservations_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите резерв для редактирования")
            return

        reserve_id = self.reservations_tree.item(selected)["values"][0]
        reservations_df = load_data("Reservations")
        reserve_row = reservations_df[reservations_df["ID резерва"] == reserve_id].iloc[0]

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактировать резерв")
        edit_window.geometry("550x600")
        edit_window.configure(bg='#ecf0f1')

        tk.Label(edit_window, text=f"Редактирование резерва #{reserve_id}",
                 font=("Arial", 12, "bold"), bg='#ecf0f1').pack(pady=10)

        # Информация о заказе (только для чтения)
        orders_df = load_data("Orders")
        order_id = int(reserve_row["ID заказа"])
        order_info = f"Заказ #{order_id}"

        if not orders_df.empty:
            order_row = orders_df[orders_df["ID заказа"] == order_id]
            if not order_row.empty:
                customer = order_row.iloc[0]["Заказчик"]
                order_name = order_row.iloc[0]["Название заказа"]
                order_info = f"{customer} | {order_name}"

        info_frame = tk.LabelFrame(edit_window, text="Информация о заказе (не редактируется)",
                                   bg='#e8f4f8', font=("Arial", 9, "bold"))
        info_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(info_frame, text=order_info, bg='#e8f4f8', font=("Arial", 10)).pack(padx=10, pady=5)

        # Деталь (только для чтения)
        detail_name = reserve_row.get("Название детали", "Не указана")
        if pd.isna(detail_name) or detail_name == "" or detail_name == "Не указана":
            detail_name = "Без привязки к детали"

        tk.Label(info_frame, text=f"Деталь: {detail_name}", bg='#e8f4f8', font=("Arial", 9)).pack(padx=10, pady=2)

        # Материал (только для чтения)
        material_info = f"{reserve_row['Марка']} {reserve_row['Толщина']}мм {reserve_row['Ширина']}x{reserve_row['Длина']}"
        tk.Label(info_frame, text=f"Материал: {material_info}", bg='#e8f4f8', font=("Arial", 9)).pack(padx=10, pady=2)

        # Редактируемое поле: Количество зарезервировано
        qty_frame = tk.Frame(edit_window, bg='#ecf0f1')
        qty_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(qty_frame, text="Зарезервировано (шт):", width=25, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.insert(0, str(int(reserve_row["Зарезервировано штук"])))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Информация о списании
        written_off = int(reserve_row["Списано"])
        remainder = int(reserve_row["Остаток к списанию"])

        stats_frame = tk.LabelFrame(edit_window, text="Статистика", bg='#fff3cd', font=("Arial", 9, "bold"))
        stats_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(stats_frame, text=f"Уже списано: {written_off} шт",
                 bg='#fff3cd', font=("Arial", 9)).pack(anchor='w', padx=10, pady=2)
        tk.Label(stats_frame, text=f"Остаток к списанию: {remainder} шт",
                 bg='#fff3cd', font=("Arial", 9)).pack(anchor='w', padx=10, pady=2)

        # Предупре��дение
        warning_frame = tk.Frame(edit_window, bg='#ffcccc', relief=tk.RIDGE, borderwidth=2)
        warning_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(warning_frame, text="ВАЖНО!", font=("Arial", 9, "bold"),
                 bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=5, pady=2)
        tk.Label(warning_frame, text="• Нельзя уменьшить количество ниже уже списанного",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)
        tk.Label(warning_frame, text="• Изменение количества пересчитает остаток к списанию",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)
        tk.Label(warning_frame, text="• Изменение влияет на баланс материалов на складе",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)

        def save_changes():
            try:
                new_qty = int(qty_entry.get().strip())

                if new_qty < written_off:
                    messagebox.showerror("Ошибка",
                                         f"Нельзя установить количество ({new_qty}) меньше уже списанного ({written_off})!")
                    return

                if new_qty <= 0:
                    messagebox.showerror("Ошибка", "Количество должно быть больше нуля!")
                    return

                old_qty = int(reserve_row["Зарезервировано штук"])
                difference = new_qty - old_qty

                if difference == 0:
                    messagebox.showinfo("Информация", "Изменений не было")
                    edit_window.destroy()
                    return

                # Подтверждение
                if not messagebox.askyesno("Подтверждение",
                                           f"Изменить количество с {old_qty} на {new_qty} шт?\n\n"
                                           f"Разница: {'+' if difference > 0 else ''}{difference} шт\n"
                                           f"Новый остаток к списанию: {new_qty - written_off} шт"):
                    return

                # Обновляем резерв
                new_remainder = new_qty - written_off
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Зарезервировано штук"] = new_qty
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = new_remainder
                save_data("Reservations", reservations_df)

                # Обновляем материал на складе (если не вручную добавленный)
                material_id = int(reserve_row["ID материала"])
                if material_id != -1:
                    materials_df = load_data("Materials")
                    if not materials_df[materials_df["ID"] == material_id].empty:
                        mat_row = materials_df[materials_df["ID"] == material_id].iloc[0]
                        current_reserved = int(mat_row["Зарезервировано"])
                        current_available = int(mat_row["Доступно"])

                        new_reserved = current_reserved + difference
                        new_available = current_available - difference

                        materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved
                        materials_df.loc[materials_df["ID"] == material_id, "Доступно"] = new_available
                        save_data("Materials", materials_df)
                        self.refresh_materials()

                self.refresh_reservations()
                self.refresh_balance()
                edit_window.destroy()
                messagebox.showinfo("Успех",
                                    f"Резерв обновлен!\n\n"
                                    f"Новое количество: {new_qty} шт\n"
                                    f"Остаток к списанию: {new_remainder} шт")

            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить резерв: {e}")
                import traceback
                traceback.print_exc()

        tk.Button(edit_window, text="Сохранить изменения", bg='#f39c12', fg='white',
                  font=("Arial", 12, "bold"), command=save_changes).pack(pady=15)

    def export_laser_task(self):
        """Формирование задания на лазер из резервов"""
        try:
            # Загружаем данные
            orders_df = load_data("Orders")
            reservations_df = load_data("Reservations")
            order_details_df = load_data("OrderDetails")

            if orders_df.empty:
                messagebox.showwarning("Предупреждение", "Нет заказов в базе!")
                return

            # Фильтруем заказы "В работе"
            active_orders = orders_df[orders_df["Статус"] == "В работе"]

            if active_orders.empty:
                messagebox.showwarning("Предупреждение", "Нет заказов со статусом 'В работе'!")
                return

            # Проверяем наличие резервов
            if reservations_df.empty:
                messagebox.showwarning("Предупреждение", "Нет зарезервированных материалов!")
                return

            # Окно выбора заказов
            select_window = tk.Toplevel(self.root)
            select_window.title("Выбор заказов для задания на лазер")
            select_window.geometry("700x600")
            select_window.configure(bg='#ecf0f1')

            tk.Label(select_window, text="Формирование задания на лазер",
                     font=("Arial", 14, "bold"), bg='#ecf0f1', fg='#e67e22').pack(pady=10)

            tk.Label(select_window, text="Выберите заказы (статус: В работе)",
                     font=("Arial", 10), bg='#ecf0f1').pack(pady=5)

            # Фрейм со списком заказов
            list_frame = tk.Frame(select_window, bg='#ecf0f1')
            list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

            scroll_y = tk.Scrollbar(list_frame, orient=tk.VERTICAL)

            # Создаем Listbox с множественным выбором
            orders_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE,
                                        font=("Arial", 10), yscrollcommand=scroll_y.set)
            scroll_y.config(command=orders_listbox.yview)
            scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
            orders_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            # Заполняем список заказов "В работе"
            order_map = {}
            orders_without_reserves = []

            for _, order in active_orders.iterrows():
                order_id = order["ID заказа"]
                order_name = order["Название заказа"]
                customer = order["Заказчик"]

                # Проверяем наличие резервов
                has_reserves = not reservations_df[reservations_df["ID заказа"] == order_id].empty

                if has_reserves:
                    display_text = f"ID:{int(order_id)} | {customer} | {order_name}"
                    orders_listbox.insert(tk.END, display_text)
                    order_map[display_text] = order_id
                else:
                    orders_without_reserves.append(f"{customer} - {order_name}")

            if orders_listbox.size() == 0:
                messagebox.showwarning("Предупреждение",
                                       "Нет заказов 'В работе' с зарезервированными материалами!")
                select_window.destroy()
                return

            # Кнопки выбора
            btn_frame = tk.Frame(select_window, bg='#ecf0f1')
            btn_frame.pack(fill=tk.X, padx=20, pady=5)

            def select_all():
                orders_listbox.select_set(0, tk.END)

            def deselect_all():
                orders_listbox.select_clear(0, tk.END)

            tk.Button(btn_frame, text="Выбрать все", bg='#3498db', fg='white',
                      font=("Arial", 9), command=select_all).pack(side=tk.LEFT, padx=5)
            tk.Button(btn_frame, text="Снять выбор", bg='#95a5a6', fg='white',
                      font=("Arial", 9), command=deselect_all).pack(side=tk.LEFT, padx=5)

            # Информация
            info_frame = tk.Frame(select_window, bg='#d1ecf1', relief=tk.RIDGE, borderwidth=2)
            info_frame.pack(fill=tk.X, padx=20, pady=10)
            tk.Label(info_frame, text="Информация:", font=("Arial", 9, "bold"),
                     bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=5, pady=2)
            tk.Label(info_frame, text="- Отображаются только заказы со статусом 'В работе'",
                     font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
            tk.Label(info_frame, text="- Для каждого резерва создается отдельная строка",
                     font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
            tk.Label(info_frame, text="- Формат: Заказчик | Название заявки | Деталь | Металл",
                     font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)
            tk.Label(info_frame, text="- Если деталь не привязана - 'Без учета деталей'",
                     font=("Arial", 8), bg='#d1ecf1', fg='#0c5460').pack(anchor='w', padx=10)

            # Предупреждение о заказах без резервов
            if orders_without_reserves:
                warning_frame = tk.Frame(select_window, bg='#fff3cd', relief=tk.RIDGE, borderwidth=2)
                warning_frame.pack(fill=tk.X, padx=20, pady=5)
                tk.Label(warning_frame, text="Внимание! Заказы 'В работе' без резервов:",
                         font=("Arial", 8, "bold"), bg='#fff3cd', fg='#856404').pack(anchor='w', padx=5, pady=2)
                for order_name in orders_without_reserves[:3]:
                    tk.Label(warning_frame, text=f"  - {order_name}",
                             font=("Arial", 7), bg='#fff3cd', fg='#856404').pack(anchor='w', padx=10)
                if len(orders_without_reserves) > 3:
                    tk.Label(warning_frame, text=f"  ... и ещё {len(orders_without_reserves) - 3}",
                             font=("Arial", 7), bg='#fff3cd', fg='#856404').pack(anchor='w', padx=10)

            def generate_file():
                selected_indices = orders_listbox.curselection()
                if not selected_indices:
                    messagebox.showwarning("Предупреждение", "Выберите хотя бы один заказ!")
                    return

                # Получаем выбранные ID заказов
                selected_order_ids = []
                for index in selected_indices:
                    display_text = orders_listbox.get(index)
                    selected_order_ids.append(order_map[display_text])

                # Формируем данные для экспорта
                export_data = []
                warnings = []

                for order_id in selected_order_ids:
                    # Получаем информацию о заказе
                    order_row = orders_df[orders_df["ID заказа"] == order_id]
                    if order_row.empty:
                        continue

                    customer = order_row.iloc[0]["Заказчик"]
                    order_name = order_row.iloc[0]["Название заказа"]

                    # Получаем резервы этого заказа
                    order_reserves = reservations_df[reservations_df["ID заказа"] == order_id]

                    if order_reserves.empty:
                        warnings.append(f"{customer} - {order_name}: нет резервов")
                        continue

                    for _, reserve in order_reserves.iterrows():
                        # Формируем название детали
                        detail_id = reserve.get("ID детали", -1)
                        detail_name = reserve.get("Название детали", "Без учета деталей")

                        # Проверяем корректность привязки детали
                        if pd.isna(detail_name) or detail_name == "" or detail_name == "Не указана" or detail_id == -1:
                            detail_name = "Без учета деталей"

                        # Формируем описание металла
                        metal_str = f"{reserve['Марка']} {reserve['Толщина']}мм {reserve['Ширина']}x{reserve['Длина']}"

                        # Добавляем строку
                        export_data.append({
                            "Заказчик": customer,
                            "Название заявки": order_name,
                            "Название детали": detail_name,
                            "Металл": metal_str
                        })

                if not export_data:
                    messagebox.showwarning("Предупреждение", "Нет данных для экспорта!")
                    return

                # Проверяем наличие строк "Без учета деталей"
                rows_without_details = sum(1 for row in export_data if row["Название детали"] == "Без учета деталей")

                if rows_without_details > 0:
                    if not messagebox.askyesno("Предупреждение",
                                               f"В таблице будет {rows_without_details} строк(и) без привязки к деталям!\n\n"
                                               "Это материалы, зарезервированные без указания конкретной детали.\n\n"
                                               "Продолжить формирование?"):
                        return

                # Диалог сохранения файла
                file_path = filedialog.asksaveasfilename(
                    title="Сохранить задание на лазер",
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    initialfile=f"zadanie_na_laser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                )

                if not file_path:
                    return

                # Создаём DataFrame и сохраняем
                export_df = pd.DataFrame(export_data)

                # Сохраняем с автоподбором ширины
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='Задание на лазер')
                    worksheet = writer.sheets['Задание на лазер']

                    # Автоподбор ширины колонок
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 60)
                        worksheet.column_dimensions[column_letter].width = adjusted_width

                select_window.destroy()

                result_msg = f"Задание на лазер успешно создано!\n\n"
                result_msg += f"Заказов обработано: {len(selected_order_ids)}\n"
                result_msg += f"Строк в таблице: {len(export_data)}\n"
                result_msg += f"Строк без деталей: {rows_without_details}\n\n"
                result_msg += f"Файл сохранен:\n{file_path}"

                messagebox.showinfo("Успех", result_msg)

            # Кнопка формирования
            tk.Button(select_window, text="Сформировать файл", bg='#e67e22', fg='white',
                      font=("Arial", 12, "bold"), command=generate_file).pack(pady=15)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать задание на лазер:\n{e}")
            import traceback
            traceback.print_exc()

    def setup_writeoffs_tab(self):
        header = tk.Label(self.writeoffs_frame, text="Списание зарезервированных материалов",
                          font=("Arial", 16, "bold"), bg='white', fg='#2c3e50')
        header.pack(pady=10)
        tree_frame = tk.Frame(self.writeoffs_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        self.writeoffs_tree = ttk.Treeview(tree_frame,
                                           columns=("ID", "ID резерва", "Заказ", "Деталь", "Материал", "Марка",
                                                    "Толщина",
                                                    "Размер", "Количество", "Дата", "Комментарий"),
                                           show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.writeoffs_tree.yview)
        scroll_x.config(command=self.writeoffs_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        columns_config = {"ID": 50, "ID резерва": 80, "Заказ": 200, "Деталь": 150, "Материал": 80,
                          "Марка": 90, "Толщина": 70, "Размер": 110, "Количество": 90,
                          "Дата": 140, "Комментарий": 180}
        for col, width in columns_config.items():
            self.writeoffs_tree.heading(col, text=col)
            self.writeoffs_tree.column(col, width=width, anchor=tk.CENTER)
        self.writeoffs_tree.pack(fill=tk.BOTH, expand=True)

        # Панель фильтрации
        self.writeoffs_filters = self.create_filter_panel(
            self.writeoffs_frame,
            self.writeoffs_tree,
            ["ID", "ID резерва", "Заказ", "Деталь", "Марка", "Толщина", "Количество"],
            self.refresh_writeoffs
        )

        buttons_frame = tk.Frame(self.writeoffs_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        btn_style = {"font": ("Arial", 10), "width": 18, "height": 2}
        tk.Button(buttons_frame, text="Списать материал", bg='#e67e22', fg='white', command=self.add_writeoff,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Удалить списание", bg='#e74c3c', fg='white', command=self.delete_writeoff,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Редактировать", bg='#f39c12', fg='white', command=self.edit_writeoff,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(buttons_frame, text="Обновить", bg='#95a5a6', fg='white', command=self.refresh_writeoffs,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        self.refresh_writeoffs()

    def refresh_writeoffs(self):
        for i in self.writeoffs_tree.get_children():
            self.writeoffs_tree.delete(i)

        writeoffs_df = load_data("WriteOffs")
        orders_df = load_data("Orders")
        reservations_df = load_data("Reservations")

        if not writeoffs_df.empty:
            for index, row in writeoffs_df.iterrows():
                # Получаем информацию о заказе
                order_id = int(row["ID заказа"])
                order_display = f"#{order_id}"

                if not orders_df.empty:
                    order_row = orders_df[orders_df["ID заказа"] == order_id]
                    if not order_row.empty:
                        customer = order_row.iloc[0]["Заказчик"]
                        order_name = order_row.iloc[0]["Название заказа"]
                        order_display = f"{customer} | {order_name}"

                # Получаем информацию о детали из резерва
                reserve_id = int(row["ID резерва"])
                detail_display = "Без детали"

                if not reservations_df.empty:
                    reserve_row = reservations_df[reservations_df["ID резерва"] == reserve_id]
                    if not reserve_row.empty:
                        detail_name = reserve_row.iloc[0].get("Название детали", "Без детали")
                        detail_id = reserve_row.iloc[0].get("ID детали", -1)

                        if pd.notna(
                                detail_name) and detail_name != "" and detail_name != "Не указана" and detail_id != -1:
                            detail_display = detail_name

                size_str = f"{row['Ширина']}x{row['Длина']}"

                values = [
                    row["ID списания"],
                    row["ID резерва"],
                    order_display,  # Заказчик | Название
                    detail_display,  # Название детали
                    row["ID материала"],
                    row["Марка"],
                    row["Толщина"],
                    size_str,
                    row["Количество"],
                    row["Дата списания"],
                    row["Комментарий"]
                ]

                self.writeoffs_tree.insert("", "end", values=values)

            self.auto_resize_columns(self.writeoffs_tree)

    def add_writeoff(self):
        reservations_df = load_data("Reservations")
        if reservations_df.empty:
            messagebox.showwarning("Предупреждение", "Нет зарезервированных материалов для списания!")
            return
        available_reserves = reservations_df[reservations_df["Остаток к списанию"] > 0]
        if available_reserves.empty:
            messagebox.showwarning("Предупреждение", "Все зарезервированные материалы уже списаны!")
            return
        add_window = tk.Toplevel(self.root)
        add_window.title("Списать материал")
        add_window.geometry("600x450")
        add_window.configure(bg='#fff3e0')
        tk.Label(add_window, text="Списание зарезервированного материала", font=("Arial", 12, "bold"), bg='#fff3e0',
                 fg='#e67e22').pack(pady=10)
        reserve_frame = tk.Frame(add_window, bg='#fff3e0')
        reserve_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(reserve_frame, text="Резерв:", width=20, anchor='w', bg='#fff3e0', font=("Arial", 10)).pack(
            side=tk.LEFT)
        reserve_options = []
        for _, row in available_reserves.iterrows():
            reserve_text = f"ID:{int(row['ID резерва'])} | Заказ:{int(row['ID заказа'])} | {row['Марка']} {row['Толщина']}мм {row['Ширина']}x{row['Длина']} | Доступно:{int(row['Остаток к списанию'])} шт"
            reserve_options.append(reserve_text)
        reserve_var = tk.StringVar()
        reserve_combo = ttk.Combobox(reserve_frame, textvariable=reserve_var, values=reserve_options, font=("Arial", 9),
                                     state="readonly", width=60)
        reserve_combo.pack(side=tk.RIGHT, expand=True, fill=tk.X)
        if reserve_options:
            reserve_combo.current(0)
        qty_frame = tk.Frame(add_window, bg='#fff3e0')
        qty_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(qty_frame, text="Количество (шт):", width=20, anchor='w', bg='#fff3e0',
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
        comment_frame = tk.Frame(add_window, bg='#fff3e0')
        comment_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(comment_frame, text="Комментарий:", width=20, anchor='w', bg='#fff3e0', font=("Arial", 10)).pack(
            side=tk.LEFT)
        comment_entry = tk.Entry(comment_frame, font=("Arial", 10))
        comment_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
        info_label = tk.Label(add_window, text="⚠ Списание уменьшит резерв и количество материала на складе!",
                              font=("Arial", 9, "italic"), bg='#fff3e0', fg='#d35400')
        info_label.pack(pady=10)

    def add_writeoff(self):
        reservations_df = load_data("Reservations")
        if reservations_df.empty:
            messagebox.showwarning("Предупреждение", "Нет резервов для списания!")
            return

        active_reserves = reservations_df[reservations_df["Остаток к списанию"] > 0]
        if active_reserves.empty:
            messagebox.showwarning("Предупреждение", "Нет резервов с остатком для списания!")
            return

        add_window = tk.Toplevel(self.root)
        add_window.title("Списание материала")
        add_window.geometry("550x500")
        add_window.configure(bg='#ecf0f1')

        tk.Label(add_window, text="Списание материала с резерва", font=("Arial", 12, "bold"), bg='#ecf0f1').pack(
            pady=10)

        # РЕЗЕРВ С ПОИСКОМ
        reserve_frame = tk.Frame(add_window, bg='#ecf0f1')
        reserve_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(reserve_frame, text="Резерв (поиск):", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)

        all_reserve_options = []

        # Загружаем заказы для отображения заказчика и названия
        orders_df = load_data("Orders")

        for _, row in active_reserves.iterrows():
            order_id = int(row['ID заказа'])

            # Ищем информацию о заказе
            order_info = ""
            if not orders_df.empty:
                order_row = orders_df[orders_df["ID заказа"] == order_id]
                if not order_row.empty:
                    customer = order_row.iloc[0]["Заказчик"]
                    order_name = order_row.iloc[0]["Название заказа"]
                    order_info = f"{customer} | {order_name}"
                else:
                    order_info = f"Заказ #{order_id}"
            else:
                order_info = f"Заказ #{order_id}"

            # Получаем название детали
            detail_name = row.get("Название детали", "Без учета деталей")
            detail_id = row.get("ID детали", -1)

            # Проверяем, привязана ли деталь
            if pd.isna(detail_name) or detail_name == "" or detail_name == "Не указана" or detail_id == -1:
                detail_info = "Без детали"
            else:
                detail_info = f"Деталь: {detail_name}"

            # Формируем строку с информацией о детали
            reserve_str = f"Резерв #{int(row['ID резерва'])} | {order_info} | {detail_info} | {row['Марка']} {row['Толщина']}мм | Осталось: {int(row['Остаток к списанию'])} шт"
            all_reserve_options.append(reserve_str)

        search_container = tk.Frame(reserve_frame, bg='#ecf0f1')
        search_container.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        reserve_search_var = tk.StringVar()
        selected_reserve = {"value": None}

        reserve_search_entry = tk.Entry(search_container, textvariable=reserve_search_var, font=("Arial", 10))
        reserve_search_entry.pack(fill=tk.X)

        # Listbox для результатов поиска
        search_results_frame = tk.Frame(add_window, bg='#ecf0f1')
        search_results_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        scroll_results = tk.Scrollbar(search_results_frame, orient=tk.VERTICAL)
        results_listbox = tk.Listbox(search_results_frame, height=8, font=("Arial", 9),
                                     yscrollcommand=scroll_results.set)
        scroll_results.config(command=results_listbox.yview)
        scroll_results.pack(side=tk.RIGHT, fill=tk.Y)
        results_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for option in all_reserve_options:
            results_listbox.insert(tk.END, option)

        def on_search_change(*args):
            search_text = reserve_search_var.get().lower()
            results_listbox.delete(0, tk.END)
            for option in all_reserve_options:
                if search_text in option.lower():
                    results_listbox.insert(tk.END, option)

        def on_select_reserve(event):
            try:
                selection = results_listbox.get(results_listbox.curselection())
                selected_reserve["value"] = selection
                reserve_search_var.set(selection)
            except:
                pass

        reserve_search_var.trace('w', on_search_change)
        results_listbox.bind('<<ListboxSelect>>', on_select_reserve)
        results_listbox.bind('<Double-Button-1>', on_select_reserve)

        # Количество
        qty_frame = tk.Frame(add_window, bg='#ecf0f1')
        qty_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(qty_frame, text="Количество (шт):", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Комментарий
        comment_frame = tk.Frame(add_window, bg='#ecf0f1')
        comment_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(comment_frame, text="Комментарий:", width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(
            side=tk.LEFT)
        comment_entry = tk.Entry(comment_frame, font=("Arial", 10))
        comment_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        def save_writeoff():
            try:
                reserve_value = selected_reserve["value"] or reserve_search_var.get()
                if not reserve_value:
                    messagebox.showwarning("Предупреждение", "Выберите резерв!")
                    return

                # Парсим ID из формата "Резерв #123 | ..."
                reserve_id = int(reserve_value.split("Резерв #")[1].split(" | ")[0])
                quantity = int(qty_entry.get())
                comment = comment_entry.get().strip()

                # Проверяем резерв
                reservations_df = load_data("Reservations")
                reservation = reservations_df[reservations_df["ID резерва"] == reserve_id].iloc[0]
                remainder = int(reservation["Остаток к списанию"])

                if quantity > remainder:
                    messagebox.showerror("Ошибка", f"Нельзя списать больше чем осталось!\nОсталось: {remainder} шт")
                    return

                # Добавляем списание
                writeoffs_df = load_data("WriteOffs")
                new_id = 1 if writeoffs_df.empty else int(writeoffs_df["ID списания"].max()) + 1

                new_row = pd.DataFrame([{
                    "ID списания": new_id,
                    "ID резерва": reserve_id,
                    "ID заказа": reservation["ID заказа"],
                    "ID материала": reservation["ID материала"],
                    "Марка": reservation["Марка"],
                    "Толщина": reservation["Толщина"],
                    "Длина": reservation["Длина"],
                    "Ширина": reservation["Ширина"],
                    "Количество": quantity,
                    "Дата списания": datetime.now().strftime("%Y-%m-%d"),
                    "Комментарий": comment
                }])

                writeoffs_df = pd.concat([writeoffs_df, new_row], ignore_index=True)
                save_data("WriteOffs", writeoffs_df)

                # Обновляем резервирование
                reservations_df = load_data("Reservations")
                reservation = reservations_df[reservations_df["ID резерва"] == reserve_id].iloc[0]

                new_written_off = int(reservation["Списано"]) + quantity
                new_remainder = int(reservation["Зарезервировано штук"]) - new_written_off

                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Списано"] = new_written_off
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = new_remainder
                save_data("Reservations", reservations_df)

                # Обновляем материал (ИСПРАВЛЕНО: уменьшаем И наличие И резерв)
                material_id = int(reservation["ID материала"])
                if material_id != -1:
                    materials_df = load_data("Materials")
                    material = materials_df[materials_df["ID"] == material_id].iloc[0]

                    # Уменьшаем количество в наличии
                    new_qty = int(material["Количество штук"]) - quantity

                    # Уменьшаем зарезервировано
                    new_reserved = int(material["Зарезервировано"]) - quantity

                    # Доступно НЕ меняется (т.к. было уже зарезервировано)

                    materials_df.loc[materials_df["ID"] == material_id, "Количество штук"] = new_qty
                    materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved

                    # Пересчитываем площадь
                    area_per_piece = float(material["Длина"]) * float(material["Ширина"]) / 1_000_000
                    new_area = new_qty * area_per_piece
                    materials_df.loc[materials_df["ID"] == material_id, "Общая площадь"] = round(new_area, 2)

                    save_data("Materials", materials_df)
                    self.refresh_materials()

                self.refresh_reservations()
                self.refresh_writeoffs()
                self.refresh_balance()
                add_window.destroy()
                messagebox.showinfo("Успех", f"✅ Списание #{new_id} успешно создано!\nСписано: {quantity} шт")

            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось создать списание: {e}")
                import traceback
                traceback.print_exc()

        tk.Button(add_window, text="Списать", bg='#e74c3c', fg='white', font=("Arial", 12, "bold"),
                  command=save_writeoff).pack(pady=15)

    def delete_writeoff(self):
        selected = self.writeoffs_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите списания для удаления")
            return
        count = len(selected)
        if messagebox.askyesno("Подтверждение",
                               f"Удалить выбранные списания ({count} шт)?\n\nВнимание: Материал вернется в резерв и на склад!"):
            writeoffs_df = load_data("WriteOffs")
            reservations_df = load_data("Reservations")
            materials_df = load_data("Materials")
            for item in selected:
                writeoff_id = self.writeoffs_tree.item(item)["values"][0]
                writeoff_row = writeoffs_df[writeoffs_df["ID списания"] == writeoff_id].iloc[0]
                reserve_id = writeoff_row["ID резерва"]
                material_id = writeoff_row["ID материала"]
                quantity_to_return = int(writeoff_row["Количество"])
                if not reservations_df[reservations_df["ID резерва"] == reserve_id].empty:
                    res_row = reservations_df[reservations_df["ID резерва"] == reserve_id].iloc[0]
                    reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Списано"] = int(
                        res_row["Списано"]) - quantity_to_return
                    reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = int(
                        res_row["Остаток к списанию"]) + quantity_to_return
                if material_id != -1:
                    if not materials_df[materials_df["ID"] == material_id].empty:
                        mat_row = materials_df[materials_df["ID"] == material_id].iloc[0]
                        current_qty = int(mat_row["Количество штук"])
                        current_reserved = int(mat_row["Зарезервировано"])
                        new_qty = current_qty + quantity_to_return
                        new_reserved = current_reserved + quantity_to_return
                        area = (float(mat_row["Длина"]) * float(mat_row["Ширина"]) * new_qty) / 1000000
                        materials_df.loc[materials_df["ID"] == material_id, "Количество штук"] = new_qty
                        materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved
                        materials_df.loc[materials_df["ID"] == material_id, "Общая площадь"] = round(area, 2)
                        materials_df.loc[materials_df["ID"] == material_id, "Доступно"] = new_qty - new_reserved
                writeoffs_df = writeoffs_df[writeoffs_df["ID списания"] != writeoff_id]
            save_data("WriteOffs", writeoffs_df)
            save_data("Reservations", reservations_df)
            save_data("Materials", materials_df)
            self.refresh_materials()
            self.refresh_reservations()
            self.refresh_writeoffs()
            self.refresh_balance()
            messagebox.showinfo("Успех", f"Отменено списаний: {count}")

    def edit_writeoff(self):
        """Редактирование списания"""
        selected = self.writeoffs_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите списание для редактирования")
            return

        writeoff_id = self.writeoffs_tree.item(selected)["values"][0]
        writeoffs_df = load_data("WriteOffs")
        writeoff_row = writeoffs_df[writeoffs_df["ID списания"] == writeoff_id].iloc[0]

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактировать списание")
        edit_window.geometry("550x650")
        edit_window.configure(bg='#ecf0f1')

        tk.Label(edit_window, text=f"Редактирование списания #{writeoff_id}",
                 font=("Arial", 12, "bold"), bg='#ecf0f1').pack(pady=10)

        # Информация о резерве (только для чтения)
        reserve_id = int(writeoff_row["ID резерва"])
        reservations_df = load_data("Reservations")
        orders_df = load_data("Orders")

        reserve_info = f"Резерв #{reserve_id}"
        order_info = ""
        detail_info = ""

        if not reservations_df.empty:
            reserve_row = reservations_df[reservations_df["ID резерва"] == reserve_id]
            if not reserve_row.empty:
                reserve_data = reserve_row.iloc[0]
                order_id = int(reserve_data["ID заказа"])

                if not orders_df.empty:
                    order_row = orders_df[orders_df["ID заказа"] == order_id]
                    if not order_row.empty:
                        customer = order_row.iloc[0]["Заказчик"]
                        order_name = order_row.iloc[0]["Название заказа"]
                        order_info = f"{customer} | {order_name}"

                detail_name = reserve_data.get("Название детали", "Без детали")
                if pd.notna(detail_name) and detail_name != "" and detail_name != "Не указана":
                    detail_info = f"Деталь: {detail_name}"
                else:
                    detail_info = "Без привязки к детали"

        info_frame = tk.LabelFrame(edit_window, text="Информация (не редактируется)",
                                   bg='#e8f4f8', font=("Arial", 9, "bold"))
        info_frame.pack(fill=tk.X, padx=20, pady=10)

        if order_info:
            tk.Label(info_frame, text=f"Заказ: {order_info}", bg='#e8f4f8', font=("Arial", 9)).pack(padx=10, pady=2,
                                                                                                    anchor='w')
        if detail_info:
            tk.Label(info_frame, text=detail_info, bg='#e8f4f8', font=("Arial", 9)).pack(padx=10, pady=2, anchor='w')

        material_info = f"{writeoff_row['Марка']} {writeoff_row['Толщина']}мм {writeoff_row['Ширина']}x{writeoff_row['Длина']}"
        tk.Label(info_frame, text=f"Материал: {material_info}", bg='#e8f4f8', font=("Arial", 9)).pack(padx=10, pady=2,
                                                                                                      anchor='w')
        tk.Label(info_frame, text=f"Дата списания: {writeoff_row['Дата списания']}",
                 bg='#e8f4f8', font=("Arial", 9)).pack(padx=10, pady=2, anchor='w')

        # Редактируемое поле: Количество
        qty_frame = tk.Frame(edit_window, bg='#ecf0f1')
        qty_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(qty_frame, text="Количество (шт):", width=25, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.insert(0, str(int(writeoff_row["Количество"])))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Редактируемое поле: Комментарий
        comment_frame = tk.Frame(edit_window, bg='#ecf0f1')
        comment_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(comment_frame, text="Комментарий:", width=25, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
        comment_entry = tk.Entry(comment_frame, font=("Arial", 10))
        comment_entry.insert(0, str(writeoff_row["Комментарий"]))
        comment_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # Информация о резерве
        if not reservations_df.empty and not reserve_row.empty:
            reserve_data = reserve_row.iloc[0]
            reserve_total = int(reserve_data["Зарезервировано штук"])
            reserve_written = int(reserve_data["Списано"])
            reserve_remainder = int(reserve_data["Остаток к списанию"])

            stats_frame = tk.LabelFrame(edit_window, text="Статистика резерва",
                                        bg='#fff3cd', font=("Arial", 9, "bold"))
            stats_frame.pack(fill=tk.X, padx=20, pady=10)
            tk.Label(stats_frame, text=f"Всего в резерве: {reserve_total} шт",
                     bg='#fff3cd', font=("Arial", 9)).pack(anchor='w', padx=10, pady=2)
            tk.Label(stats_frame, text=f"Списано всего: {reserve_written} шт",
                     bg='#fff3cd', font=("Arial", 9)).pack(anchor='w', padx=10, pady=2)
            tk.Label(stats_frame, text=f"Остаток к списанию: {reserve_remainder} шт",
                     bg='#fff3cd', font=("Arial", 9)).pack(anchor='w', padx=10, pady=2)

        # Предупреждение
        warning_frame = tk.Frame(edit_window, bg='#ffcccc', relief=tk.RIDGE, borderwidth=2)
        warning_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(warning_frame, text="ВАЖНО!", font=("Arial", 9, "bold"),
                 bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=5, pady=2)
        tk.Label(warning_frame, text="• Изменение количества пересчитает баланс материалов",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)
        tk.Label(warning_frame, text="• Изменение влияет на остаток резерва к списанию",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)

        def save_changes():
            try:
                new_qty = int(qty_entry.get().strip())
                new_comment = comment_entry.get().strip()

                if new_qty <= 0:
                    messagebox.showerror("Ошибка", "Количество должно быть больше нуля!")
                    return

                old_qty = int(writeoff_row["Количество"])
                difference = new_qty - old_qty

                # Проверяем, не превысит ли новое количество доступный остаток резерва
                if not reservations_df.empty and not reserve_row.empty:
                    reserve_data = reserve_row.iloc[0]
                    reserve_remainder = int(reserve_data["Остаток к списанию"])

                    # Доступно = текущий остаток + старое списание
                    max_available = reserve_remainder + old_qty

                    if new_qty > max_available:
                        messagebox.showerror("Ошибка",
                                             f"Нельзя списать {new_qty} шт!\n"
                                             f"Максимально доступно: {max_available} шт")
                        return

                if difference == 0 and new_comment == str(writeoff_row["Комментарий"]):
                    messagebox.showinfo("Информация", "Изменений не было")
                    edit_window.destroy()
                    return

                # Подтверждение
                msg = f"Сохранить и��менения?\n\n"
                if difference != 0:
                    msg += f"Количество: {old_qty} → {new_qty} шт (разница: {'+' if difference > 0 else ''}{difference})\n"
                if new_comment != str(writeoff_row["Комментарий"]):
                    msg += f"Комментарий изменен"

                if not messagebox.askyesno("Подтверждение", msg):
                    return

                # Обновляем списание
                writeoffs_df.loc[writeoffs_df["ID списания"] == writeoff_id, "Количество"] = new_qty
                writeoffs_df.loc[writeoffs_df["ID списания"] == writeoff_id, "Комментарий"] = new_comment
                save_data("WriteOffs", writeoffs_df)

                # Если количество изменилось - обновляем резерв и материал
                if difference != 0:
                    # Обновляем резерв
                    if not reservations_df.empty and not reserve_row.empty:
                        reserve_data = reserve_row.iloc[0]
                        current_written = int(reserve_data["Списано"])
                        current_remainder = int(reserve_data["Остаток к списанию"])

                        new_written = current_written + difference
                        new_remainder = current_remainder - difference

                        reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Списано"] = new_written
                        reservations_df.loc[
                            reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = new_remainder
                        save_data("Reservations", reservations_df)

                    # Обновляем материал (если не вручную добавленный)
                    material_id = int(writeoff_row["ID материала"])
                    if material_id != -1:
                        materials_df = load_data("Materials")
                        if not materials_df[materials_df["ID"] == material_id].empty:
                            mat_row = materials_df[materials_df["ID"] == material_id].iloc[0]
                            current_qty = int(mat_row["Количество штук"])
                            current_reserved = int(mat_row["Зарезервировано"])

                            # Разница списания влияет на количество и резерв
                            new_mat_qty = current_qty - difference
                            new_reserved = current_reserved - difference

                            materials_df.loc[materials_df["ID"] == material_id, "Количество штук"] = new_mat_qty
                            materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved

                            # Пересчитываем площадь
                            area_per_piece = float(mat_row["Длина"]) * float(mat_row["Ширина"]) / 1_000_000
                            new_area = new_mat_qty * area_per_piece
                            materials_df.loc[materials_df["ID"] == material_id, "Общая площадь"] = round(new_area, 2)

                            save_data("Materials", materials_df)
                            self.refresh_materials()

                self.refresh_reservations()
                self.refresh_writeoffs()
                self.refresh_balance()
                edit_window.destroy()
                messagebox.showinfo("Успех", f"Списание #{writeoff_id} обновлено!")

            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить списание: {e}")
                import traceback
                traceback.print_exc()

        tk.Button(edit_window, text="Сохранить изменения", bg='#f39c12', fg='white',
                  font=("Arial", 12, "bold"), command=save_changes).pack(pady=15)

    def create_visibility_toggles(self, parent_frame, tree_widget, toggles_config, refresh_callback):
        """Создание переключателей видимости для таблиц"""
        toggles_frame = tk.Frame(parent_frame, bg='white')
        toggles_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(toggles_frame, text="Фильтры:", bg='white', font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)

        toggle_vars = {}

        for key, label_text in toggles_config.items():
            var = tk.BooleanVar(value=True)
            checkbox = tk.Checkbutton(toggles_frame, text=label_text, variable=var,
                                      bg='white', font=("Arial", 9),
                                      command=refresh_callback)
            checkbox.pack(side=tk.LEFT, padx=10)
            toggle_vars[key] = var

        return toggle_vars

    def setup_balance_tab(self):
        header = tk.Label(self.balance_frame, text="Баланс материалов", font=("Arial", 16, "bold"), bg='white',
                          fg='#2c3e50')
        header.pack(pady=10)
        info_label = tk.Label(self.balance_frame, text="Красный - не хватает | Желтый - на нуле | Зеленый - в наличии",
                              font=("Arial", 10), bg='white', fg='#7f8c8d')
        info_label.pack(pady=5)
        tree_frame = tk.Frame(self.balance_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        self.balance_tree = ttk.Treeview(tree_frame,
                                         columns=("Материал", "Марка", "Толщина", "Размер", "В наличии",
                                                  "Зарезервировано", "Итого"),
                                         show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.balance_tree.yview)
        scroll_x.config(command=self.balance_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        columns_config = {"Материал": 100, "Марка": 120, "Толщина": 100, "Размер": 150,
                          "В наличии": 100, "Зарезервировано": 130, "Итого": 100}
        for col, width in columns_config.items():
            self.balance_tree.heading(col, text=col)
            self.balance_tree.column(col, width=width, anchor=tk.CENTER)
        self.balance_tree.pack(fill=tk.BOTH, expand=True)
        self.balance_toggles = self.create_visibility_toggles(
            self.balance_frame,
            self.balance_tree,
            {
                'show_negative': '🔴 Показать отрицательные',
                'show_zero': '🟡 Показать нулевые',
                'show_positive': '🟢 Показать положительные'
            },
            self.refresh_balance
        )
        self.balance_tree.tag_configure('negative', background='#ffcccc')
        self.balance_tree.tag_configure('zero', background='#fff9c4')
        self.balance_tree.tag_configure('positive', background='#c8e6c9')


        buttons_frame = tk.Frame(self.balance_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)
        btn_style = {"font": ("Arial", 10), "width": 15, "height": 2}
        tk.Button(buttons_frame, text="Обновить", bg='#95a5a6', fg='white', command=self.refresh_balance,
                  **btn_style).pack(side=tk.LEFT, padx=5)
        self.refresh_balance()

    def refresh_balance(self):
        # Удаляем ВСЕ строки из баланса
        for i in self.balance_tree.get_children():
            self.balance_tree.delete(i)

        # Загружаем актуальные данные со склада
        materials_df = load_data("Materials")

        # Если материалов нет - выходим (таблица уже пустая)
        if materials_df.empty:
            return

        # Получаем список существующих ID материалов
        existing_material_ids = set(materials_df["ID"].astype(int).tolist())

        # Получаем состояния фильтров
        show_negative = True
        show_zero = True
        show_positive = True

        if hasattr(self, 'balance_toggles') and self.balance_toggles:
            show_negative = self.balance_toggles.get('show_negative', tk.BooleanVar(value=True)).get()
            show_zero = self.balance_toggles.get('show_zero', tk.BooleanVar(value=True)).get()
            show_positive = self.balance_toggles.get('show_positive', tk.BooleanVar(value=True)).get()

        # Проходим ТОЛЬКО по материалам, которые есть на складе
        for _, row in materials_df.iterrows():
            material_id = int(row["ID"])

            # ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА: Материал должен быть в списке существующих
            if material_id not in existing_material_ids:
                continue

            qty = int(row["Количество штук"])
            reserved = int(row["Зарезервировано"])

            # Итого = В наличии - Зарезервировано
            total = qty - reserved

            # Применяем фильтры
            if total < 0 and not show_negative:
                continue
            if total == 0 and not show_zero:
                continue
            if total > 0 and not show_positive:
                continue

            size_str = f"{row['Ширина']} x {row['Длина']}"

            values = [
                f"ID: {material_id}",
                row["Марка"],
                f"{row['Толщина']} мм",
                size_str,
                qty,
                reserved,
                total
            ]

            # Определяем цвет строки
            if total < 0:
                tag = 'negative'
            elif total == 0:
                tag = 'zero'
            else:
                tag = 'positive'

            self.balance_tree.insert("", "end", values=values, tags=(tag,))

        print(
            f"[Баланс] Обновлено. Материалов на складе: {len(materials_df)}, Отображено в балансе: {len(self.balance_tree.get_children())}")

if __name__ == "__main__":
    try:
        initialize_database()
        root = tk.Tk()
        app = ProductionApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Критическая ошибка: {e}")
        import traceback

        traceback.print_exc()
        messagebox.showerror("Критическая ошибка", str(e))