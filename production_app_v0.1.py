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

        # 🆕 Инициализация данных для импорта от лазерщиков
        self.laser_table_data = []

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

        self.laser_import_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.laser_import_frame, text='Импорт от лазерщиков')
        self.setup_laser_import_tab()

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

        details_label = tk.Label(self.orders_frame,
                                 text="Детали выбранного заказа (дважды кликните «Порезано» или «Погнуто» для редактирования)",
                                 font=("Arial", 11, "bold"), bg='white', fg='#2c3e50')
        details_label.pack(pady=5)
        details_tree_frame = tk.Frame(self.orders_frame, bg='white')
        details_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y2 = tk.Scrollbar(details_tree_frame, orient=tk.VERTICAL)
        self.order_details_tree = ttk.Treeview(details_tree_frame,
                                               columns=("ID", "ID заказа", "Название детали", "Количество", "Порезано",
                                                        "Погнуто"),
                                               show="headings", yscrollcommand=scroll_y2.set)
        scroll_y2.config(command=self.order_details_tree.yview)
        scroll_y2.pack(side=tk.RIGHT, fill=tk.Y)
        for col in ["ID", "ID заказа", "Название детали", "Количество", "Порезано", "Погнуто"]:
            self.order_details_tree.heading(col, text=col)
            self.order_details_tree.column(col, width=150, anchor=tk.CENTER)
        self.order_details_tree.pack(fill=tk.BOTH, expand=True)

        # Привязываем двойной клик для редактирования
        self.order_details_tree.bind('<Double-1>', self.on_detail_double_click)

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

    def on_detail_double_click(self, event):
        """Обработка двойного клика по детали для редактирования прямо в таблице"""
        try:
            region = self.order_details_tree.identify("region", event.x, event.y)
            if region != "cell":
                return

            # Определяем колонку
            column = self.order_details_tree.identify_column(event.x)
            if not column:
                return

            # Преобразуем #1, #2, #3 в индекс 0, 1, 2
            column_index = int(column.replace('#', '')) - 1

            # Проверяем что индекс в пределах
            columns = self.order_details_tree['columns']
            if column_index < 0 or column_index >= len(columns):
                return

            column_name = columns[column_index]

            # Разрешаем редактировать только Порезано и Погнуто
            if column_name not in ["Порезано", "Погнуто"]:
                return

            # Определяем строку
            item = self.order_details_tree.identify_row(event.y)
            if not item:
                return

            # Получаем данные строки
            values = self.order_details_tree.item(item, 'values')
            if not values or len(values) < 6:
                return

            try:
                detail_id = int(values[0])
            except (ValueError, TypeError):
                messagebox.showerror("Ошибка", "Не удалось определить ID детали")
                return

            # СРАЗУ ПРОВЕРЯЕМ существование детали в базе
            df = load_data("OrderDetails")
            if df.empty:
                messagebox.showwarning("Предупреждение", "Таблица деталей пуста")
                return

            detail_exists = df[df["ID"] == detail_id]
            if detail_exists.empty:
                messagebox.showerror("Ошибка",
                                     f"Деталь ID {detail_id} не найдена в базе данных!\n\n"
                                     f"Возможно данные устарели. Нажмите 'Обновить'.")
                self.refresh_order_details()
                return

            detail_name = values[2]

            try:
                total_qty = int(values[3])
                current_cut = int(values[4]) if values[4] and str(values[4]).strip() != '' else 0
                current_bent = int(values[5]) if values[5] and str(values[5]).strip() != '' else 0
            except (ValueError, IndexError):
                messagebox.showerror("Ошибка", "Не удалось прочитать значения детали")
                return

            # Получаем координаты ячейки
            x, y, width, height = self.order_details_tree.bbox(item, column)

            # Создаем Entry для редактирования
            edit_entry = tk.Entry(self.order_details_tree, font=("Arial", 10))
            edit_entry.place(x=x, y=y, width=width, height=height)

            # Вставляем текущее значение
            current_value = values[column_index]
            edit_entry.insert(0, str(current_value))
            edit_entry.select_range(0, tk.END)
            edit_entry.focus()

            def save_cell_edit(event=None):
                try:
                    new_value_str = edit_entry.get().strip()
                    if not new_value_str:
                        new_value = 0
                    else:
                        new_value = int(new_value_str)

                    if new_value < 0:
                        messagebox.showerror("Ошибка", "Значение не может быть отрицательным!")
                        edit_entry.destroy()
                        return

                    # ПЕРЕЗАГРУЖАЕМ данные для актуальности
                    df = load_data("OrderDetails")
                    if df.empty:
                        messagebox.showerror("Ошибка", "Не удалось загрузить детали")
                        edit_entry.destroy()
                        return

                    # ПРОВЕРЯЕМ существование детали ЕЩЕ РАЗ
                    detail_row = df[df["ID"] == detail_id]
                    if detail_row.empty:
                        messagebox.showerror("Ошибка",
                                             f"Деталь ID {detail_id} была удалена!\n\n"
                                             f"Обновите список деталей.")
                        edit_entry.destroy()
                        self.refresh_order_details()
                        return

                    # Получаем актуальные данные из базы
                    actual_row = detail_row.iloc[0]
                    actual_cut = int(actual_row.get("Порезано", 0)) if pd.notna(actual_row.get("Порезано")) else 0
                    actual_bent = int(actual_row.get("Погнуто", 0)) if pd.notna(actual_row.get("Погнуто")) else 0
                    actual_qty = int(actual_row["Количество"])

                    # Определяем что редактируем
                    if column_name == "Порезано":
                        new_cut = new_value
                        new_bent = actual_bent

                        if new_cut < new_bent:
                            if not messagebox.askyesno("Предупреждение",
                                                       f"Порезано ({new_cut}) меньше погнутого ({new_bent}).\n"
                                                       f"Это означает, что погнуто больше заготовок чем есть.\n\n"
                                                       f"Продолжить?"):
                                edit_entry.destroy()
                                return

                        if new_cut > actual_qty:
                            if not messagebox.askyesno("Предупреждение",
                                                       f"Порезано ({new_cut}) больше общего количества ({actual_qty}).\n"
                                                       f"Возможно есть излишки заготовок.\n\n"
                                                       f"Продолжить?"):
                                edit_entry.destroy()
                                return

                        df.loc[df["ID"] == detail_id, "Порезано"] = new_cut

                    elif column_name == "Погнуто":
                        new_cut = actual_cut
                        new_bent = new_value

                        if new_bent > new_cut:
                            if not messagebox.askyesno("Предупреждение",
                                                       f"Погнуто ({new_bent}) больше порезанного ({new_cut}).\n"
                                                       f"Нужно сначала порезать заготовки.\n\n"
                                                       f"Продолжить?"):
                                edit_entry.destroy()
                                return

                        df.loc[df["ID"] == detail_id, "Погнуто"] = new_bent

                    # Сохраняем
                    save_data("OrderDetails", df)
                    self.refresh_order_details()
                    edit_entry.destroy()

                    # Показываем краткое уведомление
                    to_cut = actual_qty - new_cut
                    to_bend = new_cut - new_bent

                    status_msg = f"✅ {detail_name}\n"
                    status_msg += f"Порезано: {new_cut}/{actual_qty} (осталось: {to_cut})\n"
                    status_msg += f"Погнуто: {new_bent}/{new_cut} (осталось: {to_bend})"

                    self.show_status_tooltip(status_msg)

                except ValueError:
                    messagebox.showerror("Ошибка", "Введите корректное число!")
                    edit_entry.destroy()
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Не удалось обновить: {e}")
                    edit_entry.destroy()
                    import traceback
                    traceback.print_exc()

            # Привязываем события
            edit_entry.bind('<Return>', save_cell_edit)
            edit_entry.bind('<FocusOut>', save_cell_edit)
            edit_entry.bind('<Escape>', lambda e: edit_entry.destroy())

        except Exception as e:
            print(f"Ошибка в on_detail_double_click: {e}")
            import traceback
            traceback.print_exc()

    def show_status_tooltip(self, message):
        """Показывает временное всплывающее окно со статусом"""
        try:
            tooltip = tk.Toplevel(self.root)
            tooltip.wm_overrideredirect(True)

            # Позиционируем окно рядом с курсором
            x = self.root.winfo_pointerx() + 10
            y = self.root.winfo_pointery() + 10
            tooltip.wm_geometry(f"+{x}+{y}")

            label = tk.Label(tooltip, text=message, background="#d4edda",
                             foreground="#155724", relief=tk.SOLID, borderwidth=1,
                             font=("Arial", 9), padx=10, pady=5, justify=tk.LEFT)
            label.pack()

            # Автоматически закрываем через 2 секунды
            tooltip.after(2000, tooltip.destroy)
        except Exception as e:
            print(f"Ошибка в show_status_tooltip: {e}")

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
        """Редактирование резервирования с возможностью изменения заказа и детали"""
        selected = self.reservations_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите резерв для редактирования")
            return

        reserve_id = self.reservations_tree.item(selected)["values"][0]
        reservations_df = load_data("Reservations")
        reserve_row = reservations_df[reservations_df["ID резерва"] == reserve_id].iloc[0]

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактировать резерв")
        edit_window.geometry("650x800")
        edit_window.configure(bg='#ecf0f1')

        tk.Label(edit_window, text=f"Редактирование резерва #{reserve_id}",
                 font=("Arial", 12, "bold"), bg='#ecf0f1', fg='#2c3e50').pack(pady=10)

        # Загружаем данные
        orders_df = load_data("Orders")
        order_details_df = load_data("OrderDetails")

        # Текущие данные резерва
        current_order_id = int(reserve_row["ID заказа"])
        current_detail_id = reserve_row.get("ID детали", -1)
        if pd.isna(current_detail_id):
            current_detail_id = -1
        else:
            current_detail_id = int(current_detail_id)

        written_off = int(reserve_row["Списано"])

        # === ЗАКАЗ ===
        order_frame = tk.LabelFrame(edit_window, text="Заказ", bg='#ecf0f1', font=("Arial", 10, "bold"))
        order_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(order_frame, text="Выберите заказ:", bg='#ecf0f1', font=("Arial", 9)).pack(anchor='w', padx=10, pady=5)

        # Формируем список заказов
        all_order_options = []
        order_map = {}

        if not orders_df.empty:
            for _, row in orders_df.iterrows():
                order_id = int(row['ID заказа'])
                display_text = f"ID:{order_id} | {row['Заказчик']} | {row['Название заказа']}"
                all_order_options.append(display_text)
                order_map[display_text] = order_id

        order_search_var = tk.StringVar()
        order_search_entry = tk.Entry(order_frame, textvariable=order_search_var, font=("Arial", 9))
        order_search_entry.pack(fill=tk.X, padx=10, pady=5)

        order_listbox = tk.Listbox(order_frame, height=4, font=("Arial", 9))
        order_listbox.pack(fill=tk.BOTH, padx=10, pady=5)

        for option in all_order_options:
            order_listbox.insert(tk.END, option)

        selected_order = {"value": None, "id": current_order_id}

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
                selected_order["id"] = order_map[selection]
                order_search_var.set(selection)
                update_details_list()
            except:
                pass

        order_search_var.trace('w', on_order_search)
        order_listbox.bind('<<ListboxSelect>>', on_select_order)
        order_listbox.bind('<Double-Button-1>', on_select_order)

        # Устанавливаем текущий заказ
        for i, option in enumerate(all_order_options):
            if order_map[option] == current_order_id:
                order_listbox.selection_set(i)
                order_listbox.see(i)
                order_search_var.set(option)
                selected_order["value"] = option
                break

        # === ДЕТАЛЬ ===
        detail_frame = tk.LabelFrame(edit_window, text="Деталь", bg='#ecf0f1', font=("Arial", 10, "bold"))
        detail_frame.pack(fill=tk.X, padx=20, pady=10)

        tk.Label(detail_frame, text="Выберите деталь:", bg='#ecf0f1', font=("Arial", 9)).pack(anchor='w', padx=10,
                                                                                              pady=5)

        detail_var = tk.StringVar()
        detail_combo = ttk.Combobox(detail_frame, textvariable=detail_var, font=("Arial", 9),
                                    state="readonly", width=50)
        detail_combo.pack(fill=tk.X, padx=10, pady=5)

        selected_detail = {"id": current_detail_id, "name": None}

        def update_details_list():
            detail_combo['values'] = []
            detail_var.set("")
            selected_detail["id"] = -1
            selected_detail["name"] = None

            order_id = selected_order["id"]
            if not order_id:
                return

            try:
                if not order_details_df.empty:
                    details = order_details_df[order_details_df["ID заказа"] == order_id]

                    if not details.empty:
                        detail_options = ["[Без привязки к детали]"]
                        detail_options.extend([f"ID:{int(row['ID'])} - {row['Название детали']}"
                                               for _, row in details.iterrows()])
                        detail_combo['values'] = detail_options

                        # Пытаемся установить текущую деталь
                        if current_detail_id != -1:
                            for opt in detail_options:
                                if opt.startswith(f"ID:{current_detail_id} -"):
                                    detail_combo.set(opt)
                                    selected_detail["id"] = current_detail_id
                                    selected_detail["name"] = opt.split(" - ")[1]
                                    break
                        else:
                            detail_combo.current(0)
                    else:
                        detail_combo['values'] = ["[Нет деталей у заказа]"]
                        detail_combo.current(0)
                else:
                    detail_combo['values'] = ["[Нет деталей у заказа]"]
                    detail_combo.current(0)
            except Exception as e:
                print(f"Ошибка обновления списка деталей: {e}")

        def on_detail_select(event):
            value = detail_var.get()
            if value and value.startswith("ID:"):
                try:
                    selected_detail["id"] = int(value.split("ID:")[1].split(" - ")[0])
                    selected_detail["name"] = value.split(" - ")[1]
                except:
                    selected_detail["id"] = -1
                    selected_detail["name"] = None
            else:
                selected_detail["id"] = -1
                selected_detail["name"] = None

        detail_combo.bind('<<ComboboxSelected>>', on_detail_select)

        # Инициализируем список деталей
        update_details_list()

        # === МАТЕРИАЛ (только для чтения) ===
        material_frame = tk.LabelFrame(edit_window, text="Материал (не редактируется)",
                                       bg='#e8f4f8', font=("Arial", 9, "bold"))
        material_frame.pack(fill=tk.X, padx=20, pady=10)

        material_info = f"{reserve_row['Марка']} {reserve_row['Толщина']}мм {reserve_row['Ширина']}x{reserve_row['Длина']}"
        tk.Label(material_frame, text=material_info, bg='#e8f4f8', font=("Arial", 10)).pack(padx=10, pady=5)

        # === КОЛИЧЕСТВО ===
        qty_frame = tk.Frame(edit_window, bg='#ecf0f1')
        qty_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(qty_frame, text="Зарезервировано (шт):", width=25, anchor='w',
                 bg='#ecf0f1', font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        qty_entry = tk.Entry(qty_frame, font=("Arial", 10))
        qty_entry.insert(0, str(int(reserve_row["Зарезервировано штук"])))
        qty_entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)

        # === СТАТИСТИКА ===
        remainder = int(reserve_row["Остаток к списанию"])

        stats_frame = tk.LabelFrame(edit_window, text="Статистика", bg='#fff3cd', font=("Arial", 9, "bold"))
        stats_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(stats_frame, text=f"Уже списано: {written_off} шт",
                 bg='#fff3cd', font=("Arial", 9)).pack(anchor='w', padx=10, pady=2)
        tk.Label(stats_frame, text=f"Остаток к списанию: {remainder} шт",
                 bg='#fff3cd', font=("Arial", 9)).pack(anchor='w', padx=10, pady=2)

        # === ПРЕДУПРЕЖДЕНИЕ ===
        warning_frame = tk.Frame(edit_window, bg='#ffcccc', relief=tk.RIDGE, borderwidth=2)
        warning_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(warning_frame, text="⚠ ВАЖНО!", font=("Arial", 9, "bold"),
                 bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=5, pady=2)
        tk.Label(warning_frame, text="• Нельзя уменьшить количество ниже уже списанного",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)
        tk.Label(warning_frame, text="• Можно изменить заказ и деталь",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)
        tk.Label(warning_frame, text="• Изменение количества влияет на баланс материалов",
                 font=("Arial", 8), bg='#ffcccc', fg='#c0392b').pack(anchor='w', padx=10)

        def save_changes():
            try:
                new_qty = int(qty_entry.get().strip())
                new_order_id = selected_order["id"]
                new_detail_id = selected_detail["id"]
                new_detail_name = selected_detail["name"] if selected_detail["name"] else "Не указана"

                if not new_order_id:
                    messagebox.showwarning("Предупреждение", "Выберите заказ!")
                    return

                if new_qty < written_off:
                    messagebox.showerror("Ошибка",
                                         f"Нельзя установить количество ({new_qty}) меньше уже списанного ({written_off})!")
                    return

                if new_qty <= 0:
                    messagebox.showerror("Ошибка", "Количество должно быть больше нуля!")
                    return

                old_qty = int(reserve_row["Зарезервировано штук"])
                qty_difference = new_qty - old_qty

                # Проверяем изменения
                order_changed = new_order_id != current_order_id
                detail_changed = new_detail_id != current_detail_id
                qty_changed = qty_difference != 0

                if not order_changed and not detail_changed and not qty_changed:
                    messagebox.showinfo("Информация", "Изменений не было")
                    edit_window.destroy()
                    return

                # Формируем сообщение с изменениями
                changes_msg = "Будут внесены следующие изменения:\n\n"

                if order_changed:
                    old_order = orders_df[orders_df["ID заказа"] == current_order_id].iloc[0]
                    new_order = orders_df[orders_df["ID заказа"] == new_order_id].iloc[0]
                    changes_msg += f"📋 Заказ:\n"
                    changes_msg += f"  Старый: {old_order['Заказчик']} | {old_order['Название заказа']}\n"
                    changes_msg += f"  Новый: {new_order['Заказчик']} | {new_order['Название заказа']}\n\n"

                if detail_changed:
                    old_detail_name = reserve_row.get("Название детали", "Не указана")
                    if pd.isna(old_detail_name) or old_detail_name == "":
                        old_detail_name = "Не указана"
                    changes_msg += f"🔧 Деталь:\n"
                    changes_msg += f"  Старая: {old_detail_name}\n"
                    changes_msg += f"  Новая: {new_detail_name}\n\n"

                if qty_changed:
                    changes_msg += f"📦 Количество:\n"
                    changes_msg += f"  Старое: {old_qty} шт\n"
                    changes_msg += f"  Новое: {new_qty} шт\n"
                    changes_msg += f"  Разница: {'+' if qty_difference > 0 else ''}{qty_difference} шт\n"
                    changes_msg += f"  Новый остаток к списанию: {new_qty - written_off} шт\n\n"

                changes_msg += "Продолжить?"

                if not messagebox.askyesno("Подтверждение изменений", changes_msg):
                    return

                # Обновляем резерв
                new_remainder = new_qty - written_off
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "ID заказа"] = new_order_id
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "ID детали"] = new_detail_id
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Название детали"] = new_detail_name
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Зарезервировано штук"] = new_qty
                reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = new_remainder
                save_data("Reservations", reservations_df)

                # Обновляем материал на складе (если количество изменилось и не вручную добавленный)
                if qty_changed:
                    material_id = int(reserve_row["ID материала"])
                    if material_id != -1:
                        materials_df = load_data("Materials")
                        if not materials_df[materials_df["ID"] == material_id].empty:
                            mat_row = materials_df[materials_df["ID"] == material_id].iloc[0]
                            current_reserved = int(mat_row["Зарезервировано"])
                            current_available = int(mat_row["Доступно"])

                            new_reserved = current_reserved + qty_difference
                            new_available = current_available - qty_difference

                            materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved
                            materials_df.loc[materials_df["ID"] == material_id, "Доступно"] = new_available
                            save_data("Materials", materials_df)
                            self.refresh_materials()

                self.refresh_reservations()
                self.refresh_balance()
                edit_window.destroy()

                result_msg = f"✅ Резерв #{reserve_id} обновлен!\n\n"
                if order_changed:
                    result_msg += "📋 Заказ изменен\n"
                if detail_changed:
                    result_msg += f"🔧 Деталь изменена на: {new_detail_name}\n"
                if qty_changed:
                    result_msg += f"📦 Количество: {new_qty} шт (остаток: {new_remainder} шт)\n"

                messagebox.showinfo("Успех", result_msg)

            except ValueError:
                messagebox.showerror("Ошибка", "Проверьте правильность ввода числовых значений!")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить резерв: {e}")
                import traceback
                traceback.print_exc()

        tk.Button(edit_window, text="💾 Сохранить изменения", bg='#f39c12', fg='white',
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
        """Вкладка списания материалов - РУЧНОЕ списание (совместима с импортом от лазерщиков)"""
        header = tk.Label(self.writeoffs_frame, text="Списание зарезервированных материалов",
                          font=("Arial", 16, "bold"), bg='white', fg='#2c3e50')
        header.pack(pady=10)

        tree_frame = tk.Frame(self.writeoffs_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        self.writeoffs_tree = ttk.Treeview(tree_frame,
                                           columns=("ID", "ID резерва", "Заказ", "Деталь", "Материал", "Марка",
                                                    "Толщина", "Размер", "Количество", "Дата", "Комментарий"),
                                           show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.config(command=self.writeoffs_tree.yview)
        scroll_x.config(command=self.writeoffs_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        columns_config = {
            "ID": 50, "ID резерва": 80, "Заказ": 200, "Деталь": 150,
            "Материал": 80, "Марка": 90, "Толщина": 70, "Размер": 110,
            "Количество": 90, "Дата": 140, "Комментарий": 180
        }

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

        # Кнопки управления
        buttons_frame = tk.Frame(self.writeoffs_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        btn_style = {"font": ("Arial", 10), "width": 18, "height": 2}

        tk.Button(buttons_frame, text="Списать материал", bg='#e67e22', fg='white',
                  command=self.add_writeoff, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="Удалить списание", bg='#e74c3c', fg='white',
                  command=self.delete_writeoff, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="Редактировать", bg='#f39c12', fg='white',
                  command=self.edit_writeoff, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="Обновить", bg='#95a5a6', fg='white',
                  command=self.refresh_writeoffs, **btn_style).pack(side=tk.LEFT, padx=5)

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
                    order_display,
                    detail_display,
                    row["ID материала"],
                    row["Марка"],
                    row["Толщина"],
                    size_str,
                    row["Количество"],
                    row["Дата списания"],
                    row["Комментарий"]
                ]

                self.writeoffs_tree.insert("", "end", values=values)

            self.auto_resize_columns(self.writeoffs_tree)  # ИСПРАВЛЕНО: убрана лишняя скобка

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
        """Удаление записи о списании (отмена списания)"""
        selected = self.writeoffs_tree.selection()

        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите списание для удаления!")
            return

        # Получаем информацию о выбранном списании
        try:
            values = self.writeoffs_tree.item(selected[0])['values']
            writeoff_id = int(values[0])
            reserve_id = int(values[1])
            comment = values[9] if len(values) > 9 else ""

            # Показываем информацию о списании
            info_msg = (
                f"Отменить списание?\n\n"
                f"ID списания: {writeoff_id}\n"
                f"ID резерва: {reserve_id}\n"
                f"Комментарий: {comment}\n\n"
                f"⚠️ Это действие:\n"
                f"• Вернёт материал в резерв\n"
                f"• Вернёт материал на склад\n"
                f"• Обновит таблицу импорта от лазерщиков"
            )

            if not messagebox.askyesno("Подтверждение", info_msg):
                return

            print(f"\n{'=' * 80}")
            print(f"🔵 ОТМЕНА СПИСАНИЯ ID={writeoff_id}")
            print(f"{'=' * 80}")

            # Загружаем данные
            writeoffs_df = load_data("WriteOffs")
            reservations_df = load_data("Reservations")
            materials_df = load_data("Materials")

            # Находим запись списания
            writeoff_row = writeoffs_df[writeoffs_df["ID списания"] == writeoff_id]

            if writeoff_row.empty:
                messagebox.showerror("Ошибка", f"Списание ID={writeoff_id} не найдено!")
                return

            writeoff_row = writeoff_row.iloc[0]

            reserve_id = int(writeoff_row["ID резерва"])
            quantity = int(writeoff_row["Количество"])
            material_id = int(writeoff_row["ID материала"])
            writeoff_date = writeoff_row["Дата списания"]
            writeoff_comment = writeoff_row["Комментарий"]

            print(f"📋 Информация о списании:")
            print(f"   Резерв: {reserve_id}")
            print(f"   Материал: {material_id}")
            print(f"   Количество: {quantity}")
            print(f"   Дата: {writeoff_date}")
            print(f"   Комментарий: {writeoff_comment}")

            # ========== ШАГ 1: ОБНОВЛЕНИЕ РЕЗЕРВА ==========
            reserve_row = reservations_df[reservations_df["ID резерва"] == reserve_id]

            if reserve_row.empty:
                messagebox.showerror("Ошибка", f"Резерв ID={reserve_id} не найден!")
                return

            reserve_row = reserve_row.iloc[0]
            old_written_off = int(reserve_row["Списано"])
            old_remainder = int(reserve_row["Остаток к списанию"])

            new_written_off = old_written_off - quantity
            new_remainder = old_remainder + quantity

            reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Списано"] = new_written_off
            reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = new_remainder

            print(f"\n✅ Резерв обновлён:")
            print(f"   Было списано: {old_written_off} → {new_written_off}")
            print(f"   Остаток: {old_remainder} → {new_remainder}")

            # ========== ШАГ 2: ОБНОВЛЕНИЕ МАТЕРИАЛА НА СКЛАДЕ ==========
            if material_id != -1:
                material = materials_df[materials_df["ID"] == material_id]

                if not material.empty:
                    material = material.iloc[0]

                    old_qty = int(material["Количество штук"])
                    old_reserved = int(material["Зарезервировано"])

                    new_qty = old_qty + quantity
                    new_reserved = old_reserved + quantity

                    materials_df.loc[materials_df["ID"] == material_id, "Количество штук"] = new_qty
                    materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved

                    # Пересчитываем площадь
                    area_per_piece = float(material["Длина"]) * float(material["Ширина"]) / 1_000_000
                    new_area = new_qty * area_per_piece
                    materials_df.loc[materials_df["ID"] == material_id, "Общая площадь"] = round(new_area, 2)

                    print(f"\n✅ Материал на складе обновлён:")
                    print(f"   Всего: {old_qty} → {new_qty}")
                    print(f"   Зарезервировано: {old_reserved} → {new_reserved}")

            # ========== ШАГ 3: ОБНОВЛЕНИЕ ТАБЛИЦЫ ИМПОРТА ОТ ЛАЗЕРЩИКОВ ==========
            # Проверяем, было ли это списание из импорта от лазерщиков
            is_laser_import = "Лазер:" in writeoff_comment or "лазерщик" in writeoff_comment.lower()

            if is_laser_import and hasattr(self, 'laser_table_data') and self.laser_table_data:
                print(f"\n🔄 Поиск соответствующей строки в таблице импорта...")

                # Ищем соответствующую строку по дате и комментарию
                writeoff_datetime = writeoff_date  # Формат: "DD.MM.YYYY HH:MM" или "YYYY-MM-DD HH:MM:SS"

                updated_count = 0
                for idx, row_data in enumerate(self.laser_table_data):
                    # Проверяем совпадение по дате списания
                    row_writeoff_date = row_data.get("Дата списания", "")

                    # Сравниваем даты (могут быть в разных форматах)
                    if row_writeoff_date and writeoff_datetime:
                        # Упрощённое сравнение по первым символам даты
                        row_date_part = row_writeoff_date[:16] if len(row_writeoff_date) >= 16 else row_writeoff_date
                        writeoff_date_part = writeoff_datetime[:16] if len(
                            writeoff_datetime) >= 16 else writeoff_datetime

                        if row_date_part == writeoff_date_part or row_data.get("Списано") in ["✓", "Да", "Yes"]:
                            # Дополнительная проверка по количеству (если есть в комментарии)
                            # Сбрасываем статус списания
                            self.laser_table_data[idx]["Списано"] = ""
                            self.laser_table_data[idx]["Дата списания"] = ""
                            updated_count += 1

                            print(f"   ✅ Обновлена строка #{idx + 1}: {row_data.get('order', 'N/A')}")

                if updated_count > 0:
                    print(f"\n✅ Обновлено строк в таблице импорта: {updated_count}")
                    # Обновляем визуа��ьное отображение
                    if hasattr(self, 'laser_import_tree'):
                        self.refresh_laser_import_table()
                else:
                    print(f"   ⚠️ Соответствующие строки в таблице импорта не найдены")

            # ========== ШАГ 3.5: ОТКАТ КОЛИЧЕСТВА ПОРЕЗАННЫХ ДЕТАЛЕЙ ==========
            print(f"\n🔄 Откат количества порезанных деталей...")

            try:
                # Извлекаем информацию из комментария списания
                # Формат: "Лазер: @username | Деталь: название | Дата импорта: DD.MM.YYYY HH:MM"
                import re

                part_name = None
                parts_qty = None

                # Пытаемся извлечь название детали из комментария
                part_match = re.search(r'Деталь:\s*([^|]+)', writeoff_comment)
                if part_match:
                    part_name = part_match.group(1).strip()
                    print(f"   📋 Название детали из комментария: '{part_name}'")

                # Пытаемся извлечь дату импорта из комментария
                date_match = re.search(r'Дата импорта:\s*(.+)', writeoff_comment)
                import_date_str = date_match.group(1).strip() if date_match else None

                # Ищем соответствующую строку в таблице импорта
                if part_name and hasattr(self, 'laser_table_data') and self.laser_table_data:
                    print(f"   🔍 Поиск строки в таблице импорта...")

                    for idx, row_data in enumerate(self.laser_table_data):
                        # Проверяем совпадение по детали
                        row_part = str(row_data.get("part", ""))

                        if part_name.lower() in row_part.lower() or row_part.lower() in part_name.lower():
                            # Дополнительная проверка по дате
                            row_date = str(row_data.get("Дата (МСК)", ""))
                            row_time = str(row_data.get("Время (МСК)", ""))
                            row_datetime = f"{row_date} {row_time}"

                            # Проверяем совпадение дат
                            date_match_found = False
                            if import_date_str:
                                # Сравниваем первые символы (дата без секунд)
                                if row_datetime[:16] == import_date_str[:16]:
                                    date_match_found = True
                            else:
                                # Если даты нет в комментарии, проверяем по дате списания
                                row_writeoff_date = str(row_data.get("Дата списания", ""))
                                if row_writeoff_date[:16] == writeoff_date[:16]:
                                    date_match_found = True

                            if date_match_found:
                                try:
                                    parts_qty = int(row_data.get("part_quantity", 0))
                                    print(f"   ✅ Найдена строка #{idx + 1}:")
                                    print(f"      Деталь: {row_part}")
                                    print(f"      Количество деталей: {parts_qty}")
                                    break
                                except ValueError:
                                    print(f"   ⚠️ Не удалось преобразовать количество: {row_data.get('part_quantity')}")

                # Если не нашли в таблице импорта, пробуем альтернативный способ
                if parts_qty is None or parts_qty == 0:
                    print(f"   ⚠️ Количество деталей не найдено в таблице импорта")

                    # Пробуем найти в базе данных по дате списания
                    # (если списывали несколько раз одну и ту же деталь)
                    print(f"   🔍 Попытка найти через базу WriteOffs...")

                    writeoffs_df_check = load_data("WriteOffs")
                    similar_writeoffs = writeoffs_df_check[
                        (writeoffs_df_check["ID заказа"] == int(writeoff_row["ID заказа"])) &
                        (writeoffs_df_check["Дата списания"] == writeoff_date) &
                        (writeoffs_df_check["Комментарий"].str.contains(part_name, case=False, na=False))
                        ]

                    if len(similar_writeoffs) > 0:
                        print(f"   ℹ️ Найдено похожих списаний: {len(similar_writeoffs)}")
                        print(f"   ⚠️ Невозможно точно определить количество деталей")
                        # Не откатываем, если не уверены
                        parts_qty = None

                # Если нашли количество, откатываем
                if parts_qty and parts_qty > 0 and part_name:
                    print(f"   🔄 Откат количества: {parts_qty} шт для детали '{part_name}'")

                    # Загружаем детали заказа
                    order_details_df = load_data("OrderDetails")
                    order_id = int(writeoff_row["ID заказа"])

                    print(f"   🔍 Поиск детали в заказе ID={order_id}...")

                    # Ищем деталь в заказе
                    detail_match = order_details_df[
                        (order_details_df["ID заказа"] == order_id) &
                        (order_details_df["Название детали"].str.contains(part_name, case=False, na=False))
                        ]

                    if not detail_match.empty:
                        detail_id = int(detail_match.iloc[0]["ID"])
                        detail_name_full = detail_match.iloc[0]["Название детали"]
                        old_cut = int(detail_match.iloc[0].get("Порезано", 0))
                        total_qty = int(detail_match.iloc[0].get("Количество", 0))

                        # Откатываем количество (не даём уйти в минус)
                        new_cut = max(0, old_cut - parts_qty)

                        order_details_df.loc[order_details_df["ID"] == detail_id, "Порезано"] = new_cut

                        print(f"   ✅ Деталь '{detail_name_full}' откачена:")
                        print(f"      ID детали: {detail_id}")
                        print(f"      Всего требуется: {total_qty}")
                        print(f"      Было порезано: {old_cut}")
                        print(f"      Откачено: -{parts_qty}")
                        print(f"      Стало порезано: {new_cut}")
                        print(f"      Осталось порезать: {total_qty - new_cut}")

                        # Сохраняем изменения
                        save_data("OrderDetails", order_details_df)

                        print(f"   💾 Изменения сохранены в OrderDetails")
                    else:
                        print(f"   ❌ Деталь '{part_name}' не найдена в заказе ID={order_id}")
                        print(f"   📋 Доступные детали в заказе:")

                        # Показываем список деталей для отладки
                        order_details = order_details_df[order_details_df["ID заказа"] == order_id]
                        for _, detail in order_details.iterrows():
                            print(f"      - {detail['Название детали']}")
                else:
                    print(f"   ⚠️ Откат детали пропущен (недостаточно данных)")
                    print(f"      Деталь: {part_name if part_name else 'не найдена'}")
                    print(f"      Количество: {parts_qty if parts_qty else 'не найдено'}")

            except Exception as e:
                print(f"   💥 Ошибка отката детали: {e}")
                import traceback
                traceback.print_exc()

            # ========== ШАГ 4: УДАЛЕНИЕ ЗАПИСИ О СПИСАНИИ ==========
            writeoffs_df = writeoffs_df[writeoffs_df["ID списания"] != writeoff_id]

            print(f"\n🗑️ Запись о списании ID={writeoff_id} удалена")

            # ========== ШАГ 5: СОХРАНЕНИЕ ИЗМЕНЕНИЙ ==========
            print(f"\n💾 Сохранение изменений в базу данных...")

            save_data("WriteOffs", writeoffs_df)
            save_data("Reservations", reservations_df)
            save_data("Materials", materials_df)

            print(f"✅ Данные сохранены")

            # ========== ШАГ 6: ОБНОВЛЕНИЕ ИНТЕРФЕЙСА ==========
            print(f"\n🔄 Обновление интерфейса...")

            self.refresh_writeoffs()
            self.refresh_reservations()
            self.refresh_materials()
            self.refresh_balance()

            # 🆕 ОБНОВЛЯЕМ ВКЛАДКУ ЗАКАЗОВ
            if hasattr(self, 'refresh_orders'):
                self.refresh_orders()
            if hasattr(self, 'refresh_order_details'):
                self.refresh_order_details()

            print(f"✅ Интерфейс обновлён")

            print(f"\n{'=' * 80}")
            print(f"✅ ОТМЕНА СПИСАНИЯ ЗАВЕРШЕНА УСПЕШНО")
            print(f"{'=' * 80}\n")

            messagebox.showinfo("Успех",
                                f"✅ Списание отменено!\n\n"
                                f"Возвращено в резерв: {quantity} шт\n"
                                f"Резерв ID: {reserve_id}\n"
                                f"Остаток к списанию: {new_remainder} шт\n\n"
                                f"{'Таблица импорта обновлена' if updated_count > 0 else 'Таблица импорта не затронута'}")

        except Exception as e:
            print(f"\n💥 КРИТИЧЕСКАЯ ОШИБКА: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось отменить списание:\n{e}")

    def find_laser_import_row_by_writeoff(self, writeoff_data):
        """
        Поиск строки в таблице импорта по данным списания

        Args:
            writeoff_data: dict с ключами 'Дата списания', 'Комментарий', 'Количество'

        Returns:
            list: индексы найденных строк в laser_table_data
        """
        if not hasattr(self, 'laser_table_data') or not self.laser_table_data:
            return []

        writeoff_date = writeoff_data.get('Дата списания', '')
        writeoff_comment = writeoff_data.get('Комментарий', '')
        writeoff_qty = writeoff_data.get('Количество', 0)

        # Извлекаем информацию из комментария
        # Формат: "Лазер: @username | Деталь: название_детали"
        import re
        username_match = re.search(r'Лазер:\s*(@?\w+)', writeoff_comment)
        part_match = re.search(r'Деталь:\s*(.+?)(?:\||$)', writeoff_comment)

        username = username_match.group(1) if username_match else None
        part_name = part_match.group(1).strip() if part_match else None

        print(f"   🔍 Критерии поиска:")
        print(f"      Дата: {writeoff_date}")
        print(f"      Пользователь: {username}")
        print(f"      Деталь: {part_name}")
        print(f"      Количество: {writeoff_qty}")

        matching_indices = []

        for idx, row_data in enumerate(self.laser_table_data):
            # Проверяем только списанные строки
            if row_data.get("Списано") not in ["✓", "Да", "Yes"]:
                continue

            match_score = 0

            # Сопоставление по дате списания (приоритет 3)
            row_writeoff_date = row_data.get("Дата списания", "")
            if row_writeoff_date and writeoff_date:
                # Сравниваем первые 16 символов (дата + время без секунд)
                if row_writeoff_date[:16] == writeoff_date[:16]:
                    match_score += 3

            # Сопоставление по пользователю (приоритет 2)
            if username:
                row_username = row_data.get("username", "")
                if username.lower() in row_username.lower() or row_username.lower() in username.lower():
                    match_score += 2

            # Сопоставление по детали (приоритет 2)
            if part_name:
                row_part = row_data.get("part", "")
                if part_name.lower() in row_part.lower() or row_part.lower() in part_name.lower():
                    match_score += 2

            # Сопоставление по количеству (приоритет 1)
            try:
                row_qty = int(row_data.get("metal_quantity", 0))
                if row_qty == writeoff_qty:
                    match_score += 1
            except:
                pass

            # Если набрали достаточно совпадений (минимум 3 балла)
            if match_score >= 3:
                matching_indices.append((idx, match_score))
                print(f"      ✓ Строка #{idx + 1}: score={match_score}")

        # Сортируем по убыванию score и возвращаем индексы
        matching_indices.sort(key=lambda x: x[1], reverse=True)
        return [idx for idx, score in matching_indices]
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
                msg = f"Сохранить изменения?\n\n"
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

    def setup_laser_import_tab(self):
        """Вкладка импорта от лазерщиков - ЕДИНСТВЕННАЯ ВЕРСИЯ"""

        # Очищаем фрейм на случай повторного вызова
        for widget in self.laser_import_frame.winfo_children():
            widget.destroy()

        # Заголовок
        header = tk.Label(self.laser_import_frame, text="📥 Импорт данных от лазерщиков",
                          font=("Arial", 16, "bold"), bg='white', fg='#e67e22')
        header.pack(pady=10)

        # Инструкция
        info_frame = tk.LabelFrame(self.laser_import_frame, text="ℹ️ Информация",
                                   bg='#d1ecf1', font=("Arial", 10, "bold"))
        info_frame.pack(fill=tk.X, padx=20, pady=10)

        instructions = """
    📋 Формат файла CSV:
    • Колонки: Дата (МСК), Время (МСК), username, order, metal, metal_quantity, part, part_quantity

    📌 Что делает импорт:
    1. Читает файл от лазерщиков
    2. Отображает все строки в таблице
    3. Позволяет выбрать строки для списания
    4. Автоматически находит резервы и списывает материал
        """

        tk.Label(info_frame, text=instructions, bg='#d1ecf1',
                 font=("Arial", 9), justify=tk.LEFT).pack(padx=10, pady=5)

        # Кнопки управления
        buttons_frame = tk.Frame(self.laser_import_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=20, pady=10)

        btn_style = {"font": ("Arial", 10, "bold"), "width": 20, "height": 2}

        tk.Button(buttons_frame, text="📁 Импорт файла", bg='#3498db', fg='white',
                  command=self.import_laser_table, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="✅ Списать выбранные", bg='#27ae60', fg='white',
                  command=self.writeoff_laser_row, **btn_style).pack(side=tk.LEFT, padx=5)
        # 🆕 НОВАЯ КНОПКА
        tk.Button(buttons_frame, text="🔵 Пометить вручную", bg='#2196F3', fg='white',
                  command=self.mark_manual_writeoff, **btn_style).pack(side=tk.LEFT, padx=5)

        # 🆕 КНОПКА СНЯТИЯ ПОМЕТКИ
        tk.Button(buttons_frame, text="↩️ Снять пометку", bg='#9E9E9E', fg='white',
                  command=self.unmark_manual_writeoff, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="🗑️ Удалить строки", bg='#e74c3c', fg='white',
                  command=self.delete_laser_row, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="💾 Экспорт таблицы", bg='#9b59b6', fg='white',
                  command=self.export_laser_table, **btn_style).pack(side=tk.LEFT, padx=5)

        # Метка таблицы
        table_label = tk.Label(self.laser_import_frame,
                               text="📊 Импортированные данные (выберите строки для списания)",
                               font=("Arial", 11, "bold"), bg='white', fg='#2c3e50')
        table_label.pack(pady=5)

        # Фрейм для таблицы
        tree_frame = tk.Frame(self.laser_import_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Scrollbars
        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        # 🆕 СОЗДАНИЕ TREEVIEW С ЯВНЫМИ ПАРАМЕТРАМИ
        self.laser_import_tree = ttk.Treeview(
            tree_frame,
            columns=("Дата", "Время", "Пользователь", "Заказ", "Металл", "Кол-во", "Деталь", "Кол-во деталей",
                     "Списано", "Дата списания"),
            show="headings",
            height=20,  # 🆕 ЯВНАЯ ВЫСОТА
            selectmode='extended',
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        scroll_y.config(command=self.laser_import_tree.yview)
        scroll_x.config(command=self.laser_import_tree.xview)

        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Настройка колонок
        columns_config = {
            "Дата": 100,
            "Время": 80,
            "Пользователь": 120,
            "Заказ": 200,
            "Металл": 200,
            "Кол-во": 80,
            "Деталь": 200,
            "Кол-во деталей": 120,
            "Списано": 80,
            "Дата списания": 150
        }

        for col, width in columns_config.items():
            self.laser_import_tree.heading(col, text=col)
            self.laser_import_tree.column(col, width=width, anchor=tk.CENTER)

        # 🆕 ВАЖНО: pack() ПОСЛЕ настройки колонок
        self.laser_import_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Цветовая индикация
        self.laser_import_tree.tag_configure('written_off', background='#c8e6c9', foreground='#1b5e20')
        self.laser_import_tree.tag_configure('manual', background='#bbdefb', foreground='#0d47a1')  # Светло-синий
        self.laser_import_tree.tag_configure('pending', background='#fff9c4', foreground='#000000')
        self.laser_import_tree.tag_configure('error', background='#ffcccc', foreground='#b71c1c')

        # Статусная строка
        self.laser_status_label = tk.Label(
            self.laser_import_frame,
            text="📂 Импортируйте файл для начала работы",
            font=("Arial", 10),
            bg='#ecf0f1',
            fg='#2c3e50',
            relief=tk.SUNKEN,
            anchor='w',
            padx=10,
            pady=5
        )
        self.laser_status_label.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=10)

        print("✅ setup_laser_import_tab() выполнен успешно")  # DEBUG

    def import_laser_writeoff_table(self):
        """Импорт таблицы от лазерщиков"""
        file_path = filedialog.askopenfilename(
            title="Выберите таблицу от лазерщиков",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            import_df = pd.read_excel(file_path, engine='openpyxl')

            # Проверка колонок
            required_cols = ["Дата (МСК)", "Время (МСК)", "username", "order", "metal",
                             "metal_quantity", "part", "part_quantity"]
            missing = [col for col in required_cols if col not in import_df.columns]

            if missing:
                messagebox.showerror("Ошибка", f"Отсутствуют колонки:\n{', '.join(missing)}")
                return

            # Сохраняем данные
            self.laser_import_data = import_df.to_dict('records')

            # Отображаем
            self.refresh_laser_import_table()

            messagebox.showinfo("Успех", f"Загружено {len(self.laser_import_data)} записей")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать:\n{e}")

    def refresh_laser_import_table(self):
        """Обновление таблицы импорта от лазерщиков"""
        # Очищаем таблицу
        for item in self.laser_import_tree.get_children():
            self.laser_import_tree.delete(item)

        # 🆕 ЗАЩИТА ОТ ОШИБКИ
        if not hasattr(self, 'laser_table_data') or self.laser_table_data is None:
            self.laser_table_data = []
            return

        # Заполняем таблицу
        for row_data in self.laser_table_data:
            date_val = row_data.get("Дата (МСК)", "")
            time_val = row_data.get("Время (МСК)", "")
            username = row_data.get("username", "")
            order = row_data.get("order", "")
            metal = row_data.get("metal", "")
            metal_qty = row_data.get("metal_quantity", "")
            part = row_data.get("part", "")
            part_qty = row_data.get("part_quantity", "")
            written_off = row_data.get("Списано", "")
            writeoff_date = row_data.get("Дата списания", "")

            values = (date_val, time_val, username, order, metal, metal_qty, part, part_qty, written_off, writeoff_date)

            # Определяем тег для цветовой индикации
            if written_off == "Вручную":
                tag = 'manual'  # Светло-синий
            elif written_off in ["Да", "✓", "Yes"]:
                tag = 'written_off'  # Зелёный
            else:
                tag = 'pending'  # Жёлтый

            self.laser_import_tree.insert("", "end", values=values, tags=(tag,))

        self.auto_resize_columns(self.laser_import_tree)

    def writeoff_selected_laser_row(self):
        """Списание выбранной строки"""
        selected = self.laser_import_tree.selection()
        if not selected:
            messagebox.showwarning("Предупреждение", "Выберите строку для списания")
            return

        item = selected[0]
        row_index = self.laser_import_tree.index(item)

        self.process_laser_writeoff(row_index)
        self.refresh_laser_import_table()

    def writeoff_all_laser_rows(self):
        """Массовое списание всех строк"""
        if not hasattr(self, 'laser_import_data') or not self.laser_import_data:
            messagebox.showwarning("Предупреждение", "Таблица пуста")
            return

        if not messagebox.askyesno("Подтверждение",
                                   f"Списать все записи ({len(self.laser_import_data)} шт)?"):
            return

        success = 0
        errors = 0

        for idx in range(len(self.laser_import_data)):
            if self.process_laser_writeoff(idx, silent=True):
                success += 1
            else:
                errors += 1

        self.refresh_laser_import_table()
        self.refresh_writeoffs()
        self.refresh_reservations()
        self.refresh_materials()

        messagebox.showinfo("Результат", f"✅ Списано: {success}\n❌ Ошибок: {errors}")

    def process_laser_writeoff(self, row_index, silent=False):
        """Обработка одной строки списания"""
        if row_index >= len(self.laser_import_data):
            return False

        row_data = self.laser_import_data[row_index]

        # Проверка: уже списано?
        if row_data.get("_status") == "✅ Списано":
            if not silent:
                messagebox.showwarning("Внимание", "Уже списано!")
            return False

        try:
            # Извлекаем данные
            order_name = str(row_data.get("order", "")).strip()
            metal_description = str(row_data.get("metal", "")).strip()

            try:
                metal_qty = int(float(row_data.get("metal_quantity", 0)))
            except:
                row_data["_status"] = "Ошибка: некорректное количество"
                return False

            part_name = str(row_data.get("part", "")).strip()
            username = str(row_data.get("username", "")).strip()
            date_str = str(row_data.get("Дата (МСК)", ""))
            time_str = str(row_data.get("Время (МСК)", ""))

            # Поиск заказа
            orders_df = load_data("Orders")
            import re
            match = re.search(r'УП-(\d+)', order_name)
            order_id = None

            if match:
                up_number = match.group(1)
                order_match = orders_df[orders_df["Название заказа"].str.contains(
                    f"УП-{up_number}", case=False, na=False, regex=False)]
                if not order_match.empty:
                    order_id = int(order_match.iloc[0]["ID заказа"])

            if not order_id:
                row_data["_status"] = f"Ошибка: заказ '{order_name}' не найден"
                return False

            # Парсинг размеров
            metal_parts = metal_description.split()
            thickness = None
            width = None
            length = None

            for part in metal_parts:
                match = re.search(r'(\d+(?:\.\d+)?)[хxХX](\d+(?:\.\d+)?)[хxХX](\d+(?:\.\d+)?)', part)
                if match:
                    thickness = float(match.group(1))
                    width = float(match.group(2))
                    length = float(match.group(3))
                    break

            if not thickness:
                row_data["_status"] = f"Ошибка: не определены размеры"
                return False

            # Поиск резерва
            reservations_df = load_data("Reservations")
            order_reserves = reservations_df[reservations_df["ID заказа"] == order_id]

            if order_reserves.empty:
                row_data["_status"] = f"Ошибка: нет резервов"
                return False

            suitable_reserve = None
            tolerance = 0.01

            for _, reserve in order_reserves.iterrows():
                thickness_match = abs(float(reserve["Толщина"]) - thickness) < tolerance

                if width and length:
                    width_match = abs(float(reserve["Ширина"]) - width) < tolerance
                    length_match = abs(float(reserve["Длина"]) - length) < tolerance

                    if thickness_match and width_match and length_match and int(reserve["Остаток к списанию"]) > 0:
                        suitable_reserve = reserve
                        break
                else:
                    if thickness_match and int(reserve["Остаток к списанию"]) > 0:
                        suitable_reserve = reserve
                        break

            if suitable_reserve is None:
                row_data["_status"] = f"Ошибка: резерв не найден"
                return False

            reserve_id = int(suitable_reserve["ID резерва"])
            remainder = int(suitable_reserve["Остаток к списанию"])

            if metal_qty > remainder:
                row_data["_status"] = f"Ошибка: недостаточно ({remainder} шт)"
                return False

            # СПИСАНИЕ
            writeoffs_df = load_data("WriteOffs")
            new_writeoff_id = 1 if writeoffs_df.empty else int(writeoffs_df["ID списания"].max()) + 1

            comment = f"Оператор: {username} | Деталь: {part_name}"
            writeoff_datetime = f"{date_str} {time_str}"

            new_writeoff = pd.DataFrame([{
                "ID списания": new_writeoff_id,
                "ID резерва": reserve_id,
                "ID заказа": order_id,
                "ID материала": int(suitable_reserve["ID материала"]),
                "Марка": suitable_reserve["Марка"],
                "Толщина": thickness,
                "Длина": length,
                "Ширина": width,
                "Количество": metal_qty,
                "Дата списания": writeoff_datetime,
                "Комментарий": comment
            }])

            writeoffs_df = pd.concat([writeoffs_df, new_writeoff], ignore_index=True)
            save_data("WriteOffs", writeoffs_df)

            # Обновляем резерв
            new_written_off = int(suitable_reserve["Списано"]) + metal_qty
            new_remainder = remainder - metal_qty

            reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Списано"] = new_written_off
            reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = new_remainder
            save_data("Reservations", reservations_df)

            # Обновляем склад
            material_id = int(suitable_reserve["ID материала"])
            if material_id != -1:
                materials_df = load_data("Materials")
                if not materials_df[materials_df["ID"] == material_id].empty:
                    mat_row = materials_df[materials_df["ID"] == material_id].iloc[0]
                    old_qty = int(mat_row["Количество штук"])
                    new_qty = old_qty - metal_qty

                    materials_df.loc[materials_df["ID"] == material_id, "Количество штук"] = new_qty

                    reserved = int(mat_row["Зарезервировано"])
                    new_reserved = max(0, reserved - metal_qty)
                    materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved
                    materials_df.loc[materials_df["ID"] == material_id, "Доступно"] = new_qty - new_reserved

                    save_data("Materials", materials_df)

            row_data["_status"] = "✅ Списано"
            return True

        except Exception as e:
            row_data["_status"] = f"Ошибка: {str(e)}"
            return False

    def clear_laser_table(self):
        """Очистка таблицы импорта"""
        if hasattr(self, 'laser_import_data'):
            self.laser_import_data = []

        for i in self.laser_import_tree.get_children():
            self.laser_import_tree.delete(i)

        messagebox.showinfo("Успех", "Таблица очищена")

    # ==================== МЕТОДЫ ДЛЯ ВКЛАДКИ ИМПОРТА ОТ ЛАЗЕРЩИКОВ ====================

    def import_laser_table(self):
        """Импорт таблицы от лазерщиков с сохранением статусов существующих записей"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл от лазерщиков",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            # Загрузка файла
            if file_path.endswith('.csv'):
                try:
                    laser_df = pd.read_csv(file_path, sep=';', encoding='utf-8')
                except:
                    laser_df = pd.read_csv(file_path, sep=';', encoding='cp1251')
            else:
                laser_df = pd.read_excel(file_path, engine='openpyxl')

            # Проверка обязательных колонок
            required = ["Дата (МСК)", "Время (МСК)", "username", "order", "metal", "metal_quantity", "part",
                        "part_quantity"]
            missing = [col for col in required if col not in laser_df.columns]

            if missing:
                messagebox.showerror("Ошибка", f"Отсутствуют колонки:\n{', '.join(missing)}")
                return

            # Добавляем колонки статуса если их нет
            if "Списано" not in laser_df.columns:
                laser_df["Списано"] = ""
            if "Дата списания" not in laser_df.columns:
                laser_df["Дата списания"] = ""

            # 🆕 СОЗДАЁМ УНИКАЛЬНЫЙ КЛЮЧ ДЛЯ КАЖДОЙ СТРОКИ
            # Используем комбинацию: дата + время + заказ + металл + деталь
            def create_row_key(row):
                """Создание уникального ключа для строки"""
                return (
                    str(row.get("Дата (МСК)", "")),
                    str(row.get("Время (МСК)", "")),
                    str(row.get("username", "")),
                    str(row.get("order", "")),
                    str(row.get("metal", "")),
                    str(row.get("metal_quantity", "")),
                    str(row.get("part", "")),
                    str(row.get("part_quantity", ""))
                )

            # 🆕 СОЗДАЁМ СЛОВАРЬ СУЩЕСТВУЮЩИХ СТРОК С ИХ СТАТУСАМИ
            existing_rows = {}
            if hasattr(self, 'laser_table_data') and self.laser_table_data:
                for row_data in self.laser_table_data:
                    key = create_row_key(row_data)
                    existing_rows[key] = {
                        "Списано": row_data.get("Списано", ""),
                        "Дата списания": row_data.get("Дата списания", "")
                    }

            # 🆕 ОБРАБАТЫВАЕМ НОВЫЙ ФАЙЛ
            new_rows = []
            updated_rows = 0
            duplicate_rows = 0

            for _, row in laser_df.iterrows():
                row_dict = row.to_dict()
                key = create_row_key(row_dict)

                # Проверяем, существует ли уже эта строка
                if key in existing_rows:
                    # Строка уже есть - сохраняем её статус
                    row_dict["Списано"] = existing_rows[key]["Списано"]
                    row_dict["Дата списания"] = existing_rows[key]["Дата списания"]
                    updated_rows += 1
                else:
                    # Новая строка - оставляем пустой статус
                    if not row_dict.get("Списано"):
                        row_dict["Списано"] = ""
                    if not row_dict.get("Дата списания"):
                        row_dict["Дата списания"] = ""

                new_rows.append(row_dict)

            # 🆕 ОБЪЕДИНЯЕМ: СНАЧАЛА СТАРЫЕ (С СОХРАНЕННЫМИ СТАТУСАМИ), ПОТОМ НОВЫЕ
            merged_data = []
            new_count = 0

            # Создаём множество ключей из нового файла
            new_keys = set()
            for row_dict in new_rows:
                new_keys.add(create_row_key(row_dict))

            # Добавляем старые строки, если они есть в новом файле
            if hasattr(self, 'laser_table_data') and self.laser_table_data:
                for old_row in self.laser_table_data:
                    old_key = create_row_key(old_row)
                    if old_key in new_keys:
                        # Строка есть в новом файле - берём из старой таблицы (с сохраненным статусом)
                        merged_data.append(old_row)
                    # Если строки нет в новом файле - не добавляем (она удалена из источника)

            # Добавляем только НОВЫЕ строки из импортированного файла
            for new_row in new_rows:
                new_key = create_row_key(new_row)

                # Проверяем, была ли эта строка в старых данных
                is_new = new_key not in existing_rows

                if is_new:
                    merged_data.append(new_row)
                    new_count += 1

            # Сохраняем объединённые данные
            self.laser_table_data = merged_data

            # Обновляем таблицу
            self.refresh_laser_import_table()

            # Принудительное обновление
            self.laser_import_tree.update_idletasks()
            self.laser_import_frame.update()

            # Автоширина колонок
            self.auto_resize_columns(self.laser_import_tree)

            # Обновляем статус
            items_count = len(self.laser_import_tree.get_children())

            if hasattr(self, 'laser_status_label'):
                self.laser_status_label.config(
                    text=f"✅ Всего записей: {items_count} | 🆕 Новых: {new_count} | 🔄 Обновлено статусов: {updated_rows}",
                    bg='#d4edda',
                    fg='#155724'
                )

            # Формируем сообщение
            result_msg = (
                f"✅ Импорт завершён!\n\n"
                f"📊 Всего записей: {items_count}\n"
                f"🆕 Новых записей: {new_count}\n"
                f"🔄 Сохранено статусов: {updated_rows}\n\n"
            )

            # Считаем статистику по статусам
            if self.laser_table_data:
                auto_count = sum(1 for r in self.laser_table_data if r.get("Списано") in ["✓", "Да", "Yes"])
                manual_count = sum(1 for r in self.laser_table_data if r.get("Списано") == "Вручную")
                pending_count = sum(1 for r in self.laser_table_data if not r.get("Списано"))

                result_msg += (
                    f"📈 Статистика:\n"
                    f"  • ✅ Автоматически списано: {auto_count}\n"
                    f"  • 🔵 Помечено вручную: {manual_count}\n"
                    f"  • 🟡 Ожидает обработки: {pending_count}"
                )

            messagebox.showinfo("Успех", result_msg)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось импортировать:\n{e}")
            import traceback
            traceback.print_exc()



    def refresh_laser_import_table(self):
        """Обновление таблицы импорта от лазерщиков"""
        # Очищаем таблицу
        for item in self.laser_import_tree.get_children():
            self.laser_import_tree.delete(item)

        # Заполняем таблицу
        for row_data in self.laser_table_data:
            date_val = row_data.get("Дата (МСК)", "")
            time_val = row_data.get("Время (МСК)", "")
            username = row_data.get("username", "")
            order = row_data.get("order", "")
            metal = row_data.get("metal", "")
            metal_qty = row_data.get("metal_quantity", "")
            part = row_data.get("part", "")
            part_qty = row_data.get("part_quantity", "")
            written_off = row_data.get("Списано", "")
            writeoff_date = row_data.get("Дата списания", "")

            values = (date_val, time_val, username, order, metal, metal_qty, part, part_qty, written_off, writeoff_date)

            # Определяем тег для цветовой индикации
            if written_off == "Да" or written_off == "✓":
                tag = 'written_off'
            else:
                tag = 'pending'

            self.laser_import_tree.insert("", "end", values=values, tags=(tag,))

        self.auto_resize_columns(self.laser_import_tree)

    def test_add_rows(self):
        """Тестовая функция для проверки отображения строк"""
        print("\n🧪 ТЕСТ: Добавление тестовых строк...")

        # Очищаем
        for item in self.laser_import_tree.get_children():
            self.laser_import_tree.delete(item)

        # Добавляем 3 тестовые строки
        test_data = [
            ("01.01.2026", "10:00", "@test1", "Тест заказ 1", "Ст3 10х1500х3000", "5", "Деталь A", "100", "", ""),
            ("02.01.2026", "11:00", "@test2", "Тест заказ 2", "Ст3 12х1500х3000", "3", "Деталь B", "50", "", ""),
            ("03.01.2026", "12:00", "@test3", "Тест заказ 3", "09Г2С 8х1500х3000", "2", "Деталь C", "75", "", "")
        ]

        for idx, values in enumerate(test_data, 1):
            item_id = self.laser_import_tree.insert("", "end", values=values, tags=('pending',))
            print(f"  ✓ Тестовая строка {idx} добавлена: ID={item_id}")

        # Проверка
        items_count = len(self.laser_import_tree.get_children())
        print(f"✅ ТЕСТ: В таблице {items_count} элементов")

        # Принудительное обновление
        self.laser_import_tree.update_idletasks()

        messagebox.showinfo("Тест", f"Добавлено тестовых строк: {items_count}")

    def writeoff_laser_row(self):
        """Списание выбранных строк с точным сопоставлением заказа, материала и детали"""
        selected_items = self.laser_import_tree.selection()

        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите строки для списания!")
            return

        # Проверяем, что выбранные строки еще не списаны
        rows_to_writeoff = []
        already_written_off = []

        for item in selected_items:
            values = self.laser_import_tree.item(item)['values']
            if values[8] in ["Да", "✓", "Yes"]:  # Колонка "Списано"
                already_written_off.append(values[3])  # order
            else:
                rows_to_writeoff.append((item, values))

        if already_written_off:
            messagebox.showinfo("Информация",
                                f"Некоторые строки уже списаны:\n" + "\n".join(already_written_off[:5]))

        if not rows_to_writeoff:
            messagebox.showwarning("Предупреждение", "Нет строк для списания!")
            return

        # Подтверждение
        if not messagebox.askyesno("Подтверждение",
                                   f"Списать выбранные строки ({len(rows_to_writeoff)} шт)?"):
            return

        try:
            # Загружаем данные
            orders_df = load_data("Orders")
            reservations_df = load_data("Reservations")
            materials_df = load_data("Materials")
            writeoffs_df = load_data("WriteOffs")
            order_details_df = load_data("OrderDetails")

            success_count = 0
            errors = []

            print(f"\n{'=' * 80}")
            print(f"🔵 НАЧАЛО СПИСАНИЯ: {len(rows_to_writeoff)} строк(и)")
            print(f"{'=' * 80}")

            for item, values in rows_to_writeoff:
                try:
                    date_val, time_val, username, order_name, metal_desc, metal_qty_str, part_name, part_qty = values[
                        :8]

                    print(f"\n📋 Обработка строки:")
                    print(f"   Заказ: {order_name}")
                    print(f"   Металл: {metal_desc}")
                    print(f"   Количество металла: {metal_qty_str}")
                    print(f"   Деталь: {part_name}")

                    # ========== ШАГ 1: ПОИСК ЗАКАЗА ==========
                    # Ищем по точному совпадению или по номеру УП-XXX
                    order_match = None

                    # Пробуем найти номер УП-XXX
                    import re
                    up_match = re.search(r'УП-(\d+)', order_name)
                    if up_match:
                        up_number = up_match.group(1)
                        order_match = orders_df[
                            orders_df["Название заказа"].str.contains(f"УП-{up_number}", case=False, na=False)]
                        print(f"   🔍 Поиск по УП-{up_number}")

                    # Если не нашли, ищем по частичному совпадению названия
                    if order_match is None or order_match.empty:
                        order_match = orders_df[
                            orders_df["Название заказа"].str.contains(order_name, case=False, na=False)]
                        print(f"   🔍 Поиск по названию: {order_name}")

                    if order_match.empty:
                        errors.append(f"❌ Заказ '{order_name}' не найден в базе")
                        print(f"   ❌ Заказ не найден")
                        continue

                    order_id = int(order_match.iloc[0]["ID заказа"])
                    print(f"   ✅ Заказ найден: ID={order_id}")

                    # ========== ШАГ 2: ПАРСИНГ МАТЕРИАЛА ==========
                    # Пример: "ГК Ст.3 6х1500х3000" → марка="ГК Ст.3", толщина=6, ширина=1500, длина=3000
                    metal_parts = metal_desc.strip().split()

                    # Ищем размеры (формат: NxMxK или NхMхK)
                    thickness = None
                    width = None
                    length = None
                    marka = None

                    for part in metal_parts:
                        # Проверяем на размеры
                        size_match = re.search(r'(\d+(?:\.\d+)?)[хxХX](\d+(?:\.\d+)?)[хxХX](\d+(?:\.\d+)?)', part)
                        if size_match:
                            thickness = float(size_match.group(1))
                            width = float(size_match.group(2))
                            length = float(size_match.group(3))
                            # Марка - всё до размеров
                            marka_parts = metal_desc.split(part)[0].strip().split()
                            marka = " ".join(marka_parts)
                            break

                    if not thickness or not marka:
                        errors.append(f"❌ Не удалось распарсить материал: {metal_desc}")
                        print(f"   ❌ Ошибка парсинга материала")
                        continue

                    print(f"   📦 Распарсенный материал:")
                    print(f"      Марка: {marka}")
                    print(f"      Толщина: {thickness} мм")
                    print(f"      Размер: {width}x{length}")

                    # ========== ШАГ 3: ПОИСК ДЕТАЛИ В ЗАКАЗЕ ==========
                    detail_id = None
                    detail_match = order_details_df[
                        (order_details_df["ID заказа"] == order_id) &
                        (order_details_df["Название детали"].str.contains(part_name, case=False, na=False))
                        ]

                    if not detail_match.empty:
                        detail_id = int(detail_match.iloc[0]["ID"])
                        print(f"   🔧 Деталь найдена: ID={detail_id}, Название='{part_name}'")
                    else:
                        print(f"   ⚠️ Деталь '{part_name}' не найдена в заказе (списание без привязки)")

                    # ========== ШАГ 4: ПОИСК РЕЗЕРВА С УЧЕТОМ МАТЕРИАЛА И ДЕТАЛИ ==========
                    # Ищем резервы этого заказа
                    order_reserves = reservations_df[
                        (reservations_df["ID заказа"] == order_id) &
                        (reservations_df["Остаток к списанию"] > 0)
                        ]

                    if order_reserves.empty:
                        errors.append(f"❌ Нет доступных резервов для заказа '{order_name}'")
                        print(f"   ❌ Резервы не найдены")
                        continue

                    print(f"   🔍 Найдено резервов для заказа: {len(order_reserves)}")

                    # Фильтруем резервы по материалу
                    suitable_reserves = order_reserves[
                        (order_reserves["Марка"].str.contains(marka, case=False, na=False)) &
                        (order_reserves["Толщина"] == thickness)
                        ]

                    # Если указаны размеры, фильтруем и по ним
                    if width and length:
                        suitable_reserves = suitable_reserves[
                            (suitable_reserves["Ширина"] == width) &
                            (suitable_reserves["Длина"] == length)
                            ]

                    print(f"   🔍 Подходящих по материалу: {len(suitable_reserves)}")

                    # Если найдена деталь, фильтруем по детали
                    if detail_id:
                        detail_reserves = suitable_reserves[suitable_reserves["ID детали"] == detail_id]
                        if not detail_reserves.empty:
                            suitable_reserves = detail_reserves
                            print(f"   ✅ Резервы с привязкой к детали ID={detail_id}: {len(suitable_reserves)}")

                    if suitable_reserves.empty:
                        errors.append(
                            f"❌ Не найден резерв для:\n"
                            f"   Заказ: {order_name}\n"
                            f"   Материал: {marka} {thickness}мм {width}x{length}\n"
                            f"   Деталь: {part_name}"
                        )
                        print(f"   ❌ Подходящий резерв не найден")
                        continue

                    # Берём первый подходящий резерв
                    reserve_row = suitable_reserves.iloc[0]
                    reserve_id = int(reserve_row["ID резерва"])
                    remainder = int(reserve_row["Остаток к списанию"])

                    print(f"   ✅ Выбран резерв ID={reserve_id}, остаток={remainder} шт")

                    # ========== ШАГ 5: КОЛИЧЕСТВО ДЛЯ СПИСАНИЯ ==========
                    try:
                        qty_to_writeoff = int(metal_qty_str)
                    except:
                        qty_to_writeoff = 1

                    if qty_to_writeoff > remainder:
                        errors.append(
                            f"⚠️ Недостаточно материала в резерве #{reserve_id}:\n"
                            f"   Запрошено: {qty_to_writeoff}, Доступно: {remainder}"
                        )
                        print(f"   ⚠️ Недостаточно материала: нужно {qty_to_writeoff}, есть {remainder}")
                        # Списываем сколько есть
                        qty_to_writeoff = remainder

                    print(f"   📝 Будет списано: {qty_to_writeoff} шт")

                    # ========== ШАГ 6: СОЗДАНИЕ СПИСАНИЯ ==========
                    new_writeoff_id = 1 if writeoffs_df.empty else int(writeoffs_df["ID списания"].max()) + 1

                    # 🆕 УЛУЧШЕННЫЙ КОММЕНТАРИЙ для связи с таблицей импорта
                    comment_text = (
                        f"Лазер: {username} | "
                        f"Деталь: {part_name} | "
                        f"Дата импорта: {date_val} {time_val}"
                    )

                    new_writeoff = pd.DataFrame([{
                        "ID списания": new_writeoff_id,
                        "ID резерва": reserve_id,
                        "ID заказа": reserve_row["ID заказа"],
                        "ID материала": reserve_row["ID материала"],
                        "Марка": reserve_row["Марка"],
                        "Толщина": reserve_row["Толщина"],
                        "Длина": reserve_row["Длина"],
                        "Ширина": reserve_row["Ширина"],
                        "Количество": qty_to_writeoff,
                        "Дата списания": f"{date_val} {time_val}",  # 🆕 СОХРАНЯЕМ ИСХОДНУЮ ДАТУ
                        "Комментарий": comment_text  # 🆕 РАСШИРЕННЫЙ КОММЕНТАРИЙ
                    }])

                    writeoffs_df = pd.concat([writeoffs_df, new_writeoff], ignore_index=True)

                    # ========== ШАГ 7: ОБНОВЛЕНИЕ РЕЗЕРВА ==========
                    new_written_off = int(reserve_row["Списано"]) + qty_to_writeoff
                    new_remainder = int(reserve_row["Зарезервировано штук"]) - new_written_off

                    reservations_df.loc[reservations_df["ID резерва"] == reserve_id, "Списано"] = new_written_off
                    reservations_df.loc[
                        reservations_df["ID резерва"] == reserve_id, "Остаток к списанию"] = new_remainder

                    print(f"   ✅ Резерв обновлен: Списано={new_written_off}, Остаток={new_remainder}")

                    # ========== ШАГ 8: ОБНОВЛЕНИЕ МАТЕРИАЛА НА СКЛАДЕ ==========
                    material_id = int(reserve_row["ID материала"])
                    if material_id != -1:
                        material = materials_df[materials_df["ID"] == material_id]
                        if not material.empty:
                            material = material.iloc[0]

                            new_qty = int(material["Количество штук"]) - qty_to_writeoff
                            new_reserved = int(material["Зарезервировано"]) - qty_to_writeoff

                            materials_df.loc[materials_df["ID"] == material_id, "Количество штук"] = new_qty
                            materials_df.loc[materials_df["ID"] == material_id, "Зарезервировано"] = new_reserved

                            # Пересчитываем площадь
                            area_per_piece = float(material["Длина"]) * float(material["Ширина"]) / 1_000_000
                            new_area = new_qty * area_per_piece
                            materials_df.loc[materials_df["ID"] == material_id, "Общая площадь"] = round(new_area, 2)

                            print(f"   ✅ Склад обновлен: Всего={new_qty}, Зарезервировано={new_reserved}")

                    # ========== ШАГ 9: ОБНОВЛЕНИЕ ДЕТАЛИ В ЗАКАЗЕ (ПОРЕЗАНО) ==========
                    if detail_id:
                        try:
                            # Загружаем детали заказа (если ещё не загружены)
                            if 'order_details_df' not in locals():
                                order_details_df = load_data("OrderDetails")

                            detail_row = order_details_df[order_details_df["ID"] == detail_id]

                            if not detail_row.empty:
                                detail_row = detail_row.iloc[0]
                                detail_name_full = detail_row["Название детали"]

                                old_cut = int(detail_row.get("Порезано", 0))

                                # Количество деталей из импорта
                                try:
                                    parts_qty = int(part_qty)
                                except:
                                    parts_qty = 0

                                new_cut = old_cut + parts_qty

                                # Обновляем количество порезанных деталей
                                order_details_df.loc[order_details_df["ID"] == detail_id, "Порезано"] = new_cut

                                # Проверяем общее количество
                                total_qty = int(detail_row.get("Количество", 0))

                                print(f"   📐 Деталь '{detail_name_full}' обновлена:")
                                print(f"      ID детали: {detail_id}")
                                print(f"      Всего требуется: {total_qty}")
                                print(f"      Было порезано: {old_cut}")
                                print(f"      Добавлено: +{parts_qty}")
                                print(f"      Стало порезано: {new_cut}")

                                # Сохраняем изменения
                                save_data("OrderDetails", order_details_df)

                                print(f"      💾 OrderDetails сохранён")

                                # Если порезано больше или равно требуемому - показываем уведомление
                                if new_cut >= total_qty:
                                    print(f"      ✅ Деталь полностью порезана! ({new_cut}/{total_qty})")
                                else:
                                    remaining = total_qty - new_cut
                                    print(f"      ⏳ Осталось порезать: {remaining} шт")
                            else:
                                print(f"   ⚠️ Деталь ID={detail_id} не найдена в OrderDetails")

                        except Exception as e:
                            print(f"   ⚠️ Ошибка обновления детали: {e}")
                            import traceback
                            traceback.print_exc()
                    else:
                        print(f"   ℹ️ Деталь не найдена в базе, пропускаем обновление 'Порезано'")

                    # ========== ШАГ 10: ОБНОВЛЕНИЕ СТАТУСА В ТАБЛИЦЕ ИМПОРТА ==========
                    item_index = self.laser_import_tree.index(item)
                    self.laser_table_data[item_index]["Списано"] = "✓"
                    self.laser_table_data[item_index]["Дата списания"] = datetime.now().strftime("%Y-%m-%d %H:%M")

                    # ========== ШАГ 9: ОБНОВЛЕНИЕ СТАТУСА В ТАБЛИЦЕ ИМПОРТА ==========
                    item_index = self.laser_import_tree.index(item)
                    self.laser_table_data[item_index]["Списано"] = "✓"
                    self.laser_table_data[item_index]["Дата списания"] = datetime.now().strftime("%Y-%m-%d %H:%M")

                    success_count += 1
                    print(f"   ✅ СПИСАНИЕ ВЫПОЛНЕНО УСПЕШНО")

                except Exception as e:
                    error_msg = f"❌ Ошибка обработки строки '{order_name}': {str(e)}"
                    errors.append(error_msg)
                    print(f"   {error_msg}")
                    import traceback
                    traceback.print_exc()

            # ========== СОХРАНЕНИЕ ИЗМЕНЕНИЙ ==========
            print(f"\n{'=' * 80}")
            print(f"💾 СОХРАНЕНИЕ ИЗМЕНЕНИЙ В БАЗУ ДАННЫХ")
            print(f"{'=' * 80}")

            save_data("WriteOffs", writeoffs_df)
            save_data("Reservations", reservations_df)
            save_data("Materials", materials_df)

            print(f"✅ Данные сохранены")

            # ========== ОБНОВЛЕНИЕ ИНТЕРФЕЙСА ==========
            print(f"\n🔄 Обновление интерфейса...")
            self.refresh_laser_import_table()
            self.refresh_materials()
            self.refresh_reservations()
            self.refresh_writeoffs()
            self.refresh_balance()

            # 🆕 ОБНОВЛЯЕМ ВКЛАДКУ ЗАКАЗОВ
            if hasattr(self, 'refresh_orders'):
                self.refresh_orders()
            if hasattr(self, 'refresh_order_details'):
                self.refresh_order_details()


            print(f"✅ Интерфейс обновлен")

            # ========== РЕЗУЛЬТАТ ==========
            print(f"\n{'=' * 80}")
            print(f"✅ СПИСАНИЕ ЗАВЕРШЕНО")
            print(f"   Успешно: {success_count}")
            print(f"   Ошибок: {len(errors)}")
            print(f"{'=' * 80}\n")

            result_msg = f"✅ Успешно списано: {success_count} записей"
            if errors:
                result_msg += f"\n\n⚠ Ошибки ({len(errors)}):\n" + "\n".join(errors[:10])
                if len(errors) > 10:
                    result_msg += f"\n... и еще {len(errors) - 10}"

            messagebox.showinfo("Результат списания", result_msg)

        except Exception as e:
            print(f"\n💥 КРИТИЧЕСКАЯ ОШИБКА: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось выполнить списание:\n{e}")

    def mark_manual_writeoff(self):
        """Пометка строк как 'списано вручную' без фактического списания"""
        selected_items = self.laser_import_tree.selection()

        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите строки для пометки!")
            return

        # Проверяем, что строки еще не списаны
        rows_to_mark = []
        already_marked = []

        for item in selected_items:
            values = self.laser_import_tree.item(item)['values']
            status = values[8]  # Колонка "Списано"

            if status in ["✓", "Да", "Yes"]:
                already_marked.append(f"{values[3]} (автоматически)")
            elif status == "Вручную":
                already_marked.append(f"{values[3]} (уже помечено вручную)")
            else:
                rows_to_mark.append((item, values))

        if already_marked:
            messagebox.showinfo("Информация",
                                f"Некоторые строки уже обработаны:\n" + "\n".join(already_marked[:5]))

        if not rows_to_mark:
            messagebox.showwarning("Предупреждение", "Нет строк для пометки!")
            return

        # Подтверждение
        confirm_msg = (
            f"Пометить {len(rows_to_mark)} строк(и) как 'списано вручную'?\n\n"
            f"⚠️ Это НЕ спишет материал с резервов!\n"
            f"Это только пометит строки для последующего ручного списания.\n\n"
            f"Строки окрасятся в светло-синий цвет."
        )

        if not messagebox.askyesno("Подтверждение", confirm_msg):
            return

        try:
            marked_count = 0

            for item, values in rows_to_mark:
                # Обновляем статус в таблице данных
                item_index = self.laser_import_tree.index(item)

                if item_index < len(self.laser_table_data):
                    self.laser_table_data[item_index]["Списано"] = "Вручную"
                    self.laser_table_data[item_index]["Дата списания"] = datetime.now().strftime("%Y-%m-%d %H:%M")

                    # Обновляем визуальное отображение
                    new_values = list(values)
                    new_values[8] = "Вручную"  # Колонка "Списано"
                    new_values[9] = datetime.now().strftime("%Y-%m-%d %H:%M")  # Колонка "Дата списания"

                    self.laser_import_tree.item(item, values=new_values, tags=('manual',))
                    marked_count += 1

            messagebox.showinfo("Успех",
                                f"✅ Помечено строк: {marked_count}\n\n"
                                f"🔵 Строки окрашены в светло-синий цвет\n"
                                f"📝 Не забудьте списать материал вручную!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось пометить строки:\n{e}")
            import traceback
            traceback.print_exc()

    def unmark_manual_writeoff(self):
        """Снятие пометки 'списано вручную'"""
        selected_items = self.laser_import_tree.selection()

        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите строки для снятия пометки!")
            return

        # Проверяем, что строки помечены вручную
        rows_to_unmark = []

        for item in selected_items:
            values = self.laser_import_tree.item(item)['values']
            status = values[8]  # Колонка "Списано"

            if status == "Вручную":
                rows_to_unmark.append((item, values))

        if not rows_to_unmark:
            messagebox.showwarning("Предупреждение",
                                   "Выбранные строки не помечены вручную!\n\n"
                                   "Снять можно только пометку 'Вручную'.\n"
                                   "Автоматические списания удаляются через вкладку 'Списание материалов'.")
            return

        # Подтверждение
        if not messagebox.askyesno("Подтверждение",
                                   f"Снять пометку с {len(rows_to_unmark)} строк(и)?"):
            return

        try:
            unmarked_count = 0

            for item, values in rows_to_unmark:
                # Обновляем статус в таблице данных
                item_index = self.laser_import_tree.index(item)

                if item_index < len(self.laser_table_data):
                    self.laser_table_data[item_index]["Списано"] = ""
                    self.laser_table_data[item_index]["Дата списания"] = ""

                    # Обновляем визуальное отображение
                    new_values = list(values)
                    new_values[8] = ""  # Колонка "Списано"
                    new_values[9] = ""  # Колонка "Дата списания"

                    self.laser_import_tree.item(item, values=new_values, tags=('pending',))
                    unmarked_count += 1

            messagebox.showinfo("Успех", f"✅ Снято пометок: {unmarked_count}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось снять пометку:\n{e}")

    def edit_laser_row(self):
        """Редактирование выбранной строки импорта"""
        selected = self.laser_import_tree.selection()
        if not selected or len(selected) != 1:
            messagebox.showwarning("Предупреждение", "Выберите одну строку для редактирования!")
            return

        item_index = self.laser_import_tree.index(selected[0])
        row_data = self.laser_table_data[item_index]

        # Окно редактирования
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактирование записи")
        edit_window.geometry("500x400")
        edit_window.configure(bg='#ecf0f1')

        tk.Label(edit_window, text="Редактирование записи от лазерщиков",
                 font=("Arial", 12, "bold"), bg='#ecf0f1').pack(pady=10)

        # Поля для редактирования
        fields = [
            ("Заказ:", "order"),
            ("Металл:", "metal"),
            ("Кол-во металла:", "metal_quantity"),
            ("Деталь:", "part"),
            ("Кол-во деталей:", "part_quantity")
        ]

        entries = {}
        for label_text, key in fields:
            frame = tk.Frame(edit_window, bg='#ecf0f1')
            frame.pack(fill=tk.X, padx=20, pady=5)
            tk.Label(frame, text=label_text, width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.LEFT)
            entry = tk.Entry(frame, font=("Arial", 10))
            entry.insert(0, str(row_data.get(key, "")))
            entry.pack(side=tk.RIGHT, expand=True, fill=tk.X)
            entries[key] = entry

        def save_changes():
            for key, entry in entries.items():
                self.laser_table_data[item_index][key] = entry.get()

            self.refresh_laser_import_table()
            edit_window.destroy()
            messagebox.showinfo("Успех", "Запись обновлена!")

        tk.Button(edit_window, text="💾 Сохранить", bg='#3498db', fg='white',
                  font=("Arial", 12, "bold"), command=save_changes).pack(pady=20)

    def delete_laser_row(self):
        """Удаление выбранных строк"""
        selected_items = self.laser_import_tree.selection()

        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите строки для удаления!")
            return

        if not messagebox.askyesno("Подтверждение",
                                   f"Удалить выбранные строки ({len(selected_items)} шт)?"):
            return

        # Удаляем в обратном порядке, чтобы индексы не сбивались
        indices_to_delete = sorted([self.laser_import_tree.index(item) for item in selected_items], reverse=True)

        for index in indices_to_delete:
            del self.laser_table_data[index]

        self.refresh_laser_import_table()
        messagebox.showinfo("Успех", f"Удалено записей: {len(indices_to_delete)}")

    def export_laser_table(self):
        """Экспорт таблицы обратно в Excel"""
        if not self.laser_table_data:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта!")
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить таблицу",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"laser_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not file_path:
            return

        try:
            df = pd.DataFrame(self.laser_table_data)

            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False, sep=';', encoding='utf-8')
            else:
                df.to_excel(file_path, index=False, engine='openpyxl')

            messagebox.showinfo("Успех", f"Таблица сохранена:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    # ==================== КОНЕЦ МЕТОДОВ ДЛЯ ИМПОРТА ОТ ЛАЗЕРЩИКОВ ====================

    def setup_balance_tab(self):
        """Вкладка баланса материалов"""
        header = tk.Label(self.balance_frame, text="Баланс материалов",
                          font=("Arial", 16, "bold"), bg='white', fg='#2c3e50')
        header.pack(pady=10)

        tree_frame = tk.Frame(self.balance_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        self.balance_tree = ttk.Treeview(tree_frame,
                                         columns=("Марка", "Толщина", "Размер", "Всего", "Зарезервировано",
                                                  "Списано", "Доступно", "Площадь"),
                                         show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.config(command=self.balance_tree.yview)
        scroll_x.config(command=self.balance_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        columns_config = {
            "Марка": 100, "Толщина": 80, "Размер": 120, "Всего": 80,
            "Зарезервировано": 120, "Списано": 80, "Доступно": 80, "Площадь": 100
        }

        for col, width in columns_config.items():
            self.balance_tree.heading(col, text=col)
            self.balance_tree.column(col, width=width, anchor=tk.CENTER)

        self.balance_tree.pack(fill=tk.BOTH, expand=True)

        # Панель фильтрации
        self.balance_filters = self.create_filter_panel(
            self.balance_frame,
            self.balance_tree,
            ["Марка", "Толщина", "Размер"],
            self.refresh_balance
        )

        # Переключатели видимости
        self.balance_toggles = self.create_visibility_toggles(
            self.balance_frame,
            self.balance_tree,
            {
                'show_zero_balance': '📦 Показать с нулевым балансом'
            },
            self.refresh_balance
        )

        buttons_frame = tk.Frame(self.balance_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        btn_style = {"font": ("Arial", 10), "width": 15, "height": 2}

        tk.Button(buttons_frame, text="🔄 Обновить", bg='#3498db', fg='white',
                  command=self.refresh_balance, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="📊 Экспорт в Excel", bg='#27ae60', fg='white',
                  command=self.export_balance, **btn_style).pack(side=tk.LEFT, padx=5)

        self.refresh_balance()

    def refresh_balance(self):
        """Обновление баланса материалов"""
        for item in self.balance_tree.get_children():
            self.balance_tree.delete(item)

        materials_df = load_data("Materials")
        writeoffs_df = load_data("WriteOffs")

        if materials_df.empty:
            return

        show_zero = True
        if hasattr(self, 'balance_toggles') and self.balance_toggles:
            show_zero = self.balance_toggles.get('show_zero_balance', tk.BooleanVar(value=True)).get()

        # Группируем списания по материалам
        writeoff_summary = {}
        if not writeoffs_df.empty:
            for _, row in writeoffs_df.iterrows():
                mat_id = int(row["ID материала"])
                qty = int(row["Количество"])
                writeoff_summary[mat_id] = writeoff_summary.get(mat_id, 0) + qty

        for _, row in materials_df.iterrows():
            mat_id = int(row["ID"])
            total_qty = int(row["Количество штук"])
            reserved = int(row["Зарезервировано"])
            available = int(row["Доступно"])
            written_off = writeoff_summary.get(mat_id, 0)

            # 🆕 Фильтрация по доступному (а не по total_qty)
            if not show_zero and available == 0:
                continue

            size_str = f"{row['Ширина']}x{row['Длина']}"

            values = (
                row["Марка"],
                row["Толщина"],
                size_str,
                total_qty,
                reserved,
                written_off,
                available,
                row["Общая площадь"]
            )

            # 🆕 ЦВЕТОВАЯ ИНДИКАЦИЯ
            if available < 0:
                tag = 'negative'  # Отрицательное - красный
            elif available == 0:
                tag = 'zero'  # Нулевое - жёлтый
            else:
                tag = ''  # Нормальное - без цвета

            self.balance_tree.insert("", "end", values=values, tags=(tag,))

        # 🆕 НАСТРОЙКА ЦВЕТОВ
        self.balance_tree.tag_configure('negative', background='#ffcccc', foreground='#b71c1c')  # Светло-красный
        self.balance_tree.tag_configure('zero', background='#fff9c4', foreground='#856404')  # Светло-жёлтый

        self.auto_resize_columns(self.balance_tree)

    def export_balance(self):
        """Экспорт баланса в Excel"""
        file_path = filedialog.asksaveasfilename(
            title="Сохранить баланс",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"balance_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not file_path:
            return

        try:
            # Собираем данные из таблицы
            data = []
            for item in self.balance_tree.get_children():
                values = self.balance_tree.item(item)['values']
                data.append(values)

            df = pd.DataFrame(data, columns=self.balance_tree['columns'])
            df.to_excel(file_path, index=False, engine='openpyxl')

            messagebox.showinfo("Успех", f"Баланс сохранен:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")


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