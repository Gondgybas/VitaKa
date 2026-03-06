# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from openpyxl import Workbook, load_workbook
from datetime import datetime, timedelta
from pathlib import Path
import os
import json

DATABASE_FILE = "production_database.xlsx"
DATA_PATH = Path(__file__).parent  # Папка где лежит скрипт


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
        reservations_sheet.append(
            ["ID резерва", "ID заказа", "ID детали", "Название детали", "ID материала", "Марка", "Толщина", "Длина",
             "Ширина", "Зарезервировано штук", "Списано", "Остаток к списанию", "Дата резерва"])
        writeoffs_sheet = wb.create_sheet("WriteOffs")
        writeoffs_sheet.append([
            "ID списания", "ID резерва", "ID заказа", "ID материала", "Марка", "Толщина", "Длина", "Ширина",
            "Количество", "Дата списания", "Комментарий"
        ])

        # 🆕 ЛИСТ ДЛЯ ЛОГИРОВАНИЯ ИЗМЕНЕНИЙ КОЛИЧЕСТВА МАТЕРИАЛА
        changelogs_sheet = wb.create_sheet("MaterialChangeLogs")
        changelogs_sheet.append([
            "ID лога", "Дата и время", "ID материала", "Марка", "Толщина",
            "Длина", "Ширина", "Старое кол-во", "Новое кол-во", "Изменение", "Комментарий"
        ])

        wb.save(DATABASE_FILE)
        print(f"База данных '{DATABASE_FILE}' создана!")


def get_database_path():
    """Получить путь к папке с базой данных из настроек"""
    settings_file = "app_settings.json"
    try:
        if os.path.exists(settings_file):
            with open(settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                return settings.get("database_path", os.path.dirname(os.path.abspath(__file__)))
    except:
        pass
    # По умолчанию - текущая папка
    return os.path.dirname(os.path.abspath(__file__))


def load_data(sheet_name):
    """Загрузка данных из Excel с учётом пути из настроек"""
    db_path = get_database_path()
    file_path = os.path.join(db_path, "production_database.xlsx")

    try:
        if os.path.exists(file_path):
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            return df
        else:
            print(f"⚠️ Файл базы данных не найден: {file_path}")
            return pd.DataFrame()
    except Exception as e:
        print(f"❌ Ошибка загрузки данных из {sheet_name}: {e}")
        return pd.DataFrame()


def save_data(sheet_name, df):
    """Сохранение данных в Excel с учётом пути из настроек"""
    db_path = get_database_path()
    file_path = os.path.join(db_path, "production_database.xlsx")

    try:
        if os.path.exists(file_path):
            with pd.ExcelFile(file_path, engine='openpyxl') as xls:
                sheets = {s: pd.read_excel(xls, s) for s in xls.sheet_names}
        else:
            sheets = {}

        sheets[sheet_name] = df

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for s, data in sheets.items():
                data.to_excel(writer, sheet_name=s, index=False)

        print(f"✅ Данные сохранены в {sheet_name}")
    except Exception as e:
        print(f"❌ Ошибка сохранения данных в {sheet_name}: {e}")
        messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить данные: {e}")


class ExcelStyleFilter:
    """Фильтр в стиле Excel для Treeview - выпадающее меню при клике на заголовок"""

    def __init__(self, tree, refresh_callback, columns_config=None):
        """
        tree: ttk.Treeview виджет
        refresh_callback: функция обновления данных
        columns_config: dict с настройками столбцов (опционально)
        """
        self.tree = tree
        self.refresh_callback = refresh_callback
        self.columns_config = columns_config or {}

        # Хранилище активных фильтров
        self.active_filters = {}

        # Исходные данные (до фильтрации)
        self.original_data = []

        # 🆕 ЗАЩИТА ОТ ДВОЙНОГО ВЫЗОВА
        self._filter_window_open = False
        self._last_click_time = 0

        # Привязываем клик к заголовкам
        self.tree.bind('<Button-1>', self.on_header_click)

    def on_header_click(self, event):
        """Обработка клика по заголовку столбца"""
        import time

        # 🆕 ЗАЩИТА ОТ ДВОЙНОГО КЛИКА
        current_time = time.time()
        if current_time - self._last_click_time < 0.3:
            return

        # 🆕 ЗАЩИТА: НЕ ОТКРЫВАЕМ ВТОРОЕ ОКНО
        if self._filter_window_open:
            return

        region = self.tree.identify_region(event.x, event.y)

        if region == "heading":
            column = self.tree.identify_column(event.x)
            column_id = self.tree.column(column, "id")

            self._last_click_time = current_time
            self._filter_window_open = True

            # Показываем меню фильтра
            self.show_filter_menu(event, column_id)

    def show_filter_menu(self, event, column_id):
        """Показать меню фильтра для столбца"""
        try:
            column_index = list(self.tree["columns"]).index(column_id)

            # Собираем уникальные значения
            all_unique_values = set()
            visible_unique_values = set()

            visible_items = list(self.tree.get_children(''))

            if not hasattr(self, '_all_item_cache'):
                self._all_item_cache = set()

            for item_id in visible_items:
                self._all_item_cache.add(item_id)
                values = self.tree.item(item_id)["values"]
                value = values[column_index]
                visible_unique_values.add(str(value))
                all_unique_values.add(str(value))

            for item_id in self._all_item_cache:
                if item_id not in visible_items:
                    try:
                        values = self.tree.item(item_id)["values"]
                        if values:
                            value = values[column_index]
                            all_unique_values.add(str(value))
                    except:
                        pass

            # Определяем выбранные значения
            if column_id in self.active_filters:
                currently_selected = set(self.active_filters[column_id])
            else:
                currently_selected = all_unique_values.copy()

            # Сортировка
            selected_visible = sorted(list(currently_selected & visible_unique_values))
            selected_hidden = sorted(list(currently_selected - visible_unique_values))
            unselected = sorted(list(all_unique_values - currently_selected))

            unique_values = selected_visible + selected_hidden + unselected

            # Создаём окно фильтра
            filter_window = tk.Toplevel(self.tree)
            filter_window.title(f"Фильтр: {column_id}")
            filter_window.geometry("320x600")
            filter_window.configure(bg='#ecf0f1')
            filter_window.transient(self.tree)
            filter_window.grab_set()

            x = event.x_root
            y = event.y_root + 20
            filter_window.geometry(f"+{x}+{y}")

            # Заголовок
            header_frame = tk.Frame(filter_window, bg='#3498db')
            header_frame.pack(fill=tk.X)
            tk.Label(header_frame, text=f"Фильтр: {column_id}",
                     font=("Arial", 12, "bold"), bg='#3498db', fg='white', pady=10).pack()

            # Кнопки сортировки
            sort_frame = tk.Frame(filter_window, bg='#ecf0f1')
            sort_frame.pack(fill=tk.X, padx=10, pady=10)
            tk.Label(sort_frame, text="Сортировка:", font=("Arial", 10, "bold"), bg='#ecf0f1').pack(anchor='w',
                                                                                                    pady=(0, 5))
            tk.Button(sort_frame, text="▲ По возрастанию (A→Z, 0→9)",
                      command=lambda: self.apply_sort(column_id, 'asc', filter_window),
                      bg='#3498db', fg='white', font=("Arial", 9), relief=tk.RAISED).pack(fill=tk.X, pady=2)
            tk.Button(sort_frame, text="▼ По убыванию (Z→A, 9→0)",
                      command=lambda: self.apply_sort(column_id, 'desc', filter_window),
                      bg='#3498db', fg='white', font=("Arial", 9), relief=tk.RAISED).pack(fill=tk.X, pady=2)

            tk.Frame(filter_window, height=2, bg='#95a5a6').pack(fill=tk.X, pady=5)

            # 🆕 ПОЛЕ ПОИСКА
            search_frame = tk.Frame(filter_window, bg='#ecf0f1')
            search_frame.pack(fill=tk.X, padx=10, pady=5)
            tk.Label(search_frame, text="🔍 Поиск:", font=("Arial", 10, "bold"), bg='#ecf0f1').pack(side=tk.LEFT, padx=5)

            search_var = tk.StringVar()
            search_entry = tk.Entry(search_frame, textvariable=search_var, font=("Arial", 10), width=20)
            search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

            def clear_search():
                search_var.set("")
                search_entry.focus_set()

            tk.Button(search_frame, text="✖", font=("Arial", 8), bg='#e74c3c', fg='white',
                      command=clear_search, width=3).pack(side=tk.LEFT, padx=2)

            tk.Frame(filter_window, height=2, bg='#95a5a6').pack(fill=tk.X, pady=5)

            tk.Label(filter_window, text="Фильтр по значению:",
                     font=("Arial", 10, "bold"), bg='#ecf0f1').pack(pady=(5, 5), padx=10, anchor='w')

            # Фрейм со списком
            list_frame = tk.Frame(filter_window, bg='white', relief=tk.SUNKEN, borderwidth=1)
            list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

            scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            canvas = tk.Canvas(list_frame, bg='white', yscrollcommand=scrollbar.set, highlightthickness=0)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=canvas.yview)

            checkboxes_frame = tk.Frame(canvas, bg='white')
            canvas_window = canvas.create_window((0, 0), window=checkboxes_frame, anchor='nw')

            # Прокрутка
            def _on_mousewheel(e):
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

            def _bind_to_mousewheel(e):
                canvas.bind_all("<MouseWheel>", _on_mousewheel)

            def _unbind_from_mousewheel(e):
                canvas.unbind_all("<MouseWheel>")

            canvas.bind('<Enter>', _bind_to_mousewheel)
            canvas.bind('<Leave>', _unbind_from_mousewheel)

            trace_id = None

            # Очистка при закрытии
            def cleanup_and_close():
                try:
                    _unbind_from_mousewheel(None)
                except:
                    pass

                try:
                    if trace_id is not None:
                        search_var.trace_remove('write', trace_id)
                except:
                    pass

                self._filter_window_open = False

                try:
                    filter_window.destroy()
                except:
                    pass

            filter_window.protocol("WM_DELETE_WINDOW", cleanup_and_close)

            # "Выбрать всё"
            all_selected = (len(currently_selected.intersection(unique_values)) == len(unique_values))
            select_all_var = tk.BooleanVar(value=all_selected)
            checkbox_vars = {}

            def toggle_all():
                state = select_all_var.get()
                search_text = search_var.get().lower().strip()
                for value, var in checkbox_vars.items():
                    if not search_text or search_text in value.lower():
                        var.set(state)

            select_all_frame = tk.Frame(checkboxes_frame, bg='#e8f4f8')
            select_all_frame.pack(fill=tk.X, pady=2)
            tk.Checkbutton(select_all_frame, text="✓ Выбрать всё",
                           variable=select_all_var, command=toggle_all,
                           font=("Arial", 10, "bold"), bg='#e8f4f8',
                           activebackground='#d1ecf1').pack(anchor='w', padx=5, pady=5)

            tk.Frame(checkboxes_frame, height=2, bg='#95a5a6').pack(fill=tk.X, padx=5, pady=2)

            checkbox_frames = {}

            # Создаём чекбоксы
            for value in unique_values:
                is_checked = (value in currently_selected)
                is_visible = (value in visible_unique_values)

                var = tk.BooleanVar(value=is_checked)
                checkbox_vars[value] = var

                bg_color = 'white' if is_visible else '#f8f8f8'
                cb_frame = tk.Frame(checkboxes_frame, bg=bg_color)
                cb_frame.pack(fill=tk.X, padx=2, pady=1)

                display_text = f"{value} 🔒" if not is_visible else value
                cb = tk.Checkbutton(cb_frame, text=display_text, variable=var,
                                    font=("Arial", 9, "italic" if not is_visible else "normal"),
                                    bg=bg_color, fg='#888' if not is_visible else 'black',
                                    activebackground='#e0e0e0' if not is_visible else '#f0f0f0')
                cb.pack(anchor='w', padx=10, pady=2)

                checkbox_frames[value] = cb_frame

            # 🆕 ФУНКЦИЯ ФИЛЬТРАЦИИ
            def filter_checkboxes(*args):
                try:
                    if not filter_window.winfo_exists():
                        return
                except:
                    return

                search_text = search_var.get().lower().strip()

                # Скрываем все
                for cb_frame in checkbox_frames.values():
                    cb_frame.pack_forget()

                # Показываем нужные
                for value in unique_values:
                    if value not in checkbox_vars:
                        continue

                    var = checkbox_vars[value]
                    cb_frame = checkbox_frames[value]

                    if not search_text:
                        cb_frame.pack(fill=tk.X, padx=2, pady=1)
                    elif search_text in value.lower():
                        cb_frame.pack(fill=tk.X, padx=2, pady=1)
                        var.set(True)
                    else:
                        var.set(False)

                # Обновляем "Выбрать всё"
                if search_text:
                    checked = sum(1 for v in checkbox_vars.values() if v.get())
                    select_all_var.set(checked > 0)
                else:
                    checked = sum(1 for v in checkbox_vars.values() if v.get())
                    select_all_var.set(checked == len(checkbox_vars))

                # Обновляем canvas
                try:
                    checkboxes_frame.update_idletasks()
                    canvas.configure(scrollregion=canvas.bbox("all"))
                except:
                    pass

            trace_id = search_var.trace_add('write', filter_checkboxes)
            search_entry.focus_set()

            def on_frame_configure(event=None):
                try:
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    canvas.itemconfig(canvas_window, width=canvas.winfo_width())
                except:
                    pass

            checkboxes_frame.bind("<Configure>", on_frame_configure)
            canvas.bind("<Configure>", on_frame_configure)

            # Функции для кнопок
            def apply_value_filter():
                selected_values = {value for value, var in checkbox_vars.items() if var.get()}
                if not selected_values:
                    messagebox.showwarning("Предупреждение", "Выберите хотя бы одно значение!")
                    return
                cleanup_and_close()
                self.apply_filter(column_id, selected_values, None)

            def clear_filter():
                if column_id in self.active_filters:
                    del self.active_filters[column_id]
                self.update_column_headers()
                cleanup_and_close()
                self.refresh_callback()

            # Кнопки
            buttons_frame = tk.Frame(filter_window, bg='#ecf0f1')
            buttons_frame.pack(fill=tk.X, padx=10, pady=10)

            tk.Button(buttons_frame, text="✓ Применить фильтр", command=apply_value_filter,
                      bg='#27ae60', fg='white', font=("Arial", 10, "bold"), relief=tk.RAISED, borderwidth=2).pack(
                side=tk.LEFT, padx=5, expand=True, fill=tk.X)

            tk.Button(buttons_frame, text="✗ Сбросить фильтр", command=clear_filter,
                      bg='#e74c3c', fg='white', font=("Arial", 10)).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

            tk.Button(buttons_frame, text="Отмена", command=cleanup_and_close,
                      bg='#95a5a6', fg='white', font=("Arial", 10)).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        except Exception as e:
            print(f"❌ ОШИБКА: {e}")
            import traceback
            traceback.print_exc()
            self._filter_window_open = False

    def apply_sort(self, column_id, direction, window):
        """Применить сортировку"""
        column_index = list(self.tree["columns"]).index(column_id)

        items = [(self.tree.item(item_id)["values"], item_id) for item_id in self.tree.get_children()]

        try:
            items.sort(key=lambda x: float(str(x[0][column_index]).replace(',', '.')),
                       reverse=(direction == 'desc'))
        except:
            items.sort(key=lambda x: str(x[0][column_index]).lower(),
                       reverse=(direction == 'desc'))

        for index, (values, item_id) in enumerate(items):
            self.tree.move(item_id, '', index)

        window.destroy()
        self._filter_window_open = False

    def apply_filter(self, column_id, selected_values, sort_order):
        """Применить фильтр"""
        self.active_filters[column_id] = selected_values

        column_index = list(self.tree["columns"]).index(column_id)

        # Показываем все из кэша
        for item_id in self._all_item_cache:
            try:
                self.tree.reattach(item_id, '', 'end')
            except:
                pass

        # Применяем ВСЕ активные фильтры
        visible_items = set(self.tree.get_children(''))

        for col_id, values in self.active_filters.items():
            col_index = list(self.tree["columns"]).index(col_id)

            items_to_hide = set()

            for item_id in visible_items:
                try:
                    item_values = self.tree.item(item_id)["values"]
                    value = str(item_values[col_index])

                    if value not in values:
                        items_to_hide.add(item_id)
                except:
                    pass

            for item_id in items_to_hide:
                self.tree.detach(item_id)

            visible_items -= items_to_hide

        self.update_column_headers()

        if self.refresh_callback:
            self.refresh_callback()

    def update_column_headers(self):
        """Обновить заголовки колонок (добавить/убрать индикатор фильтра)"""
        for col in self.tree["columns"]:
            current_text = col.replace(" 🔽", "")

            if col in self.active_filters or current_text in self.active_filters:
                new_text = f"{current_text} 🔽"
            else:
                new_text = current_text

            self.tree.heading(col, text=new_text)

    def reapply_all_filters(self):
        """Переприменить все активные фильтры"""
        if not self.active_filters:
            return

        # Показываем все
        for item_id in self._all_item_cache:
            try:
                self.tree.reattach(item_id, '', 'end')
            except:
                pass

        # Применяем каждый фильтр
        visible_items = set(self.tree.get_children(''))

        for column_id, selected_values in self.active_filters.items():
            column_index = list(self.tree["columns"]).index(column_id)

            items_to_hide = set()

            for item_id in visible_items:
                try:
                    values = self.tree.item(item_id)["values"]
                    value = str(values[column_index])

                    if value not in selected_values:
                        items_to_hide.add(item_id)
                except:
                    pass

            for item_id in items_to_hide:
                self.tree.detach(item_id)

            visible_items -= items_to_hide

        self.update_column_headers()

    def clear_all_filters(self):
        """очистить все фильтры"""
        self.active_filters = {}

        for item_id in self._all_item_cache:
            try:
                self.tree.reattach(item_id, '', 'end')
            except:
                pass

        self.update_column_headers()

        if self.refresh_callback:
            self.refresh_callback()



class ProductionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ООО Вита-Ка")
        self.root.geometry("1400x800")
        self.root.configure(bg='#f0f0f0')

        # Создаём верхнюю панель с заголовком и кнопкой настроек
        header_frame = tk.Frame(root, bg='#2c3e50', height=50)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)

        # Заголовок приложения
        title_label = tk.Label(
            header_frame,
            text="⚙️ Система учета производства",
            font=("Arial", 16, "bold"),
            bg='#2c3e50',
            fg='white'
        )
        title_label.pack(side=tk.LEFT, padx=20, pady=10)

        # Кнопка настроек (шестерёнка)
        settings_button = tk.Button(
            header_frame,
            text="⚙️ Настройки",
            font=("Arial", 11, "bold"),
            bg='#34495e',
            fg='white',
            activebackground='#1abc9c',
            activeforeground='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2',
            command=self.open_settings
        )
        settings_button.pack(side=tk.RIGHT, padx=20, pady=10)

        # Инициализация переменных toggles
        self.materials_toggles = {}
        self.orders_toggles = {}
        self.reservations_toggles = {}
        self.balance_toggles = {}
        self.writeoffs_toggles = {}
        self.details_toggles = {}

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

        # 🆕 НОВАЯ ВКЛАДКА
        self.details_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.details_frame, text='Учёт деталей')
        self.setup_details_tab()

        self.balance_frame = tk.Frame(self.notebook, bg='white')
        self.notebook.add(self.balance_frame, text='Баланс материалов')
        self.setup_balance_tab()

        self.load_toggle_settings()

        # Загрузка настроек и обработчик закрытия
        self.load_toggle_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.material_logs_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.material_logs_frame, text="📊 История материалов")
        self.setup_material_logs_tab()

    def load_settings(self):
        """Загрузка настроек из файла"""
        settings_file = "app_settings.json"
        default_settings = {
            "database_path": os.path.dirname(os.path.abspath(__file__))  # Текущая папка по умолчанию
        }

        try:
            if os.path.exists(settings_file):
                with open(settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    print(f"✅ Настройки загружены: {settings}")
                    return settings
            else:
                print(f"⚠️ Файл настроек не найден, используются значения по умолчанию")
                return default_settings
        except Exception as e:
            print(f"❌ Ошибка загрузки настроек: {e}")
            return default_settings

    def save_settings(self, settings):
        """Сохранение настроек в файл"""
        settings_file = "app_settings.json"
        try:
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            print(f"✅ Настройки сохранены: {settings}")
            return True
        except Exception as e:
            print(f"❌ Ошибка сохранения настроек: {e}")
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: {e}")
            return False

    def open_settings(self):
        """Открытие окна настроек"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("⚙️ Настройки программы")
        settings_window.geometry("700x300")
        settings_window.configure(bg='#ecf0f1')
        settings_window.resizable(False, False)

        # Заголовок
        header = tk.Label(
            settings_window,
            text="⚙️ Настройки системы",
            font=("Arial", 16, "bold"),
            bg='#ecf0f1',
            fg='#2c3e50'
        )
        header.pack(pady=20)

        # Загружаем текущие настройки
        current_settings = self.load_settings()

        # Путь к папке с БД
        path_frame = tk.LabelFrame(
            settings_window,
            text="📁 Путь к системной папке с данными",
            bg='#ecf0f1',
            font=("Arial", 11, "bold"),
            fg='#34495e'
        )
        path_frame.pack(fill=tk.X, padx=30, pady=15)

        path_info = tk.Label(
            path_frame,
            text="В этой папке должны находиться файлы:\n• production_database.xlsx\n• laser_import_cache.xlsx",
            bg='#ecf0f1',
            font=("Arial", 9),
            fg='#7f8c8d',
            justify=tk.LEFT
        )
        path_info.pack(anchor='w', padx=10, pady=5)

        path_entry_frame = tk.Frame(path_frame, bg='#ecf0f1')
        path_entry_frame.pack(fill=tk.X, padx=10, pady=10)

        path_var = tk.StringVar(value=current_settings.get("database_path", ""))
        path_entry = tk.Entry(
            path_entry_frame,
            textvariable=path_var,
            font=("Arial", 10),
            width=50
        )
        path_entry.pack(side=tk.LEFT, padx=5)

        def browse_folder():
            folder = filedialog.askdirectory(
                title="Выберите папку с файлами базы данных",
                initialdir=path_var.get()
            )
            if folder:
                path_var.set(folder)

        browse_button = tk.Button(
            path_entry_frame,
            text="📂 Обзор...",
            font=("Arial", 10),
            bg='#3498db',
            fg='white',
            command=browse_folder,
            cursor='hand2'
        )
        browse_button.pack(side=tk.LEFT, padx=5)

        # Кнопки Сохранить/Отмена
        buttons_frame = tk.Frame(settings_window, bg='#ecf0f1')
        buttons_frame.pack(pady=20)

        def save_and_close():
            new_path = path_var.get().strip()

            # Проверяем что папка существует
            if not os.path.exists(new_path):
                messagebox.showerror(
                    "Ошибка",
                    f"Папка не существует:\n{new_path}"
                )
                return

            # Сохраняем настройки
            new_settings = {
                "database_path": new_path
            }

            if self.save_settings(new_settings):
                messagebox.showinfo(
                    "Успех",
                    "Настройки сохранены!\n\nПерезапустите программу для применения изменений."
                )
                settings_window.destroy()

        save_button = tk.Button(
            buttons_frame,
            text="💾 Сохранить",
            font=("Arial", 11, "bold"),
            bg='#27ae60',
            fg='white',
            width=15,
            height=2,
            command=save_and_close,
            cursor='hand2'
        )
        save_button.pack(side=tk.LEFT, padx=10)

        cancel_button = tk.Button(
            buttons_frame,
            text="❌ Отмена",
            font=("Arial", 11, "bold"),
            bg='#95a5a6',
            fg='white',
            width=15,
            height=2,
            command=settings_window.destroy,
            cursor='hand2'
        )
        cancel_button.pack(side=tk.LEFT, padx=10)

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

    def auto_resize_columns(self, tree, min_width=80, max_width=None):  # ← None вместо 400
        """Автоматический подбор ширины колонок по содержимому"""
        try:
            import tkinter.font as tkfont

            try:
                font = tkfont.Font(font=tree.cget("font"))
            except:
                font = tkfont.Font(family="Arial", size=10)

            for col in tree["columns"]:
                # Измеряем ширину заголовка
                heading_text = tree.heading(col)["text"]
                heading_width = font.measure(heading_text) + 40

                # Измеряем максимальную ширину значений
                max_content_width = heading_width

                for item_id in tree.get_children():
                    try:
                        col_index = tree["columns"].index(col)
                        value = tree.item(item_id)["values"][col_index]
                        value_str = str(value)
                        value_width = font.measure(value_str) + 30

                        if value_width > max_content_width:
                            max_content_width = value_width
                    except:
                        continue

                # Применяем ограничения
                if max_width is not None:  # ← Добавить проверку
                    optimal_width = max(min_width, min(max_content_width, max_width))
                else:
                    optimal_width = max(min_width, max_content_width)

                tree.column(col, width=int(optimal_width))

                print(f"📏 Колонка '{col}': {int(optimal_width)}px")

        except Exception as e:
            print(f"⚠️ Ошибка автоподбора ширины колонок: {e}")

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
        """Обработчик закрытия приложения"""
        # 🆕 АВТОСОХРАНЕНИЕ ТАБЛИЦЫ ИМПОРТА
        print("\n💾 Сохранение данных перед закрытием...")

        if hasattr(self, 'laser_table_data') and self.laser_table_data:
            self.save_laser_import_cache()

        # Сохраняем настройки переключателей
        self.save_toggle_settings()

        print("✅ Данные сохранены")

        # Закрываем приложение
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

        # НАСТРОЙКА КОЛОНОК БЕЗ РАСТЯГИВАНИЯ
        for col in self.materials_tree["columns"]:
            self.materials_tree.heading(col, text=col)
            self.materials_tree.column(col, anchor=tk.CENTER, width=100, minwidth=80, stretch=False)

        self.materials_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ЦВЕТОВАЯ ИНДИКАЦИЯ (КАК В БАЛАНСЕ)
        self.materials_tree.tag_configure('negative', background='#f8d7da', foreground='#721c24')
        self.materials_tree.tag_configure('available', background='#d4edda', foreground='#155724')
        self.materials_tree.tag_configure('fully_reserved', background='#fff3cd', foreground='#856404')
        self.materials_tree.tag_configure('empty', background='#d1ecf1', foreground='#0c5460')

        # 🆕 ТЕСТОВАЯ ВСТАВКА ДЛЯ ПРОВЕРКИ РАБОТЫ ТЕГОВ
        print("\n🧪 === ТЕСТОВАЯ ВСТАВКА ДЛЯ ПРОВЕРКИ ТЕГОВ ===")
        test_negative = self.materials_tree.insert("", "end",
                                                   values=("TEST1", "ТЕСТ", "1.0", "1000", "1000", "10", "10", "15",
                                                           "-5", "2025-01-01"),
                                                   tags=('negative',))
        test_available = self.materials_tree.insert("", "end",
                                                    values=("TEST2", "ТЕСТ", "2.0", "2000", "2000", "20", "20", "5",
                                                            "15", "2025-01-01"),
                                                    tags=('available',))
        test_reserved = self.materials_tree.insert("", "end",
                                                   values=("TEST3", "ТЕСТ", "3.0", "3000", "3000", "30", "30", "30",
                                                           "0", "2025-01-01"),
                                                   tags=('fully_reserved',))
        test_empty = self.materials_tree.insert("", "end",
                                                values=("TEST4", "ТЕСТ", "4.0", "4000", "4000", "0", "0", "0", "0",
                                                        "2025-01-01"),
                                                tags=('empty',))

        print(f"   Вставлен TEST1 (negative): теги = {self.materials_tree.item(test_negative, 'tags')}")
        print(f"   Вставлен TEST2 (available): теги = {self.materials_tree.item(test_available, 'tags')}")
        print(f"   Вставлен TEST3 (fully_reserved): теги = {self.materials_tree.item(test_reserved, 'tags')}")
        print(f"   Вставлен TEST4 (empty): теги = {self.materials_tree.item(test_empty, 'tags')}")
        print(f"🧪 === ПРОВЕРЬТЕ ТАБЛИЦУ: ВИДНЫ ЛИ 4 ЦВЕТНЫЕ СТРОКИ? ===\n")

        # 🆕 ИНИЦИАЛИЗАЦИЯ EXCEL-ФИЛЬТРА ДЛЯ МАТЕРИАЛОВ
        self.materials_excel_filter = ExcelStyleFilter(
            tree=self.materials_tree,
            refresh_callback=self.refresh_materials
        )

        # 🆕 ИНДИКАТОР АКТИВНЫХ ФИЛЬТРОВ
        self.materials_filter_status = tk.Label(
            self.materials_frame,
            text="",
            font=("Arial", 9),
            bg='#d1ecf1',
            fg='#0c5460'
        )
        self.materials_filter_status.pack(pady=5)

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
        tk.Button(buttons_frame, text="✖ Сбросить фильтры", bg='#e67e22', fg='white',
                  command=self.clear_materials_filters, **btn_style).pack(side=tk.LEFT, padx=5)

        self.refresh_materials()

    def clear_materials_filters(self):
        """Сбросить все фильтры материалов"""
        if hasattr(self, 'materials_excel_filter'):
            self.materials_excel_filter.clear_all_filters()

    def refresh_materials(self):
        """Обновление списка материалов"""

        print(f"\n{'=' * 60}")
        print(f"🔄 НАЧАЛО refresh_materials()")
        print(f"{'=' * 60}")

        # СОХРАНЯЕМ АКТИВНЫЕ ФИЛЬТРЫ ПЕРЕД ОЧИСТКОЙ
        active_filters_backup = {}
        if hasattr(self, 'materials_excel_filter') and self.materials_excel_filter.active_filters:
            active_filters_backup = self.materials_excel_filter.active_filters.copy()
            print(f"🔍 Сохранены фильтры материалов: {list(active_filters_backup.keys())}")

        # ПОЛНОСТЬЮ ОЧИЩАЕМ ДЕРЕВО
        for i in self.materials_tree.get_children():
            self.materials_tree.delete(i)

        # ОЧИЩАЕМ КЭШ ЭЛЕМЕНТОВ
        if hasattr(self, 'materials_excel_filter'):
            self.materials_excel_filter._all_item_cache = set()

        df = load_data("Materials")

        print(f"📊 Загружено материалов из БД: {len(df)}")

        # 🆕 ПРОВЕРКА: НАСТРОЕНЫ ЛИ ТЕГИ?
        print(f"\n🎨 Проверка конфигурации тегов:")
        try:
            print(f"   negative: {self.materials_tree.tag_configure('negative')}")
            print(f"   available: {self.materials_tree.tag_configure('available')}")
            print(f"   fully_reserved: {self.materials_tree.tag_configure('fully_reserved')}")
            print(f"   empty: {self.materials_tree.tag_configure('empty')}")
        except Exception as e:
            print(f"   ❌ ОШИБКА проверки тегов: {e}")

        if not df.empty:
            show_zero_stock = True
            show_zero_available = True

            if hasattr(self, 'materials_toggles') and self.materials_toggles:
                show_zero_stock = self.materials_toggles.get('show_zero_stock', tk.BooleanVar(value=True)).get()
                show_zero_available = self.materials_toggles.get('show_zero_available', tk.BooleanVar(value=True)).get()

            print(f"\n📋 Начало вставки строк:")
            inserted_count = 0
            tag_stats = {'negative': 0, 'available': 0, 'fully_reserved': 0, 'empty': 0}

            for index, row in df.iterrows():
                # 🆕 ДЕТАЛЬНАЯ ДИАГНОСТИКА ЗНАЧЕНИЙ
                quantity_raw = row["Количество штук"]
                available_raw = row["Доступно"]

                try:
                    quantity = int(quantity_raw) if quantity_raw else 0
                except:
                    quantity = 0

                try:
                    available = int(available_raw) if available_raw else 0
                except:
                    available = 0

                if not show_zero_stock and quantity == 0:
                    continue
                if not show_zero_available and available == 0:
                    continue

                values = (row["ID"], row["Марка"], row["Толщина"], row["Длина"], row["Ширина"],
                          row["Количество штук"], row["Общая площадь"], row["Зарезервировано"],
                          row["Доступно"], row["Дата добавления"])

                # 🆕 ОПРЕДЕЛЯЕМ ТЕГ С ДЕТАЛЬНОЙ ДИАГНОСТИКОЙ
                tag = None

                if available < 0:
                    tag = 'negative'
                    tag_stats['negative'] += 1
                elif available > 0:
                    tag = 'available'
                    tag_stats['available'] += 1
                elif available == 0 and quantity > 0:
                    tag = 'fully_reserved'
                    tag_stats['fully_reserved'] += 1
                else:
                    tag = 'empty'
                    tag_stats['empty'] += 1

                # Диагностика первых 5 строк
                if inserted_count < 5:
                    print(f"   Строка {inserted_count}:")
                    print(f"      ID: {row['ID']}, Марка: {row['Марка']}")
                    print(f"      Кол-во (raw): '{quantity_raw}' (type: {type(quantity_raw).__name__})")
                    print(f"      Кол-во (parsed): {quantity} (type: {type(quantity).__name__})")
                    print(f"      Доступно (raw): '{available_raw}' (type: {type(available_raw).__name__})")
                    print(f"      Доступно (parsed): {available} (type: {type(available).__name__})")
                    print(f"      Условие: available={available}, quantity={quantity}")
                    print(f"      ✅ Тег: {tag}")

                # Вставляем с тегом
                item_id = self.materials_tree.insert("", "end", values=values, tags=(tag,))
                inserted_count += 1

                # 🆕 ПРОВЕРЯЕМ ЧТО ТЕГ ДЕЙСТВИТЕЛЬНО ПРИМЕНИЛСЯ
                if inserted_count <= 5:
                    actual_tags = self.materials_tree.item(item_id, 'tags')
                    print(f"      Проверка: фактические теги элемента = {actual_tags}")

                # СОХРАНЯЕМ item_id В КЭШ
                if hasattr(self, 'materials_excel_filter'):
                    if not hasattr(self.materials_excel_filter, '_all_item_cache'):
                        self.materials_excel_filter._all_item_cache = set()
                    self.materials_excel_filter._all_item_cache.add(item_id)

            print(f"\n✅ Вставлено строк: {inserted_count}")
            print(f"📊 Статистика по тегам:")
            print(f"   🔴 negative (красный): {tag_stats['negative']}")
            print(f"   🟢 available (зелёный): {tag_stats['available']}")
            print(f"   🟡 fully_reserved (жёлтый): {tag_stats['fully_reserved']}")
            print(f"   🔵 empty (голубой): {tag_stats['empty']}")

        # АВТОПОДБОР ШИРИНЫ КОЛОНОК
        self.auto_resize_columns(self.materials_tree, min_width=80, max_width=200)

        # ПЕРЕПРИМЕНЯЕМ ФИЛЬТРЫ ПОСЛЕ ЗАГРУЗКИ ДАННЫХ
        if active_filters_backup and hasattr(self, 'materials_excel_filter'):
            print(f"\n🔄 Переприменяю фильтры материалов: {list(active_filters_backup.keys())}")
            self.materials_excel_filter.active_filters = active_filters_backup
            self.materials_excel_filter.reapply_all_filters()

        # 🆕 ФИНАЛЬНАЯ ПРОВЕРКА: ЕСТЬ ЛИ ТЕГИ У ВИДИМЫХ ЭЛЕМЕНТОВ?
        print(f"\n🔍 Финальная проверка тегов видимых элементов:")
        visible_items = list(self.materials_tree.get_children(''))
        print(f"   Всего видимых элементов: {len(visible_items)}")

        for i, item_id in enumerate(visible_items[:3]):  # Первые 3
            tags = self.materials_tree.item(item_id, 'tags')
            values = self.materials_tree.item(item_id, 'values')
            print(f"   Элемент {i}: ID={values[0]}, Доступно={values[8]}, Теги={tags}")

        print(f"\n{'=' * 60}")
        print(f"✅ КОНЕЦ refresh_materials()")
        print(f"{'=' * 60}\n")

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

        # 🆕 СОХРАНЯЕМ СТАРОЕ КОЛИЧЕСТВО ДЛЯ СРАВНЕНИЯ
        old_quantity = int(row["Количество штук"])

        edit_window = tk.Toplevel(self.root)
        edit_window.title("Редактировать материал")
        edit_window.geometry("450x600")  # ← УВЕЛИЧИЛИ ВЫСОТУ ДЛЯ КОММЕНТАРИЯ
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

        # 🆕 ПОЛЕ ДЛЯ КОММЕНТАРИЯ (ЕСЛИ ИЗМЕНИЛОСЬ КОЛИЧЕСТВО)
        comment_frame = tk.Frame(edit_window, bg='#ecf0f1')
        comment_frame.pack(fill=tk.X, padx=20, pady=10)
        tk.Label(comment_frame, text="Комментарий\n(если меняете кол-во):",
                 width=20, anchor='w', bg='#ecf0f1', font=("Arial", 10)).pack(side=tk.TOP, anchor='w')
        comment_entry = tk.Text(comment_frame, font=("Arial", 10), height=3, width=40)
        comment_entry.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=5)

        def save_changes():
            try:
                thickness = float(entries["Толщина"].get())
                length = float(entries["Длина"].get())
                width = float(entries["Ширина"].get())
                new_quantity = int(entries["Количество штук"].get())
                reserved = int(row["Зарезервировано"])
                area = (length * width * new_quantity) / 1000000

                # 🆕 ПРОВЕРКА: ИЗМЕНИЛОСЬ ЛИ КОЛИЧЕСТВО?
                quantity_changed = (new_quantity != old_quantity)

                if quantity_changed:
                    comment_text = comment_entry.get("1.0", tk.END).strip()

                    if not comment_text:
                        response = messagebox.askyesno(
                            "Комментарий отсутствует",
                            "Количество изменилось, но комментарий не указан.\n\n"
                            "Продолжить без комментария?"
                        )
                        if not response:
                            return
                        comment_text = "(без комментария)"

                    # 🆕 ЗАПИСЫВАЕМ ЛОГ ИЗМЕНЕНИЯ
                    self.log_material_change(
                        material_id=item_id,
                        marka=entries["Марка"].get(),
                        thickness=thickness,
                        length=length,
                        width=width,
                        old_qty=old_quantity,
                        new_qty=new_quantity,
                        comment=comment_text
                    )

                # СОХРАНЯЕМ ИЗМЕНЕНИЯ В МАТЕРИАЛАХ
                df.loc[df["ID"] == item_id, "Марка"] = entries["Марка"].get()
                df.loc[df["ID"] == item_id, "Толщина"] = thickness
                df.loc[df["ID"] == item_id, "Длина"] = length
                df.loc[df["ID"] == item_id, "Ширина"] = width
                df.loc[df["ID"] == item_id, "Количество штук"] = new_quantity
                df.loc[df["ID"] == item_id, "Общая площадь"] = round(area, 2)
                df.loc[df["ID"] == item_id, "Доступно"] = new_quantity - reserved

                save_data("Materials", df)
                self.refresh_materials()
                self.refresh_balance()
                edit_window.destroy()

                if quantity_changed:
                    messagebox.showinfo("Успех",
                                        f"Материал обновлен!\n\n"
                                        f"Количество изменено: {old_quantity} → {new_quantity}\n"
                                        f"Изменение: {new_quantity - old_quantity:+d} шт.\n"
                                        f"Лог записан.")
                else:
                    messagebox.showinfo("Успех", "Материал успешно обновлен!")

            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось обновить материал: {e}")

        tk.Button(edit_window, text="Сохранить", bg='#3498db', fg='white', font=("Arial", 12, "bold"),
                  command=save_changes).pack(pady=20)

    def log_material_change(self, material_id, marka, thickness, length, width, old_qty, new_qty, comment):
        """Логирование изменения количества материала вручную"""
        try:
            # Загружаем логи (или создаём пустой DataFrame если листа нет)
            try:
                logs_df = load_data("MaterialChangeLogs")
            except:
                logs_df = pd.DataFrame(columns=[
                    "ID лога", "Дата и время", "ID материала", "Марка", "Толщина",
                    "Длина", "Ширина", "Старое кол-во", "Новое кол-во", "Изменение", "Комментарий"
                ])

            # Генерируем ID лога
            if logs_df.empty:
                log_id = 1
            else:
                log_id = int(logs_df["ID лога"].max()) + 1

            # Вычисляем изменение
            change = new_qty - old_qty

            # 🆕 ПРАВИЛЬНЫЙ ФОРМАТ: число с явным знаком
            if change > 0:
                change_str = f"+{change}"
            elif change < 0:
                change_str = str(change)  # минус уже есть
            else:
                change_str = "0"

            print(f"🔍 Логирование изменения: старое={old_qty}, новое={new_qty}, изменение='{change_str}'")

            # Создаём новую запись
            new_log = pd.DataFrame([{
                "ID лога": log_id,
                "Дата и время": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ID материала": material_id,
                "Марка": marka,
                "Толщина": thickness,
                "Длина": length,
                "Ширина": width,
                "Старое кол-во": old_qty,
                "Новое кол-во": new_qty,
                "Изменение": change_str,  # ← СТРОКА С ЗНАКОМ: "+5" или "-3"
                "Комментарий": comment
            }])

            # Добавляем в логи
            logs_df = pd.concat([logs_df, new_log], ignore_index=True)

            # Сохраняем
            save_data("MaterialChangeLogs", logs_df)

            print(
                f"✅ Лог изменения записан: ID материала={material_id}, изменение={change_str}, комментарий='{comment}'")

            # АВТОМАТИЧЕСКОЕ ОБНОВЛЕНИЕ ВКЛАДКИ "История материалов"
            if hasattr(self, 'material_logs_tree'):
                self.refresh_material_logs()

        except Exception as e:
            print(f"⚠️ Ошибка записи лога изменения материала: {e}")
            import traceback
            traceback.print_exc()

    def setup_material_logs_tab(self):
        """Вкладка истории изменений количества материалов"""

        # Заголовок
        header_frame = tk.Frame(self.material_logs_frame, bg='white')
        header_frame.pack(fill=tk.X, pady=10)

        tk.Label(header_frame, text="История изменений количества материалов",
                 font=("Arial", 16, "bold"), bg='white', fg='#2c3e50').pack()

        # 🆕 ИНДИКАТОР КОЛИЧЕСТВА ЗАПИСЕЙ
        self.material_logs_status = tk.Label(
            header_frame,
            text="Загрузка...",
            font=("Arial", 10),
            bg='#d1ecf1',
            fg='#0c5460',
            relief=tk.RIDGE,
            padx=10,
            pady=5
        )
        self.material_logs_status.pack(pady=5)

        # Таблица
        tree_frame = tk.Frame(self.material_logs_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        self.material_logs_tree = ttk.Treeview(
            tree_frame,
            columns=("ID лога", "Дата", "ID материала", "Марка", "Толщина", "Размер",
                     "Старое", "Новое", "Изменение", "Комментарий"),
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        scroll_y.config(command=self.material_logs_tree.yview)
        scroll_x.config(command=self.material_logs_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        columns_config = {
            "ID лога": 70,
            "Дата": 140,
            "ID материала": 90,
            "Марка": 100,
            "Толщина": 70,
            "Размер": 110,
            "Старое": 80,
            "Новое": 80,
            "Изменение": 90,
            "Комментарий": 250
        }

        for col, width in columns_config.items():
            self.material_logs_tree.heading(col, text=col)
            self.material_logs_tree.column(col, width=width, anchor=tk.CENTER)

        self.material_logs_tree.pack(fill=tk.BOTH, expand=True)

        # 🆕 ЦВЕТОВАЯ ИНДИКАЦИЯ СТРОК (ЗЕЛЁНЫЙ = ДОБАВЛЕНИЕ, КРАСНЫЙ = УМЕНЬШЕНИЕ)
        self.material_logs_tree.tag_configure('increase', background='#d4edda')  # Зелёный
        self.material_logs_tree.tag_configure('decrease', background='#f8d7da')  # Красный
        self.material_logs_tree.tag_configure('neutral', background='white')

        # 🆕 ИНИЦИАЛИЗАЦИЯ EXCEL-ФИЛЬТРА ДЛЯ РЕЗЕРВИРОВАНИЯ
        self.reservations_excel_filter = ExcelStyleFilter(
            tree=self.reservations_tree,
            refresh_callback=self.refresh_reservations
        )

        # 🆕 ИНДИКАТОР АКТИВНЫХ ФИЛЬТРОВ
        self.reservations_filter_status = tk.Label(
            self.reservations_frame,
            text="",
            font=("Arial", 9),
            bg='#d1ecf1',
            fg='#0c5460'
        )
        self.reservations_filter_status.pack(pady=5)

        # Кнопки управления
        buttons_frame = tk.Frame(self.material_logs_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        btn_style = {"font": ("Arial", 10), "width": 15, "height": 2}

        tk.Button(buttons_frame, text="Обновить", bg='#95a5a6', fg='white',
                  command=self.refresh_material_logs, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="Экспорт в Excel", bg='#3498db', fg='white',
                  command=self.export_material_logs, **btn_style).pack(side=tk.LEFT, padx=5)

        # Первоначальная загрузка
        self.refresh_material_logs()

    def refresh_material_logs(self):
        """Обновление таблицы логов"""
        for item in self.material_logs_tree.get_children():
            self.material_logs_tree.delete(item)

        try:
            logs_df = load_data("MaterialChangeLogs")

            if not logs_df.empty:
                # Сортируем по дате (новые сверху)
                logs_df = logs_df.sort_values("Дата и время", ascending=False)

                for _, log in logs_df.iterrows():
                    size_str = f"{int(log['Длина'])}x{int(log['Ширина'])}"

                    values = (
                        int(log["ID лога"]),
                        log["Дата и время"],
                        int(log["ID материала"]),
                        log["Марка"],
                        log["Толщина"],
                        size_str,
                        int(log["Старое кол-во"]),
                        int(log["Новое кол-во"]),
                        log["Изменение"],
                        log["Комментарий"]
                    )

                    # 🆕 ЦВЕТОВАЯ ИНДИКАЦИЯ ПО ИЗМЕНЕНИЮ
                    change_str = str(log["Изменение"])
                    if change_str.startswith('+'):
                        tag = 'increase'  # Зелёный (добавление)
                    elif change_str.startswith('-'):
                        tag = 'decrease'  # Красный (уменьшение)
                    else:
                        tag = 'neutral'

                    self.material_logs_tree.insert("", "end", values=values, tags=(tag,))

                # 🆕 ОБНОВЛЯЕМ СТАТУС
                total = len(logs_df)
                increase_count = len(logs_df[logs_df["Изменение"].str.startswith('+')])
                decrease_count = len(logs_df[logs_df["Изменение"].str.startswith('-')])

                status_text = (
                    f"📊 Всего записей: {total} | "
                    f"🟢 Добавлений: {increase_count} | "
                    f"🔴 Уменьшений: {decrease_count}"
                )
                self.material_logs_status.config(text=status_text, bg='#d1ecf1', fg='#0c5460')
            else:
                self.material_logs_status.config(
                    text="ℹ️ История изменений пуста",
                    bg='#fff3cd',
                    fg='#856404'
                )

            self.auto_resize_columns(self.material_logs_tree)

        except Exception as e:
            print(f"⚠️ Ошибка загрузки логов: {e}")
            import traceback
            traceback.print_exc()
            self.material_logs_status.config(
                text=f"❌ Ошибка загрузки: {e}",
                bg='#f8d7da',
                fg='#721c24'
            )

    def export_material_logs(self):
        """Экспорт истории изменений в Excel"""
        try:
            logs_df = load_data("MaterialChangeLogs")

            if logs_df.empty:
                messagebox.showwarning("Предупреждение", "Нет данных для экспорта!")
                return

            file_path = filedialog.asksaveasfilename(
                title="Экспорт истории изменений",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"material_changes_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )

            if not file_path:
                return

            # Сортируем по дате (старые сверху для Excel)
            logs_df = logs_df.sort_values("Дата и время", ascending=True)

            # Экспортируем
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                logs_df.to_excel(writer, index=False, sheet_name='История изменений')
                worksheet = writer.sheets['История изменений']

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
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            messagebox.showinfo("Успех", f"История изменений экспортирована:\n\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать:\n{e}")

    def refresh_material_logs(self):
        """Обновление таблицы логов"""
        for item in self.material_logs_tree.get_children():
            self.material_logs_tree.delete(item)

        try:
            logs_df = load_data("MaterialChangeLogs")

            if not logs_df.empty:
                # 🆕 ДИАГНОСТИКА: проверяем формат данных
                print(f"🔍 Загружено логов: {len(logs_df)}")
                if len(logs_df) > 0:
                    first_change = logs_df.iloc[0]["Изменение"]
                    print(f"🔍 Первая запись 'Изменение': '{first_change}' (тип: {type(first_change).__name__})")

                # Сортируем по дате (новые сверху)
                logs_df = logs_df.sort_values("Дата и время", ascending=False)

                # Счётчики
                increase_count = 0
                decrease_count = 0

                for _, log in logs_df.iterrows():
                    size_str = f"{int(log['Длина'])}x{int(log['Ширина'])}"

                    # 🆕 БЕЗОПАСНОЕ ПРЕОБРАЗОВАНИЕ "Изменение" В СТРОКУ
                    change_value = log["Изменение"]

                    if pd.isna(change_value):
                        change_str = "0"
                    else:
                        change_str = str(change_value).strip()

                    values = (
                        int(log["ID лога"]),
                        log["Дата и время"],
                        int(log["ID материала"]),
                        log["Марка"],
                        log["Толщина"],
                        size_str,
                        int(log["Старое кол-во"]),
                        int(log["Новое кол-во"]),
                        change_str,  # ← ИСПОЛЬЗУЕМ ПРЕОБРАЗОВАННУЮ СТРОКУ
                        log["Комментарий"]
                    )

                    # 🆕 ПРАВИЛЬНАЯ ЦВЕТОВАЯ ИНДИКАЦИЯ
                    try:
                        # Пробуем распарсить изменение как число
                        change_num = int(change_str.replace('+', '').replace(' ', ''))

                        if change_num > 0:
                            tag = 'increase'  # Зелёный (добавление)
                            increase_count += 1
                        elif change_num < 0:
                            tag = 'decrease'  # Красный (уменьшение)
                            decrease_count += 1
                        else:
                            tag = 'neutral'
                    except:
                        tag = 'neutral'

                    self.material_logs_tree.insert("", "end", values=values, tags=(tag,))

                # 🆕 ОБНОВЛЯЕМ СТАТУС
                total = len(logs_df)

                status_text = (
                    f"📊 Всего записей: {total} | "
                    f"🟢 Добавлений: {increase_count} | "
                    f"🔴 Уменьшений: {decrease_count}"
                )

                print(f"📊 Статистика: всего={total}, добавлений={increase_count}, уменьшений={decrease_count}")

                self.material_logs_status.config(text=status_text, bg='#d1ecf1', fg='#0c5460')
            else:
                self.material_logs_status.config(
                    text="ℹ️ История изменений пуста",
                    bg='#fff3cd',
                    fg='#856404'
                )

            self.auto_resize_columns(self.material_logs_tree)

        except Exception as e:
            print(f"⚠️ Ошибка загрузки логов: {e}")
            import traceback
            traceback.print_exc()
            self.material_logs_status.config(
                text=f"❌ Ошибка загрузки: {e}",
                bg='#f8d7da',
                fg='#721c24'
            )

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

        # ========== ТАБЛИЦА ЗАКАЗОВ ==========
        orders_label = tk.Label(self.orders_frame, text="Список заказов", font=("Arial", 12, "bold"), bg='white')
        orders_label.pack(pady=5)

        # 🆕 Фрейм таблицы заказов НА ВСЕЙ ШИРИНЕ (убрано центрирование)
        orders_tree_frame = tk.Frame(self.orders_frame, bg='white')
        orders_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scroll_y = tk.Scrollbar(orders_tree_frame, orient=tk.VERTICAL)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.orders_tree = ttk.Treeview(orders_tree_frame,
                                        columns=("ID", "Название", "Заказчик", "Дата", "Статус", "Примечания"),
                                        show="headings", yscrollcommand=scroll_y.set, height=8)
        scroll_y.config(command=self.orders_tree.yview)

        # Настройка колонок БЕЗ растягивания
        for col in self.orders_tree["columns"]:
            self.orders_tree.heading(col, text=col)
            self.orders_tree.column(col, anchor=tk.CENTER, width=100, minwidth=80, stretch=False)

        self.orders_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.orders_tree.bind('<<TreeviewSelect>>', self.on_order_select)

        # ИНИЦИАЛИЗАЦИЯ EXCEL-ФИЛЬТРА ДЛЯ ЗАКАЗОВ
        self.orders_excel_filter = ExcelStyleFilter(
            tree=self.orders_tree,
            refresh_callback=self.refresh_orders
        )

        # ИНДИКАТОР АКТИВНЫХ ФИЛЬТРОВ (ЗАКАЗЫ)
        self.orders_filter_status = tk.Label(
            self.orders_frame,
            text="",
            font=("Arial", 9),
            bg='#d1ecf1',
            fg='#0c5460'
        )
        self.orders_filter_status.pack(pady=5)

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

        # Кнопки управления заказами
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
        tk.Button(buttons_frame, text="✖ Сбросить фильтры", bg='#e67e22', fg='white',
                  command=self.clear_orders_filters, **btn_style).pack(side=tk.LEFT, padx=5)

        # ========== ТАБЛИЦА ДЕТАЛЕЙ ЗАКАЗА ==========
        details_label = tk.Label(self.orders_frame, text="Детали выбранного заказа", font=("Arial", 12, "bold"),
                                 bg='white')
        details_label.pack(pady=5)

        # 🆕 Фрейм таблицы деталей НА ВСЕЙ ШИРИНЕ (убрано центрирование)
        details_tree_frame = tk.Frame(self.orders_frame, bg='white')
        details_tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scroll_y2 = tk.Scrollbar(details_tree_frame, orient=tk.VERTICAL)
        scroll_y2.pack(side=tk.RIGHT, fill=tk.Y)

        self.order_details_tree = ttk.Treeview(details_tree_frame,
                                               columns=("ID", "ID заказа", "Название детали", "Количество", "Порезано",
                                                        "Погнуто"),
                                               show="headings", yscrollcommand=scroll_y2.set)
        scroll_y2.config(command=self.order_details_tree.yview)

        # Настройка колонок БЕЗ растягивания
        for col in self.order_details_tree["columns"]:
            self.order_details_tree.heading(col, text=col)
            self.order_details_tree.column(col, anchor=tk.CENTER, width=100, minwidth=80, stretch=False)

        self.order_details_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.order_details_tree.bind('<Double-1>', self.on_detail_double_click)

        # Привязка правого клика для копирования
        self.order_details_tree.bind('<Button-3>', self.on_detail_right_click)

        # ИНИЦИАЛИЗАЦИЯ EXCEL-ФИЛЬТРА ДЛЯ ДЕТАЛЕЙ
        self.order_details_excel_filter = ExcelStyleFilter(
            tree=self.order_details_tree,
            refresh_callback=self.refresh_order_details
        )

        # ИНДИКАТОР АКТИВНЫХ ФИЛЬТРОВ (ДЕТАЛИ)
        self.order_details_filter_status = tk.Label(
            self.orders_frame,
            text="",
            font=("Arial", 9),
            bg='#d1ecf1',
            fg='#0c5460'
        )
        self.order_details_filter_status.pack(pady=5)

        # Кнопки управления деталями
        details_buttons_frame = tk.Frame(self.orders_frame, bg='white')
        details_buttons_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(details_buttons_frame, text="Добавить деталь", bg='#27ae60', fg='white',
                  command=self.add_order_detail, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(details_buttons_frame, text="Редактировать деталь", bg='#f39c12', fg='white',
                  command=self.edit_order_detail, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(details_buttons_frame, text="Удалить деталь", bg='#e74c3c', fg='white',
                  command=self.delete_order_detail, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(details_buttons_frame, text="✖ Сбросить фильтры", bg='#e67e22', fg='white',
                  command=self.clear_order_details_filters, **btn_style).pack(side=tk.LEFT, padx=5)

        self.refresh_orders()

    def clear_orders_filters(self):
        """Сбросить все фильтры заказов"""
        if hasattr(self, 'orders_excel_filter'):
            self.orders_excel_filter.clear_all_filters()

    def clear_order_details_filters(self):
        """Сбросить все фильтры деталей заказа"""
        if hasattr(self, 'order_details_excel_filter'):
            self.order_details_excel_filter.clear_all_filters()

    def on_order_select(self, event):
        self.refresh_order_details()

    def on_order_select(self, event):
        self.refresh_order_details()

    def refresh_orders(self):
        """Обновление списка заказов"""

        # СОХРАНЯЕМ АКТИВНЫЕ ФИЛЬТРЫ ПЕРЕД ОЧИСТКОЙ
        active_filters_backup = {}
        if hasattr(self, 'orders_excel_filter') and self.orders_excel_filter.active_filters:
            active_filters_backup = self.orders_excel_filter.active_filters.copy()
            print(f"🔍 Сохранены фильтры заказов: {list(active_filters_backup.keys())}")

        # ПОЛНОСТЬЮ ОЧИЩАЕМ ДЕРЕВО
        for i in self.orders_tree.get_children():
            self.orders_tree.delete(i)

        # ОЧИЩАЕМ КЭШ ЭЛЕМЕНТОВ
        if hasattr(self, 'orders_excel_filter'):
            self.orders_excel_filter._all_item_cache = set()

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

                values = (row["ID заказа"], row["Название заказа"], row["Заказчик"],
                          row["Дата создания"], row["Статус"], row["Примечания"])

                item_id = self.orders_tree.insert("", "end", values=values)

                # СОХРАНЯЕМ item_id В КЭШ
                if hasattr(self, 'orders_excel_filter'):
                    if not hasattr(self.orders_excel_filter, '_all_item_cache'):
                        self.orders_excel_filter._all_item_cache = set()
                    self.orders_excel_filter._all_item_cache.add(item_id)

        # АВТОПОДБОР ШИРИНЫ КОЛОНОК
        self.auto_resize_columns(self.orders_tree, min_width=100, max_width=300)

        # ПЕРЕПРИМЕНЯЕМ ФИЛЬТРЫ ПОСЛЕ ЗАГРУЗКИ ДАННЫХ
        if active_filters_backup and hasattr(self, 'orders_excel_filter'):
            print(f"🔄 Переприменяю фильтры заказов: {list(active_filters_backup.keys())}")
            self.orders_excel_filter.active_filters = active_filters_backup
            self.orders_excel_filter.reapply_all_filters()

    def refresh_order_details(self):
        """Обновление деталей выбранного заказа"""

        print(f"\n🔍 refresh_order_details вызван")

        # СОХРАНЯЕМ АКТИВНЫЕ ФИЛЬТРЫ ПЕРЕД ОЧИСТКОЙ
        active_filters_backup = {}
        if hasattr(self, 'order_details_excel_filter') and self.order_details_excel_filter.active_filters:
            active_filters_backup = self.order_details_excel_filter.active_filters.copy()
            print(f"   Сохранены фильтры деталей: {list(active_filters_backup.keys())}")

        # ПОЛНОСТЬЮ ОЧИЩАЕМ ДЕРЕВО
        for i in self.order_details_tree.get_children():
            self.order_details_tree.delete(i)

        # ОЧИЩАЕМ КЭШ ЭЛЕМЕНТОВ
        if hasattr(self, 'order_details_excel_filter'):
            self.order_details_excel_filter._all_item_cache = set()

        selected = self.orders_tree.selection()
        if not selected:
            print(f"   ❌ Заказ не выбран")
            return

        order_id = self.orders_tree.item(selected[0])["values"][0]
        print(f"   ✅ Выбран заказ ID: {order_id}")

        df = load_data("OrderDetails")
        print(f"   📊 Загружено деталей всего: {len(df)}")

        if not df.empty:
            # 🆕 ДИАГНОСТИКА: ПОКАЗЫВАЕМ КОЛОНКИ
            print(f"   📋 Колонки OrderDetails: {list(df.columns)}")

            # Фильтруем детали по ID заказа (используем iloc[1] - вторая колонка)
            try:
                details = df[df.iloc[:, 1] == order_id]
            except:
                # Попытка по названию колонки
                if "ID заказа" in df.columns:
                    details = df[df["ID заказа"] == order_id]
                else:
                    details = pd.DataFrame()  # Пустой DataFrame

            print(f"   📊 Деталей для заказа {order_id}: {len(details)}")

            for index, row in details.iterrows():
                # 🆕 ИСПОЛЬЗУЕМ ИНДЕКСЫ КОЛОНОК (НАДЁЖНО)
                try:
                    detail_id = row.iloc[0]  # ID детали
                    order_id_val = row.iloc[1]  # ID заказа
                    detail_name = row.iloc[2]  # Название детали
                    quantity = row.iloc[3]  # Количество
                    cut = row.iloc[4] if len(row) > 4 else 0  # Порезано
                    bent = row.iloc[5] if len(row) > 5 else 0  # Погнуто

                    values = (detail_id, order_id_val, detail_name, quantity, cut, bent)

                    print(f"      ✅ Вставка детали: {values}")
                    item_id = self.order_details_tree.insert("", "end", values=values)

                    # СОХРАНЯЕМ item_id В КЭШ
                    if hasattr(self, 'order_details_excel_filter'):
                        if not hasattr(self.order_details_excel_filter, '_all_item_cache'):
                            self.order_details_excel_filter._all_item_cache = set()
                        self.order_details_excel_filter._all_item_cache.add(item_id)

                except Exception as e:
                    print(f"      ⚠️ Ошибка чтения детали {index}: {e}")
                    continue
        else:
            print(f"   ❌ DataFrame OrderDetails пуст")

        # Проверяем сколько элементов в дереве
        visible_items = self.order_details_tree.get_children()
        print(f"   📊 Видимых элементов в дереве: {len(visible_items)}")

        # АВТОПОДБОР ШИРИНЫ КОЛОНОК
        self.auto_resize_columns(self.order_details_tree, min_width=100, max_width=300)

        # ПЕРЕПРИМЕНЯЕМ ФИЛЬТРЫ ПОСЛЕ ЗАГРУЗКИ ДАННЫХ
        if active_filters_backup and hasattr(self, 'order_details_excel_filter'):
            print(f"   🔄 Переприменяю фильтры деталей: {list(active_filters_backup.keys())}")
            self.order_details_excel_filter.active_filters = active_filters_backup
            self.order_details_excel_filter.reapply_all_filters()

        # Финальная проверка
        final_visible = self.order_details_tree.get_children()
        print(f"   ✅ Итого видимых элементов: {len(final_visible)}\n")


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

    def on_detail_right_click(self, event):
        """Контекстное меню при правом клике на деталь"""
        # Определяем строку под курсором
        item = self.order_details_tree.identify_row(event.y)
        if not item:
            return

        # Выделяем строку
        self.order_details_tree.selection_set(item)

        # Создаём контекстное меню
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(
            label="📋 Копировать информацию о детали",
            command=lambda: self.copy_detail_info(item)
        )

        # Показываем меню
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def copy_detail_info(self, item_id):
        """Копирование информации о детали в буфер обмена"""
        try:
            # Получаем данные из строки таблицы
            values = self.order_details_tree.item(item_id)["values"]
            if not values or len(values) < 6:
                messagebox.showwarning("Предупреждение", "Не удалось получить данные детали")
                return

            detail_id = values[0]
            order_id = values[1]
            detail_name = values[2]
            quantity = int(values[3])
            cut = int(values[4])
            bent = int(values[5])

            # Рассчитываем остаток для нарезки
            remaining_to_cut = quantity - cut

            # Загружаем данные заказа
            orders_df = load_data("Orders")
            order_row = orders_df[orders_df["ID заказа"] == order_id]

            if order_row.empty:
                customer = "Неизвестно"
                order_name = "Неизвестно"
            else:
                customer = order_row.iloc[0]["Заказчик"]
                order_name = order_row.iloc[0]["Название заказа"]

            # Загружаем данные резервирования
            reservations_df = load_data("Reservations")

            # Ищем резервы для этой детали и заказа
            detail_reserves = reservations_df[
                (reservations_df["ID заказа"] == order_id) &
                (reservations_df["ID детали"] == detail_id)
                ]

            # Переменные для вывода
            material_info = ""
            material_stock = ""
            remaining_reserved_count = ""

            if not detail_reserves.empty:
                # Загружаем данные материалов для проверки остатка на складе
                materials_df = load_data("Materials")

                material_parts = []
                stock_parts = []
                remaining_count_list = []

                for _, reserve in detail_reserves.iterrows():
                    material_id = reserve["ID материала"]
                    marka = reserve["Марка"]
                    thickness = reserve["Толщина"]
                    width = reserve["Ширина"]
                    length = reserve["Длина"]
                    remaining_qty = int(reserve["Остаток к списанию"])

                    # Описание материала (без количества)
                    material_desc = f"{marka} {thickness}мм {width}x{length}"

                    # Добавляем только количество листов к списанию
                    if remaining_qty > 0:
                        remaining_count_list.append(str(remaining_qty))

                    # Ищем ОБЩИЙ фактический остаток на складе (колонка "Количество штук")
                    if material_id != -1 and not materials_df.empty:
                        material_row = materials_df[materials_df["ID"] == material_id]
                        if not material_row.empty:
                            total_quantity = int(material_row.iloc[0]["Количество штук"])
                            material_parts.append(material_desc)
                            stock_parts.append(str(total_quantity))

                material_info = "; ".join(material_parts) if material_parts else ""
                material_stock = "; ".join(stock_parts) if stock_parts else ""
                remaining_reserved_count = "; ".join(remaining_count_list) if remaining_count_list else ""

            # Формируем текст для копирования
            parts = [
                f"{customer} | {order_name}",
                f"{detail_name}",
                f"Осталось: {remaining_to_cut} шт"
            ]

            # Добавляем материал (если есть)
            if material_info:
                parts.append(f"Материал: {material_info}")
            else:
                parts.append("Материал: ")

            # Добавляем остаток на складе (если есть)
            if material_stock:
                parts.append(f"Остаток на складе: {material_stock} шт")
            else:
                parts.append("Остаток на складе: ")

            # Добавляем количество к списанию (если есть)
            if remaining_reserved_count:
                parts.append(f"Остаток порезать: {remaining_reserved_count} шт")
            else:
                parts.append("Остаток порезать: ")

            copy_text = " | ".join(parts)

            # Копируем в буфер обмена
            self.root.clipboard_clear()
            self.root.clipboard_append(copy_text)
            self.root.update()  # Обновляем буфер обмена

            # Показываем уведомление
            messagebox.showinfo(
                "Скопировано",
                f"Информация о детали скопирована в буфер обмена:\n\n{copy_text}"
            )

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось скопировать информацию: {e}")
            import traceback
            traceback.print_exc()

    def on_details_tab_right_click(self, event):
        """Контекстное меню при правом клике на деталь во вкладке 'Учёт деталей'"""
        # Определяем строку под курсором
        item = self.details_tree.identify_row(event.y)
        if not item:
            return

        # Выделяем строку
        self.details_tree.selection_set(item)

        # Создаём контекстное меню
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(
            label="📋 Копировать информацию о детали",
            command=lambda: self.copy_details_tab_info(item)
        )

        # Показываем меню
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def copy_details_tab_info(self, item_id):
        """Копирование информации о детали из вкладки 'Учёт деталей' в буфер обмена"""
        try:
            # Получаем данные из строки таблицы
            values = self.details_tree.item(item_id)["values"]
            if not values or len(values) < 9:
                messagebox.showwarning("Предупреждение", "Не удалось получить данные детали")
                return

            detail_id = values[0]
            customer = values[1]
            detail_name = values[2]
            order_name = values[3]
            quantity = int(values[4])
            cut = int(values[5])
            bent = int(values[6])
            remaining = int(values[7])

            # Загружаем данные резервирования
            reservations_df = load_data("Reservations")

            # Ищем резервы для этой детали
            detail_reserves = reservations_df[
                reservations_df["ID детали"] == detail_id
                ]

            # Переменные для вывода
            material_info = ""
            material_stock = ""
            remaining_reserved_count = ""

            if not detail_reserves.empty:
                # Загружаем данные материалов для проверки остатка на складе
                materials_df = load_data("Materials")

                material_parts = []
                stock_parts = []
                remaining_count_list = []

                for _, reserve in detail_reserves.iterrows():
                    material_id = reserve["ID материала"]
                    marka = reserve["Марка"]
                    thickness = reserve["Толщина"]
                    width = reserve["Ширина"]
                    length = reserve["Длина"]
                    remaining_qty = int(reserve["Остаток к списанию"])

                    # Описание материала (без количества)
                    material_desc = f"{marka} {thickness}мм {width}x{length}"

                    # Добавляем только количество листов к списанию
                    if remaining_qty > 0:
                        remaining_count_list.append(str(remaining_qty))

                    # Ищем ОБЩИЙ фактический остаток на складе (колонка "Количество штук")
                    if material_id != -1 and not materials_df.empty:
                        material_row = materials_df[materials_df["ID"] == material_id]
                        if not material_row.empty:
                            total_quantity = int(material_row.iloc[0]["Количество штук"])
                            material_parts.append(material_desc)
                            stock_parts.append(str(total_quantity))

                material_info = "; ".join(material_parts) if material_parts else ""
                material_stock = "; ".join(stock_parts) if stock_parts else ""
                remaining_reserved_count = "; ".join(remaining_count_list) if remaining_count_list else ""

            # Формируем текст для копирования
            parts = [
                f"{customer} | {order_name}",
                f"{detail_name}",
                f"Осталось: {remaining} шт"
            ]

            # Добавляем материал (если есть)
            if material_info:
                parts.append(f"Материал: {material_info}")
            else:
                parts.append("Материал: ")

            # Добавляем остаток на складе (если есть)
            if material_stock:
                parts.append(f"Остаток на складе: {material_stock} шт")
            else:
                parts.append("Остаток на складе: ")

            # Добавляем количество к списанию (если есть)
            if remaining_reserved_count:
                parts.append(f"Остаток порезать: {remaining_reserved_count} шт")
            else:
                parts.append("Остаток порезать: ")

            copy_text = " | ".join(parts)

            # Копируем в буфер обмена
            self.root.clipboard_clear()
            self.root.clipboard_append(copy_text)
            self.root.update()  # Обновляем буфер обмена

            # Показываем уведомление
            messagebox.showinfo(
                "Скопировано",
                f"Информация о детали скопирована в буфер обмена:\n\n{copy_text}"
            )

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось скопировать информацию: {e}")
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
                                                       "Толщина", "Размер", "Резерв", "Списано", "Остаток", "Дата"),
                                              show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.config(command=self.reservations_tree.yview)
        scroll_x.config(command=self.reservations_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Настройка колонок БЕЗ растягивания
        for col in self.reservations_tree["columns"]:
            self.reservations_tree.heading(col, text=col)
            self.reservations_tree.column(col, anchor=tk.CENTER, width=100, minwidth=80, stretch=False)

        self.reservations_tree.pack(fill=tk.BOTH, expand=True)

        # 🆕 ИНИЦИАЛИЗАЦИЯ EXCEL-ФИЛЬТРА ДЛЯ РЕЗЕРВИРОВАНИЯ
        self.reservations_excel_filter = ExcelStyleFilter(
            tree=self.reservations_tree,
            refresh_callback=self.refresh_reservations
        )

        # 🆕 ИНДИКАТОР АКТИВНЫХ ФИЛЬТРОВ
        self.reservations_filter_status = tk.Label(
            self.reservations_frame,
            text="",
            font=("Arial", 9),
            bg='#d1ecf1',
            fg='#0c5460'
        )
        self.reservations_filter_status.pack(pady=5)

        # Переключатели видимости
        self.reservations_toggles = self.create_visibility_toggles(
            self.reservations_frame,
            self.reservations_tree,
            {
                'show_fully_written_off': '📝 Показать полностью списанные'
            },
            self.refresh_reservations
        )

        # Кнопки управления
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
        tk.Button(buttons_frame, text="✖ Сбросить фильтры", bg='#e67e22', fg='white',
                  command=self.clear_reservations_filters, **btn_style).pack(side=tk.LEFT, padx=5)

        self.refresh_reservations()

    def clear_reservations_filters(self):
        """Сбросить все фильтры резервирования"""
        if hasattr(self, 'reservations_excel_filter'):
            self.reservations_excel_filter.clear_all_filters()

    def refresh_reservations(self):
        """Обновление списка резервов"""

        # СОХРАНЯЕМ АКТИВНЫЕ ФИЛЬТРЫ ПЕРЕД ОЧИСТКОЙ
        active_filters_backup = {}
        if hasattr(self, 'reservations_excel_filter') and self.reservations_excel_filter.active_filters:
            active_filters_backup = self.reservations_excel_filter.active_filters.copy()

        # ПОЛНОСТЬЮ ОЧИЩАЕМ ДЕРЕВО
        for i in self.reservations_tree.get_children():
            self.reservations_tree.delete(i)

        # ОЧИЩАЕМ КЭШ ЭЛЕМЕНТОВ
        if hasattr(self, 'reservations_excel_filter'):
            self.reservations_excel_filter._all_item_cache = set()

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
                    order_display,
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

                item_id = self.reservations_tree.insert("", "end", values=values)

                # СОХРАНЯЕМ item_id В КЭШ
                if hasattr(self, 'reservations_excel_filter'):
                    if not hasattr(self.reservations_excel_filter, '_all_item_cache'):
                        self.reservations_excel_filter._all_item_cache = set()
                    self.reservations_excel_filter._all_item_cache.add(item_id)

        # ✅ АВТОПОДБОР ШИРИНЫ КОЛОНОК (ДОЛЖЕН БЫТЬ ЗДЕСЬ!)
        self.auto_resize_columns(self.reservations_tree, min_width=80, max_width=400)

        # ПЕРЕПРИМЕНЯЕМ ФИЛЬТРЫ ПОСЛЕ ЗАГРУЗКИ ДАННЫХ
        if active_filters_backup and hasattr(self, 'reservations_excel_filter'):
            self.reservations_excel_filter.active_filters = active_filters_backup
            self.reservations_excel_filter.reapply_all_filters()

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
                # 🆕 ПРАВИЛЬНЫЙ ПАРСИНГ: "ID:123 | Заказчик | Название"
                order_str = selected_order["value"]

                if order_str.startswith("ID:"):
                    order_id = int(order_str.split("ID:")[1].split(" | ")[0].strip())
                else:
                    # Старый формат для совместимости
                    order_id = int(order_str.split(" - ")[0])

                print(f"🔍 Загрузка деталей для заказа ID={order_id}")

                order_details_df = load_data("OrderDetails")

                if not order_details_df.empty:
                    details = order_details_df[order_details_df["ID заказа"] == order_id]

                    if not details.empty:
                        detail_options = ["[Без привязки к детали]"]
                        detail_options.extend([f"ID:{int(row['ID'])} - {row['Название детали']}"
                                               for _, row in details.iterrows()])
                        detail_combo['values'] = detail_options
                        detail_combo.current(0)
                        print(f"✅ Найдено деталей: {len(details)}")
                    else:
                        detail_combo['values'] = ["[Нет деталей у заказа]"]
                        detail_combo.current(0)
                        print(f"⚠️ У заказа ID={order_id} нет деталей")
                else:
                    detail_combo['values'] = ["[Нет деталей у заказа]"]
                    detail_combo.current(0)
                    print(f"⚠️ Таблица деталей пуста")
            except Exception as e:
                print(f"❌ Ошибка обновления списка деталей: {e}")
                import traceback
                traceback.print_exc()

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

                        # 🆕 ОБЪЕДИНЯЕМ ЗАКАЗЧИКА И НАЗВАНИЕ ЗАЯВКИ В ОДИН СТОЛБЕЦ
                        combined_order = f"{customer} | {order_name}"

                        # Добавляем строку
                        export_data.append({
                            "Заказ": combined_order,  # ← ОБЪЕДИНЁННЫЙ СТОЛБЕЦ
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

        try:
            values = self.writeoffs_tree.item(selected[0])['values']
            writeoff_id = int(values[0])
            reserve_id = int(values[1])
            comment = values[9] if len(values) > 9 else ""

            info_msg = (
                f"Отменить списание?\n\n"
                f"ID списания: {writeoff_id}\n"
                f"ID резерва: {reserve_id}\n\n"
                f"⚠️ Это действие:\n"
                f"• Вернёт материал в резерв\n"
                f"• Вернёт материал на склад\n"
                f"• Обновит таблицу импорта от лазерщиков"
            )

            if not messagebox.askyesno("Подтверждение", info_msg):
                return

            writeoffs_df = load_data("WriteOffs")
            reservations_df = load_data("Reservations")
            materials_df = load_data("Materials")
            order_details_df = load_data("OrderDetails")

            writeoff_row = writeoffs_df[writeoffs_df["ID списания"] == writeoff_id]

            if writeoff_row.empty:
                messagebox.showerror("Ошибка", f"Списание ID={writeoff_id} не найдено!")
                return

            writeoff_row = writeoff_row.iloc[0]

            reserve_id = int(writeoff_row["ID резерва"])
            quantity = int(writeoff_row["Количество"])
            material_id = int(writeoff_row["ID материала"])
            writeoff_date = str(writeoff_row["Дата списания"])
            writeoff_comment = str(writeoff_row["Комментарий"])

            # ОБНОВЛЕНИЕ РЕЗЕРВА
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

            # ОБНОВЛЕНИЕ МАТЕРИАЛА
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

                    area_per_piece = float(material["Длина"]) * float(material["Ширина"]) / 1_000_000
                    new_area = new_qty * area_per_piece
                    materials_df.loc[materials_df["ID"] == material_id, "Общая площадь"] = round(new_area, 2)

            # ОБНОВЛЕНИЕ ТАБЛИЦЫ ИМПОРТА
            is_laser_import = "Лазер:" in writeoff_comment or "лазерщик" in writeoff_comment.lower()

            if is_laser_import and hasattr(self, 'laser_table_data') and self.laser_table_data:
                import re

                part_name = None
                parts_qty = None

                part_match = re.search(r'Деталь:\s*([^|]+)', writeoff_comment)
                if part_match:
                    part_name = part_match.group(1).strip()

                date_match = re.search(r'Дата импорта:\s*(.+)', writeoff_comment)
                import_date_str = date_match.group(1).strip() if date_match else None

                for idx, row_data in enumerate(self.laser_table_data):
                    row_part = str(row_data.get("part", ""))

                    if part_name and (part_name.lower() in row_part.lower() or row_part.lower() in part_name.lower()):
                        row_date = str(row_data.get("Дата (МСК)", ""))
                        row_time = str(row_data.get("Время (МСК)", ""))
                        row_datetime = f"{row_date} {row_time}"

                        date_match_found = False
                        if import_date_str and len(row_datetime) >= 16 and len(import_date_str) >= 16:
                            if row_datetime[:16] == import_date_str[:16]:
                                date_match_found = True
                        elif not import_date_str:
                            row_writeoff_date = row_data.get("Дата списания", "")

                            # Безопасное преобразование
                            if pd.isna(row_writeoff_date) or row_writeoff_date is None:
                                row_writeoff_date = ""
                            else:
                                row_writeoff_date = str(row_writeoff_date)

                            if len(row_writeoff_date) >= 16 and len(writeoff_date) >= 16:
                                if row_writeoff_date[:16] == writeoff_date[:16]:
                                    date_match_found = True

                        if date_match_found:
                            try:
                                parts_qty = int(row_data.get("part_quantity", 0))
                                self.laser_table_data[idx]["Списано"] = ""
                                self.laser_table_data[idx]["Дата списания"] = ""
                            except:
                                pass
                            break

                if hasattr(self, 'laser_import_tree'):
                    self.refresh_laser_import_table()
                    try:
                        self.save_laser_import_cache()
                    except:
                        pass

            # ОТКАТ ДЕТАЛЕЙ
            if is_laser_import and "Деталь:" in writeoff_comment and parts_qty:
                try:
                    import re
                    part_match = re.search(r'Деталь:\s*([^|]+)', writeoff_comment)

                    if part_match and parts_qty > 0:
                        part_name = part_match.group(1).strip()
                        order_id = int(writeoff_row["ID заказа"])

                        detail_match = order_details_df[
                            (order_details_df["ID заказа"] == order_id) &
                            (order_details_df["Название детали"].str.contains(part_name, case=False, na=False))
                            ]

                        if not detail_match.empty:
                            detail_id = int(detail_match.iloc[0]["ID"])
                            old_cut = int(detail_match.iloc[0].get("Порезано", 0))
                            new_cut = max(0, old_cut - parts_qty)
                            order_details_df.loc[order_details_df["ID"] == detail_id, "Порезано"] = new_cut
                            save_data("OrderDetails", order_details_df)
                except:
                    pass

            # УДАЛЕНИЕ СПИСАНИЯ
            writeoffs_df = writeoffs_df[writeoffs_df["ID списания"] != writeoff_id]

            # СОХРАНЕНИЕ
            save_data("WriteOffs", writeoffs_df)
            save_data("Reservations", reservations_df)
            save_data("Materials", materials_df)

            # ОБНОВЛЕНИЕ ИНТЕРФЕЙСА
            self.refresh_writeoffs()
            self.refresh_reservations()
            self.refresh_materials()
            self.refresh_balance()

            if hasattr(self, 'refresh_orders'):
                self.refresh_orders()
            if hasattr(self, 'refresh_order_details'):
                self.refresh_order_details()

            messagebox.showinfo("Успех",
                                f"✅ Списание отменено!\n\n"
                                f"Возвращено в резерв: {quantity} шт\n"
                                f"Резерв ID: {reserve_id}\n"
                                f"Остаток к списанию: {new_remainder} шт")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось отменить списание:\n{e}")
            import traceback
            traceback.print_exc()

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

        print("✅ setup_laser_import_tab() выполнен успешно")

        # 🆕 АВТОЗАГРУЗКА КЭША ПРИ СТАРТЕ
        self.load_laser_import_cache()

    def setup_details_tab(self):
        """Вкладка учёта деталей"""

        # Заголовок
        header = tk.Label(self.details_frame, text="📐 Учёт деталей по заказам",
                          font=("Arial", 16, "bold"), bg='white', fg='#2c3e50')
        header.pack(pady=10)

        # Информационная панель
        info_frame = tk.LabelFrame(self.details_frame, text="ℹ️ Информация",
                                   bg='#d1ecf1', font=("Arial", 10, "bold"))
        info_frame.pack(fill=tk.X, padx=20, pady=10)

        info_text = """
    📊 Отображаются все детали из заказов со статусом "В работе"
    🟢 Зелёный - деталь полностью порезана
    🟡 Жёлтый - деталь в процессе
    ⚪ Белый - деталь не начата
        """

        tk.Label(info_frame, text=info_text, bg='#d1ecf1',
                 font=("Arial", 9), justify=tk.LEFT).pack(padx=10, pady=5)

        # Фрейм для таблицы
        tree_frame = tk.Frame(self.details_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # Scrollbars
        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        # Создание таблицы
        self.details_tree = ttk.Treeview(
            tree_frame,
            columns=("ID", "Заказчик", "Название детали", "Заказ", "Количество", "Порезано", "Погнуто", "Осталось",
                     "Прогресс %"),
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        scroll_y.config(command=self.details_tree.yview)
        scroll_x.config(command=self.details_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Настройка колонок
        columns_config = {
            "ID": 60,
            "Заказчик": 200,
            "Название детали": 300,
            "Заказ": 200,
            "Количество": 100,
            "Порезано": 100,
            "Погнуто": 100,
            "Осталось": 100,
            "Прогресс %": 100
        }

        for col, width in columns_config.items():
            self.details_tree.heading(col, text=col)
            self.details_tree.column(col, width=width, anchor=tk.CENTER)

        self.details_tree.pack(fill=tk.BOTH, expand=True)

        # Привязка правого клика для копирования информации о детали
        self.details_tree.bind('<Button-3>', self.on_details_tab_right_click)

        # Цветовые теги
        self.details_tree.tag_configure('completed', background='#c8e6c9', foreground='#1b5e20')  # Зелёный
        self.details_tree.tag_configure('in_progress', background='#fff9c4', foreground='#f57f17')  # Жёлтый
        self.details_tree.tag_configure('not_started', background='#ffffff', foreground='#000000')  # Белый
        self.details_tree.tag_configure('over_cut', background='#ffcccc',
                                        foreground='#b71c1c')  # Красный (если порезано больше)

        # Панель фильтрации
        self.details_filters = self.create_filter_panel(
            self.details_frame,
            self.details_tree,
            ["Заказчик", "Название детали", "Заказ"],
            self.refresh_details
        )

        # Переключатели видимости
        toggles_frame = tk.LabelFrame(self.details_frame, text="⚙️ Настройки отображения",
                                      bg='#ecf0f1', font=("Arial", 10, "bold"))
        toggles_frame.pack(fill=tk.X, padx=20, pady=10)

        self.details_toggles['show_completed'] = tk.BooleanVar(value=True)
        self.details_toggles['show_not_started'] = tk.BooleanVar(value=True)
        self.details_toggles['show_in_progress'] = tk.BooleanVar(value=True)

        tk.Checkbutton(toggles_frame, text="🟢 Показать завершённые",
                       variable=self.details_toggles['show_completed'],
                       command=self.refresh_details, bg='#ecf0f1', font=("Arial", 10),
                       activebackground='#ecf0f1').pack(side=tk.LEFT, padx=10)

        tk.Checkbutton(toggles_frame, text="🟡 Показать в процессе",
                       variable=self.details_toggles['show_in_progress'],
                       command=self.refresh_details, bg='#ecf0f1', font=("Arial", 10),
                       activebackground='#ecf0f1').pack(side=tk.LEFT, padx=10)

        tk.Checkbutton(toggles_frame, text="⚪ Показать не начатые",
                       variable=self.details_toggles['show_not_started'],
                       command=self.refresh_details, bg='#ecf0f1', font=("Arial", 10),
                       activebackground='#ecf0f1').pack(side=tk.LEFT, padx=10)

        # Кнопки управления
        buttons_frame = tk.Frame(self.details_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=20, pady=10)

        btn_style = {"font": ("Arial", 10, "bold"), "width": 18, "height": 2}

        tk.Button(buttons_frame, text="🔄 Обновить", bg='#3498db', fg='white',
                  command=self.refresh_details, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="📊 Экспорт в Excel", bg='#27ae60', fg='white',
                  command=self.export_details, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="📈 Статистика", bg='#9c27b0', fg='white',
                  command=self.show_details_statistics, **btn_style).pack(side=tk.LEFT, padx=5)

        # Статусная строка
        self.details_status_label = tk.Label(
            self.details_frame,
            text="",
            font=("Arial", 10),
            bg='#ecf0f1',
            fg='#2c3e50',
            relief=tk.SUNKEN,
            anchor='w',
            padx=10,
            pady=5
        )
        self.details_status_label.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=10)

        # Первоначальное заполнение
        self.refresh_details()

    def refresh_details(self):
        """Обновление таблицы деталей"""

        def safe_int(value, default=0):
            if value == "" or pd.isna(value) or value is None:
                return default
            try:
                return int(value)
            except (ValueError, TypeError):
                return default

        for item in self.details_tree.get_children():
            self.details_tree.delete(item)

        order_details_df = load_data("OrderDetails")
        orders_df = load_data("Orders")

        if order_details_df.empty:
            return

        # 🆕 ОЧИСТКА ПУСТЫХ ЗНАЧЕНИЙ
        order_details_df["Количество"] = order_details_df["Количество"].replace("", 0)
        order_details_df["Порезано"] = order_details_df["Порезано"].replace("", 0)
        order_details_df["Погнуто"] = order_details_df["Погнуто"].replace("", 0)

        # Сохраняем очищенные данные
        save_data("OrderDetails", order_details_df)

        """Обновление таблицы учёта деталей"""

        # Очищаем таблицу
        for item in self.details_tree.get_children():
            self.details_tree.delete(item)

        # Загружаем данные
        orders_df = load_data("Orders")
        order_details_df = load_data("OrderDetails")

        if orders_df.empty or order_details_df.empty:
            if hasattr(self, 'details_status_label'):
                self.details_status_label.config(
                    text="⚠️ Нет данных о деталях",
                    bg='#fff3cd',
                    fg='#856404'
                )
            return

        # Читаем переключатели
        show_completed = self.details_toggles['show_completed'].get()
        show_in_progress = self.details_toggles['show_in_progress'].get()
        show_not_started = self.details_toggles['show_not_started'].get()

        # Фильтруем заказы со статусом "В работе"
        active_orders = orders_df[orders_df["Статус"] == "В работе"]

        if active_orders.empty:
            if hasattr(self, 'details_status_label'):
                self.details_status_label.config(
                    text="ℹ️ Нет заказов в работе",
                    bg='#d1ecf1',
                    fg='#0c5460'
                )
            return

        # Счётчики
        total_count = 0
        shown_count = 0
        completed_count = 0
        in_progress_count = 0
        not_started_count = 0

        # Получаем активные фильтры
        active_filters = {}
        if hasattr(self, 'details_filters') and self.details_filters:
            for col_name, filter_var in self.details_filters.items():
                filter_text = filter_var.get().strip().lower()
                if filter_text:
                    active_filters[col_name] = filter_text

        # Проходим по всем деталям активных заказов
        for _, order_row in active_orders.iterrows():
            order_id = int(order_row["ID заказа"])
            order_name = order_row["Название заказа"]
            customer_name = order_row["Заказчик"]

            # Получаем детали этого заказа
            order_details = order_details_df[order_details_df["ID заказа"] == order_id]

            for _, detail_row in order_details.iterrows():
                detail_id = int(detail_row["ID"])
                detail_name = detail_row["Название детали"]
                quantity = int(detail_row["Количество"])
                cut = int(detail_row.get("Порезано", 0))
                bent = int(detail_row.get("Погнуто", 0))

                # Рассчитываем остаток и прогресс
                remaining = quantity - cut
                progress_pct = round((cut / quantity * 100), 1) if quantity > 0 else 0

                total_count += 1

                # Определяем статус
                if cut >= quantity:
                    status = 'completed'
                    completed_count += 1
                    if not show_completed:
                        continue
                elif cut > 0:
                    status = 'in_progress'
                    in_progress_count += 1
                    if not show_in_progress:
                        continue
                else:
                    status = 'not_started'
                    not_started_count += 1
                    if not show_not_started:
                        continue

                # Формируем значения
                values = (
                    detail_id,
                    customer_name,
                    detail_name,
                    order_name,
                    quantity,
                    cut,
                    bent,
                    remaining,
                    f"{progress_pct}%"
                )

                # Применяем фильтры
                if active_filters:
                    skip_row = False

                    if "Заказчик" in active_filters:
                        if active_filters["Заказчик"] not in customer_name.lower():
                            skip_row = True

                    if "Название детали" in active_filters:
                        if active_filters["Название детали"] not in detail_name.lower():
                            skip_row = True

                    if "Заказ" in active_filters:
                        if active_filters["Заказ"] not in order_name.lower():
                            skip_row = True

                    if skip_row:
                        continue

                # Цветовая индикация с учётом перепорезки
                if cut > quantity:
                    tag = 'over_cut'  # Порезано больше чем нужно
                else:
                    tag = status

                self.details_tree.insert("", "end", values=values, tags=(tag,))
                shown_count += 1

        self.auto_resize_columns(self.details_tree)

        # Обновляем статусную строку
        if hasattr(self, 'details_status_label'):
            status_text = (
                f"📊 Отображено: {shown_count} из {total_count} | "
                f"🟢 Завершено: {completed_count} | "
                f"🟡 В процессе: {in_progress_count} | "
                f"⚪ Не начато: {not_started_count}"
            )

            self.details_status_label.config(
                text=status_text,
                bg='#d4edda',
                fg='#155724'
            )

    def export_details(self):
        """Экспорт учёта деталей в Excel"""

        if not self.details_tree.get_children():
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта!")
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить учёт деталей",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not file_path:
            return

        try:
            # Собираем данные из таблицы
            data = []
            for item in self.details_tree.get_children():
                values = self.details_tree.item(item)['values']
                data.append(values)

            # Создаём DataFrame
            columns = ["ID", "Заказчик", "Название детали", "Заказ", "Количество", "Порезано", "Погнуто", "Осталось",
                       "Прогресс %"]
            df = pd.DataFrame(data, columns=columns)

            # Сохраняем в Excel
            df.to_excel(file_path, index=False, engine='openpyxl')

            messagebox.showinfo("Успех", f"Учёт деталей сохранён:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")

    def show_details_statistics(self):
        """Показать детальную статистику по деталям"""

        orders_df = load_data("Orders")
        order_details_df = load_data("OrderDetails")

        if orders_df.empty or order_details_df.empty:
            messagebox.showinfo("Статистика", "Нет данных для анализа")
            return

        # Фильтруем заказы в работе
        active_orders = orders_df[orders_df["Статус"] == "В работе"]

        if active_orders.empty:
            messagebox.showinfo("Статистика", "Нет заказов в работе")
            return

        # Собираем статистику
        total_details = 0
        total_qty = 0
        total_cut = 0
        total_bent = 0

        completed_details = 0
        in_progress_details = 0
        not_started_details = 0

        by_customer = {}

        for _, order_row in active_orders.iterrows():
            order_id = int(order_row["ID заказа"])
            customer = order_row["Заказчик"]

            order_details = order_details_df[order_details_df["ID заказа"] == order_id]

            for _, detail_row in order_details.iterrows():
                qty = int(detail_row["Количество"])
                cut = int(detail_row.get("Порезано", 0))
                bent = int(detail_row.get("Погнуто", 0))

                total_details += 1
                total_qty += qty
                total_cut += cut
                total_bent += bent

                # Определяем статус
                if cut >= qty:
                    completed_details += 1
                elif cut > 0:
                    in_progress_details += 1
                else:
                    not_started_details += 1

                # Группируем по заказчику
                if customer not in by_customer:
                    by_customer[customer] = {
                        'details': 0,
                        'quantity': 0,
                        'cut': 0,
                        'bent': 0
                    }

                by_customer[customer]['details'] += 1
                by_customer[customer]['quantity'] += qty
                by_customer[customer]['cut'] += cut
                by_customer[customer]['bent'] += bent

        # Общий прогресс
        overall_progress = round((total_cut / total_qty * 100), 1) if total_qty > 0 else 0

        # Формируем сообщение
        stats_msg = (
            f"📊 ОБЩАЯ СТАТИСТИКА\n"
            f"{'=' * 50}\n\n"
            f"📐 Уникальных деталей: {total_details}\n"
            f"📦 Всего требуется порезать: {total_qty} шт\n"
            f"✂️ Порезано: {total_cut} шт ({overall_progress}%)\n"
            f"🔧 Погнуто: {total_bent} шт\n"
            f"⏳ Осталось порезать: {total_qty - total_cut} шт\n\n"
            f"{'=' * 50}\n\n"
            f"📈 ПО СТАТУСАМ:\n\n"
            f"🟢 Завершено: {completed_details} деталей\n"
            f"🟡 В процессе: {in_progress_details} деталей\n"
            f"⚪ Не начато: {not_started_details} деталей\n\n"
            f"{'=' * 50}\n\n"
            f"👥 ПО ЗАКАЗЧИКАМ:\n\n"
        )

        # Сортируем заказчиков по количеству деталей
        sorted_customers = sorted(by_customer.items(), key=lambda x: x[1]['quantity'], reverse=True)

        for customer, stats in sorted_customers[:10]:  # Показываем топ-10
            customer_progress = round((stats['cut'] / stats['quantity'] * 100), 1) if stats['quantity'] > 0 else 0
            stats_msg += (
                f"\n{customer}:\n"
                f"  Деталей: {stats['details']}\n"
                f"  Требуется: {stats['quantity']} шт\n"
                f"  Порезано: {stats['cut']} шт ({customer_progress}%)\n"
                f"  Погнуто: {stats['bent']} шт\n"
            )

        if len(by_customer) > 10:
            stats_msg += f"\n... и еще {len(by_customer) - 10} заказчиков"

        # Создаём окно со статистикой
        stats_window = tk.Toplevel(self.root)
        stats_window.title("📊 Статистика по деталям")
        stats_window.geometry("600x700")
        stats_window.configure(bg='#f0f0f0')

        # Текстовое поле со скроллом
        text_frame = tk.Frame(stats_window, bg='white')
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scroll = tk.Scrollbar(text_frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        text_widget = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 10),
                              yscrollcommand=scroll.set)
        text_widget.insert("1.0", stats_msg)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(fill=tk.BOTH, expand=True)

        scroll.config(command=text_widget.yview)

        # Кнопка закрытия
        tk.Button(stats_window, text="Закрыть", command=stats_window.destroy,
                  bg='#3498db', fg='white', font=("Arial", 12, "bold"),
                  width=20, height=2).pack(pady=10)

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

        if not hasattr(self, 'laser_table_data') or self.laser_table_data is None:
            self.laser_table_data = []
            return

        if not self.laser_table_data:
            return

        # 🔥 СОРТИРОВКА С ПРАВИЛЬНЫМ ФОРМАТОМ ДАТЫ
        try:
            print(f"🔄 Сортировка {len(self.laser_table_data)} записей перед отображением...")

            # Преобразуем в DataFrame
            df_display = pd.DataFrame(self.laser_table_data)

            # 🆕 ПРАВИЛЬНЫЙ ПАРСИНГ ДАТЫ: ФОРМАТ DD.MM.YYYY
            df_display['_datetime_sort'] = pd.to_datetime(
                df_display['Дата (МСК)'].astype(str) + ' ' + df_display['Время (МСК)'].astype(str),
                format='%d.%m.%Y %H:%M:%S',  # ← ЯВНО УКАЗЫВАЕМ ФОРМАТ
                errors='coerce'
            )

            # Сортируем по убыванию (новые сверху)
            df_display = df_display.sort_values('_datetime_sort', ascending=False, na_position='last')

            # Удаляем временную колонку
            df_display = df_display.drop('_datetime_sort', axis=1)

            # Преобразуем обратно в список словарей
            sorted_data = df_display.to_dict('records')

            # Показываем первую и последнюю запись
            if sorted_data:
                first = f"{sorted_data[0].get('Дата (МСК)', '')} {sorted_data[0].get('Время (МСК)', '')}"
                last = f"{sorted_data[-1].get('Дата (МСК)', '')} {sorted_data[-1].get('Время (МСК)', '')}"
                print(f"✅ Отсортировано: ПЕРВАЯ (новая) = {first}, ПОСЛЕДНЯЯ (старая) = {last}")
        except Exception as e:
            print(f"⚠️ Ошибка сортировки (формат DD.MM.YYYY): {e}")

            # 🆕 ПОПРОБУЕМ АЛЬТЕРНАТИВНЫЙ ФОРМАТ
            try:
                print("🔄 Пробуем альтернативный формат YYYY-MM-DD...")
                df_display = pd.DataFrame(self.laser_table_data)
                df_display['_datetime_sort'] = pd.to_datetime(
                    df_display['Дата (МСК)'].astype(str) + ' ' + df_display['Время (МСК)'].astype(str),
                    errors='coerce'
                )
                df_display = df_display.sort_values('_datetime_sort', ascending=False, na_position='last')
                df_display = df_display.drop('_datetime_sort', axis=1)
                sorted_data = df_display.to_dict('records')
                print("✅ Альтернативный формат сработал!")
            except Exception as e2:
                print(f"⚠️ И альтернативный формат не сработал: {e2}")
                import traceback
                traceback.print_exc()
                sorted_data = self.laser_table_data

        # СЧЁТЧИКИ
        manual_count = 0
        auto_count = 0
        pending_count = 0

        # Заполняем таблицу ОТСОРТИРОВАННЫМИ данными
        for idx, row_data in enumerate(sorted_data):
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

            # БЕЗОПАСНОЕ ПРЕОБРАЗОВАНИЕ written_off В СТРОКУ
            if pd.isna(written_off) or written_off is None:
                written_off = ""
            else:
                written_off = str(written_off).strip()

            values = (date_val, time_val, username, order, metal, metal_qty, part, part_qty, written_off, writeoff_date)

            # ЦВЕТОВАЯ ИНДИКАЦИЯ
            if written_off == "Вручную":
                tag = 'manual'
                manual_count += 1
            elif written_off in ["Да", "✓", "Yes"]:
                tag = 'written_off'
                auto_count += 1
            else:
                tag = 'pending'
                pending_count += 1

            self.laser_import_tree.insert("", "end", values=values, tags=(tag,))

        self.auto_resize_columns(self.laser_import_tree)

        print(f"📊 Отображено: 🔵 Синих={manual_count}, 🟢 Зелёных={auto_count}, 🟡 Жёлтых={pending_count}")

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
            import re
            thickness = None
            width = None
            length = None

            print(f"   🔍 Парсинг материала: '{metal_description}'")
            print(f"   📏 Длина строки: {len(metal_description)} символов")
            print(f"   🔤 Побайтово: {metal_description.encode('utf-8')}")

            # Пробуем разные паттерны
            patterns = [
                (r'(\d+(?:\.\d+)?)\s*мм\s*(\d+(?:\.\d+)?)\s*[xXхХ×]\s*(\d+(?:\.\d+)?)', "Формат: 4.0мм 1500x3000"),
                (r'(\d+(?:\.\d+)?)\s*[xXхХ×]\s*(\d+(?:\.\d+)?)\s*[xXхХ×]\s*(\d+(?:\.\d+)?)', "Формат: 4x1500x3000"),
                (r'(\d+(?:\.\d+)?)\s*мм?\s*(\d+(?:\.\d+)?)\s*[xXхХ×]?\s*(\d+(?:\.\d+)?)', "Формат гибкий"),
            ]

            for idx, (pattern, description) in enumerate(patterns, 1):
                print(f"   🧪 Тест паттерна {idx}: {description}")
                match = re.search(pattern, metal_description, re.IGNORECASE)

                if match:
                    thickness = float(match.group(1))
                    width = float(match.group(2))
                    length = float(match.group(3))
                    print(f"   ✅ УСПЕХ! Найдено: {thickness} × {width} × {length}")
                    break
                else:
                    print(f"   ❌ Не подошёл")

            if not thickness:
                # Попробуем найти хоть какие-то числа
                all_numbers = re.findall(r'\d+(?:\.\d+)?', metal_description)
                print(f"   🔢 Все числа в строке: {all_numbers}")

                # Если нашли минимум 3 числа - попробуем взять последние 3
                if len(all_numbers) >= 3:
                    try:
                        # Ищем первое число как толщину (обычно 3-10)
                        for i, num in enumerate(all_numbers):
                            val = float(num)
                            if 0.5 <= val <= 50:  # Толщина обычно от 0.5 до 50 мм
                                thickness = val
                                # Берём следующие 2 числа как размеры
                                if i + 2 < len(all_numbers):
                                    width = float(all_numbers[i + 1])
                                    length = float(all_numbers[i + 2])
                                    print(f"   ⚠️ Использована эвристика: {thickness} × {width} × {length}")
                                    break
                    except:
                        pass

            if not thickness:
                print(f"   ❌ НЕ РАСПОЗНАНО: '{metal_description}'")
                row_data["_status"] = f"Ошибка парсинга материала"
                return False

            print(f"   ✅ ИТОГО: толщина={thickness}, ширина={width}, длина={length}")

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

            # Создаём уникальный ключ для каждой строки
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

            # Создаём словарь существующих строк с их статусами
            existing_rows = {}
            if hasattr(self, 'laser_table_data') and self.laser_table_data:
                for row_data in self.laser_table_data:
                    key = create_row_key(row_data)
                    existing_rows[key] = {
                        "Списано": row_data.get("Списано", ""),
                        "Дата списания": row_data.get("Дата списания", "")
                    }

            # Обрабатываем новый файл
            new_rows = []
            updated_rows = 0

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

            # Объединяем данные
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
                        # Строка есть в новом файле - берём из старой таблицы
                        merged_data.append(old_row)

            # Добавляем только НОВЫЕ строки из импортированного файла
            for new_row in new_rows:
                new_key = create_row_key(new_row)
                is_new = new_key not in existing_rows

                if is_new:
                    merged_data.append(new_row)
                    new_count += 1

            # Сохраняем объединённые данные
            self.laser_table_data = merged_data

            # 🆕 СОРТИРОВКА: НОВЫЕ ЗАПИСИ ВВЕРХУ
            try:
                print("🔄 Сортировка импортированных данных...")
                df_merged = pd.DataFrame(self.laser_table_data)

                df_merged['_datetime_sort'] = pd.to_datetime(
                    df_merged['Дата (МСК)'].astype(str) + ' ' + df_merged['Время (МСК)'].astype(str),
                    errors='coerce'
                )
                df_merged = df_merged.sort_values('_datetime_sort', ascending=False, na_position='last')
                df_merged = df_merged.drop('_datetime_sort', axis=1)

                self.laser_table_data = df_merged.to_dict('records')
                print(f"✅ Данные отсортированы: {len(self.laser_table_data)} записей")
            except Exception as e:
                print(f"⚠️ Ошибка сортировки после импорта: {e}")

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
                    f"  • 🟡 Ожидает списания: {pending_count}\n"
                )

            # Сохраняем в кэш
            try:
                self.save_laser_import_cache()
            except Exception as cache_err:
                print(f"⚠️ Не удалось сохранить кэш: {cache_err}")

            messagebox.showinfo("Успех", result_msg)

        except Exception as e:
            messagebox.showerror("Ошибка импорта", f"Не удалось импортировать файл:\n\n{str(e)}")
            print(f"❌ Ошибка импорта: {e}")
            import traceback
            traceback.print_exc()

    def refresh_laser_import_table(self):
        """Обновление таблицы импорта от лазерщиков"""
        # Очищаем таблицу
        for item in self.laser_import_tree.get_children():
            self.laser_import_tree.delete(item)

        if not hasattr(self, 'laser_table_data') or self.laser_table_data is None:
            self.laser_table_data = []
            return

        if not self.laser_table_data:
            return

        # СОРТИРОВКА: НОВЫЕ ЗАПИСИ ВВЕРХУ
        try:
            # Преобразуем в DataFrame
            df_display = pd.DataFrame(self.laser_table_data)

            # Парсинг даты в формате DD.MM.YYYY HH:MM:SS
            df_display['_datetime_sort'] = pd.to_datetime(
                df_display['Дата (МСК)'].astype(str) + ' ' + df_display['Время (МСК)'].astype(str),
                format='%d.%m.%Y %H:%M:%S',
                errors='coerce'
            )

            # Сортируем по убыванию (новые вверху)
            df_display = df_display.sort_values('_datetime_sort', ascending=False, na_position='last')

            # Удаляем временную колонку
            df_display = df_display.drop('_datetime_sort', axis=1)

            # Преобразуем обратно в список словарей
            sorted_data = df_display.to_dict('records')

        except Exception as e:
            print(f"⚠️ Ошибка сортировки: {e}")
            sorted_data = self.laser_table_data

        # 🆕 СОХРАНЯЕМ ОТСОРТИРОВАННЫЕ ДАННЫЕ ОБРАТНО
        self.laser_table_data = sorted_data

        # СЧЁТЧИКИ
        manual_count = 0
        auto_count = 0
        pending_count = 0

        # Заполняем таблицу ОТСОРТИРОВАННЫМИ данными
        for idx, row_data in enumerate(sorted_data):
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

            # БЕЗОПАСНОЕ ПРЕОБРАЗОВАНИЕ written_off В СТРОКУ
            if pd.isna(written_off) or written_off is None:
                written_off = ""
            else:
                written_off = str(written_off).strip()

            values = (date_val, time_val, username, order, metal, metal_qty, part, part_qty, written_off, writeoff_date)

            # ЦВЕТОВАЯ ИНДИКАЦИЯ
            if written_off == "Вручную":
                tag = 'manual'
                manual_count += 1
            elif written_off in ["Да", "✓", "Yes"]:
                tag = 'written_off'
                auto_count += 1
            else:
                tag = 'pending'
                pending_count += 1

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
                    # Примеры:
                    # "ГК Ст.3 4.0мм 1500x3000" → марка="ГК Ст.3", толщина=4.0, ширина=1500, длина=3000
                    # "ГК Ст.3 6х1500х3000" → марка="ГК Ст.3", толщина=6, ширина=1500, длина=3000

                    thickness = None
                    width = None
                    length = None
                    marka = None

                    print(f"   🔍 Парсинг материала: '{metal_desc}'")

                    # 🆕 ПАТТЕРН 1: Формат с "мм" (4.0мм 1500x3000)
                    pattern1 = r'(\d+(?:\.\d+)?)\s*мм\s*(\d+(?:\.\d+)?)\s*[xXхХ×]\s*(\d+(?:\.\d+)?)'
                    match1 = re.search(pattern1, metal_desc, re.IGNORECASE)

                    if match1:
                        thickness = float(match1.group(1))
                        width = float(match1.group(2))
                        length = float(match1.group(3))
                        # Марка - всё до размеров
                        marka = metal_desc.split(match1.group(0))[0].strip()
                        print(f"   ✅ Паттерн 1 (с мм): {thickness}мм {width}x{length}, марка='{marka}'")

                    # ПАТТЕРН 2: Классический формат (6х1500х3000)
                    if not thickness:
                        pattern2 = r'(\d+(?:\.\d+)?)\s*[xXхХ×]\s*(\d+(?:\.\d+)?)\s*[xXхХ×]\s*(\d+(?:\.\d+)?)'
                        match2 = re.search(pattern2, metal_desc)

                        if match2:
                            thickness = float(match2.group(1))
                            width = float(match2.group(2))
                            length = float(match2.group(3))
                            # Марка - всё до размеров
                            marka = metal_desc.split(match2.group(0))[0].strip()
                            print(f"   ✅ Паттерн 2 (без мм): {thickness}х{width}х{length}, марка='{marka}'")

                    if not thickness or not marka:
                        errors.append(f"❌ Не удалось распарсить материал: {metal_desc}")
                        print(f"   ❌ Ошибка парсинга материала: '{metal_desc}'")
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

            # ОБНОВЛЕНИЕ ИНТЕРФЕЙСА
            self.refresh_laser_import_table()
            self.refresh_materials()
            self.refresh_reservations()
            self.refresh_writeoffs()
            self.refresh_balance()

            if hasattr(self, 'refresh_orders'):
                self.refresh_orders()
            if hasattr(self, 'refresh_order_details'):
                self.refresh_order_details()

            # 🆕 ОБНОВЛЯЕМ ВКЛАДКУ УЧЁТА ДЕТАЛЕЙ
            if hasattr(self, 'refresh_details'):
                self.refresh_details()


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

            # 🆕 АВТОСОХРАНЕНИЕ ПОСЛЕ СПИСАНИЯ
            self.save_laser_import_cache()

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

        # 🆕 АВТОСОХРАНЕНИЕ ПОСЛЕ УДАЛЕНИЯ
        self.save_laser_import_cache()

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

    # 🆕 НОВЫЙ МЕТОД - СОХРАНЕНИЕ КЭША
    def save_laser_import_cache(self):
        """Автоматическое сохранение таблицы импорта в кэш-файл"""
        if not hasattr(self, 'laser_table_data') or not self.laser_table_data:
            return

        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            cache_file = os.path.join(script_dir, "laser_import_cache.xlsx")

            print(f"💾 Сохранение {len(self.laser_table_data)} записей в кэш...")

            df = pd.DataFrame(self.laser_table_data)
            df.to_excel(cache_file, index=False, engine='openpyxl')

            print(f"✅ Кэш импорта сохранён: {len(self.laser_table_data)} записей")
        except Exception as e:
            print(f"⚠️ Ошибка сохранения кэша: {e}")

    # 🆕 НОВЫЙ МЕТОД - ЗАГРУЗКА КЭША
    def load_laser_import_cache(self):
        """Автоматическая загрузка таблицы импорта из кэш-файла"""
        try:
            # Используем текущую директорию скрипта
            script_dir = os.path.dirname(os.path.abspath(__file__))
            cache_file = os.path.join(script_dir, "laser_import_cache.xlsx")

            if not os.path.exists(cache_file):
                print(f"ℹ️ Кэш импорта не найден: {cache_file}")
                return

            # Загружаем из Excel
            df = pd.read_excel(cache_file, engine='openpyxl')

            if df.empty:
                print("ℹ️ Кэш импорта пуст")
                return

            # Проверяем наличие необходимых колонок
            required = ["Дата (МСК)", "Время (МСК)", "username", "order", "metal", "metal_quantity", "part",
                        "part_quantity"]

            if all(col in df.columns for col in required):
                # 🆕 ВАЖНО: Преобразуем NaN в пустые строки перед конвертацией
                df = df.fillna("")

                def load_laser_import_cache(self):
                    """Автоматическая загрузка таблицы импорта из кэш-файла"""
                    try:
                        script_dir = os.path.dirname(os.path.abspath(__file__))
                        cache_file = os.path.join(script_dir, "laser_import_cache.xlsx")

                        if not os.path.exists(cache_file):
                            print(f"ℹ️ Кэш импорта не найден: {cache_file}")
                            return

                        df = pd.read_excel(cache_file, engine='openpyxl')

                        if df.empty:
                            print("ℹ️ Кэш импорта пуст")
                            return

                        required = ["Дата (МСК)", "Время (МСК)", "username", "order", "metal", "metal_quantity", "part",
                                    "part_quantity"]

                        if all(col in df.columns for col in required):
                            df = df.fillna("")

                            # 🆕 СОРТИРОВКА: НОВЫЕ ЗАПИСИ ВВЕРХУ
                            try:
                                print("🔄 Сортировка кэша...")
                                df['_datetime_sort'] = pd.to_datetime(
                                    df['Дата (МСК)'].astype(str) + ' ' + df['Время (МСК)'].astype(str),
                                    errors='coerce'
                                )
                                df = df.sort_values('_datetime_sort', ascending=False, na_position='last')
                                df = df.drop('_datetime_sort', axis=1)

                                if not df.empty:
                                    first = f"{df.iloc[0]['Дата (МСК)']} {df.iloc[0]['Время (МСК)']}"
                                    last = f"{df.iloc[-1]['Дата (МСК)']} {df.iloc[-1]['Время (МСК)']}"
                                    print(f"✅ Отсортировано: первая={first}, последняя={last}")
                            except Exception as e:
                                print(f"⚠️ Ошибка сортировки: {e}")

                            # Преобразуем в список словарей
                            self.laser_table_data = df.to_dict('records')

                            # Дополнительная очистка
                            for row in self.laser_table_data:
                                if "Списано" in row:
                                    if pd.isna(row["Списано"]) or row["Списано"] is None:
                                        row["Списано"] = ""
                                    else:
                                        row["Списано"] = str(row["Списано"]).strip()

                                if "Дата списания" in row:
                                    if pd.isna(row["Дата списания"]) or row["Дата списания"] is None:
                                        row["Дата списания"] = ""
                                    else:
                                        row["Дата списания"] = str(row["Дата списания"]).strip()

                            print(f"✅ Загружен кэш импорта: {len(self.laser_table_data)} записей из {cache_file}")

                            # Обновляем таблицу
                            if hasattr(self, 'laser_import_tree'):
                                self.refresh_laser_import_table()

                                if hasattr(self, 'laser_status_label'):
                                    items_count = len(self.laser_import_tree.get_children())
                                    auto_count = sum(1 for r in self.laser_table_data if
                                                     r.get("Списано", "").strip() in ["✓", "Да", "Yes"])
                                    manual_count = sum(
                                        1 for r in self.laser_table_data if r.get("Списано", "").strip() == "Вручную")
                                    pending_count = sum(
                                        1 for r in self.laser_table_data if not r.get("Списано", "").strip())

                                    status_text = (
                                        f"📂 Загружено из кэша: {items_count} | "
                                        f"✅ Списано: {auto_count} | "
                                        f"🔵 Вручную: {manual_count} | "
                                        f"🟡 Ожидает: {pending_count}"
                                    )
                                    self.laser_status_label.config(text=status_text, bg='#d1ecf1', fg='#0c5460')

                    except Exception as e:
                        pass
                # Преобразуем в список словарей
                self.laser_table_data = df.to_dict('records')

                # 🆕 ДОПОЛНИТЕЛЬНАЯ ОЧИСТКА: убедимся что все значения - строки или числа
                for row in self.laser_table_data:
                    # Преобразуем "Списано" в строку
                    if "Списано" in row:
                        if pd.isna(row["Списано"]) or row["Списано"] is None:
                            row["Списано"] = ""
                        else:
                            row["Списано"] = str(row["Списано"]).strip()

                    # Преобразуем "Дата списания" в строку
                    if "Дата списания" in row:
                        if pd.isna(row["Дата списания"]) or row["Дата списания"] is None:
                            row["Дата списания"] = ""
                        else:
                            row["Дата списания"] = str(row["Дата списания"]).strip()

                print(f"✅ Загружен кэш импорта: {len(self.laser_table_data)} записей из {cache_file}")

                # 🆕 ДИАГНОСТИКА: выведем первые строки для проверки
                if self.laser_table_data:
                    print("\n🔍 Проверка загруженных данных:")
                    for i, row in enumerate(self.laser_table_data[:3]):
                        status = row.get("Списано", "")
                        print(f"   Строка {i + 1}: Списано = '{status}' (тип: {type(status).__name__})")

                # Обновляем таблицу
                if hasattr(self, 'laser_import_tree'):
                    self.refresh_laser_import_table()

                    # Обновляем статус
                    if hasattr(self, 'laser_status_label'):
                        items_count = len(self.laser_import_tree.get_children())

                        # Считаем статистику
                        auto_count = sum(
                            1 for r in self.laser_table_data if r.get("Списано", "").strip() in ["✓", "Да", "Yes"])
                        manual_count = sum(
                            1 for r in self.laser_table_data if r.get("Списано", "").strip() == "Вручную")
                        pending_count = sum(1 for r in self.laser_table_data if not r.get("Списано", "").strip())

                        status_text = (
                            f"📂 Загружено из кэша: {items_count} | "
                            f"✅ Списано: {auto_count} | "
                            f"🔵 Вручную: {manual_count} | "
                            f"🟡 Ожидает: {pending_count}"
                        )

                        self.laser_status_label.config(
                            text=status_text,
                            bg='#d1ecf1',
                            fg='#0c5460'
                        )
            else:
                print("⚠️ Кэш импорта имеет неправильную структуру")

        except Exception as e:
            print(f"⚠️ Ошибка загрузки кэша: {e}")
            import traceback
            traceback.print_exc()

    def setup_balance_tab(self):
        """Вкладка баланса материалов"""

        # Заголовок
        header = tk.Label(self.balance_frame, text="Баланс материалов по маркам и толщинам",
                          font=("Arial", 16, "bold"), bg='white', fg='#2c3e50')
        header.pack(pady=10)

        # 🆕 Фрейм таблицы НА ВСЕЙ ШИРИНЕ (убрано центрирование)
        tree_frame = tk.Frame(self.balance_frame, bg='white')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scroll_y = tk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)

        # ТАБЛИЦА
        self.balance_tree = ttk.Treeview(tree_frame,
                                         columns=("Марка", "Толщина", "Размер", "Всего", "Зарезервировано", "Доступно"),
                                         show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.config(command=self.balance_tree.yview)
        scroll_x.config(command=self.balance_tree.xview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # НАСТРОЙКА КОЛОНОК БЕЗ РАСТЯГИВАНИЯ
        for col in self.balance_tree["columns"]:
            self.balance_tree.heading(col, text=col)
            self.balance_tree.column(col, anchor=tk.CENTER, width=100, minwidth=80, stretch=False)

        # ТАБЛИЦА ЗАПОЛНЯЕТ ВСЮ ВЫСОТУ И ШИРИНУ ФРЕЙМА
        self.balance_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ИНИЦИАЛИЗАЦИЯ ФИЛЬТРА В СТИЛЕ EXCEL
        self.balance_excel_filter = ExcelStyleFilter(
            tree=self.balance_tree,
            refresh_callback=self.refresh_balance
        )

        # ЦВЕТОВАЯ ИНДИКАЦИЯ
        self.balance_tree.tag_configure('negative', background='#f8d7da', foreground='#721c24')
        self.balance_tree.tag_configure('available', background='#d4edda', foreground='#155724')
        self.balance_tree.tag_configure('fully_reserved', background='#fff3cd', foreground='#856404')
        self.balance_tree.tag_configure('empty', background='#d1ecf1', foreground='#0c5460')

        # ИНДИКАТОР АКТИВНЫХ ФИЛЬТРОВ
        self.balance_filter_status = tk.Label(
            self.balance_frame,
            text="",
            font=("Arial", 9),
            bg='#d1ecf1',
            fg='#0c5460'
        )
        self.balance_filter_status.pack(pady=5)

        # ЛЕГЕНДА ЦВЕТОВ
        legend_frame = tk.Frame(self.balance_frame, bg='white')
        legend_frame.pack(pady=5)

        tk.Label(legend_frame, text="Легенда:", font=("Arial", 10, "bold"), bg='white').pack(side=tk.LEFT, padx=5)

        # Отрицательное значение (проблема)
        negative_label = tk.Label(legend_frame, text="  Отрицательное  ", bg='#f8d7da', fg='#721c24',
                                  font=("Arial", 9, "bold"), relief=tk.RAISED, borderwidth=1)
        negative_label.pack(side=tk.LEFT, padx=3)

        # Полностью зарезервировано
        reserved_label = tk.Label(legend_frame, text="  Полностью зарезервировано  ", bg='#fff3cd', fg='#856404',
                                  font=("Arial", 9, "bold"), relief=tk.RAISED, borderwidth=1)
        reserved_label.pack(side=tk.LEFT, padx=3)

        # Есть доступно
        available_label = tk.Label(legend_frame, text="  Есть доступно  ", bg='#d4edda', fg='#155724',
                                   font=("Arial", 9, "bold"), relief=tk.RAISED, borderwidth=1)
        available_label.pack(side=tk.LEFT, padx=3)

        # Нет в наличии
        empty_label = tk.Label(legend_frame, text="  Нет в наличии  ", bg='#d1ecf1', fg='#0c5460',
                               font=("Arial", 9, "bold"), relief=tk.RAISED, borderwidth=1)
        empty_label.pack(side=tk.LEFT, padx=3)

        # Кнопки управления
        buttons_frame = tk.Frame(self.balance_frame, bg='white')
        buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(buttons_frame, text="🔄 Обновить", bg='#3498db', fg='white',
                  font=("Arial", 10), command=self.refresh_balance).pack(side=tk.LEFT, padx=5)

        tk.Button(buttons_frame, text="✖ Сбросить все фильтры", bg='#e67e22', fg='white',
                  font=("Arial", 10), command=self.clear_balance_filters).pack(side=tk.LEFT, padx=5)

        # Первичная загрузка данных
        self.refresh_balance()

    def clear_balance_filters(self):
        """Сбросить все фильтры баланса"""
        if hasattr(self, 'balance_excel_filter'):
            self.balance_excel_filter.clear_all_filters()

    def refresh_balance(self):
        """Обновление таблицы баланса материалов"""

        # СОХРАНЯЕМ АКТИВНЫЕ ФИЛЬТРЫ ПЕРЕД ОЧИСТКОЙ
        active_filters_backup = {}
        if hasattr(self, 'balance_excel_filter') and self.balance_excel_filter.active_filters:
            active_filters_backup = self.balance_excel_filter.active_filters.copy()
            print(f"🔍 Сохранены фильтры: {list(active_filters_backup.keys())}")

        # ПОЛНОСТЬЮ ОЧИЩАЕМ ДЕРЕВО
        for i in self.balance_tree.get_children():
            self.balance_tree.delete(i)

        # ОЧИЩАЕМ КЭШ ЭЛЕМЕНТОВ
        if hasattr(self, 'balance_excel_filter'):
            self.balance_excel_filter._all_item_cache = set()

        df = load_data("Materials")

        if not df.empty:
            # Группируем по марке, толщине и размеру
            balance_data = {}

            for index, row in df.iterrows():
                marka = row["Марка"]
                thickness = row["Толщина"]
                length = row["Длина"]
                width = row["Ширина"]
                size_key = f"{int(length)}x{int(width)}"

                key = (marka, thickness, size_key)

                if key not in balance_data:
                    balance_data[key] = {
                        "total": 0,
                        "reserved": 0,
                        "available": 0
                    }

                balance_data[key]["total"] += int(row["Количество штук"])
                balance_data[key]["reserved"] += int(row["Зарезервировано"])
                balance_data[key]["available"] += int(row["Доступно"])

            # Заполняем таблицу с цветовой индикацией
            for (marka, thickness, size), data in sorted(balance_data.items()):
                total = data["total"]
                reserved = data["reserved"]
                available = data["available"]

                # ОБЫЧНЫЕ VALUES БЕЗ ПУСТОЙ КОЛОНКИ
                values = (marka, thickness, size, total, reserved, available)

                # ЦВЕТОВАЯ ИНДИКАЦИЯ
                if available < 0:
                    tag = 'negative'
                elif available > 0:
                    tag = 'available'
                elif available == 0 and total > 0:
                    tag = 'fully_reserved'
                else:
                    tag = 'empty'

                item_id = self.balance_tree.insert("", "end", values=values, tags=(tag,))

                # СОХРАНЯЕМ item_id В КЭШ
                if hasattr(self, 'balance_excel_filter'):
                    if not hasattr(self.balance_excel_filter, '_all_item_cache'):
                        self.balance_excel_filter._all_item_cache = set()
                    self.balance_excel_filter._all_item_cache.add(item_id)

        # АВТОПОДБОР ШИРИНЫ КОЛОНОК С ОГРАНИЧЕНИЯМИ
        self.auto_resize_columns(self.balance_tree, min_width=100, max_width=300)

        # ПЕРЕПРИМЕНЯЕМ ФИЛЬТРЫ ПОСЛЕ ЗАГРУЗКИ ДАННЫХ
        if active_filters_backup and hasattr(self, 'balance_excel_filter'):
            print(f"🔄 Переприменяю фильтры: {list(active_filters_backup.keys())}")
            self.balance_excel_filter.active_filters = active_filters_backup
            self.balance_excel_filter.reapply_all_filters()

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