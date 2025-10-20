import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
import scanner
import excel_exporter
import logging
import time
import tkinterdnd2
import sys
import os


class ArchiveScannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Сканер архивов v0.6")
        self.center_window(900, 700)

        self.scanner = scanner.ArchiveScanner()
        self.exporter = excel_exporter.ExcelExporter()
        self.scan_thread = None
        self.scan_results = None
        self.is_scanning = False

        self.setup_ui()
        self.setup_logging()

        try:
            self.root.drop_target_register(tkinterdnd2.DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop)
        except:
            pass

    def center_window(self, width, height):
        """Центрирует окно на экране"""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def setup_logging(self):
        self.log_handler = TextHandler(self.log_text)
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)

        # Область для перетаскивания
        ttk.Label(main_frame, text="Перетащите папку сюда или выберите вручную:",
                  font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        self.drop_frame = ttk.LabelFrame(main_frame, text="Область перетаскивания", padding="10")
        self.drop_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.drop_frame.columnconfigure(0, weight=1)

        try:
            self.drop_frame.drop_target_register(tkinterdnd2.DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        except:
            pass

        drop_label = ttk.Label(self.drop_frame,
                               text="🗂️ Перетащите папку сюда\n(или используйте кнопку 'Обзор' ниже)",
                               font=('Arial', 9),
                               foreground="gray",
                               justify=tk.CENTER)
        drop_label.grid(row=0, column=0, pady=20)

        # Остальной UI без изменений
        ttk.Label(main_frame, text="Целевая папка:").grid(row=2, column=0, sticky=tk.W, pady=5)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(main_frame, textvariable=self.path_var, width=60)
        self.path_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 0))

        self.browse_btn = ttk.Button(main_frame, text="Обзор...", command=self.browse_folder)
        self.browse_btn.grid(row=2, column=2, padx=(5, 0), pady=5)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))

        self.scan_btn = ttk.Button(button_frame, text="Сканировать", command=self.start_scanning)
        self.scan_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_btn = ttk.Button(button_frame, text="Остановить", command=self.stop_scanning, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.export_btn = ttk.Button(button_frame, text="Экспорт в Excel", command=self.export_to_excel,
                                     state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.preview_btn = ttk.Button(button_frame, text="Предпросмотр", command=self.show_preview, state=tk.DISABLED)
        self.preview_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.clear_btn = ttk.Button(button_frame, text="Очистить", command=self.clear_results)
        self.clear_btn.pack(side=tk.LEFT)

        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)

        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E))

        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.grid(row=0, column=1, padx=(5, 0))

        ttk.Label(main_frame, text="Лог выполнения:").grid(row=5, column=0, sticky=(tk.W, tk.N), pady=(10, 0))

        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=15, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))

    # ⬅️ ОСТАЛЬНЫЕ МЕТОДЫ БЕЗ ИЗМЕНЕНИЙ
    def on_drop(self, event):
        try:
            path = event.data.strip()
            if path.startswith('{') and path.endswith('}'):
                path = path[1:-1]
            path_obj = Path(path)
            if path_obj.is_dir():
                self.path_var.set(str(path_obj))
                self.status_var.set(f"Папка выбрана: {path_obj.name}")
                logging.info(f"Папка выбрана через перетаскивание: {path_obj}")
            else:
                messagebox.showwarning("Предупреждение", "Перетащите папку, а не файл")
        except Exception as e:
            logging.error(f"Ошибка при обработке перетаскивания: {e}")
            messagebox.showerror("Ошибка", f"Не удалось обработать перетаскивание: {e}")

    def show_preview(self):
        if not self.scan_results:
            messagebox.showwarning("Предупреждение", "Нет данных для предпросмотра")
            return
        PreviewWindow(self.root, self.scan_results)

    def _scan_completed(self):
        self.is_scanning = False
        self.scan_btn.config(state=tk.NORMAL)
        self.browse_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.NORMAL)
        self.preview_btn.config(state=tk.NORMAL)
        count = len(self.scan_results) if self.scan_results else 0
        self.status_var.set(f"Сканирование завершено. Найдено элементов: {count}")
        self.progress_label.config(text="100%")
        self.progress_var.set(100)
        logging.info(f"Сканирование завершено. Обработано элементов: {count}")

    def clear_results(self):
        self.scan_results = None
        self.log_text.delete(1.0, tk.END)
        self.export_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        self.status_var.set("Готов к работе")
        logging.info("Результаты очищены")

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.path_var.set(folder_path)

    def update_progress(self, value: float):
        self.progress_var.set(value)
        self.progress_label.config(text=f"{value:.1f}%")

    def start_scanning(self):
        folder_path = self.path_var.get()
        if not folder_path or not Path(folder_path).exists():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите существующую папку")
            return
        self.scan_btn.config(state=tk.DISABLED)
        self.browse_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        self.status_var.set("Подсчет файлов...")
        self.is_scanning = True
        self.scanner.set_progress_callback(self.update_progress)
        self.scan_thread = threading.Thread(target=self._scan_thread, args=(folder_path,))
        self.scan_thread.daemon = True
        self.scan_thread.start()

    def _scan_thread(self, folder_path):
        try:
            self.scan_results = self.scanner.scan_directory(folder_path)
            if self.is_scanning:
                self.root.after(0, self._scan_completed)
        except Exception as e:
            if self.is_scanning:
                self.root.after(0, self._scan_failed, str(e))

    def stop_scanning(self):
        if self.is_scanning and self.scan_thread:
            self.is_scanning = False
            self.scanner.stop_scanning()
            self.status_var.set("Остановка...")
            logging.info("Остановка сканирования...")
            self.scan_thread.join(timeout=2.0)
            self.scan_btn.config(state=tk.NORMAL)
            self.browse_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.status_var.set("Сканирование остановлено пользователем")
            if self.scan_results:
                self.export_btn.config(state=tk.NORMAL)
                self.preview_btn.config(state=tk.NORMAL)

    def _scan_failed(self, error_msg):
        self.is_scanning = False
        self.scan_btn.config(state=tk.NORMAL)
        self.browse_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        self.status_var.set("Ошибка сканирования")
        messagebox.showerror("Ошибка", f"Произошла ошибка при сканировании:\n{error_msg}")
        logging.error(f"Ошибка сканирования: {error_msg}")

    def export_to_excel(self):
        if not self.scan_results:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта. Сначала выполните сканирование.")
            return
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Сохранить как..."
            )
            if file_path:
                self.status_var.set("Экспорт в Excel...")
                self.export_btn.config(state=tk.DISABLED)
                export_thread = threading.Thread(target=self._export_thread, args=(file_path,))
                export_thread.daemon = True
                export_thread.start()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте: {e}")
            self.export_btn.config(state=tk.NORMAL)

    def _export_thread(self, file_path):
        try:
            self.exporter.export_to_excel(self.scan_results, file_path)
            self.root.after(0, self._export_completed, file_path)
        except Exception as e:
            self.root.after(0, self._export_failed, str(e))

    def _export_completed(self, file_path):
        self.export_btn.config(state=tk.NORMAL)
        self.status_var.set(f"Экспорт завершен: {Path(file_path).name}")
        logging.info(f"Экспорт в Excel завершен: {file_path}")
        messagebox.showinfo("Успех", f"Данные успешно экспортированы в:\n{file_path}")

    def _export_failed(self, error_msg):
        self.export_btn.config(state=tk.NORMAL)
        self.status_var.set("Ошибка экспорта")
        messagebox.showerror("Ошибка", f"Ошибка при экспорте в Excel:\n{error_msg}")
        logging.error(f"Ошибка экспорта: {error_msg}")


# ⬅️ ОБНОВЛЕННЫЙ КЛАСС ПРЕДПРОСМОТРА С ЦВЕТАМИ И ОТСТУПАМИ
class PreviewWindow(tk.Toplevel):
    def __init__(self, parent, scan_results):
        super().__init__(parent)
        self.title("Предпросмотр результатов")
        self.center_window(1300, 700)
        self.resizable(True, True)
        self.scan_results = scan_results

        # ⬅️ ЦВЕТА ДЛЯ РАЗНЫХ ТИПОВ ФАЙЛОВ
        self.colors = {
            'file': '#FFFFFF',  # Белый - обычные файлы
            'directory': '#E3F2FD',  # Голубой - папки
            'archive': '#FFF3E0',  # Оранжевый - архивы
            'archive_file': '#F3E5F5',  # Фиолетовый - файлы в архивах
            'archive_directory': '#E8F5E8',  # Зеленый - папки в архивах ⬅️ НОВЫЙ ЦВЕТ
            'nested_archive': '#FFEBEE',  # Красный - вложенные архивы
            'archive_error': '#FFCDD2'  # Розовый - ошибки архивов
        }

        self.setup_ui()
        self.load_data()

    def center_window(self, width, height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Предпросмотр результатов сканирования",
                  font=('Arial', 12, 'bold')).pack(pady=(0, 10))

        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview
        columns = ('number', 'name', 'type', 'path', 'size', 'depth', 'archive_depth', 'crc32')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='tree headings', height=20)

        # Настраиваем заголовки
        self.tree.heading('number', text='№')
        self.tree.heading('name', text='Имя')
        self.tree.heading('type', text='Тип')
        self.tree.heading('path', text='Путь')
        self.tree.heading('size', text='Размер')
        self.tree.heading('depth', text='Глубина')
        self.tree.heading('archive_depth', text='Ур. архива')
        self.tree.heading('crc32', text='Контр. сумма')

        # Настраиваем ширину колонок
        self.tree.column('number', width=40, anchor='center')
        self.tree.column('name', width=300)
        self.tree.column('type', width=100)  # ⬅️ Увеличили для длинных названий
        self.tree.column('path', width=350)
        self.tree.column('size', width=80)
        self.tree.column('depth', width=60)
        self.tree.column('archive_depth', width=60)
        self.tree.column('crc32', width=100)

        # Скрываем дефолтную колонку #0
        self.tree['show'] = 'headings'

        # ⬅️ НАСТРАИВАЕМ ТЕГИ ДЛЯ ЦВЕТОВ
        self._setup_tree_tags()

        # Скроллбары
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self.status_var = tk.StringVar(value=f"Загружено записей: {len(self.scan_results)}")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, pady=(5, 0))

    def _setup_tree_tags(self):
        """Настраивает цветовые теги для Treeview"""
        for item_type, color in self.colors.items():
            self.tree.tag_configure(item_type, background=color)

    def load_data(self):
        """Загружает данные в таблицу с нумерацией и правильной иерархией"""
        # Сортируем результаты для правильного порядка
        sorted_results = sorted(self.scan_results, key=lambda x: x['path'])

        # Создаем словарь для быстрого доступа к элементам по пути
        path_to_id = {}
        row_number = 1

        # Сначала добавляем все элементы
        for item in sorted_results:
            item_id = self.tree.insert(
                '',  # корневой уровень
                'end',
                values=(
                    row_number,
                    self.format_display_name(item),
                    self.get_type_display(item['type']),  # ⬅️ ПРАВИЛЬНОЕ ОТОБРАЖЕНИЕ ТИПА
                    item['path'],
                    item.get('size', 0),
                    item['depth'],
                    item.get('archive_depth', 0),
                    item.get('crc32', '')
                ),
                tags=(item['type'],)  # ⬅️ ДОБАВЛЯЕМ ТЕГ ДЛЯ ЦВЕТА
            )
            path_to_id[item['path']] = item_id
            row_number += 1

        # Перестраиваем дерево с правильной иерархией
        self._rebuild_tree_with_hierarchy(path_to_id)

    def _rebuild_tree_with_hierarchy(self, path_to_id):
        """Перестраивает дерево с правильной иерархией и отступами"""
        # Создаем временный словарь для хранения элементов по уровням
        levels = {}

        # Группируем элементы по уровням вложенности
        for item_path, item_id in path_to_id.items():
            level = item_path.count('\\') + item_path.count('/')
            if level not in levels:
                levels[level] = []
            levels[level].append((item_path, item_id))

        # Сортируем уровни по возрастанию
        sorted_levels = sorted(levels.keys())

        # Перестраиваем дерево от корня к листьям
        for level in sorted_levels:
            for item_path, item_id in levels[level]:
                item_data = self.tree.item(item_id)

                # Находим родительский путь
                parent_path = str(Path(item_path).parent)

                # Если родитель существует, перемещаем элемент
                if parent_path in path_to_id and parent_path != item_path:
                    parent_id = path_to_id[parent_path]

                    # Удаляем элемент и добавляем его к родителю
                    self.tree.delete(item_id)
                    new_item_id = self.tree.insert(
                        parent_id,
                        'end',
                        values=item_data['values'],
                        tags=item_data['tags']  # ⬅️ СОХРАНЯЕМ ТЕГИ
                    )
                    path_to_id[item_path] = new_item_id

        # ⬅️ ДОБАВЛЯЕМ ОТСТУПЫ ДЛЯ ФАЙЛОВ В ПАПКАХ
        self._add_indentation()

        # Автоматически раскрываем первые уровни
        for item_id in self.tree.get_children():
            self.tree.item(item_id, open=True)
            for child_id in self.tree.get_children(item_id):
                self.tree.item(child_id, open=True)

    def _add_indentation(self):
        """Добавляет визуальные отступы для файлов в папках"""
        for item_id in self.tree.get_children():
            parent_data = self.tree.item(item_id)
            parent_values = list(parent_data['values'])

            # ⬅️ ДЛЯ КОРНЕВЫХ ЭЛЕМЕНТОВ - БЕЗ ОТСТУПА
            parent_values[1] = parent_values[1]  # имя без изменений

            self.tree.item(item_id, values=tuple(parent_values))

            # ⬅️ ДЛЯ ДОЧЕРНИХ ЭЛЕМЕНТОВ - ДОБАВЛЯЕМ ОТСТУП
            for child_id in self.tree.get_children(item_id):
                child_data = self.tree.item(child_id)
                child_values = list(child_data['values'])

                # Добавляем отступ к имени
                child_values[1] = "    " + child_values[1]

                self.tree.item(child_id, values=tuple(child_values))

                # ⬅️ РЕКУРСИВНО ДЛЯ ВЛОЖЕННЫХ ЭЛЕМЕНТОВ
                for grandchild_id in self.tree.get_children(child_id):
                    grandchild_data = self.tree.item(grandchild_id)
                    grandchild_values = list(grandchild_data['values'])

                    # Добавляем двойной отступ
                    grandchild_values[1] = "        " + grandchild_values[1]

                    self.tree.item(grandchild_id, values=tuple(grandchild_values))

    def format_display_name(self, item):
        """Форматирует имя для отображения в таблице"""
        name = item['name']

        # ⬅️ ПРАВИЛЬНЫЕ ПРЕФИКСЫ ДЛЯ РАЗНЫХ ТИПОВ
        if item['type'] == 'directory':
            return f"[ПАПКА] {name}"
        elif item['type'] == 'archive':
            return f"[АРХИВ] {name}"
        elif item['type'] == 'nested_archive':
            return f"[ВЛОЖ. АРХИВ] {name}"
        elif item['type'] == 'archive_file':
            return f"[ФАЙЛ В АРХИВЕ] {name}"
        elif item['type'] == 'archive_directory':
            return f"[ПАПКА В АРХИВЕ] {name}"  # ⬅️ НОВЫЙ ПРЕФИКС
        else:
            return f"[ФАЙЛ] {name}"

    def get_type_display(self, item_type):
        """Возвращает правильное отображение типа"""
        type_map = {
            'file': 'Файл',
            'directory': 'Папка',
            'archive': 'Архив',
            'archive_file': 'Файл в архиве',
            'archive_directory': 'Папка в архиве',  # ⬅️ НОВЫЙ ТИП
            'nested_archive': 'Влож. архив',
            'archive_error': 'Ошибка архива'
        }
        return type_map.get(item_type, item_type)

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.update_idletasks()

        self.text_widget.after(0, append)

def main():
    try:
        import tkinterdnd2
        root = tkinterdnd2.TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
        logging.warning("tkinterdnd2 не установлен, перетаскивание недоступно")

    app = ArchiveScannerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()