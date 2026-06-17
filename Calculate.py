import tkinter as tk
from tkinter import ttk, messagebox
import win32com.client
import time
import threading
import pythoncom
import queue
import math

def calculate_item_length(item):
    """Расчет длины объекта в миллиметрах"""
    t = float(item.Length) * 0.352778 if hasattr(item, 'Length') else 0
    if hasattr(item, 'PageItems'):
        for s in item.PageItems:
            t += calculate_item_length(s)
    return t

# Переменные для хранения результата
current_selection_length = 0.0
is_running = True
illustrator_available = False

# Очередь для безопасного обновления UI
ai_event_queue = queue.Queue(maxsize=100)

# Список возможных ProgID для разных версий Illustrator (от CS3 до 30)
ILLUSTRATOR_PROGIDS = [
    "Illustrator.Application",
    "Illustrator.Application.30",
    "Illustrator.Application.29",
    "Illustrator.Application.28",
    "Illustrator.Application.27",
    "Illustrator.Application.26",
    "Illustrator.Application.25",
    "Illustrator.Application.24",
    "Illustrator.Application.23",
    "Illustrator.Application.22",
    "Illustrator.Application.21",
    "Illustrator.Application.2024",
    "Illustrator.Application.2023",
    "Illustrator.Application.2022",
    "Illustrator.Application.2021",
    "Illustrator.Application.2020",
    "Illustrator.Application.2019",
    "Illustrator.Application.2018",
    "Illustrator.Application.2017",
    "Illustrator.Application.CC",
    "Illustrator.Application.CS6",
    "Illustrator.Application.CS5",
    "Illustrator.Application.CS4",
    "Illustrator.Application.CS3"
]

def check_illustrator_without_starting():
    """Проверяет, запущен ли Illustrator, перебирая все возможные ProgID"""
    try:
        pythoncom.CoInitialize()
        for progid in ILLUSTRATOR_PROGIDS:
            try:
                ai = win32com.client.GetActiveObject(progid)
                version = getattr(ai, 'Version', 'неизвестная')
                ai = None
                return True, f"Illustrator версии {version} запущен (подключение через {progid})"
            except:
                continue
        return False, "Illustrator не запущен или используется неподдерживаемая версия"
    except Exception as e:
        return False, f"Ошибка проверки: {e}"
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

def get_illustrator_application():
    """Получает объект Illustrator Application, перебирая все возможные ProgID"""
    for progid in ILLUSTRATOR_PROGIDS:
        try:
            ai = win32com.client.GetActiveObject(progid)
            return ai
        except:
            continue
    return None

def monitor_selection():
    """Основная функция мониторинга выделения в Illustrator"""
    global current_selection_length, is_running, illustrator_available
    
    ai_local = None
    monitoring = False
    
    print("Мониторинг Illustrator запущен")
    
    while is_running:
        try:
            illustrator_running, message = check_illustrator_without_starting()
            
            if not illustrator_running:
                illustrator_available = False
                current_selection_length = 0.0
                if monitoring:
                    monitoring = False
                    print("Illustrator закрыт или недоступен")
                    try:
                        ai_event_queue.put(('status', False), timeout=0.1)
                    except:
                        pass
                time.sleep(3)
                continue
            
            if not monitoring:
                print(f"Illustrator обнаружен: {message}")
                monitoring = True
                illustrator_available = True
                try:
                    ai_event_queue.put(('status', True), timeout=0.1)
                except:
                    pass
            
            try:
                pythoncom.CoInitialize()
                ai_local = get_illustrator_application()
                
                if ai_local is None:
                    print("Не удалось подключиться к Illustrator")
                    time.sleep(2)
                    continue
                
                selection = list(ai_local.Selection)
                
                if selection:
                    total_length = sum(calculate_item_length(i) for i in selection)
                    new_length = round(total_length / 1000, 1)
                    
                    if abs(new_length - current_selection_length) > 0.01:
                        current_selection_length = new_length
                        try:
                            ai_event_queue.put(('length_update', new_length), timeout=0.1)
                        except:
                            pass
                        check_interval = 0.5
                    else:
                        check_interval = 1.0
                else:
                    if current_selection_length != 0.0:
                        current_selection_length = 0.0
                        try:
                            ai_event_queue.put(('length_update', 0.0), timeout=0.1)
                        except:
                            pass
                    check_interval = 2.0
                
                ai_local = None
                pythoncom.CoUninitialize()
                time.sleep(check_interval)
                
            except Exception as e:
                print(f"Ошибка в цикле мониторинга: {e}")
                try:
                    ai_local = None
                    pythoncom.CoUninitialize()
                except:
                    pass
                time.sleep(1)
                
        except Exception as e:
            print(f"Общая ошибка: {e}")
            time.sleep(3)


class CombinedCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Калькулятор by Kirill V15")
        
        # Размеры для каждой вкладки
        self.tab_sizes = {
            0: (600, 570),   # Ножи
            1: (1200, 1050),  # Вкладыши
            2: (500, 330),   # % Выгоды
            3: (700, 550)    # Расчет тиража
        }
        
        self.root.geometry("600x570")
        
        # Флаги для работы с Illustrator
        self.ai_value_inserted = False
        self.last_ai_value = 0.0
        self.last_update_time = 0
        self.min_update_interval = 0.3
        
        # Создаем Notebook (вкладки)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Создаем вкладки
        self.create_knife_calculator_tab()
        self.create_inlay_calculator_tab()
        self.create_percentage_calculator_tab()
        self.create_edition_calculator_tab()
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.notebook.select(0)
        self.root.resizable(True, True)
        self.current_tab = 0
        
        # ===== ГЛОБАЛЬНЫЙ ОБРАБОТЧИК КОПИРОВАНИЯ ПО СКАН-КОДУ (РАБОТАЕТ НА ЛЮБОЙ РАСКЛАДКЕ) =====
        self.root.bind('<Key>', self.on_global_key_press)
        
        self.root.after(100, self.process_ai_queue)
        self.root.after(1000, self.delayed_start)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    # ===== ОБРАБОТЧИК ГЛОБАЛЬНЫХ КЛАВИШ =====
    def on_global_key_press(self, event):
        """Обрабатывает нажатия клавиш, проверяет Ctrl+C (скан-код 67) независимо от раскладки"""
        # Проверяем, что зажат Ctrl (бит 0x0004) и нажата клавиша 'C' (скан-код 67)
        if (event.state & 0x0004) and event.keycode == 67:
            widget = self.root.focus_get()
            if not widget:
                return
            # Пытаемся получить выделенный текст
            try:
                if hasattr(widget, 'selection_get'):
                    selected = widget.selection_get()
                    if selected:
                        self.root.clipboard_clear()
                        self.root.clipboard_append(selected)
                        self.show_copy_notification()
                        return "break"
            except tk.TclError:
                # Нет выделения
                pass
        # Для остальных клавиш пропускаем событие (не блокируем)
        return None
    
    def process_ai_queue(self):
        """Обработка событий из очереди Illustrator"""
        try:
            processed = 0
            while processed < 10:
                try:
                    event_type, data = ai_event_queue.get_nowait()
                    
                    if event_type == 'length_update' and self.current_tab == 0:
                        current_time = time.time()
                        if current_time - self.last_update_time > self.min_update_interval:
                            self.last_update_time = current_time
                            
                            formatted_value = f"{data:.1f}"
                            
                            if abs(data - self.last_ai_value) > 0.01:
                                self.last_ai_value = data
                                
                                if not self.ai_value_inserted:
                                    self.input_value.set(formatted_value)
                                    self.length_input.icursor(tk.END)
                                    self.knife_calculate()
                                    
                                    if data > 0:
                                        self.show_ai_notification(f"Из Illustrator: {formatted_value} мм")
                    
                    elif event_type == 'status':
                        status_text = "Illustrator: ✅ Запущен" if data else "Illustrator: ❌ Не запущен"
                        self.ai_status_var.set(status_text)
                        
                    processed += 1
                    
                except queue.Empty:
                    break
                    
        except Exception as e:
            print(f"Ошибка обработки очереди: {e}")
        finally:
            self.root.after(50, self.process_ai_queue)
    
    def delayed_start(self):
        """Запускает потоки с задержкой"""
        self.start_illustrator_monitor()
        self.root.after(2000, self.check_ai_status)
    
    def on_closing(self):
        """Обработчик закрытия окна"""
        global is_running
        is_running = False
        time.sleep(0.5)
        self.root.destroy()
    
    def start_illustrator_monitor(self):
        """Запускает поток для мониторинга выделения в Illustrator"""
        self.monitor_thread = threading.Thread(target=monitor_selection, daemon=True)
        self.monitor_thread.start()
        print("Запущен поток мониторинга Illustrator")
    
    def show_ai_notification(self, message):
        """Показывает всплывающее уведомление"""
        try:
            if hasattr(self, 'ai_popup') and self.ai_popup:
                try:
                    self.ai_popup.destroy()
                except:
                    pass
            
            self.ai_popup = tk.Toplevel(self.root)
            self.ai_popup.wm_overrideredirect(True)
            self.ai_popup.configure(bg='#e6f3ff')
            
            if hasattr(self, 'length_input'):
                x = self.length_input.winfo_rootx()
                y = self.length_input.winfo_rooty() - 30
                self.ai_popup.geometry(f"+{x}+{y}")
            
            label = tk.Label(self.ai_popup, text=message, bg='#e6f3ff', fg='#0066cc', 
                             font=("Arial", 8), padx=8, pady=3)
            label.pack()
            
            self.ai_popup.after(1500, self.ai_popup.destroy)
        except:
            pass
    
    def on_tab_changed(self, event=None):
        """Обработчик смены вкладки"""
        selected_index = self.notebook.index(self.notebook.select())
        self.current_tab = selected_index
        self.ai_value_inserted = False
        
        width, height = self.tab_sizes.get(selected_index, (600, 400))
        current_x = self.root.winfo_x()
        current_y = self.root.winfo_y()
        self.root.geometry(f"{width}x{height}+{current_x}+{current_y}")
        
        if selected_index == 1 and hasattr(self, 'left_canvas'):
            self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
    
    def check_ai_status(self):
        """Проверяет статус подключения к Illustrator"""
        def check():
            illustrator_running, message = check_illustrator_without_starting()
            
            if illustrator_running:
                try:
                    ai_event_queue.put(('status', True), timeout=0.1)
                except:
                    pass
            else:
                try:
                    ai_event_queue.put(('status', False), timeout=0.1)
                except:
                    pass
            
            self.root.after(10000, self.check_ai_status)
        
        threading.Thread(target=check, daemon=True).start()
    
    def insert_from_illustrator(self):
        """Вручную вставляет значение из Illustrator"""
        global current_selection_length
        
        illustrator_running, message = check_illustrator_without_starting()
        
        if not illustrator_running:
            messagebox.showwarning("Illustrator не запущен", 
                                 "Adobe Illustrator не запущен или недоступен.")
            return
        
        if current_selection_length > 0:
            formatted_value = f"{current_selection_length:.1f}"
            self.input_value.set(formatted_value)
            self.ai_value_inserted = True
            self.knife_calculate()
            self.show_ai_notification(f"Значение вставлено: {formatted_value} мм")
        else:
            messagebox.showinfo("Информация", 
                              "В Illustrator ничего не выделено.")
    
    # ===== ВКЛАДКА "Расчет штампа" =====
    def create_knife_calculator_tab(self):
        knife_tab = ttk.Frame(self.notebook)
        self.notebook.add(knife_tab, text='Расчет штампа')
        
        self.input_value = tk.StringVar()
        self.result_value = tk.StringVar()
        self.selected_numbers = {
            4: tk.BooleanVar(value=False),
            5: tk.BooleanVar(value=False),
            6: tk.BooleanVar(value=False),
            7: tk.BooleanVar(value=False)
        }
        self.quantity_vars = {
            4: tk.StringVar(value="0"),
            5: tk.StringVar(value="0"),
            6: tk.StringVar(value="0"),
            7: tk.StringVar(value="0")
        }
        
        self.punching_costs = {4:360,5:360,6:400,7:450}
        self.diameter_names = {4:"ø 4 мм",5:"ø 5 мм",6:"ø 6-8 мм",7:"ø 9-10 мм"}
        self.quantity_entries = {}
        self.result_entry = None
        self.setup_knife_ui(knife_tab)
    
    def setup_knife_ui(self, parent):
        main_frame = ttk.Frame(parent, padding="15")
        main_frame.pack(fill='both', expand=True)
        
        title_label = ttk.Label(main_frame, text="Калькулятор стоимости ножей", font=("Arial",14,"bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0,15))
        
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=0, column=1, sticky=tk.E, pady=(0,15))
        self.ai_status_var = tk.StringVar(value="Illustrator: ❌ Не запущен")
        self.ai_status_label = ttk.Label(status_frame, textvariable=self.ai_status_var, font=("Arial",8))
        self.ai_status_label.pack(side=tk.RIGHT)
        
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0,5))
        ttk.Label(input_frame, text="Введите длину ножей (мм):").pack(side=tk.LEFT, padx=(0,10))
        ai_button = ttk.Button(input_frame, text="Вставить из Illustrator", command=self.insert_from_illustrator, width=20)
        ai_button.pack(side=tk.RIGHT, padx=(5,0))
        
        self.length_input = ttk.Entry(main_frame, textvariable=self.input_value, width=20, font=("Arial",10))
        self.length_input.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0,15))
        def validate_input(char): return char in "0123456789,."
        self.length_input.config(validate="key", validatecommand=(self.length_input.register(validate_input), '%S'))
        
        def on_key_press(event):
            if event.char or event.keysym in ['BackSpace','Delete','Left','Right']:
                self.ai_value_inserted = False
            return True
        self.length_input.bind('<KeyPress>', on_key_press)
        
        result_frame = ttk.Frame(main_frame)
        result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0,15))
        ttk.Label(result_frame, text="ИТОГОВАЯ ЦЕНА:", font=("Arial",12,"bold")).grid(row=0, column=0, sticky=tk.W)
        self.result_entry = ttk.Entry(result_frame, textvariable=self.result_value, font=("Arial",14,"bold"), foreground="green", state='readonly', width=20, justify='left')
        self.result_entry.grid(row=1, column=0, sticky=tk.W, pady=(5,0))
        # Контекстное меню для копирования (дополнительно)
        self.create_context_menu(self.result_entry)
        
        selection_frame = ttk.LabelFrame(main_frame, text="Выберите пробивку", padding="10")
        selection_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0,15))
        
        ttk.Label(selection_frame, text="Диаметр, мм", font=("Arial",9,"bold")).grid(row=0, column=0, padx=(0,25))
        ttk.Label(selection_frame, text="Добавить", font=("Arial",9,"bold")).grid(row=0, column=1, padx=(0,25))
        ttk.Label(selection_frame, text="Количество шт.", font=("Arial",9,"bold")).grid(row=0, column=2, padx=(0,25))
        ttk.Label(selection_frame, text="Стоимость", font=("Arial",9,"bold")).grid(row=0, column=3)
        
        self.cost_labels = {}
        diameters = [4,5,6,7]
        for i, d in enumerate(diameters,1):
            ttk.Label(selection_frame, text=self.diameter_names[d]).grid(row=i, column=0, padx=(0,25), pady=6)
            cb = ttk.Checkbutton(selection_frame, variable=self.selected_numbers[d], command=self.toggle_quantity_entry)
            cb.grid(row=i, column=1, padx=(0,25), pady=6)
            qe = ttk.Entry(selection_frame, textvariable=self.quantity_vars[d], width=8, font=("Arial",9), state='disabled')
            qe.grid(row=i, column=2, padx=(0,25), pady=6)
            self.quantity_entries[d] = qe
            qe.config(validate="key", validatecommand=(qe.register(lambda char: char.isdigit()), '%S'))
            cl = ttk.Label(selection_frame, text="0 руб", foreground="blue")
            cl.grid(row=i, column=3, pady=6)
            self.cost_labels[d] = cl
        
        info_frame = ttk.Frame(main_frame)
        info_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0,10))
        ttk.Label(info_frame, text="* Базовая стоимость: 4300 руб + 1700 руб/мм длины", font=("Arial",8), foreground="gray").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(info_frame, text="* ø4 мм: 360 руб за штуку", font=("Arial",8), foreground="gray").grid(row=1, column=0, sticky=tk.W)
        ttk.Label(info_frame, text="* ø5 мм: 360 руб за штуку", font=("Arial",8), foreground="gray").grid(row=2, column=0, sticky=tk.W)
        ttk.Label(info_frame, text="* ø6-8 мм: 400 руб за штуку", font=("Arial",8), foreground="gray").grid(row=3, column=0, sticky=tk.W)
        ttk.Label(info_frame, text="* ø9-10 мм: 450 руб за штуку", font=("Arial",8), foreground="gray").grid(row=4, column=0, sticky=tk.W)
        
        self.length_input.bind('<KeyRelease>', self.on_input_change)
        for var in self.quantity_vars.values():
            var.trace('w', self.on_quantity_change)
        
        main_frame.columnconfigure(0, weight=1); main_frame.columnconfigure(1, weight=1)
        selection_frame.columnconfigure(3, weight=1)
        self.knife_calculate()
    
    def show_copy_notification(self):
        """Показывает всплывающее уведомление о копировании"""
        popup = tk.Toplevel(self.root)
        popup.wm_overrideredirect(True)
        popup.configure(bg='lightgreen')
        # Позиционируем относительно активного виджета
        widget = self.root.focus_get()
        if widget and widget.winfo_exists():
            x = widget.winfo_rootx() + 10
            y = widget.winfo_rooty() + widget.winfo_height() + 5
        else:
            x = self.root.winfo_rootx() + 100
            y = self.root.winfo_rooty() + 100
        popup.geometry(f"+{x}+{y}")
        label = tk.Label(popup, text="Скопировано!", bg='lightgreen', fg='black', font=("Arial",9))
        label.pack(padx=10, pady=5)
        popup.after(1500, popup.destroy)
    
    def create_context_menu(self, widget):
        """Создаёт контекстное меню с опцией 'Копировать' для виджета"""
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Копировать", command=lambda: self.copy_widget_text(widget))
        def show_menu(event): menu.post(event.x_root, event.y_root)
        widget.bind("<Button-3>", show_menu)
    
    def copy_widget_text(self, widget):
        """Копирует выделенный текст из виджета или весь текст, если выделения нет"""
        try:
            selected = widget.selection_get()
        except tk.TclError:
            # Если нет выделения, берём весь текст (для Entry/Text)
            if isinstance(widget, (tk.Entry, ttk.Entry)):
                selected = widget.get()
            elif isinstance(widget, tk.Text):
                selected = widget.get("1.0", tk.END).rstrip("\n")
            else:
                return
        if selected:
            self.root.clipboard_clear()
            self.root.clipboard_append(selected)
            self.show_copy_notification()
    
    def parse_float(self, value):
        if not value: return 0.0
        try: return float(value.replace(',','.'))
        except ValueError: return 0.0
    
    def toggle_quantity_entry(self):
        for d in [4,5,6,7]:
            if self.selected_numbers[d].get():
                self.quantity_entries[d].config(state='normal')
            else:
                self.quantity_entries[d].config(state='disabled')
                self.quantity_vars[d].set("0")
        self.knife_calculate()
    
    def on_input_change(self, event=None):
        if event and event.char: self.ai_value_inserted = False
        self.knife_calculate()
    
    def on_quantity_change(self, *args):
        self.knife_calculate()
    
    def update_punching_costs(self):
        for d in [4,5,6,7]:
            try:
                q = int(self.quantity_vars[d].get() or "0")
                if q>0 and self.selected_numbers[d].get():
                    cost = self.punching_costs[d]*q
                    self.cost_labels[d].config(text=f"{cost} руб")
                else:
                    self.cost_labels[d].config(text="0 руб")
            except:
                self.cost_labels[d].config(text="0 руб")
    
    def knife_calculate(self):
        try:
            length_text = self.input_value.get().strip()
            if not length_text:
                self.result_value.set("Введите длину ножей")
                return
            length = self.parse_float(length_text)
            if length <= 0:
                self.result_value.set("Длина должна быть > 0")
                return
            base_cost = 4300 + (length*1700)
            punching_cost = 0
            for d, var in self.selected_numbers.items():
                if var.get():
                    try:
                        q = int(self.quantity_vars[d].get() or "0")
                        if q>0: punching_cost += self.punching_costs[d]*q
                    except: pass
            total = base_cost + punching_cost
            self.result_value.set(f"{total:,.2f} руб".replace(","," "))
            self.update_punching_costs()
        except ValueError:
            self.result_value.set("Ошибка: введите число")
    
    # ===== ВКЛАДКА "Расчет вкладышей" =====
    def create_inlay_calculator_tab(self):
        inlay_tab = ttk.Frame(self.notebook)
        self.notebook.add(inlay_tab, text='Расчет вкладышей')
        
        # Стандартные форматы бумаги
        self.paper_formats = {
            "620x940": (940, 620),
            "640x900": (900, 640),
            "700x1000": (1000, 700),
            "720x1040": (1040, 720)
        }
        
        # Стандартные форматы вкладышей (ISO 216)
        self.inlay_standart_formats = {
            "A0": (841, 1189),
            "A1": (594, 841),
            "A2": (420, 594),
            "A3": (297, 420),
            "A4": (210, 297),
            "A5": (148, 210),
            "A6": (105, 148),
            "A7": (74, 105),
            "A8": (52, 74)
        }
        
        self.selected_format = tk.StringVar(value="620x940")
        self.custom_width = tk.StringVar()
        self.custom_height = tk.StringVar()
        self.use_custom = tk.BooleanVar(value=False)
        self.inlay_width = tk.StringVar()
        self.inlay_height = tk.StringVar()
        self.selected_inlay_format = tk.StringVar()
        
        self.print_type = tk.StringVar(value="односторонняя")
        self.color_mode = tk.StringVar(value="color")
        self.auto_rotate = tk.BooleanVar(value=True)
        self.auto_calculate_all = tk.BooleanVar(value=True)
        self.shrink_mode = tk.BooleanVar(value=False)
        
        self.total_size_text = tk.StringVar(value="Общий размер вкладышей: -")
        self.best_layout_text = tk.StringVar(value="Оптимальное расположение: -")
        self.best_paper_text = tk.StringVar(value="Лучший формат: -")
        
        self.setup_inlay_ui(inlay_tab)
    
    def setup_inlay_ui(self, parent):
        main_panel = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        main_panel.pack(fill='both', expand=True)
        
        left_frame = ttk.Frame(main_panel)
        main_panel.add(left_frame, weight=1)
        
        self.left_canvas = tk.Canvas(left_frame, highlightthickness=0)
        left_scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.left_canvas.yview)
        self.left_scrollable = ttk.Frame(self.left_canvas)
        
        self.left_scrollable.bind(
            "<Configure>",
            lambda e: self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))
        )
        
        self.left_canvas.create_window((0, 0), window=self.left_scrollable, anchor="nw")
        self.left_canvas.configure(yscrollcommand=left_scrollbar.set)
        self.left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")
        
        right_frame = ttk.Frame(main_panel)
        main_panel.add(right_frame, weight=2)
        
        right_container = ttk.Frame(right_frame)
        right_container.pack(fill='both', expand=True, padx=5, pady=5)
        
        right_scrollbar = ttk.Scrollbar(right_container)
        right_scrollbar.pack(side="right", fill="y")
        
        self.result_textbox = tk.Text(right_container, wrap="word", yscrollcommand=right_scrollbar.set,
                                      font=("Consolas", 9))
        self.result_textbox.pack(side="left", fill="both", expand=True)
        right_scrollbar.config(command=self.result_textbox.yview)
        
        # Горячие клавиши для текстового поля: Ctrl+A для выделения всего
        def select_all_text(event):
            self.result_textbox.tag_add('sel', '1.0', 'end')
            return 'break'
        self.result_textbox.bind('<Control-a>', select_all_text)
        self.result_textbox.bind('<Control-A>', select_all_text)
        # Ctrl+C обрабатывается глобально, так что ничего не привязываем
        
        # Контекстное меню для текстового поля
        def copy_from_textbox():
            try:
                selected = self.result_textbox.selection_get()
                self.root.clipboard_clear()
                self.root.clipboard_append(selected)
                self.show_copy_notification()
            except tk.TclError:
                pass
        textbox_menu = tk.Menu(self.result_textbox, tearoff=0)
        textbox_menu.add_command(label="Копировать", command=copy_from_textbox)
        self.result_textbox.bind("<Button-3>", lambda e: textbox_menu.post(e.x_root, e.y_root))
        
        title_label = ttk.Label(self.left_scrollable, text="Калькулятор размещения вкладышей", font=("Arial",14,"bold"))
        title_label.pack(anchor='w', pady=10, padx=10)
        
        format_frame = ttk.LabelFrame(self.left_scrollable, text="Формат бумаги", padding=10)
        format_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(format_frame, text="Стандартные форматы:", font=("Arial",9,"bold")).pack(anchor="w")
        for name in self.paper_formats.keys():
            rb = ttk.Radiobutton(format_frame, text=name, variable=self.selected_format, value=name)
            rb.pack(anchor="w", padx=20)
        ttk.Checkbutton(format_frame, text="Произвольный размер", variable=self.use_custom, command=self.toggle_custom_format).pack(anchor="w", pady=(10,5))
        
        custom_frame = ttk.Frame(format_frame)
        custom_frame.pack(fill="x", padx=20, pady=5)
        ttk.Label(custom_frame, text="Ширина листа (мм):").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.custom_width_entry = ttk.Entry(custom_frame, textvariable=self.custom_width, width=10, state='disabled')
        self.custom_width_entry.grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(custom_frame, text="Высота листа (мм):").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.custom_height_entry = ttk.Entry(custom_frame, textvariable=self.custom_height, width=10, state='disabled')
        self.custom_height_entry.grid(row=1, column=1, padx=5, pady=2)
        
        inlay_frame = ttk.LabelFrame(self.left_scrollable, text="Размер вкладыша (в мм)", padding=10)
        inlay_frame.pack(fill="x", padx=10, pady=5)
        
        format_sel_frame = ttk.Frame(inlay_frame)
        format_sel_frame.pack(fill="x", pady=5)
        ttk.Label(format_sel_frame, text="Стандартный формат:").pack(side=tk.LEFT, padx=(0,5))
        self.inlay_format_combobox = ttk.Combobox(format_sel_frame, textvariable=self.selected_inlay_format,
                                                  values=[""] + list(self.inlay_standart_formats.keys()),
                                                  state="readonly", width=10)
        self.inlay_format_combobox.pack(side=tk.LEFT, padx=5)
        self.inlay_format_combobox.bind("<<ComboboxSelected>>", self.on_inlay_format_selected)
        self.inlay_format_size_label = ttk.Label(format_sel_frame, text="", foreground="blue")
        self.inlay_format_size_label.pack(side=tk.LEFT, padx=10)
        
        wf = ttk.Frame(inlay_frame); wf.pack(fill="x", pady=5)
        ttk.Label(wf, text="Ширина:").pack(side="left", padx=5)
        self.inlay_width_entry = ttk.Entry(wf, textvariable=self.inlay_width, width=10)
        self.inlay_width_entry.pack(side="left")
        ttk.Label(wf, text="мм").pack(side="left", padx=5)
        
        hf = ttk.Frame(inlay_frame); hf.pack(fill="x", pady=5)
        ttk.Label(hf, text="Высота:").pack(side="left", padx=5)
        self.inlay_height_entry = ttk.Entry(hf, textvariable=self.inlay_height, width=10)
        self.inlay_height_entry.pack(side="left")
        ttk.Label(hf, text="мм").pack(side="left", padx=5)
        
        self.inlay_width_entry.bind('<KeyRelease>', self.on_inlay_manual_change)
        self.inlay_height_entry.bind('<KeyRelease>', self.on_inlay_manual_change)
        
        print_frame = ttk.LabelFrame(self.left_scrollable, text="Тип печати и оборота", padding=10)
        print_frame.pack(fill="x", padx=10, pady=5)
        ttk.Radiobutton(print_frame, text="📄 Односторонняя", variable=self.print_type, value="односторонняя").pack(anchor="w",pady=2)
        ttk.Radiobutton(print_frame, text="🔄 Двусторонняя - ЧУЖОЙ оборот", variable=self.print_type, value="двусторонняя_чужой").pack(anchor="w",pady=2)
        ttk.Radiobutton(print_frame, text="⚡ Двусторонняя - СВОЙ оборот", variable=self.print_type, value="двусторонняя_свой").pack(anchor="w",pady=2)
        
        color_frame = ttk.LabelFrame(self.left_scrollable, text="Тип печати по цвету", padding=10)
        color_frame.pack(fill="x", padx=10, pady=5)
        ttk.Radiobutton(color_frame, text="🎨 Цветная (с припусками +4 мм)", variable=self.color_mode, value="color").pack(anchor="w",pady=2)
        ttk.Radiobutton(color_frame, text="⚫ Черно-белая (без припусков)", variable=self.color_mode, value="bw").pack(anchor="w",pady=2)
        
        opt_frame = ttk.LabelFrame(self.left_scrollable, text="Опции оптимизации", padding=10)
        opt_frame.pack(fill="x", padx=10, pady=5)
        ttk.Checkbutton(opt_frame, text="🔄 Автоматически переворачивать вкладыши", variable=self.auto_rotate).pack(anchor="w",pady=2)
        ttk.Checkbutton(opt_frame, text="📊 Автоматически рассчитывать все форматы", variable=self.auto_calculate_all).pack(anchor="w",pady=2)
        ttk.Checkbutton(opt_frame, text="📦 Сжать до формата (уменьшить отступы/припуски)", variable=self.shrink_mode).pack(anchor="w",pady=2)
        
        ttk.Button(self.left_scrollable, text="🧮 Рассчитать", command=self.inlay_calculate).pack(pady=10)
        
        summary_frame = ttk.LabelFrame(self.left_scrollable, text="Краткие результаты", padding=10)
        summary_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(summary_frame, text="🏆 Лучший формат бумаги:", font=("Arial",9,"bold")).pack(anchor="w")
        self.best_paper_label = ttk.Label(summary_frame, textvariable=self.best_paper_text, foreground="green")
        self.best_paper_label.pack(anchor="w", padx=10)
        ttk.Label(summary_frame, text="📐 Оптимальное расположение:", font=("Arial",9,"bold")).pack(anchor="w", pady=(5,0))
        self.best_layout_label = ttk.Label(summary_frame, textvariable=self.best_layout_text)
        self.best_layout_label.pack(anchor="w", padx=10)
        ttk.Label(summary_frame, text="📏 Общий размер вкладышей:", font=("Arial",9,"bold")).pack(anchor="w", pady=(5,0))
        self.total_size_label = ttk.Label(summary_frame, textvariable=self.total_size_text)
        self.total_size_label.pack(anchor="w", padx=10)
        
        # Контекстное меню для ярлыков результатов (если нужно)
        for lbl in [self.best_paper_label, self.best_layout_label, self.total_size_label]:
            self.create_context_menu(lbl)
        
        def on_mousewheel(event):
            self.left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.left_canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    def toggle_custom_format(self):
        if self.use_custom.get():
            self.custom_width_entry.config(state='normal')
            self.custom_height_entry.config(state='normal')
        else:
            self.custom_width_entry.config(state='disabled')
            self.custom_height_entry.config(state='disabled')
    
    def on_inlay_format_selected(self, event=None):
        fmt = self.selected_inlay_format.get()
        if fmt in self.inlay_standart_formats:
            w, h = self.inlay_standart_formats[fmt]
            self.inlay_width.set(str(w))
            self.inlay_height.set(str(h))
            self.inlay_format_size_label.config(text=f"{w}×{h} мм")
        else:
            self.inlay_format_size_label.config(text="")
    
    def on_inlay_manual_change(self, event=None):
        try:
            w = float(self.inlay_width.get().replace(',','.'))
            h = float(self.inlay_height.get().replace(',','.'))
        except ValueError:
            if self.selected_inlay_format.get():
                self.selected_inlay_format.set("")
                self.inlay_format_size_label.config(text="")
            return
        found = None
        for fmt, (fw, fh) in self.inlay_standart_formats.items():
            if abs(w - fw) < 0.5 and abs(h - fh) < 0.5:
                found = fmt
                break
        if found:
            if self.selected_inlay_format.get() != found:
                self.selected_inlay_format.set(found)
                self.inlay_format_size_label.config(text=f"{fw}×{fh} мм")
        else:
            if self.selected_inlay_format.get():
                self.selected_inlay_format.set("")
                self.inlay_format_size_label.config(text="")
    
    def inlay_calculate_layout(self, paper_width, paper_height, inlay_h, inlay_w, rotated=False, 
                               color_mode="color", top_bottom_margin=25, side_margin=10, 
                               extra_margin=4, height_reduction=0):
        available_width = paper_width - side_margin
        available_height = paper_height - top_bottom_margin
        
        if color_mode == "color":
            inlay_w_eff = inlay_w + extra_margin
            inlay_h_eff = inlay_h + extra_margin
        else:
            inlay_w_eff = inlay_w
            inlay_h_eff = inlay_h - height_reduction
        
        if rotated:
            w_eff = inlay_h_eff
            h_eff = inlay_w_eff
        else:
            w_eff = inlay_w_eff
            h_eff = inlay_h_eff
        
        count_width = int(available_width // w_eff) if w_eff > 0 else 0
        count_height = int(available_height // h_eff) if h_eff > 0 else 0
        
        total_count = count_width * count_height
        total_width = count_width * w_eff
        total_height = count_height * h_eff
        
        used_area = total_width * total_height
        available_area = available_width * available_height
        efficiency = (used_area / available_area * 100) if available_area > 0 else 0
        
        if rotated:
            orientation = f"{inlay_w}×{inlay_h} мм"
        else:
            orientation = f"{inlay_h}×{inlay_w} мм"
        
        return {
            'paper_width': paper_width,
            'paper_height': paper_height,
            'rotated': rotated,
            'count_width': count_width,
            'count_height': count_height,
            'total_count': total_count,
            'total_width': total_width,
            'total_height': total_height,
            'inlay_w': inlay_w,
            'inlay_h': inlay_h,
            'inlay_w_eff': inlay_w_eff,
            'inlay_h_eff': inlay_h_eff,
            'available_width': available_width,
            'available_height': available_height,
            'efficiency': efficiency,
            'orientation': orientation,
            'top_bottom_margin': top_bottom_margin,
            'side_margin': side_margin,
            'extra_margin': extra_margin if color_mode=='color' else 0,
            'height_reduction': height_reduction if color_mode=='bw' else 0,
            'color_mode': color_mode
        }
    
    def calculate_final_count(self, layout, print_type):
        total = layout['total_count']
        cw = layout['count_width']
        ch = layout['count_height']
        if print_type == "односторонняя":
            return total
        elif print_type == "двусторонняя_чужой":
            return total-1 if total%2 else total
        elif print_type == "двусторонняя_свой":
            if cw%2:
                return (cw-1)*ch
            return total
        return total
    
    def inlay_calculate(self):
        try:
            self.result_textbox.delete(1.0, tk.END)
            
            w = float(self.inlay_width.get().replace(",","."))
            h = float(self.inlay_height.get().replace(",","."))
            if w<=0 or h<=0:
                raise ValueError("Размеры вкладыша должны быть положительными")
            
            ptype = self.print_type.get()
            color = self.color_mode.get()
            shrink = self.shrink_mode.get()
            
            if self.use_custom.get():
                try:
                    pw = float(self.custom_width.get().replace(",","."))
                    ph = float(self.custom_height.get().replace(",","."))
                    if pw <= 0 or ph <= 0:
                        raise ValueError("Размеры листа должны быть положительными")
                    if ph < 360:
                        messagebox.showerror("Ошибка", "Высота листа не может быть меньше 360 мм")
                        return
                    if pw > 1038:
                        messagebox.showerror("Ошибка", "Ширина листа не может быть больше 1038 мм")
                        return
                except ValueError:
                    messagebox.showerror("Ошибка", "Введите корректные размеры произвольного листа")
                    return
                formats_to_check = [(pw, ph)]
                selected_paper_repr = f"{pw}×{ph}"
                current_pw, current_ph = pw, ph
            else:
                selected_paper = self.selected_format.get()
                current_pw, current_ph = self.paper_formats[selected_paper]
                formats_to_check = list(self.paper_formats.keys())
                selected_paper_repr = selected_paper
            
            variants = []
            if color == "color":
                variants.append(('color', 25, 10, 4, 0))
                if shrink:
                    variants.append(('color', 22, 10, 4, 0))
                    variants.append(('color', 22, 10, 3, 0))
            else:
                variants.append(('bw', 25, 10, 0, 0))
                if shrink:
                    variants.append(('bw', 22, 10, 0, 1))
            
            best_layout = None
            best_final_count = -1
            
            for (c_mode, top_bottom, side, extra, height_red) in variants:
                layout_orig = self.inlay_calculate_layout(
                    current_pw, current_ph, h, w, rotated=False,
                    color_mode=c_mode,
                    top_bottom_margin=top_bottom,
                    side_margin=side,
                    extra_margin=extra,
                    height_reduction=height_red
                )
                layout_orig['original_count'] = layout_orig['total_count']
                layout_orig['original_width'] = layout_orig['count_width']
                layout_orig['final_count'] = self.calculate_final_count(layout_orig, ptype)
                layout_orig['variant_params'] = (top_bottom, extra, height_red)
                
                if layout_orig['final_count'] > best_final_count:
                    best_final_count = layout_orig['final_count']
                    best_layout = layout_orig
                
                if self.auto_rotate.get():
                    layout_rot = self.inlay_calculate_layout(
                        current_pw, current_ph, h, w, rotated=True,
                        color_mode=c_mode,
                        top_bottom_margin=top_bottom,
                        side_margin=side,
                        extra_margin=extra,
                        height_reduction=height_red
                    )
                    layout_rot['original_count'] = layout_rot['total_count']
                    layout_rot['original_width'] = layout_rot['count_width']
                    layout_rot['final_count'] = self.calculate_final_count(layout_rot, ptype)
                    layout_rot['variant_params'] = (top_bottom, extra, height_red)
                    
                    if layout_rot['final_count'] > best_final_count:
                        best_final_count = layout_rot['final_count']
                        best_layout = layout_rot
            
            if best_layout is None:
                messagebox.showerror("Ошибка", "Не удалось рассчитать ни одного варианта")
                return
            
            ptype_names = {
                "односторонняя": "📄 Односторонняя",
                "двусторонняя_чужой": "🔄 Двусторонняя (ЧУЖОЙ оборот)",
                "двусторонняя_свой": "⚡ Двусторонняя (СВОЙ оборот)"
            }
            ptype_name = ptype_names.get(ptype, ptype)
            color_text = "цветная" if color=="color" else "черно-белая"
            
            self.result_textbox.insert(tk.END, "="*90+"\n")
            self.result_textbox.insert(tk.END, f"📋 ВЫБРАННЫЙ ФОРМАТ: {selected_paper_repr}\n")
            self.result_textbox.insert(tk.END, f"🖨️ ТИП ПЕЧАТИ: {ptype_name}\n")
            self.result_textbox.insert(tk.END, f"🎨 ЦВЕТНОСТЬ: {color_text}\n")
            if shrink:
                self.result_textbox.insert(tk.END, f"📦 РЕЖИМ СЖАТИЯ: Включен\n")
            else:
                self.result_textbox.insert(tk.END, f"📦 РЕЖИМ СЖАТИЯ: Отключен\n")
            self.result_textbox.insert(tk.END, "="*90+"\n\n")
            
            if shrink:
                if color == "color":
                    if best_layout.get('extra_margin', 4) == 3:
                        self.result_textbox.insert(tk.END, "📌 Режим сжатия: отступ сверху/снизу уменьшен до 22 мм, припуск уменьшен до 3 мм\n")
                    else:
                        self.result_textbox.insert(tk.END, "📌 Режим сжатия: отступ сверху/снизу уменьшен до 22 мм (припуск 4 мм)\n")
                else:
                    if best_layout.get('height_reduction', 0) > 0:
                        self.result_textbox.insert(tk.END, f"📌 Режим сжатия: отступ сверху/снизу уменьшен до 22 мм, высота вкладыша уменьшена на {best_layout['height_reduction']} мм\n")
                    else:
                        self.result_textbox.insert(tk.END, "📌 Режим сжатия: отступ сверху/снизу уменьшен до 22 мм\n")
            else:
                self.result_textbox.insert(tk.END, "📌 Режим сжатия: стандартные отступы 25 мм\n")
            
            self.result_textbox.insert(tk.END, f"📐 Размер вкладыша: {h}×{w} мм\n")
            self.result_textbox.insert(tk.END, f"📏 Размер с припуском: {best_layout['inlay_h_eff']}×{best_layout['inlay_w_eff']} мм\n")
            self.result_textbox.insert(tk.END, f"📄 Доступная область: {best_layout['available_height']}×{best_layout['available_width']} мм\n")
            self.result_textbox.insert(tk.END, f"📐 Отступы: верх/низ {best_layout['top_bottom_margin']} мм, бока {best_layout['side_margin']} мм\n\n")
            
            self.result_textbox.insert(tk.END, f"{'🔄' if best_layout['rotated'] else '✓'} ОРИЕНТАЦИЯ: {best_layout['orientation']}\n")
            self.result_textbox.insert(tk.END, f"   Исходное расположение: {best_layout['count_height']} рядов × {best_layout['original_width']} колонок = {best_layout['original_count']} шт.\n")
            
            if ptype != "односторонняя":
                if ptype == "двусторонняя_свой" and best_layout['original_width']%2:
                    new_w = best_layout['original_width']-1
                    self.result_textbox.insert(tk.END, f"   ⚠️ СВОЙ оборот: нечетное количество по ширине ({best_layout['original_width']})\n")
                    self.result_textbox.insert(tk.END, f"   ➖ Убираем 1 колонку\n")
                    self.result_textbox.insert(tk.END, f"   ✅ Расположение: {best_layout['count_height']}×{new_w} = {best_layout['final_count']} шт.\n")
                elif ptype == "двусторонняя_чужой" and best_layout['original_count']%2:
                    self.result_textbox.insert(tk.END, f"   ⚠️ ЧУЖОЙ оборот: нечетное общее количество ({best_layout['original_count']})\n")
                    self.result_textbox.insert(tk.END, f"   ➖ Убираем 1 вкладыш\n")
                    self.result_textbox.insert(tk.END, f"   ✅ Количество: {best_layout['final_count']} шт.\n")
            
            self.result_textbox.insert(tk.END, f"\n✅ ИТОГ: {best_layout['final_count']} вкладышей на листе\n")
            
            total_size_info = f"{best_layout['total_height']:.1f}×{best_layout['total_width']:.1f} мм"
            self.total_size_text.set(total_size_info)
            
            if ptype == "двусторонняя_свой" and best_layout['original_width']%2:
                new_w = best_layout['original_width']-1
                layout_info = f"{'🔄' if best_layout['rotated'] else '✓'} {best_layout['orientation']} → {best_layout['count_height']}×{new_w} = {best_layout['final_count']} шт."
            elif ptype == "двусторонняя_чужой" and best_layout['original_count']%2:
                layout_info = f"{'🔄' if best_layout['rotated'] else '✓'} {best_layout['orientation']} → {best_layout['final_count']} шт. (было {best_layout['original_count']})"
            else:
                layout_info = f"{'🔄' if best_layout['rotated'] else '✓'} {best_layout['orientation']} → {best_layout['final_count']} шт."
            self.best_layout_text.set(layout_info)
            
            self.result_textbox.insert(tk.END, f"📊 Общий размер: {total_size_info}\n")
            self.result_textbox.insert(tk.END, f"📈 Эффективность: {best_layout['efficiency']:.1f}%\n\n")
            
            # СРАВНЕНИЕ ВСЕХ ФОРМАТОВ - ровная таблица
            if self.auto_calculate_all.get() and not self.use_custom.get():
                all_results = []
                for fmt in formats_to_check:
                    pw, ph = self.paper_formats[fmt]
                    best_for_fmt = None
                    best_fc = -1
                    for (c_mode, top_bottom, side, extra, height_red) in variants:
                        for rot in [False, True] if self.auto_rotate.get() else [False]:
                            layout = self.inlay_calculate_layout(
                                pw, ph, h, w, rotated=rot,
                                color_mode=c_mode,
                                top_bottom_margin=top_bottom,
                                side_margin=side,
                                extra_margin=extra,
                                height_reduction=height_red
                            )
                            layout['original_count'] = layout['total_count']
                            layout['original_width'] = layout['count_width']
                            layout['final_count'] = self.calculate_final_count(layout, ptype)
                            layout['paper_format'] = fmt
                            if layout['final_count'] > best_fc:
                                best_fc = layout['final_count']
                                best_for_fmt = layout
                    if best_for_fmt:
                        all_results.append(best_for_fmt)
                
                if all_results:
                    all_results.sort(key=lambda x: x['final_count'], reverse=True)
                    self.result_textbox.insert(tk.END, "="*90+"\n")
                    self.result_textbox.insert(tk.END, "📊 СРАВНЕНИЕ ВСЕХ ФОРМАТОВ\n")
                    self.result_textbox.insert(tk.END, "="*90+"\n\n")
                    
                    col_widths = [12, 14, 16, 11, 12]
                    header = (f"{'Формат':<{col_widths[0]}} | "
                              f"{'Ориентация':<{col_widths[1]}} | "
                              f"{'Расположение':<{col_widths[2]}} | "
                              f"{'Количество':<{col_widths[3]}} | "
                              f"{'Эффективность':<{col_widths[4]}}")
                    separator = "-" * len(header)
                    self.result_textbox.insert(tk.END, header + "\n")
                    self.result_textbox.insert(tk.END, separator + "\n")
                    
                    for r in all_results:
                        marker = "🏆 " if r == all_results[0] else "   "
                        fmt_str = f"{marker}{r['paper_format']:<{col_widths[0]-2}}"
                        orient_str = f"{'🔄' if r['rotated'] else '✓'} {r['orientation']:<{col_widths[1]-3}}"
                        if ptype == "двусторонняя_свой" and r['original_width'] % 2:
                            new_w = r['original_width'] - 1
                            lay_str = f"{r['count_height']}×{new_w:<{col_widths[2]-3}}"
                        else:
                            lay_str = f"{r['count_height']}×{r['count_width']:<{col_widths[2]-3}}"
                        cnt_str = f"{r['final_count']:<{col_widths[3]}}"
                        if r['final_count'] != r['original_count']:
                            cnt_str = f"{r['final_count']} ({r['original_count']})"
                        eff_str = f"{r['efficiency']:.1f}%"
                        row = (f"{fmt_str:<{col_widths[0]}} | "
                               f"{orient_str:<{col_widths[1]}} | "
                               f"{lay_str:<{col_widths[2]}} | "
                               f"{cnt_str:<{col_widths[3]}} | "
                               f"{eff_str:<{col_widths[4]}}")
                        self.result_textbox.insert(tk.END, row + "\n")
                    
                    self.result_textbox.insert(tk.END, "\n")
                    
                    best = all_results[0]
                    self.result_textbox.insert(tk.END, "✨ " + "★"*86 + " ✨\n")
                    self.result_textbox.insert(tk.END, f"   🏆 ЛУЧШИЙ ФОРМАТ: {best['paper_format']}\n")
                    self.result_textbox.insert(tk.END, f"   📦 Количество: {best['final_count']} вкладышей на листе\n")
                    self.result_textbox.insert(tk.END, f"   🔄 Ориентация: {'ПЕРЕВЕРНУТО' if best['rotated'] else 'ОРИГИНАЛЬНАЯ'} ({best['orientation']})\n")
                    if ptype == "двусторонняя_свой" and best['original_width'] % 2:
                        new_w = best['original_width'] - 1
                        self.result_textbox.insert(tk.END, f"   📐 Расположение: {best['count_height']}×{best['original_width']} → {best['count_height']}×{new_w}\n")
                    else:
                        self.result_textbox.insert(tk.END, f"   📐 Расположение: {best['count_height']}×{best['count_width']}\n")
                    self.result_textbox.insert(tk.END, f"   📊 Эффективность: {best['efficiency']:.1f}%\n")
                    self.result_textbox.insert(tk.END, "✨ " + "★"*86 + " ✨\n")
                    
                    self.best_paper_text.set(f"{best['paper_format']} → {best['final_count']} шт. ({best['efficiency']:.1f}%)")
                else:
                    self.best_paper_text.set("Лучший формат: -")
            elif self.use_custom.get():
                self.best_paper_text.set("Произвольный формат")
            else:
                self.best_paper_text.set("Лучший формат: расчет отключен")
            
            self.result_textbox.see(1.0)
            
        except ValueError as e:
            messagebox.showerror("Ошибка", f"Некорректный ввод: {e}")
            self.total_size_text.set("Общий размер вкладышей: -")
            self.best_layout_text.set("Оптимальное расположение: -")
            self.best_paper_text.set("Лучший формат: -")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")
            self.total_size_text.set("Общий размер вкладышей: -")
            self.best_layout_text.set("Оптимальное расположение: -")
            self.best_paper_text.set("Лучший формат: -")
    
    # ===== ВКЛАДКА "% выгоды" =====
    def create_percentage_calculator_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text='Расчет % выгоды')
        self.entry1 = tk.StringVar()
        self.entry2 = tk.StringVar()
        self.entry3 = tk.StringVar()
        self.entry4 = tk.StringVar()
        self.result_percent = tk.StringVar(value="Результат: ")
        self.setup_percent_ui(tab)
    
    def setup_percent_ui(self, parent):
        f = ttk.Frame(parent, padding="20")
        f.pack(fill='both', expand=True)
        f.columnconfigure(0, weight=1); f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Калькулятор % выгоды", font=("Arial",14,"bold")).grid(row=0, column=0, columnspan=2, pady=(0,20))
        ttk.Label(f, text="высота листа №1: (с роля)").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(f, textvariable=self.entry1, width=20).grid(row=2, column=0, padx=5, pady=(0,10), sticky=tk.W+tk.E)
        ttk.Label(f, text="ширина листа №1:").grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(f, textvariable=self.entry2, width=20).grid(row=2, column=1, padx=5, pady=(0,10), sticky=tk.W+tk.E)
        ttk.Label(f, text="высота листа №2: (с роля)").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(f, textvariable=self.entry3, width=20).grid(row=4, column=0, padx=5, pady=(0,10), sticky=tk.W+tk.E)
        ttk.Label(f, text="ширина листа №2:").grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Entry(f, textvariable=self.entry4, width=20).grid(row=4, column=1, padx=5, pady=(0,10), sticky=tk.W+tk.E)
        ttk.Button(f, text="Расчет процента выгоды", command=self.percent_calculate).grid(row=5, column=0, columnspan=2, pady=20)
        self.result_percent_label = ttk.Label(f, textvariable=self.result_percent, font=("Arial",12,"bold"))
        self.result_percent_label.grid(row=6, column=0, columnspan=2, pady=10)
        # Контекстное меню для результата (чтобы копировать)
        self.create_context_menu(self.result_percent_label)
    
    def percent_calculate(self):
        try:
            v1 = float(self.entry1.get().replace(',','.'))
            v2 = float(self.entry2.get().replace(',','.'))
            v3 = float(self.entry3.get().replace(',','.'))
            v4 = float(self.entry4.get().replace(',','.'))
            first = v1*v2
            second = v3*v4
            if first > second:
                diff = first - second
                res = (diff/first)*100
                m = f'второй вариант {v3} x {v4}'
            else:
                diff = second - first
                res = (diff/second)*100
                m = f'первый вариант {v1} x {v2}'
            self.result_percent.set(f"На: {res:.2f} % выгоднее " + m)
        except ValueError:
            self.result_percent.set("Ошибка! Введите числа")
    
    # ===== ВКЛАДКА "Расчет тиража" =====
    def create_edition_calculator_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text='Расчет тиража')
        self.edition_input = tk.StringVar()
        self.edition_input2 = tk.StringVar()
        self.edition_result = tk.StringVar(value="Минимальное количество пачек: -")
        self.setup_edition_ui(tab)
    
    def setup_edition_ui(self, parent):
        f = ttk.Frame(parent, padding="30")
        f.pack(fill='both', expand=True)
        ttk.Label(f, text="Калькулятор тиража", font=("Arial",14,"bold")).pack(pady=(0,20))
        ttk.Label(f, text="Проверка раскладки", font=("Arial",14,"bold")).pack(pady=(0,1))
        ttk.Label(f, text="Расчет минимального количества пачек на тираж \nРезультат округляется до ЧЕТНОГО числа В БОЛЬШУЮ сторону",font=("Arial",9), foreground="gray", justify=tk.CENTER).pack(pady=(0,20))
        inp = ttk.Frame(f)
        inp.pack(pady=10)

        ttk.Label(inp, text="Введите общий тираж (кол-во пачек):", font=("Arial",10)).pack(side=tk.LEFT, padx=(0,10))
        entry = ttk.Entry(inp, textvariable=self.edition_input, width=15, font=("Arial",11))
        entry.pack(side=tk.LEFT) 

        inp2 = ttk.Frame(f)
        inp2.pack(pady=1)
    
        ttk.Label(inp2, text="Введите кол-во знаков на листе:       ", font=("Arial",10)).pack(side=tk.LEFT, padx=(0,10))
        entry2 = ttk.Entry(inp2, textvariable=self.edition_input2, width=15, font=("Arial",11))
        entry2.pack(side=tk.LEFT)
        
        def validate_edition(char): return char.isdigit() or char == ""
        entry.config(validate="key", validatecommand=(entry.register(validate_edition), '%P'))
        ttk.Button(f, text="🧮 Рассчитать", command=self.edition_calculate, width=20).pack(pady=20)
        res_frame = ttk.LabelFrame(f, text="Результат расчета", padding=15)
        res_frame.pack(fill="x", pady=10)
        self.edition_result_label = ttk.Label(res_frame, textvariable=self.edition_result, font=("Arial",12,"bold"), foreground="blue")
        self.edition_result_label.pack()
        # Контекстное меню для результата
        self.create_context_menu(self.edition_result_label)
        info = ttk.Frame(f)
        info.pack(pady=(20,0))
    
    def edition_calculate(self):
        try:
            txt = self.edition_input.get().strip()
            txt2 = self.edition_input2.get().strip()
            if not txt or not txt2:
                self.edition_result.set("⚠️ Введите значение")
                return
            edition = int(txt)
            edition2 = int(txt2)
            if edition <= 0 or edition2 <= 0:
                self.edition_result.set("⚠️ Значение должно быть больше 0")
                return
            packs_needed = edition / 5000
            packs_int = math.ceil(packs_needed)
            orig = packs_int
            if packs_int % 2 != 0:
                packs_int += 1
            sheets = math.ceil(edition / packs_int)
            total_packs = sheets * packs_int
            result = f"📦 Минимальное количество пачек: {packs_int} шт. на лист"
            dobavlenie = packs_int - edition2
            if edition2 < packs_int:
                result += f"\n!!! слишком мало пачек, добавьте минимум +{dobavlenie} шт !!!" 
            else:
                result += f"\nКоличество пачек верное"   

            result += f"\n\n📦 Тираж: {edition:,} пачек".replace(",", " ")
            result += f"\n📄 Листов: {sheets:,} штук".replace(",", " ")
            result += f"\n✅ Итого : {total_packs:,} пачек".replace(",", " ")
            self.edition_result.set(result)
        except ValueError:
            self.edition_result.set("❌ Ошибка! Введите целое число")
        except Exception as e:
            self.edition_result.set(f"❌ Ошибка: {e}")

def main():
    root = tk.Tk()
    app = CombinedCalculatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    print("="*90)
    print("Калькулятор by Kirill V14")
    print("="*90)
    main()