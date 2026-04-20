import time
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import datetime
import os
import getpass
import socket
import ctypes
import segno
from openpyxl.drawing.image import Image
import xlwings as xw
import pyodbc
import logging
import subprocess
import sys


def to_eng():
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
        current = ctypes.windll.user32.GetKeyboardLayout(tid) & 0xFFFF

        max_attempts = 9
        attempts = 0

        while (current & 0x00FF) != 0x09 and attempts < max_attempts:
            ctypes.windll.user32.PostMessageW(hwnd, 0x0050, 0, 0)
            time.sleep(0.05)

            hwnd = ctypes.windll.user32.GetForegroundWindow()
            tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
            current = ctypes.windll.user32.GetKeyboardLayout(tid) & 0xFFFF
            attempts += 1

        return (current & 0x00FF) == 0x09
    except:
        return False


def is_eng():
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
        layout = ctypes.windll.user32.GetKeyboardLayout(thread_id)
        layout_id = layout & 0xFFFF

        # английская раскладка - код 0x0409 или 0x09
        return (layout_id & 0x00FF) == 0x09 or layout_id == 0x0409
    except:
        return False


# логирование
def setup_logging():
    os.makedirs('logs/', exist_ok=True)
    logging.basicConfig(
        filename='logs/logs.log',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )


setup_logging()

qr_path = r'QRCODES\\'

# внешний вид
ctk.set_default_color_theme('dark-blue')

# переключение на en с проверкой
if not is_eng():
    to_eng()


class ThemeColors:
    def __init__(self, mode='dark'):
        self.update(mode)

    def update(self, mode):
        self.border_color_base = '#808080'
        self.border_color_green = '#2E8B57'
        self.border_color_yellow = '#DAA520'
        self.border_color_red = '#DC143C'
        self.notification_color_green = '#2E8B57'
        self.notification_color_yellow = '#DAA520'
        self.notification_color_red = '#800000'

        if mode == 'light':
            self.fg_color_disable = '#e6e6e6'
            self.fg_color_enable = '#ffffff'
        else:
            self.fg_color_disable = '#202121'
            self.fg_color_enable = '#343536'


class DatabaseManager:
    def __init__(self):
        self.SQL_SERVER = 'сервер'
        self.SQL_DB = 'база данных (не таблица)'
        self.SQL_USER = 'юзер бд'
        self.SQL_PASSWORD = 'пароль'
        self.conn_str = f'DRIVER={{SQL Server}};SERVER={self.SQL_SERVER};DATABASE={self.SQL_DB};UID={self.SQL_USER};PWD={self.SQL_PASSWORD}'

    def check_connection(self):
        try:
            conn = pyodbc.connect(self.conn_str)
            conn.close()
            return True
        except pyodbc.Error as e:
            logging.error(f'Ошибка проверки подключения к бд: {e}')
            return False

    def get_connection(self):
        try:
            return pyodbc.connect(self.conn_str)
        except pyodbc.Error as e:
            logging.error(f'Ошибка подключения к бд: {e}')
            return None

    def execute_query(self, query, params=None, fetch_one=False, fetch_all=False, commit=False):
        conn = None
        cursor = None
        try:
            conn = self.get_connection()
            if not conn:
                return None
            cursor = conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            if commit:
                conn.commit()
                return True
            if fetch_one:
                return cursor.fetchone()
            elif fetch_all:
                return cursor.fetchall()
            return None
        except pyodbc.Error as e:
            return None
        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    def add_record(self, barcode, excise, user_name, computer_name, qr_name, created_date=None):

        query = '''
            INSERT INTO barcodb (barcode, excise, user_name, computer_name, qr_name, created_date)
            VALUES (?, ?, ?, ?, ?, ?)
        '''

        params = (barcode, excise, user_name, computer_name, qr_name, created_date)
        return self.execute_query(query, params, commit=True)

    def get_data(self):
        query = 'SELECT * FROM barcodb order by created_date'
        return self.execute_query(query, fetch_all=True)

    def check_exists(self, excise):
        query = 'SELECT 1 FROM barcodb WHERE excise = ?'
        result = self.execute_query(query, params=(excise,), fetch_one=True)

        return 1 if result else 0


class ReportGenerator:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title('Scanner')
        self.db = DatabaseManager()
        self.colors = ThemeColors()
        self.appearance_mode = 'dark'
        ctk.set_appearance_mode(self.appearance_mode)

        # задаем размеры окна
        self.width = 650
        self.height = 300
        self.root.geometry(f'{self.width}x{self.height}')

        self.root.resizable(True, False)
        self._notification_timer = None

        # таймеры для полей
        self.barcode_timer = None
        self.excise_timer = None
        self.scanner_timeout = 200

        # ждем применения геометрии
        self.root.update_idletasks()

        # получаем реальные размеры окна после применения масштабирования
        actual_width = self.root.winfo_width()

        # находим dpi scaler
        scaler = round(actual_width / self.width, 2)

        # получаем размеры экрана
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # вычисляем позицию для центрирования
        x = (screen_width - self.width) // 2
        y = (screen_height - self.height) // 2

        # устанавливаем окончательную позицию
        self.root.geometry(f'+{round(x * scaler)}+{round(y * scaler)}')

        self.create_widgets()

        self.root.after(100, self.entry_barcode.focus)

    def create_widgets(self):
        # основной контейнер
        main_container = ctk.CTkFrame(self.root, fg_color='transparent')
        main_container.pack(fill='both', expand=True, padx=30, pady=10)

        # верхняя панель с индикатором и тогглером
        top_bar = ctk.CTkFrame(main_container, fg_color='transparent')
        top_bar.pack(fill='x', pady=(0, 5))

        # индикатор подключения к БД
        self.connection_indicator = ctk.CTkLabel(
            top_bar,
            text='●',
            font=ctk.CTkFont(size=17)
        )
        self.connection_indicator.pack(side='left', padx=10)
        self.update_connection_indicator()

        # свитчер переключения темы
        self.theme_switch = ctk.CTkSwitch(
            top_bar,
            text='Сменить тему',
            command=self.toggle_theme,
            font=ctk.CTkFont(size=15)
        )
        self.theme_switch.pack(side='right', padx=0)
        self.theme_switch.select() if self.appearance_mode == 'dark' else self.theme_switch.deselect()

        # рамка для баркода
        self.barcode_frame = ctk.CTkFrame(main_container, border_width=2, border_color=self.colors.border_color_base,
                                          corner_radius=10)
        self.barcode_frame.pack(fill='x', pady=(0, 20))

        label_barcode = ctk.CTkLabel(self.barcode_frame, text='Баркод:', font=ctk.CTkFont(size=17, weight='bold'))
        label_barcode.pack(side='left', padx=(20, 10), pady=20)

        self.entry_barcode = ctk.CTkEntry(self.barcode_frame, width=300, height=35,
                                          font=ctk.CTkFont(size=20, weight='bold'))
        self.entry_barcode.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # рамка для акциза
        self.excise_frame = ctk.CTkFrame(main_container, border_width=2, border_color=self.colors.border_color_base,
                                         corner_radius=10)
        self.excise_frame.pack(fill='x', pady=(0, 10))

        label_excise = ctk.CTkLabel(self.excise_frame, text='Акциз:  ', font=ctk.CTkFont(size=17, weight='bold'))
        label_excise.pack(side='left', padx=(20, 10), pady=20)

        self.entry_excise = ctk.CTkEntry(self.excise_frame, width=300, height=35, state='disabled',
                                         fg_color=self.colors.fg_color_disable,
                                         font=ctk.CTkFont(size=16, weight='bold'))
        self.entry_excise.pack(side='left', padx=(0, 20), pady=20, fill='x', expand=True)

        # кнопка формирования отчёта
        self.button_generate = ctk.CTkButton(
            main_container,
            text='Сформировать отчет',
            command=self.generate_report,
            width=200,
            height=45,
            fg_color=self.colors.border_color_green,
            text_color='white',
            font=ctk.CTkFont(size=17, weight='bold'),
            hover_color='#3CB371'
        )
        self.button_generate.pack(side='right')

        # метка для уведомлений
        self.notification_label = ctk.CTkLabel(
            main_container,
            text='',
            font=ctk.CTkFont(size=17),
            fg_color=self.colors.border_color_green,
            text_color='white',
            corner_radius=6,
            padx=25,
            pady=15
        )

        # привязываем события ввода
        self.entry_barcode.bind('<KeyRelease>', self.on_barcode_change)

    def toggle_theme(self):
        if self.appearance_mode == 'dark':
            self.appearance_mode = 'light'
            ctk.set_appearance_mode('light')
        else:
            self.appearance_mode = 'dark'
            ctk.set_appearance_mode('dark')

        # обновляем цвета в классе ThemeColors
        self.colors.update(self.appearance_mode)

        # обновляем состояние полей
        if self.entry_excise.cget('state') == 'disabled':
            self.entry_excise.configure(fg_color=self.colors.fg_color_disable)
        else:
            self.entry_excise.configure(fg_color=self.colors.fg_color_enable)

    def update_connection_indicator(self):
        (color, state) = ('#2E8B57', 'connected') if self.db.check_connection() else ('#DC143C', 'disconnected')
        self.connection_indicator.configure(text_color=color, text=f'●  {state}')

    def send_data(self, barcode, excise):
        try:
            # проверка уникальности
            if not self.db.check_connection():
                self.update_connection_indicator()
                self.show_notification(f'Нет соединения с БД', label_bg=self.colors.notification_color_red)
                return False
            if self.db.check_exists(excise):
                self.show_notification(f'Данный акциз уже существует', label_bg=self.colors.notification_color_red)
                return 'exist'

            # получаем данные о пользователе и компьютере
            user_name = getpass.getuser()
            computer_name = socket.gethostname()
            created_date = datetime.datetime.now()

            try:
                # проверяем существует ли папка
                if not os.path.exists(qr_path):
                    os.makedirs(qr_path, exist_ok=True)
                else:
                    pass

                # проверяем доступ на запись
                test_file = os.path.join(qr_path, 'test.txt')
                with open(test_file, 'w') as f:
                    f.write('test')
                os.remove(test_file)
            except Exception as e:
                logging.error(f'Ошибка доступа к папке: {e}')
                self.show_notification(f'Нет доступа к папке с QR', label_bg=self.colors.notification_color_red)
                return False

            # генерируем qr
            try:
                os.makedirs(qr_path, exist_ok=True)
            except Exception as e:
                logging.error(f'Ошибка при создании папки для QR: {e}')

            qr_name = round(time.time())
            qr = segno.make(str(excise), error='h')
            qr.save(rf'{qr_path}{qr_name}.png',
                    scale=10,  # размер модуля в пикселях
                    border=2,  # отступ вокруг QR
                    dark='black',  # цвет темных модулей
                    light='white')  # цвет светлых модулей

            # отправляем insert
            if self.db.add_record(barcode, excise, user_name, computer_name, qr_name, created_date):
                self.show_notification(f'Данные успешно добавлены', label_bg=self.colors.notification_color_green)
            else:
                self.show_notification(f'Ошибка в insert запросе', label_bg=self.colors.notification_color_red)

            return True
        except Exception as e:
            logging.error(f'Ошибка в отправке данных: {e}')
            self.show_notification(f'Ошибка в отправке данных', label_bg=self.colors.notification_color_red)
            return False

    def on_barcode_change(self, event=None):
        # отменяем предыдущий таймер
        if self.barcode_timer:
            self.root.after_cancel(self.barcode_timer)

        barcode = self.entry_barcode.get()

        if not is_eng():
            self.show_notification(f'Неверная раскладка. Переключите на EN',
                                   label_bg=self.colors.notification_color_red,
                                   duration=1500)
            self.entry_barcode.delete(0, tk.END)
            return

        # ждем 100мс после последнего символа перед проверкой
        self.barcode_timer = self.root.after(self.scanner_timeout, lambda: self.check_barcode(barcode))

    def check_barcode(self, barcode):
        # проверка баркода
        barcode_len = len(barcode)

        if barcode_len == 0:
            self.barcode_frame.configure(border_color=self.colors.border_color_base)
            self.entry_excise.delete(0, tk.END)
            self.entry_excise.configure(fg_color=self.colors.fg_color_disable)
            self.entry_excise.configure(state='disabled')
            self.excise_frame.configure(border_color=self.colors.border_color_base)

        elif (barcode_len == 13 or barcode_len == 12) and barcode.isdigit():
            if barcode_len == 12:
                self.entry_barcode.insert(0, 0)
            self.barcode_frame.configure(border_color=self.colors.border_color_green)
            self.entry_excise.configure(state='normal')
            self.entry_excise.configure(fg_color=self.colors.fg_color_enable)
            self.entry_excise.focus()
            self.excise_frame.configure(border_color=self.colors.border_color_yellow)

            if not hasattr(self, '_excise_bound'):
                self.entry_excise.bind('<KeyRelease>', self.on_excise_change)
                self._excise_bound = True
        else:

            self.barcode_frame.configure(border_color=self.colors.border_color_red)
            self.entry_excise.delete(0, tk.END)
            self.entry_excise.configure(fg_color=self.colors.fg_color_disable)
            self.entry_excise.configure(state='disabled')
            self.excise_frame.configure(border_color=self.colors.border_color_base)

            # отображение ошибки после конца ввода
            if barcode_len > 0:
                if not barcode.isdigit():
                    self.entry_barcode.delete(0, tk.END)
                    self.show_notification(f'В EAN должны быть только цифры',
                                           label_bg=self.colors.notification_color_red)
                elif barcode_len != 13:
                    self.entry_barcode.delete(0, tk.END)
                    self.show_notification(f'Неверный EAN (длина {barcode_len} из 13)',
                                           label_bg=self.colors.notification_color_red)

    def show_notification(self, message, duration=3000, label_bg=None):
        if label_bg is None:
            label_bg = self.colors.notification_color_green

        if self._notification_timer is not None:
            try:
                self.root.after_cancel(self._notification_timer)
            except:
                pass
            self._notification_timer = None

        self.hide_notification()

        self.notification_label.configure(text=message, fg_color=label_bg)
        self.notification_label.pack(side='left', pady=(5, 0))

        self._notification_timer = self.root.after(duration, self.hide_notification)

    def hide_notification(self):
        self.notification_label.pack_forget()
        self.notification_label.configure(text='')

        if self._notification_timer is not None:
            try:
                self.root.after_cancel(self._notification_timer)
            except:
                pass
            self._notification_timer = None

    def on_excise_change(self, event=None):
        # отменяем предыдущий таймер
        if self.excise_timer:
            self.root.after_cancel(self.excise_timer)

        excise = self.entry_excise.get()
        barcode = self.entry_barcode.get()

        if not is_eng():
            self.show_notification(f'Неверная раскладка. Переключите на EN',
                                   label_bg=self.colors.notification_color_red,
                                   duration=1500)
            self.entry_excise.delete(0, tk.END)
            return

        # ждем 100мс после последнего символа перед проверкой
        self.excise_timer = self.root.after(self.scanner_timeout, lambda: self.check_excise(barcode, excise))

    def check_excise(self, barcode, excise):
        # проверка акциза после задержки
        excise_len = len(excise)
        barcode_len = len(barcode)

        if barcode_len == 0:
            self.excise_frame.configure(border_color=self.colors.border_color_base)
            self.entry_excise.delete(0, tk.END)
            self.entry_excise.configure(state='disabled')
            self.entry_barcode.focus()

        elif excise_len == 0:
            self.excise_frame.configure(border_color=self.colors.border_color_yellow)

        elif (excise_len == 13 or excise_len == 12) and excise.isdigit():
            self.entry_barcode.delete(0, tk.END)
            if excise_len == 12:
                self.entry_barcode.insert(0, f'{0}{excise}')
            else:
                self.entry_barcode.insert(0, excise)
            self.entry_excise.delete(0, tk.END)
            self.excise_frame.configure(border_color=self.colors.border_color_yellow)
            self.show_notification(f'Баркод обновлён', label_bg=self.colors.notification_color_yellow)

        elif 0 < excise_len < 150:
            self.entry_excise.delete(0, tk.END)
            self.excise_frame.configure(border_color=self.colors.border_color_red)
            self.show_notification(f'Неверный акциз (длина {excise_len} из 150)',
                                   label_bg=self.colors.notification_color_red)

        else:  # excise_len >= 150
            self.excise_frame.configure(border_color=self.colors.border_color_green)

            if barcode_len == 13 and barcode.isdigit():
                # отправляем данные
                code = self.send_data(barcode, excise)

                # очищаем поля
                if code == 'exist':
                    self.entry_excise.delete(0, tk.END)
                    self.excise_frame.configure(border_color=self.colors.border_color_red)
                    self.entry_excise.focus()
                else:
                    self.entry_barcode.delete(0, tk.END)
                    self.entry_excise.delete(0, tk.END)

                    # деактивируем поле акциза
                    self.entry_excise.configure(state='disabled')
                    self.entry_excise.configure(fg_color=self.colors.fg_color_disable)

                    # сбрасываем цвета рамок
                    self.barcode_frame.configure(border_color=self.colors.border_color_base)
                    self.excise_frame.configure(border_color=self.colors.border_color_base)

                    # фокус на баркод
                    self.entry_barcode.focus()

    def generate_report(self):
        try:
            if not self.db.check_connection():
                self.update_connection_indicator()
                return self.show_notification(f'Нет соединения с БД', label_bg=self.colors.notification_color_red)

            data = []
            if self.db.check_connection():
                all_data = self.db.get_data()
                if all_data:
                    for row in all_data:
                        data.append(row)

            report_time = time.time()
            if not os.path.exists(r'reports\\'):
                os.makedirs(r'reports\\', exist_ok=True)
            filename = os.path.abspath(fr'reports\report_{round(report_time)}.xlsx')

            # создаем новую книгу Excel
            wb = xw.Book()
            ws = wb.sheets[0]
            ws.name = 'report'

            # заголовки
            headers = ['Баркод', 'Акциз', 'QR Акциза', 'Путь к QR', 'Компьютер', 'Пользователь', 'Дата добавления']
            for col, header in enumerate(headers, 1):
                ws.range((1, col)).value = header
                ws.range((1, col)).font.bold = True
                # центрируем заголовки по горизонтали
                ws.range((1, col)).api.HorizontalAlignment = -4108

            # устанавливаем текстовый формат для колонок A и B
            ws.range('A:A').api.NumberFormat = '@'
            ws.range('B:B').api.NumberFormat = '@'
            ws.range('D:D').api.NumberFormat = '@'
            ws.range('E:E').api.NumberFormat = '@'
            ws.range('F:F').api.NumberFormat = '@'
            ws.range('G:G').api.NumberFormat = '@'

            # центрируем столбец A (баркод) по горизонтали
            ws.range('A:A').api.HorizontalAlignment = -4108
            ws.range('D:D').api.HorizontalAlignment = -4108
            ws.range('E:E').api.HorizontalAlignment = -4108
            ws.range('F:F').api.HorizontalAlignment = -4108
            ws.range('G:G').api.HorizontalAlignment = -4108

            # включаем перенос текста и уменьшение шрифта для колонки B
            ws.range('B:B').api.WrapText = True
            ws.range('B:B').api.ShrinkToFit = True
            ws.range('D:D').api.WrapText = True

            # заполняем данные и вставляем QR
            for idx, row in enumerate(data, start=2):
                barcode, excise, user_name, computer_name, qr_name, created_date = row[1:]

                # преобразуем в абсолютный путь
                absolute_qr_path = os.path.abspath(qr_path + str(qr_name).strip() + '.png')

                # записываем данные
                ws.range((idx, 1)).value = str(barcode)
                ws.range((idx, 2)).value = str(excise)
                ws.range((idx, 4)).value = absolute_qr_path
                ws.range((idx, 5)).value = str(computer_name)
                ws.range((idx, 6)).value = str(user_name)
                ws.range((idx, 7)).value = str(created_date).split('.')[0]

                # вставляем готовый QR-код
                if os.path.exists(absolute_qr_path):
                    cell = ws.range((idx, 3))

                    pic = ws.pictures.add(absolute_qr_path,
                                          left=cell.left + 4,
                                          top=cell.top + 2,
                                          width=77,
                                          height=77)
                    ws.range((idx, 3)).row_height = 80
                else:
                    ws.range((idx, 3)).value = 'QR не найден'

            # настраиваем ширину колонок
            ws.range('A:A').column_width = 20
            ws.range('B:B').column_width = 50
            ws.range('C:C').column_width = 15
            ws.range('D:D').column_width = 50
            ws.range('E:E').column_width = 15
            ws.range('F:F').column_width = 15
            ws.range('G:G').column_width = 20

            # сохраняем и закрываем
            wb.save(filename)
            wb.close()

            messagebox.showinfo('Success', f'Отчет сохранен как {filename}')
            return filename
        except Exception as e:
            logging.error(f'Ошибка при создании отчета: {e}')
            messagebox.showerror('Error', f'Ошибка при создании отчета: {e}')

    def run(self):
        self.root.mainloop()


if __name__ == '__main__':
    app = ReportGenerator()
    app.run()
