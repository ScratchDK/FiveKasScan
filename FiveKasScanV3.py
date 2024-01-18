import pandas as pand
import PyPDF2
import simpleaudio
from datetime import date, timedelta
#import openpyxl
import time
import xlsxwriter
import keyboard   # Управление клавиатурой
#import re
#import py_win_keyboard_layout

from tkinter import *
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import messagebox
from tkinter import ttk

from ctypes import *   # Блок: клава, мышь

keyboard.block_key('Tab')

pand.set_option('display.max_rows', None)   # Ограничения на показ количества столбцов в DataFrame
pand.set_option('display.max_columns', None)   # Ограничения на показ количества столбцов в DataFrame

start_work = time.time()

leftovers = []   # Остатки ШПИ
found = []   # Найденные ШПИ
not_found = []   # Не найденные ШПИ
numbers_used = []   # Уже отсканированные ШПИ

get_switch = 0   # Переключатель рестров (один, все)

switch_reg = 1

next = 0

count_found_all = 0   # Счетчик совпадений ШПИ во всех реестрах
count_found_one = 0
count_leftovers_all = 0   # Счетчик остатков ШПИ во всех реестрах
count_leftovers_one = 0

count_f = 0

absentee_counter = 0

qr = 0

status_barcode = ''

barcode_qr = ''
#_______________________________________________________________________________________________________________________


# Классы
class ChangeSettings:

    def about_prog(self):
        text_info = '''"FiveKasScan"
Версия программы: 3.011724;
Разработчик: Коваль Дмитрий Владимирович;
Программа создана специально для 
Пятого кассационного суда!'''

        messagebox.showinfo("О программе", text_info)

    def saveSettings(self):

        clean()

        text_search = self.text_search.get()
        text_range = self.text_range.get()
        text_select = self.selected()
        text_split = text_range.split(',')

        text_split.insert(0, text_search)
        text_split.insert(1, text_select)

        messagebox.showinfo('Поиск', f'Колонка поиска "{text_search}" сохранена')

        with open('SettingsV2.txt', 'w') as fileRecord:
            for i in text_split:
                i = str(i).replace(' ', '')
                i += '\n'
                fileRecord.write(i)

        openSettings()


    def selected(self):

        self.select = openSettings()
        self.select_scan = self.select[1]

        self.status = self.switch.get()

        if self.status == 0:
            self.status = self.select_scan
        else:
            self.status = self.switch.get()

        return self.status


    def settings(self):

        self.window_settings = Toplevel()
        self.window_settings.attributes("-topmost", True)
        self.window_settings.iconbitmap('Minenergo_logo.ico')
        self.window_settings.resizable(True, True)  # Запрещает разворачивать окно программы на весь экран если False
        self.window_settings.title('Настройки')

        self.geometry = {'padx': 10, 'pady': 10}  # , 'fill': tkinter.BOTH


        self.switch = IntVar()
        self.switch.set(self.selected())

        self.rad = Radiobutton(self.window_settings, text=f'14-значный сканер\n'
                                                          f'Пример: 361624-86-18100-2', command=self.selected, value=1,
                               variable=self.switch)
        self.rad.grid(column=0, row=2, **self.geometry)

        self.rad2 = Radiobutton(self.window_settings, text=f'13-значный сканер\n'
                                                           f'Пример: 361624-86-18100', command=self.selected, value=2,
                                variable=self.switch)
        self.rad2.grid(column=1, row=2, **self.geometry)


        label = Label(self.window_settings,
                      text=f'Введите в поле справа\n'
                           f'точное название колонки\n'
                           f'по которой вы будете вести поиск',
                      font=("Times New Roman", 10), fg="black")
        label.grid(column=0, row=0, **geometry)

        label2 = Label(self.window_settings,
                      text=f'Введите в поле справа, через ","\n'
                           f'минимальный и максимальный диапазон строк\n'
                           f'в которых может находиться название колонки\n'
                           f'по которой вы будете вести поиск',
                      font=("Times New Roman", 10), fg="black")
        label2.grid(column=0, row=1, **geometry)

        self.text_range = Entry(self.window_settings, width=15, state='normal')
        self.text_range.grid(column=1, row=1, **self.geometry)

        self.text_search = Entry(self.window_settings, width=15,
                                 state='normal')  # state - доступно поле для ввода или нет
        self.text_search.grid(column=1, row=0, **self.geometry)


        label_space = Label(self.window_settings, font=("Times New Roman", 10), fg="black")
        label_space.grid(column=1, row=3)


        button1 = Button(self.window_settings, command=self.about_prog, text="О программе",  bg="white", fg="black")
        button1.grid(column=1, row=4, **self.geometry)

        button2 = Button(self.window_settings, command=self.saveSettings, text="Сохранить настройки", bg="white", fg="black")
        button2.grid(column=0, row=4, **self.geometry)

        self.window_settings.mainloop()


class GetData:

    global leftovers
    global found
    global type_scan

    def document_processing(self):

        self.skip_count += 1

        self.column_seach = self.skip[0]

        if self.column_seach == 'Бандероль':
            self.column_seach = 'Идентификатор'
        else:
            pass

        select_scan = self.skip[1]

        str_search = self.column_seach

        for self.file in self.files:

            list_items = []
            list_name = []

            self.file = self.file.split('/')
            self.file = '\\\\'.join(self.file)

            list_items.append(f'...{self.file[-27:]}')  # Название документа перед столбцем

            try:
                find_index = pand.read_excel(self.file, skiprows=self.skip_count)
            except ValueError:
                continue

            try:
                for i in find_index:
                    c = i.lower()
                    if self.column_seach.lower() in i or self.column_seach.lower() in c:
                        str_search = i

                list = find_index[str_search].tolist()

                for i in list:
                    len_i = len(str(i))

                    if select_scan == '2':

                        if len_i == 17 or len_i == 14:
                            i = str(i).replace(' ', '')
                            i = i[:-1]
                        try:
                            i = int(i)
                        except ValueError:
                            pass
                        else:
                            list_items.append(str(i))

                    elif select_scan == '1':

                        if len_i == 17:
                            i = str(i).replace(' ', '')
                        try:
                            i = int(i)
                        except ValueError:
                            pass
                        else:
                            list_items.append(str(i))


                if list_items[0] in leftovers:
                    pass
                else:
                    leftovers.append(list_items)
                    list_name.append(list_items[0])
                    found.append(list_name)

            except UnboundLocalError:
                self.skip_count += 1
                return

            except KeyError:
                continue

            except FileNotFoundError:  # Если был закрыт проводник до выбора файла
                pass

    def openFilesExcel(self):

        self.skip = openSettings()
        print(f'excel = {self.skip}')

        ask_user = messagebox.askyesno(title="Формирование документа", message="При нажатии кнопки 'Да', будет нельзя\n"
                                                                               "продолжить работу с данным реестром!")

        if ask_user == False:
            return

        clean()

        global count_leftovers_all

        self.files = filedialog.askopenfilenames()   # Получаем все пути к выбраным файлам

        start_open = time.time()

        self.skip_count = self.skip[2]   # Счетчик пропуска строки

        while self.skip_count < self.skip[3]:

            self.document_processing()

        scroll_all.config(state=NORMAL)

        for element in leftovers:
            count_leftovers_all += int(len(element) - 1)
            for i in element:
                scroll_all.insert(INSERT, '\n' + i)
            scroll_all.insert(INSERT, '\n' + '_' * 35)
            scroll_all.insert(INSERT, '\n \n')

        scroll_all.config(state=DISABLED)

        label2.config(text=f'ОСТАЛОСЬ: {count_leftovers_all}')

        end_open = time.time() - start_open
        messagebox.showinfo('Загрузка', f'Загруженные документы: "{len(leftovers)}"\n'
                                        f'Время загрузки документов: {round(end_open, 2)} ')

    # Добавлено позже
    def openFilesPDF(self):

        global count_leftovers_all
        global switch_reg
        global next

        search = openSettings()
        print(f'PDF = {search}')
        column_seach = search[0]
        select_scan = search[1]

        if column_seach == 'Идентификатор':
            column_seach = 'Бандероль'
        else:
            pass

        ask_user = messagebox.askyesno(title="Формирование документа", message="При нажатии кнопки 'Да', будет нельзя\n"
                                                                               "продолжить работу с данным реестром!")

        if ask_user == False:
            return

        clean()

        all_doc = []
        list_name = []

        files = filedialog.askopenfilenames()

        start_open = time.time()

        for file in files:

            file = file.split('/')
            file = '\\\\'.join(file)

            openPdf = open(file, 'rb')

            one_doc = ''  # Собираем все страницы в один документ
            one_doc += file + ' '

            pdfReader = PyPDF2.PdfReader(openPdf)

            pageObj = len(pdfReader.pages)

            for i in range(pageObj):
                getStr = pdfReader.pages[i]
                el = getStr.extract_text()
                one_doc += el

            one_doc = one_doc.split()
            one_doc.insert(0, file)
            all_doc.append(one_doc)

            openPdf.close()

        for doc in all_doc:

            switch_reg = 1

            new_doc = []
            name = ''
            short = []
            name = doc[0]
            name = f'...{name[-27:]}'   # Обрезаем название до 27 символов с конца
            short.append(name)   # Создаем отдельный список только с названием документа
            new_doc.append(name)   # Добавляем название в начало списка документа

            full_num = ''

            for i in doc:

                if column_seach in i:   #Определяем тип реестра
                    switch_reg = 0

                if next == 1 and select_scan == '1':
                    i = i[0]
                    full_num += i   # Добавляем последнюю цифру к ШПИ
                    next = 0
                    try:
                        full_num = int(full_num)
                    except ValueError:
                        pass
                    else:
                        full_num = str(full_num)
                        new_doc.append(full_num)

                elif next == 1 and select_scan == '2':
                    next = 0
                    try:
                        full_num = int(full_num)
                    except ValueError:
                        pass
                    else:
                        full_num = str(full_num)
                        new_doc.append(full_num)


                if switch_reg == 1:
                    if len(i) == 6:
                        full_num += i
                    elif len(i) == 2:
                        full_num += i
                    elif len(i) == 5:
                        full_num += i
                        next = 1
                    else:
                        full_num = ''

                elif switch_reg == 0:   #Если совпадение по ключевому слову то выполняеться часть кода
                    a = i[0:13]
                    if len(a) == 13:
                        try:
                            a = int(a)
                        except ValueError:
                            pass
                        else:
                            a = str(a)
                            new_doc.append(a)

            if short in leftovers:
                pass
            else:
                leftovers.append(new_doc)
                found.append(short)

        scroll_all.config(state=NORMAL)

        for element in leftovers:
            count_leftovers_all += int(len(element) - 1)
            for i in element:
                scroll_all.insert(INSERT, '\n' + i)
            scroll_all.insert(INSERT, '\n' + '_' * 35)
            scroll_all.insert(INSERT, '\n \n')

        scroll_all.config(state=DISABLED)

        label2.config(text=f'ОСТАЛОСЬ: {count_leftovers_all}')

        end_open = time.time() - start_open
        messagebox.showinfo('Загрузка', f'Загруженные документы: "{len(leftovers)}"\n'
                                        f'Время загрузки документов: {round(end_open, 2)} ')
#_______________________________________________________________________________________________________________________


# Функции
def select_check():

    status = switch.get()

    return status


def openSettings():

    with open('SettingsV2.txt', 'r') as fileOpen:
        fileRead = fileOpen.read()

    fileRead = fileRead.split('\n')
    if len(fileRead) > 4:
        fileRead.pop(4)

    id = fileRead[0]
    if id == '':
        id = 'Бандероль'

    select_scan = fileRead[1]
    if select_scan == '':
        select_scan = '2'

    min_range = fileRead[2]
    if min_range == '':
        min_range = 18

    max_range = fileRead[3]
    if max_range == '':
        max_range = 25

    return id, select_scan, min_range, max_range


def clean():

    global leftovers
    global found
    global not_found
    global numbers_used

    global count_found_all
    global count_found_one
    global count_leftovers_all
    global count_leftovers_one

    scroll_found_all.config(state=NORMAL)
    scroll_found_one.config(state=NORMAL)
    scroll_search_found.config(state=NORMAL)

    scroll_not_found.config(state=NORMAL)
    scroll_search_not.config(state=NORMAL)

    scroll_all.config(state=NORMAL)
    scroll_one.config(state=NORMAL)
    scroll_search_leftovers.config(state=NORMAL)

    scroll_found_all.delete(1.0, END)
    scroll_found_one.delete(1.0, END)
    scroll_search_found.delete(1.0, END)

    scroll_not_found.delete(1.0, END)
    scroll_search_not.delete(1.0, END)

    scroll_all.delete(1.0, END)
    scroll_one.delete(1.0, END)
    scroll_search_leftovers.delete(1.0, END)

    leftovers = []  # Остатки ШПИ
    found = []  # Найденные ШПИ
    not_found = []  # Не найденные ШПИ
    numbers_used = []  # Уже отсканированные ШПИ

    count_found_all = 0  # Счетчик совпадений ШПИ во всех реестрах
    count_found_one = 0
    count_leftovers_all = 0  # Счетчик остатков ШПИ во всех реестрах
    count_leftovers_one = 0

    label.config(text=f'НАЙДЕНО: {0}')
    label1.config(text=f'ОТСУТСТВУЕТ РЕЕСТР: {0}')
    label2.config(text=f'ОСТАЛОСЬ: {0}')
    label4.config(text=f'Последний ШПИ: {0}')

    scroll_found_all.config(state=DISABLED)
    scroll_found_one.config(state=DISABLED)
    scroll_search_found.config(state=DISABLED)

    scroll_not_found.config(state=DISABLED)
    scroll_search_not.config(state=DISABLED)

    scroll_all.config(state=DISABLED)
    scroll_one.config(state=DISABLED)
    scroll_search_leftovers.config(state=DISABLED)


def play_audio(track):

    wave_object = simpleaudio.WaveObject.from_wave_file(track)
    play = wave_object.play()
    play.wait_done()


def check_field(event):

    entry_field.config(background='white')


def check_search(event):

    entry_field.config(background='IndianRed1')


def disable_verification():

    global absentee_counter

    status = switch1.get()

    if status == 0:

        absentee_counter = 0

    return status


def barcode_not_supported(barcode):

    filename = 'signal.wav'

    keyboard.block_key('Return')
    keyboard.block_key('Prnt Scrn')

    play_audio(filename)

    messagebox.showwarning('Предупреждение', f'Данный вид ШПИ не поддерживается!\n'
                                             f'Чтобы продолжить работу, нажмите "Ok",\n'
                                             f'или клавиши: "Пробел" или "esc"')
    keyboard.unblock_key('Return')
    keyboard.unblock_key('Prnt Scrn')
    entry_field.delete(0, END)


def barcode_number_used(barcode):

    repeat = 'repeat.wav'

    keyboard.block_key('Return')
    keyboard.block_key('Prnt Scrn')

    play_audio(repeat)

    messagebox.showwarning('Предупреждение', f'Вы уже сканировали этот ШПИ: "{barcode}"!\n'
                                             f'Чтобы продолжить работу, нажмите "Ok",\n'
                                             f'или клавиши: "Пробел" или "esc"')
    keyboard.unblock_key('Return')
    keyboard.unblock_key('Prnt Scrn')


def barcode_not_found(barcode):

    filename = 'signal.wav'

    keyboard.block_key('Return')
    keyboard.block_key('Prnt Scrn')

    play_audio(filename)

    messagebox.showwarning('Предупреждение', f'ШПИ "{barcode}", не найден!\n'
                                             f'Чтобы продолжить работу, нажмите "Ok",\n'
                                             f'или клавиши: "Пробел" или "esc"')

    keyboard.unblock_key('Return')
    keyboard.unblock_key('Prnt Scrn')

    entry_field.delete(0, END)


def get_barcode(event):

    #py_win_keyboard_layout.change_foreground_window_keyboard_layout(0x04090409)

    global qr
    global barcode_qr

    verification = disable_verification()  # Отключить проверку

    filename = 'signal.wav'

    barcode = entry_field.get()  # Получаем ШПИ из поля ввода
    entry_field.delete(0, END)

    print(f'barcode = {barcode}')

    # Начало qr кода
    if 'ID' in barcode:
        qr = 1
        keyboard.block_key('Prnt Scrn')
        print(f'qr ID = {barcode}')
    elif 'ШВ' in barcode:
        qr = 1
        keyboard.block_key('Prnt Scrn')
        print(f'qr ID = {barcode}')

#_______________________________________________________________________________________________________________________
    if qr != 1:
        try:
            barcode = int(barcode)
        except ValueError:
            barcode_not_supported(barcode)
            return
        else:
            barcode = str(barcode)
#_______________________________________________________________________________________________________________________

    if 'Barcode' in barcode:
        barcode = barcode[9:23]
        barcode_qr = barcode
        print(f'qr Barcode = {barcode}')
    elif 'ИфксщвуЖ' in barcode:
        barcode = barcode[9:23]
        barcode_qr = barcode
        print(f'qr Barcode = {barcode}')

    if 'Delivery postcode' in barcode:
        qr = 0
        keyboard.unblock_key('Prnt Scrn')
        print(f'qr Delivery postcode = {barcode}')

        if status_barcode == 1:
            if verification == 0:
                barcode_not_found(barcode_qr)
                barcode_qr = ''
                return
            else:
                barcode_qr = ''
                return
        elif status_barcode == 2:
            if verification == 0:
                barcode_number_used(barcode_qr)
                barcode_qr = ''
                return
            else:
                barcode_qr = ''
                return

    elif 'Вудшмукн зщыесщвуЖ' in barcode:
        qr = 0
        keyboard.unblock_key('Prnt Scrn')
        print(f'qr Delivery postcode = {barcode}')

        if status_barcode == 1:
            if verification == 0:
                barcode_not_found(barcode_qr)
                barcode_qr = ''
                return
            else:
                barcode_qr = ''
                return
        elif status_barcode == 2:
            if verification == 0:
                barcode_number_used(barcode_qr)
                barcode_qr = ''
                return
            else:
                barcode_qr = ''
                return
    # Конец qr кода

    if len(barcode) == 14 or len(barcode) == 13:
        print(f'get_barcode = {barcode}')
        check_barcode(barcode)
    else:
        return


def check_barcode(barcode):

    global count_f
    count_f += 1
    print(f'цикл: {count_f}')

    certificate = select_check()   # Акт о повреждении

    verification = disable_verification()   # Отключить проверку

    start_click = time.time()

    global leftovers
    global found
    global not_found
    global count_found_all
    global count_found_one
    global count_leftovers_all
    global count_leftovers_one
    global absentee_counter
    global status_barcode

    status = False  # Проверям был ли найден ШПИ

#_______________________________________________________________________________________________________________________
    label4.config(text = f' Последний ШПИ: {barcode} ')

    entry_field.delete(0, END)
#_______________________________________________________________________________________________________________________
    if certificate == 1:
        for list in found:
            if barcode in list:

                scroll_found_all.config(state=NORMAL)
                scroll_found_one.config(state=NORMAL)

                scroll_found_all.delete(1.0, END)

                list.remove(barcode)

                scroll_one.delete(1.0, END)
                scroll_found_one.delete(1.0, END)

                list.append(barcode + ' (С актом)')
                count_found_all += 1
                count_found_one = len(list) - 1
                label.config(text=f'НАЙДЕНО: {count_found_one}')

                for i in list:
                    scroll_found_one.insert(INSERT, '\n' + i)  # Найдено в одном реестре

                for element in found:
                    for i in element:
                        scroll_found_all.insert(INSERT, '\n' + i)
                    scroll_found_all.insert(INSERT, '\n' + '_' * 35)
                    scroll_found_all.insert(INSERT, '\n \n')

                scroll_found_all.config(state=DISABLED)
                scroll_found_one.config(state=DISABLED)

                rad.deselect()

                end_click = time.time() - start_click

                label6.config(text=f'Время сканирования: {round(end_click, 2)}')

        if barcode in not_found:

            scroll_not_found.config(state=NORMAL)

            scroll_not_found.delete(1.0, END)

            not_found.remove(barcode)

            not_found.append(barcode + ' (С актом)')

            for element in reversed(not_found):
                scroll_not_found.insert(INSERT, '\n' + element)

            scroll_not_found.config(state=DISABLED)

            rad.deselect()

            end_click = time.time() - start_click

            label6.config(text=f'Время сканирования: {round(end_click, 2)}')
#_______________________________________________________________________________________________________________________
# Проверяем был ли использован ШПИ
    if barcode in numbers_used and verification == 0:

        if qr == 1:
            status_barcode = 2
            return
        else:
            barcode_number_used(barcode)
            return

    elif barcode in numbers_used and verification == 1:
        return
    else:
        numbers_used.append(barcode)

    print(f'numbers_used = {numbers_used}')

    scroll_found_all.config(state=NORMAL)
    scroll_found_one.config(state=NORMAL)
    scroll_not_found.config(state=NORMAL)
    scroll_all.config(state=NORMAL)
    scroll_one.config(state=NORMAL)

    scroll_found_all.delete(1.0, END)
    scroll_not_found.delete(1.0, END)
    scroll_all.delete(1.0, END)


    for element in leftovers:
        if barcode in element:
            element.remove(barcode)
            count_leftovers_all -= 1
            status = True
            count_leftovers_one = len(element) - 1
            label2.config(text=f'ОСТАЛОСЬ: {count_leftovers_one}')
            scroll_one.delete(1.0, END)
            scroll_found_one.delete(1.0, END)

            for list in found:
                if list[0] == element[0]:
                    if certificate == 1:
                        list.append(barcode + ' (С актом)')
                        count_found_all += 1
                        count_found_one = len(list) - 1
                        label.config(text=f'НАЙДЕНО: {count_found_one}')

                        for i in list:
                            scroll_found_one.insert(INSERT, '\n' + i)  # Найдено в одном реестре

                    elif certificate == 0:
                        list.append(barcode)
                        count_found_all += 1
                        count_found_one = len(list) - 1
                        label.config(text=f'НАЙДЕНО: {count_found_one}')

                        for i in list:
                            scroll_found_one.insert(INSERT, '\n' + i)  # Найдено в одном реестре

            for i in element:
                scroll_one.insert(INSERT, '\n' + i)   # Остатки в одном реестре


    # Добавляем ШПИ в список (не найдены)
    if status == False and barcode not in not_found:
        if verification == 0 and qr == 1:

            status_barcode = 1

            if certificate == 1:
                not_found.append(barcode + ' (С актом)')

            elif certificate == 0:
                not_found.append(barcode)

        elif verification == 0:

            if certificate == 1:
                not_found.append(barcode + ' (С актом)')
                barcode_not_found(barcode)

            elif certificate == 0:
                not_found.append(barcode)
                barcode_not_found(barcode)

        else:

            if certificate == 1:
                not_found.append(barcode + ' (С актом)')
                absentee_counter += 1

            elif certificate == 0:
                not_found.append(barcode)
                absentee_counter += 1

        # if verification == 0:   #Проверка включена
        #     # windll.user32.BlockInput(True)   # Блокирует клавиатуру и мышь. Не работает на 64 системе. Не блочит тачпанель
        #     keyboard.block_key('Return')
        #     play_audio(filename)
        #
        #     messagebox.showwarning('Предупреждение', f'Данный вид ШПИ не поддерживается!\n'
        #                                              f'Чтобы продолжить работу, нажмите "Ok",\n'
        #                                              f'или клавиши: "Пробел" или "esc"')
        #     keyboard.unblock_key('Return')
        #     entry_field.delete(0, END)
        # else:
        #     not_found.append(barcode)

    for element in leftovers:
        for i in element:
            scroll_all.insert(INSERT, '\n' + i)
        scroll_all.insert(INSERT, '\n' + '_' * 35)
        scroll_all.insert(INSERT, '\n \n')

    for element in found:
        for i in element:
            scroll_found_all.insert(INSERT, '\n' + i)
        scroll_found_all.insert(INSERT, '\n' + '_' * 35)
        scroll_found_all.insert(INSERT, '\n \n')

    for element in reversed(not_found):
        scroll_not_found.insert(INSERT, '\n' + element)

    label1.config(text=f'ОТСУТСТВУЕТ РЕЕСТР: {len(not_found)}')

    if verification == 1:
        label1.config(text=f'ОТСУТСТВУЕТ РЕЕСТР: {len(not_found)}({absentee_counter})')

    if get_switch == 0:
        label.config(text=f'НАЙДЕНО: {count_found_all}')
        label2.config(text=f'ОСТАЛОСЬ: {count_leftovers_all}')
    else:
        pass

    scroll_found_all.config(state=DISABLED)
    scroll_found_one.config(state=DISABLED)
    scroll_not_found.config(state=DISABLED)
    scroll_all.config(state=DISABLED)
    scroll_one.config(state=DISABLED)

    rad.deselect()

    end_click = time.time() - start_click

    label6.config(text=f'Время сканирования: {round(end_click, 2)}')
#_______________________________________________________________________________________________________________________


def change_registry():

    global get_switch

    if get_switch == 0:
        get_switch = 1

        scroll_all.grid_remove()
        scroll_found_all.grid_remove()
        scroll_search_found.grid_remove()
        scroll_search_leftovers.grid_remove()
        scroll_search_not.grid_remove()
        scroll_one.grid(column=1, row=6, **geometry)
        scroll_found_one.grid(column=0, row=6, **geometry)
        label.config(text=f'НАЙДЕНО: {count_found_one}')
        label2.config(text=f'ОСТАЛОСЬ: {count_leftovers_one}')
        button2.config(text='Переключить (Все реестры)')


    else:
        get_switch = 0

        scroll_one.grid_remove()
        scroll_found_one.grid_remove()
        scroll_search_found.grid_remove()
        scroll_search_leftovers.grid_remove()
        scroll_search_not.grid_remove()
        scroll_all.grid(column=1, row=6, **geometry)
        scroll_found_all.grid(column=0, row=6, **geometry)
        label.config(text=f'НАЙДЕНО: {count_found_all}')
        label2.config(text=f'ОСТАЛОСЬ: {count_leftovers_all}')
        button2.config(text='Переключить (Один реестр)')
#_______________________________________________________________________________________________________________________


def search():

    scroll_search_found.config(state=NORMAL)
    scroll_search_not.config(state=NORMAL)
    scroll_search_leftovers.config(state=NORMAL)

    scroll_search_found.delete(1.0, END)
    scroll_search_not.delete(1.0, END)
    scroll_search_leftovers.delete(1.0, END)

    get_search = entry_search.get()
    entry_search.delete(0, END)

    for element in leftovers:
        if get_search in element:
            scroll_all.grid_remove()
            scroll_one.grid_remove()
            scroll_search_leftovers.insert(INSERT, '\n' + element[0])
            scroll_search_leftovers.insert(INSERT, '\n' + get_search)

    for element in found:
        for i in element:
            if get_search == i[:14]:
                scroll_found_all.grid_remove()
                scroll_found_one.grid_remove()
                scroll_search_found.insert(INSERT, '\n' + element[0])
                scroll_search_found.insert(INSERT, '\n' + i)

    if get_search in not_found:
        scroll_search_not.grid_remove()
        scroll_search_not.insert(INSERT, '\n' + get_search)

    scroll_search_leftovers.grid(column=1, row=6, **geometry)
    scroll_search_found.grid(column=0, row=6, **geometry)
    scroll_search_not.grid(column=2, row=6, **geometry)

    scroll_search_found.config(state=DISABLED)
    scroll_search_not.config(state=DISABLED)
    scroll_search_leftovers.config(state=DISABLED)
#_______________________________________________________________________________________________________________________


def creacteDoc():

    ask_user = messagebox.askyesno(title="Формирование документа", message="При нажатии кнопки 'Да', будет нельзя\n"
                                                                           "продолжить работу с данным реестром!")

    if ask_user == False:
        return

    today = date.today()
    one_day = timedelta(1)
    todayDate = today + one_day
    todayDate = todayDate.strftime("%d-%m-%Y")

    file = filedialog.asksaveasfilename()
    file = file.split('/')
    file = '\\\\'.join(file)
    xlsx = f' ({str(todayDate)}).xlsx'

    if file == '':
        return

    if file[-5:] != '.xlsx':
        file += f' ({str(todayDate)}).xlsx'
    else:
        file = file[:-5]
        file += f' ({str(todayDate)}).xlsx'

    start_create = time.time()

    file_write = file

    count_leftovers_write = 3
    count_found_write = 3
    count_notfound_write = 3
    count_notfound_write_sheet1 = 9

    count_page = 0


    with xlsxwriter.Workbook(file_write) as book:

        sheet = book.add_worksheet()

        sheet1 = book.add_worksheet()

        cell_format_cap = book.add_format()
        cell_format_cap.set_bold()  # Жирный шрифт
        cell_format_cap.set_font_size(12)  # Размер шрифта
        cell_format_cap.set_align('center')  # Размещение текста по центру

        cell_format_cap1 = book.add_format()
        cell_format_cap1.set_bold()  # Жирный шрифт
        cell_format_cap1.set_font_size(12)  # Размер шрифта
        cell_format_cap1.set_align('center')  # Размещение текста по центру
        cell_format_cap1.set_border()

        cell_format_name = book.add_format()
        cell_format_name.set_border()   # Граница ячейки
        cell_format_name.set_bold()   # Жирный шрифт
        cell_format_name.set_font_size(12)   # Размер шрифта
        cell_format_name.set_align('center')   # Размещение текста по центру

        cell_format = book.add_format()
        cell_format.set_border()   # Граница ячейки
        cell_format.set_font_size(10)
        cell_format.set_align('right')

        sheet.set_column(0, 2, 35)   # Длина строки в Exсel
        sheet1.set_column(0, 2, 35)

        sheet.merge_range('A1:C1', '')
        sheet.merge_range('A2:C2', f'РЕЕСТР ПОЧТОВЫХ ОТПРАВЛЕНИЙ ЗА: "{todayDate}"', cell_format_cap)  # Объединить ячейки в одну
        sheet.merge_range('A3:C3', '')
        sheet.write(3, 0, f'НАЙДЕНО: {count_found_all}', cell_format_name)
        sheet.write(3, 1, f'ОСТАЛОСЬ: {count_leftovers_all}', cell_format_name)
        sheet.write(3, 2, f'ОТСУТСТВУЕТ РЕЕСТР: {len(not_found)}', cell_format_name)

        sheet1.merge_range('A1:C1', '')
        sheet1.merge_range('A2:C2', f'ПЯТЫЙ КАССАЦИОННЫЙ СУД ОБЩЕЙ ЮРИСДИКЦИИ', cell_format_cap)  # Объединить ячейки в одну
        sheet1.merge_range('A3:C3', f'РЕЕСТР ПОЧТОВЫХ ОТПРАВЛЕНИЙ ЗА: "{todayDate}"', cell_format_cap)  # Объединить ячейки в одну
        sheet1.merge_range('A4:C4', '')
        sheet1.merge_range('A5:C5', '')
        sheet1.write('A6', 'Сдал: ____________________', cell_format_cap)
        sheet1.write('C6', 'Принял: ____________________', cell_format_cap)
        sheet1.merge_range('A7:C7', '')
        sheet1.merge_range('A8:C8', '')
        sheet1.merge_range('A9:C9', f'Идентификатор почтового отправления, номер почтового перевода,'
                                    f' количество ТМЦ, поручений', cell_format_cap1)  # Объединить ячейки в одну

        max_length = 0


        for element in leftovers:

            for i in element:

                count_leftovers_write += 1

                if i == element[0]:
                    sheet.write(count_leftovers_write, 1, i, cell_format_name)

                    if int(len(element)) <= 1:
                        count_leftovers_write += 1
                        sheet.write(count_leftovers_write, 1, "(Пустой)", cell_format_name)

                elif i != element[0] and i != '':

                    sheet.write(count_leftovers_write, 1, i, cell_format)


        for element in found:

            count_found_end = len(element) - 1

            for i in element:

                count_found_write += 1

                if i == element[0]:

                    i = f'{i} ({str(count_found_end)})'

                    sheet.write(count_found_write, 0, i, cell_format_name)

                    if int(len(element)) <= 1:
                        count_found_write += 1
                        sheet.write(count_found_write, 0, "(Пустой)", cell_format_name)

                elif i != element[0]:

                    sheet.write(count_found_write, 0, i, cell_format)

        for element in not_found:

            count_notfound_write += 1
            count_notfound_write_sheet1 += 1
            count_page += 1

            sheet.write(count_notfound_write, 2, element, cell_format)
            sheet1.write(count_notfound_write_sheet1 - 1, 0, count_page, cell_format)

            if len(element) == 14:
                element = f'{element[0:6]} {element[6:8]} {element[8:13]} {element[13]}'

            sheet1.merge_range(f'B{count_notfound_write_sheet1}:C{count_notfound_write_sheet1}', element, cell_format)

        # Ищем максимальную длину столбца
        if count_found_write > max_length:
            max_length = count_found_write

        if count_leftovers_write > max_length:
            max_length = count_leftovers_write

        if count_notfound_write > max_length:
            max_length = count_notfound_write
        #____________________________________

        # Добавляем пустые ячейки чтобы сформировать одинаковую по длине таблицу
        while count_found_write < max_length:
            count_found_write += 1
            sheet.write(count_found_write, 0, " ", cell_format_name)

        while count_leftovers_write < max_length:
            count_leftovers_write += 1
            sheet.write(count_leftovers_write, 1, " ", cell_format_name)

        while count_notfound_write < max_length:
            count_notfound_write += 1
            sheet.write(count_notfound_write, 2, " ", cell_format_name)

    clean()

    end_create = time.time() - start_create
    end_work = time.time() - start_work
    end_work = end_work / 60

    messagebox.showinfo("Время формирования документа", f"Документ софрмирован за: {round(end_create, 2)}\n"
                                                        f"Время работы программы: {round(end_work, 2)} мин")
#_______________________________________________________________________________________________________________________


# Вызов классов
call_GetData = GetData()
call_Settings = ChangeSettings()
#_______________________________________________________________________________________________________________________


# Начало блока оформления программы
def fullscreen(e = None):
    if window.attributes('-fullscreen'): # Проверяем режим окна
       window.attributes('-fullscreen', False) # Меняем режим окна
    else:
       window.attributes('-fullscreen', True)

# def is_valid(newval):
#     return re.match("^\d{0,30}$", newval) is not None
#_______________________________________________________________________________________________________________________


# Параметры главного окна программы
window = Tk()
window.iconbitmap('Minenergo_logo.ico')
window.resizable(True, True)                                                                                            #Запрещает разворачивать окно программы на весь экран
window.title('FiveKasScan')

geometry = {'padx': 10, 'pady': 10}
#_______________________________________________________________________________________________________________________


# Radiobutton
switch = IntVar()

rad = Checkbutton(window, text='С актом о повреждении', command=select_check, variable=switch)
rad.grid(column=1, row=3, **geometry)

switch1 = IntVar()

rad1 = Checkbutton(window, text='Отключить проверку', command=disable_verification, variable=switch1)
rad1.grid(column=2, row=3, **geometry)
#_______________________________________________________________________________________________________________________


# Entry

#check = (window.register(is_valid), "%P")

entry_field = Entry(window, width=30, state = 'normal')                                                                        #state - доступно поле для ввода или нет
entry_field = ttk.Entry(validate="key")
#entry_field = ttk.Entry(validate="key", validatecommand=check)
entry_field.grid(column=1, row=1)
entry_field.focus()
entry_field.bind('<Return>', get_barcode)
entry_field.bind('<Button-1>', check_field)

entry_search = Entry(window, width=15, state = 'normal')                                                                        #state - доступно поле для ввода или нет
entry_search.grid(column=1, row=8)
entry_search.bind('<Button-1>', check_search)
#_______________________________________________________________________________________________________________________


# Buttons
button1 = Button(window, text="Загрузить документы (Excel)", command=call_GetData.openFilesExcel, bg="lightgreen", fg="black", width=27)
button1.grid(column=0, row=0, **geometry)

button2 = Button(window, text="Переключить (Один реестр)", command=change_registry, bg="white", fg="black", width=27)
button2.grid(column=0, row=2, **geometry)

button3 = Button(window, text="Загрузить документы (PDF)", command=call_GetData.openFilesPDF, bg="IndianRed1", fg="black", width=27)
button3.grid(column=0, row=1, **geometry)

button4 = Button(window, text="Формирование документа", command=creacteDoc, bg="white", fg="black", width=27)
button4.grid(column=2, row=0, **geometry)

button5 = Button(window, text="Поиск", command=search, bg="white", fg="black", width=27)
button5.grid(column=2, row=8, **geometry)

button6 = Button(window, text="Настройки", command=call_Settings.settings, bg="white", fg="black", width=27)
button6.grid(column=2, row=1, **geometry)
#_______________________________________________________________________________________________________________________


# Labels
label = Label(window, text = f'НАЙДЕНО: {0}', justify='center', font=("Times New Roman", 12), fg="Green")
label.grid(column=0, row=5, padx=20, pady=20)

label1 = Label(window, text = f'ОТСУТСТВУЕТ РЕЕСТР: {0}', justify='center', font=("Times New Roman", 12), fg="Red")
label1.grid(column=2, row=5, **geometry)

label2 = Label(window, text = f'ОСТАЛОСЬ: {0}', justify='center', font=("Times New Roman", 12), fg="black")
label2.grid(column=1, row=5, **geometry)

label3 = Label(window, text = 'Для сканирования ШПИ,\n'
                              'курсор должен находиться в поле ввода!', font=("Times New Roman", 14), fg="black")
label3.grid(column=1, row=0, **geometry)

label4 = Label(window, text = f' Последний ШПИ: {0} ', font=("Times New Roman", 14), fg="black", relief=RIDGE)
label4.grid(column=1, row=2, **geometry)

label5 = Label(window)
label5.grid(column=0, row=3)

label6 = Label(window, text = f'Время сканирования:', font=("Times New Roman", 14), fg="black")
label6.grid(column=2, row=2, **geometry)

label7 = Label(window, text = f'Для поиска, введите ШПИ в поле справа', font=("Times New Roman", 12), fg="black")
label7.grid(column=0, row=8, **geometry)
#_______________________________________________________________________________________________________________________


# Scrolltext
# Найдено
scroll_found_all = scrolledtext.ScrolledText(window, width=35, height=22, state=DISABLED)
scroll_found_all.grid(column=0, row=6, **geometry)
scroll_found_all.bind('<Button-1>', check_search)

scroll_found_one = scrolledtext.ScrolledText(window, width=35, height=22, state=DISABLED)
scroll_found_one.bind('<Button-1>', check_search)

scroll_search_found = scrolledtext.ScrolledText(window, width=35, height=22, state=DISABLED)
scroll_search_found.bind('<Button-1>', check_search)

# Остатки
scroll_all = scrolledtext.ScrolledText(window, width=35, height=22, state=DISABLED)
scroll_all.grid(column=1, row=6, **geometry)
scroll_all.bind('<Button-1>', check_search)

scroll_one = scrolledtext.ScrolledText(window, width=35, height=22, state=DISABLED)
scroll_one.bind('<Button-1>', check_search)

scroll_search_leftovers = scrolledtext.ScrolledText(window, width=35, height=22, state=DISABLED)
scroll_search_leftovers.bind('<Button-1>', check_search)

# Отсутствует
scroll_not_found = scrolledtext.ScrolledText(window, width=30, height=22, state=DISABLED)
scroll_not_found.grid(column=2, row=6, **geometry)
scroll_not_found.bind('<Button-1>', check_search)

scroll_search_not = scrolledtext.ScrolledText(window, width=30, height=22, state=DISABLED)
scroll_search_not.bind('<Button-1>', check_search)
#_______________________________________________________________________________________________________________________


window.mainloop()
