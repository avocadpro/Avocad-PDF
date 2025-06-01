# -*- coding: utf-8 -*-
# Импортируем нужные библиотеки
import pythoncom
from win32com.client import Dispatch, gencache # Для запуска и управления КОМПАС-3D
import tkinter as tk  # Библиотека для создания окон
from tkinter import ttk, filedialog, messagebox # Кнопки, диалоговые окна, сообщения

# Функция запуска и подготовки КОМПАС-3D
def start_init():
    global api7, iApplication, iConverter,root # Используем переменные глобально

    # Подключаемся к уже запущенному КОМПАС-3D или запускаем новый
    api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    iApplication = api7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(api7.IApplication.CLSID, pythoncom.IID_IDispatch))
    iApplication.Visible = True

    # Создаём главное окно приложения
    root = tk.Tk()

    # Путь к папке, где лежит нужная библиотека PDF
    kompas_dir_path = "C:\\Program Files\\ASCON\\KOMPAS-3D v23\\Bin"
    dll_path = kompas_dir_path+"\\Pdf2d.dll"

    # Загружаем PDF-конвертер
    iConverter = iApplication.Converter(dll_path)
    # Если получилось загрузить — возвращаем True
    if iConverter:
        return True
    # Если не получилось — показываем сообщение об ошибке
    messagebox.showerror("Ошибка", "Возникла ошибка при инициализации компонентов Компас")
    return False

# Функция: создать PDF из текущего открытого документа
def create_pdf():
    iKompasDocument = iApplication.ActiveDocument
    # Проверяем, что это чертёж, фрагмент или спецификация
    if iKompasDocument and iKompasDocument.DocumentType in range(1,4):
        doc_filename = iKompasDocument.PathName # Получаем путь к файлу
        pdf_filename = doc_filename[:-4]+".pdf" # Меняем расширение на .pdf
        # Преобразуем в PDF
        iConverter.Convert(doc_filename, pdf_filename, 0, True)
        # Показываем сообщение, что файл создан
        iApplication.MessageBoxEx("Создан файл {}".format(pdf_filename),
                        "PDF Module", 48)
    else:
        # Если документ неподходящий — выводим предупреждение
        iApplication.MessageBoxEx("Активный документ не является чертежом, фрагментом или спецификацией!",
                        "PDF Module", 64)

# Функция: выбрать несколько документов на диске и сохранить их в PDF
def create_many_pdf():
    file_paths = filedialog.askopenfilenames(title = "Выбор документов",
                                            filetypes = [("КОМПАС-Документы", ("*.cdw", "*.frw ", "*.spw"))])
    # Получаем список выбранных файлов
    file_list = root.tk.splitlist(file_paths)
    # Преобразуем каждый файл в PDF
    for doc_filename in file_list:
        pdf_filename = doc_filename[:-4]+".pdf"
        iConverter.Convert(doc_filename, pdf_filename, 0, True)
        iApplication.MessageBoxEx("Создан файл {}".format(pdf_filename),
                        "PDF Module", 48)

# Функция: преобразовать все открытые документы в PDF
def active_docs_to_pdf():
    iDocuments = iApplication.Documents # Получаем интерфейс для работы с документами

    # Перебираем каждый документ
    for i in range(iDocuments.Count):
        iKompasDocument = iDocuments.Item(i)
        # Проверяем, что это чертёж, фрагмент или спецификация
        if iKompasDocument.DocumentType in range(1,4):
            try:
                doc_filename = iKompasDocument.PathName
                pdf_filename = doc_filename[:-4]+".pdf"
                # Преобразуем в PDF
                iConverter.Convert(doc_filename, pdf_filename, 0, True)
                iApplication.MessageBoxEx("Создан файл {}".format(pdf_filename),
                        "PDF Module", 48)
            except Exception:
                # Если произошла ошибка — сообщаем
                iApplication.MessageBoxEx("Произошла ошибка при сохранении файла {}".format(pdf_filename))

# Функция: создаёт окно с кнопками
def get_window():
    root.title("AvocadPDF") # Заголовок окна
    root.resizable(False, False) # Запрещаем менять размер
    root.attributes("-topmost", True)  # Всегда поверх других окон

    # Рамка вокруг кнопок
    frame1 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = '')
    frame1.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = 'nsew')

    # Кнопка 1 — текущий документ в PDF
    button1 = ttk.Button(frame1, text="Создать PDF",
                         command=create_pdf,width=40,
                         state="normal")
    button1.grid(row=0, column=0)

    # Кнопка 2 — выбрать документы с диска и сохранить
    button2 = ttk.Button(frame1, text="Выбрать документы с диска",
                         command=create_many_pdf,width=40,
                         state="normal")
    button2.grid(row=1, column=0)

    # Кнопка 3 — все открытые документы сохранить
    button3 = ttk.Button(frame1, text="Активные документы в PDF",
                         command=active_docs_to_pdf,width=40,
                         state="normal")
    button3.grid(row=2, column=0)

    root.mainloop() # Запускаем главное окно (цикл обработки событий)

# Запускаем программу, если этот файл запущен напрямую
if __name__=="__main__":
    start_status = start_init() # Инициализируем КОМПАС
    if not start_status:
        pass # Если не получилось — пропускаем и выходим из программы
    else:
        get_window()  # Иначе — запускаем окно с кнопками
