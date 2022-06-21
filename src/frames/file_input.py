import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog as fd
from tkinter.messagebox import showinfo

from services import ReportSumService


class FileInputFrame(ttk.Frame):
    filetypes = (
        ('Excel', '*.xlsx'),
        ('Excel 97', '*.xls')
    )

    def __init__(self, container):
        super().__init__(container)

        options = {'padx': 5, 'pady': 5}

        ## WB row

        # label report
        self.input_file_label = ttk.Label(self, text='Отчет WB')
        self.input_file_label.grid(column=0, row=0, sticky=tk.W, **options)

        self.output_file_label = ttk.Label(self, text='Файл результата обработки')
        self.output_file_label.grid(column=0, row=1, sticky=tk.W, **options)

        # temperature entry
        self.input_file_name = tk.StringVar()
        self.input_file_name_entry = ttk.Entry(self, textvariable=self.input_file_name)
        self.input_file_name_entry.grid(column=1, row=0, **options)
        self.input_file_name_entry.focus()

        self.output_file_name = tk.StringVar()
        self.output_file_name_entry = ttk.Entry(self, textvariable=self.output_file_name)
        self.output_file_name_entry.grid(column=1, row=1, **options)
        self.output_file_name_entry.focus()

        # button
        self.button = ttk.Button(self, text='Выбрать')
        self.button['command'] = (
            lambda: self.input_file_name.set(fd.askopenfilename(filetypes=self.filetypes))
        )
        self.button.grid(column=2, row=0, **options)

        self.output_button = ttk.Button(self, text='Сосхранить в')
        self.output_button['command'] = (
            lambda: self.output_file_name.set(fd.asksaveasfilename(filetypes=self.filetypes))
        )
        self.output_button.grid(column=2, row=1, **options)

        self.run_button = ttk.Button(self, text='Начать обработку')
        self.run_button['command'] = self.start_processing
        self.run_button.grid(column=2, row=4, **options)
        # show the frame on the container
        self.pack(**options)

    def start_processing(self):
        input_file_path = Path(self.input_file_name.get())
        output_file_path = Path(self.output_file_name.get())

        if not input_file_path.is_file():
            showinfo('Ошибка', 'Неверно задан файл отчета')
            return

        if not self.output_file_name.get():
            showinfo('Ошибка', 'Не задан файл куда сохранять')
            return

        ReportSumService().get_sum_report(input_file_path, output_file_path)
        showinfo('Выполнено', f'результаты в файле {output_file_path.name} ')
