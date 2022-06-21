import tkinter as tk

from frames import FileInputFrame
from settings import chek_default_config


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # configure the root window
        self.title('DEMO отчеты WB в 1С')
        self.geometry('600x150')


if __name__ == '__main__':
    chek_default_config()
    app = App()
    FileInputFrame(app)
    app.mainloop()
