import os
import sys
import tkinter as tk
from tkinter import ttk
from ctypes import windll
from tkinter import filedialog
import pandas as pd


class App():
    def __init__(self):
        windll.shcore.SetProcessDpiAwareness(1)
        self.root = tk.Tk()
        self.params()
        self.root.title(self.TITLE_WINDOW)
        self.root.geometry(self.GEOMETRY_WINDOW)
        self.root.configure(bg=self.COLOR_BACKGROUND)
        self.root.iconbitmap(self.PATH_ICON)

        self.add_label(img=self.IMG_CSV, anchor="e")
        self.add_label(img=self.IMG_RIGHT_ARROW, col=1)
        self.add_label(img=self.IMG_XLS, col=2, anchor="w")

        self.var_notify = tk.StringVar()
        self.add_label_nofity()

        self.button_csv = self.add_button(
            text=self.TEXT_BUTTON_CSV,
            command=self.click_button_csv)
        self.button_xls = self.add_button(
            text=self.TEXT_BUTTON_XLS,
            command=self.click_button_xls,
            col=1,
            state="disabled")
        self.button_convert = self.add_button(
            text=self.TEXT_BUTTON_CONVERT,
            command=self.click_button_convert,
            col=2,
            state="disabled")

        self.root.mainloop()

    def params(self):
        # images
        PATH_FOLDER_PROJET = os.path.abspath(".")

        try:
            PATH_FOLDER_IMG = sys._MEIPASS + "\\imgs"
        except Exception:
            PATH_FOLDER_IMG = PATH_FOLDER_PROJET + "\\imgs"

        PATH_IMG_CSV = PATH_FOLDER_IMG + "\\csv.png"
        PATH_IMG_RIGHT_ARROW = PATH_FOLDER_IMG + "\\right-arrow.png"
        PATH_IMG_XLS = PATH_FOLDER_IMG + "\\xls.png"
        self.PATH_ICON = PATH_FOLDER_IMG + "\\convert.ico"

        self.IMG_CSV = tk.PhotoImage(
            file=PATH_IMG_CSV).subsample(5, 5)
        self.IMG_RIGHT_ARROW = tk.PhotoImage(
            file=PATH_IMG_RIGHT_ARROW).subsample(10, 10)
        self.IMG_XLS = tk.PhotoImage(
            file=PATH_IMG_XLS).subsample(5, 5)

        # title
        self.TITLE_WINDOW = "CSV TO XLS"

        # geometry
        WIDTH_WINDOW = 960
        HEIGHT_WINDOW = 540

        self.root.update_idletasks()
        SCREEN_DPI = self.root.winfo_fpixels("1i")
        SCREEN_WIDTH = self.root.winfo_screenwidth()
        SCREEN_HEIGHT = self.root.winfo_screenheight()

        x = (SCREEN_WIDTH - WIDTH_WINDOW) // 2
        y = (SCREEN_HEIGHT - HEIGHT_WINDOW) // 2

        self.GEOMETRY_WINDOW = f"{WIDTH_WINDOW}x{HEIGHT_WINDOW}+{x}+{y}"

        # font
        size_16 = round((16 * 96) / SCREEN_DPI)
        size_30 = round((30 * 96) / SCREEN_DPI)

        self.FONT_LABEL_NOFITY = ("Arial", size_30, "italic")
        self.FONT_TBUTTON = ("Arial", size_16, "italic")

        # color
        self.COLOR_BACKGROUND = "light blue"

        # text
        self.TEXT_BUTTON_CSV = "Chọn tệp CSV"
        self.TEXT_BUTTON_XLS = "Chọn thư mục lưu tệp XLS"
        self.TEXT_BUTTON_CONVERT = "Chuyển đổi"

        self.TEXT_NOTIFY_1 = "Ứng dụng chuyển đổi CSV sang XLS"
        self.TEXT_NOTIFY_2 = "Đang chuyển đổi..."
        self.TEXT_NOTIFY_3 = "Hoàn thành chuyển đổi!"

        # style
        style = ttk.Style()
        style.configure(
            'TButton',
            padding=(5, 5, 5, 5),
            background=self.COLOR_BACKGROUND,
            font=self.FONT_TBUTTON)

    def add_label(self, img, row=0, col=0, anchor="center"):
        label = tk.Label(
            self.root,
            bg=self.COLOR_BACKGROUND,
            image=img,
            anchor=anchor
        )
        label.grid(row=row, column=col, sticky="nesw")
        self.root.rowconfigure(row, weight=1)
        self.root.columnconfigure(col, weight=1)

    def add_label_nofity(self, row=1):
        label = tk.Label(
            self.root,
            bg=self.COLOR_BACKGROUND,
            textvariable=self.var_notify,
            font=self.FONT_LABEL_NOFITY,
            anchor="center"
        )
        label.grid(row=row, columnspan=3, sticky="nesw")
        self.var_notify.set(self.TEXT_NOTIFY_1)

    def add_button(self, text, command, row=2, col=0, state="normal"):
        button = ttk.Button(
            self.root,
            text=text,
            width=22,
            style="TButton",
            state=state,
            command=command
        )
        button.grid(row=row, column=col)
        self.root.rowconfigure(row, weight=1)
        self.root.columnconfigure(col, weight=1)
        return button

    def click_button_csv(self):
        self.var_notify.set(self.TEXT_NOTIFY_1)

        self.PATH_FILE_CSV = filedialog.askopenfilename(
            title=self.TEXT_BUTTON_CSV,
            filetypes=[("CSV files", "*.csv")],
            defaultextension=".csv"
        )

        if self.PATH_FILE_CSV:
            self.NAME_FILE_CSV = os.path.basename(self.PATH_FILE_CSV)
            self.PATH_FOLDER_CSV = os.path.dirname(self.PATH_FILE_CSV)
            self.PATH_FOLDER_XLS = self.PATH_FOLDER_CSV
            NAME_FOLDER_CSV = os.path.basename(self.PATH_FOLDER_CSV)

            self.button_csv.config(text=self.NAME_FILE_CSV)
            self.button_xls.config(state="normal", text=NAME_FOLDER_CSV)
            self.button_convert.config(state="normal")

    def click_button_xls(self):
        self.PATH_FOLDER_XLS = filedialog.askdirectory(
            title=self.TEXT_BUTTON_XLS
        )

        if self.PATH_FOLDER_XLS:
            NAME_FOLDER_XLS = os.path.basename(self.PATH_FOLDER_XLS)
            self.button_xls.config(text=NAME_FOLDER_XLS)

    def click_button_convert(self):
        self.var_notify.set(self.TEXT_NOTIFY_2)
        self.button_csv.config(state="disabled")
        self.button_xls.config(state="disabled")
        self.button_convert.config(state="disabled")

        NAME_FILE_CSV_NOT_EXTENSION = self.NAME_FILE_CSV[:-4]
        NAME_FILE_XLS = rf"{NAME_FILE_CSV_NOT_EXTENSION}.xlsx"
        PATH_FILE_XLS = rf"{self.PATH_FOLDER_XLS}//{NAME_FILE_XLS}"
        
        df = pd.read_csv(self.PATH_FILE_CSV)
        df.to_excel(PATH_FILE_XLS, index=False)
        self.var_notify.set(self.TEXT_NOTIFY_3)

        self.button_csv.config(state="normal", text=self.TEXT_BUTTON_CSV)
        self.button_xls.config(text=self.TEXT_BUTTON_XLS)


App()
