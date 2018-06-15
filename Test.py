import tkinter as tk
from tkinter import filedialog


root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
workbook = xlrd.open_workbook(file_path)

sheet = workbook.open_by_index(0)
