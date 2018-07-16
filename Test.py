import tkinter as tk
from openpyxl import load_workbook
from tkinter import filedialog


root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
wb = load_workbook (file_path)
date = input('Enter the date of the exam: ')
subject = input ('Enter the subject/topic of the exam: ')
ws = wb.active
for row in ws:
	if str(row[0].value).lower().strip() != 'name':
		message = 'Your ward ' + str(row[0].value) +' has scored ' +str(row[2].value) + ' on the test of '+ subject  +' held on '+ date
		print(message)