import pandas as pd
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()

reader = open(file_path)
writer = open(file_path + 'without_blanks.csv', 'w')

delimiter = ','

for line in reader:
    if all(col != '' for col in line.split(delimiter)):
        writer.write(line)