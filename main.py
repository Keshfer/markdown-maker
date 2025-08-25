import pypandoc
import tkinter as tk
from tkinter import filedialog
from ctypes import windll
import os
import re

windll.shcore.SetProcessDpiAwareness(2)

#create and hide root window
# root = tk.Tk()
# root.withdraw()

with filedialog.askopenfile() as file:
	print(file)
	print(file.name)
	file_name = os.path.basename(file.name)
	file_name = re.sub(r"\.[^.]+$", ".md", file_name, 1) #replace format extension with .md
	print(file_name)
	pypandoc.convert_file(file.name, "md", outputfile=file_name)	
