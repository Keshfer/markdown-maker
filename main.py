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

allowed_files = [
	("Word Document", "*.docx"),
	("OpenDocument Text", "*.odt"),
	("Portable Document", "*.pdf"),
	("Text Document", "*.txt")
]

with filedialog.askopenfile(filetypes=allowed_files) as file:
	print(file)
	print(file.name)
	file_name = os.path.basename(file.name)
	dir_path = os.path.dirname(file.name)
	#replace format extension with .md. [^.] means any character except ".". We use \. since . by itself is wildcard in regex 
	file_name = re.sub(r"\.[^.]+$", ".md", file_name, 1) 
	print(dir_path)
	print(file_name)
	output_path = dir_path + "/" + file_name #same directory as file
	print(output_path)
	pypandoc.convert_file(file.name, "md", outputfile=output_path)	
