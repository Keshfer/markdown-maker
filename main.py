import pypandoc
import tkinter as tk
from tkinter import filedialog
from ctypes import windll
import os
import re

MD_DIR = "md_output"
#MEDIA_DIR = "media"


def make_output_dir(dir_path) -> bool:
	#make dir to contain output files
	path = dir_path + "/" + MD_DIR
	try:
		os.makedirs(path)
		print("made dir")
		return True
	except FileExistsError:
		print(f"directory already exists in {path}")
		return True
	except PermissionError:
		print(f"Unable to create directory {path}")
		return False
	except Exception as e:
		print(f"error occurred: {e}")
		return False

# def make_media_dir(dir_path) -> bool:
# 	#make the media dir inside output dir, assuming it exists
# 	path = dir_path + "/" + MEDIA_DIR
# 	try:
# 		os.makedirs(path)
# 		return True
# 	except FileExistsError:
# 		print(f"directory already exists in {path}")
# 		return True
# 	except PermissionError:
# 		print(f"Unable to create directory {path}")
# 		return False
# 	except Exception as e:
# 		print(f"error occurred: {e}")
# 		return False

def convert_file(org_file_name, file_path, imgs_dir):
	pypandoc.convert_file(org_file_name, "md", outputfile=file_path, extra_args=[f"--extract-media={imgs_dir}"])

def main():
	windll.shcore.SetProcessDpiAwareness(2)

	#create and hide root window
	# root = tk.Tk()
	# root.withdraw()

	allowed_files = [
		("Word Document", "*.docx"),
		("OpenDocument Text", "*.odt"),
	]

	with filedialog.askopenfile(filetypes=allowed_files) as file:
		file_name = os.path.basename(file.name)
		dir_path = os.path.dirname(file.name)
		#replace format extension with .md. [^.] means any character except ".". We use \. since . by itself is wildcard in regex 
		file_name = re.sub(r"\.[^.]+$", ".md", file_name, 1)
		status = make_output_dir(dir_path)
		if not status:
			print("storing output in " + dir_path)
			output_path = dir_path + "/" + file_name
			output_imgs_path = dir_path
			#pypandoc.convert_file(file.name, "md", outputfile=output_path, extra_args=[f"--extract-media={output_imgs_path}"])
			convert_file(file.name, output_path, output_imgs_path)
			return
		#at this point the directory md_output exists
		# status = make_media_dir(dir_path + "/md_output")
		# if not status:
		# 	output_path = dir_path + "/" + MD_DIR + "/" + file_name
		# 	output_imgs_path = dir_path + "/" + MD_DIR
		# 	#pypandoc.convert_file(file.name, "md", outputfile=output_path, extra_args=[f"--extract-media={output_imgs_path}"])
		# 	convert_file(file.name, output_path, output_imgs_path)
		# 	return
		#Both md_output and media directory exists beyond here
		output_path = dir_path + "/" + MD_DIR + "/" + file_name
		output_imgs_path =  dir_path + "/" + MD_DIR
		print(output_path)
		print(output_imgs_path)
		#pypandoc.convert_file(file.name, "md", outputfile=output_path, extra_args=[f"--extract-media={output_imgs_path}"])	
		convert_file(file.name, output_path, output_imgs_path)

main()