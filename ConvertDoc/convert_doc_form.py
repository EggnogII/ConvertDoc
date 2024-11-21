import os
import shutil
import pygit2
import subprocess
import pypandoc
import convert_doc_controller
from shutil import copy, move
from tkinter import filedialog, simpledialog, ttk, messagebox, Frame, TOP, BOTTOM
from tkinter import *
from resolutions import horizontal_resolution, vertical_resolution

class convert_doc_gui:
    def __init__(self, master):

        self.master = master
        master.title("ConvertDoc")
        self.label = Label(master, text="ConvertDoc")
        self.label.pack()
        top_frame = Frame(master)
        top_frame.pack(side=TOP)
        bottom_frame = Frame(master)
        bottom_frame.pack(side=BOTTOM)
        west_frame = Frame(master)
        west_frame.pack(side=LEFT)
        east_frame = Frame(master)
        east_frame.pack(side=RIGHT)

        self.progress_bar = ttk.Progressbar(bottom_frame, orient='horizontal', mode='determinate', maximum=100)
        self.convert_to_md_button = Button(master, text="Convert DOCX to MD", command=lambda : convert_doc_controller.convert_docx_to_markdown(self.progress_bar))
        self.convert_to_md_button.pack(padx=10, pady=10, expand=True, side=LEFT)
        self.convert_to_docx_button = Button(master, text="Convert MD to DOCX", command=lambda : convert_doc_controller.convert_markdown_to_docx_button(self.progress_bar))
        self.convert_to_docx_button.pack(padx=10, pady=10, expand=True, side=LEFT)
        self.progress_bar.pack(fill=BOTH, padx=10, pady=10, anchor='center', expand=True)


    
