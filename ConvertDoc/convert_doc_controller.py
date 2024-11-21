from tkinter import filedialog

def document_selector(selection_title, filetype):
    filename = filedialog.askopenfilename(title=selection_title, filetypes=
    (("Documents", f"*.{filetype}"), ("all files", "*.*")))

    return filename



