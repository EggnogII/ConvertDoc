from tkinter import filedialog

def markdown_document_selector():

    filename = filedialog.askopenfilename(title="Select Markdown File", filetypes=
    (("Markdown files", "*.md"), ("all files", "*.*")))

    return filename


