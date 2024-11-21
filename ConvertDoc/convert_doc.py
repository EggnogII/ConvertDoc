from convert_doc_form import convert_doc_gui
from resolutions import horizontal_resolution, vertical_resolution
from tkinter import *

def main():
    root = Tk()
    gui = convert_doc_gui(root)
    minwidth = 854
    minheight = 480

    screenwidth = minwidth
    screenheight = minheight

    fullscreen_width = root.winfo_screenwidth()
    fullscreen_height = root.winfo_screenheight()

    if fullscreen_width == horizontal_resolution.UHD.value and fullscreen_height == vertical_resolution.UHD.value:
        screenwidth = horizontal_resolution.QVGA.value
        screenheight = vertical_resolution.QVGA.value

    if fullscreen_width < horizontal_resolution.QHD.value and fullscreen_height < vertical_resolution.QHD.value:
        screenwidth = horizontal_resolution.VGA.value
        screenheight = vertical_resolution.VGA.value


    root.geometry('%sx%s' % (screenwidth, screenheight))
    root.mainloop()

if __name__ == '__main__':
    main()