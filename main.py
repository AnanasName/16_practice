import PySimpleGUI as gui
import os
from enum import Enum
from docx2pdf import convert
from pdf2docx import Converter
from PIL import Image


class Event(Enum):
    files = 0
    open_folder = 1
    select_folder = 2
    pdfToDocx = 3
    docxToPdf = 4
    compress = 5
    delete = 6
    exit = 7
    parent_folder = 8

def main():
    folder = os.getcwd()
    first_column = [[
        gui.In(size=(28, 1), enable_events=True, key=Event.open_folder),
        gui.FolderBrowse(button_text="Select Folder"),
    ], [
        gui.Listbox(values=os.listdir(folder), size=(
            40, 10), select_mode=gui.LISTBOX_SELECT_MODE_MULTIPLE, key=Event.files, enable_events=True)
    ]]

    second_column = [
        [gui.Button("Open parent folder", key=Event.parent_folder)],
        [gui.Button("Open selected folder", key=Event.select_folder, disabled=True)],
        [gui.Button("DOCX to PDF", key=Event.docxToPdf, disabled=True)],
        [gui.Button("PDF to DOCX", key=Event.pdfToDocx, disabled=True)],
        [gui.Button("Delete", key=Event.delete, disabled=True)],
        [gui.Button("Compress", key=Event.compress, disabled=True)],
        [gui.Button("Exit", key=Event.exit)]
    ]

    layout = [[
        gui.Column(first_column),
        gui.Column(second_column)
    ]]

    window = gui.Window("File Explorer", layout)

    while True:
        event, values = window.read()

        selected_files = values[Event.files]
        open_folder = values[Event.open_folder]

        only_pdf = len(selected_files) > 0 and all(
            map(lambda file_name: file_name.lower().endswith(".pdf"), selected_files))

        only_docx = len(selected_files) > 0 and all(
            map(lambda file_name: file_name.lower().endswith(".docx"), selected_files))

        only_images = len(selected_files) > 0 and all(
            map(lambda file_name: file_name.lower().endswith((".jpeg", ".png", "jpg")), selected_files))

        only_folder = len(selected_files) == 1 and os.path.isdir(
            os.path.join(folder, selected_files[0]))

        window[Event.pdfToDocx].update(disabled=not only_pdf)
        window[Event.docxToPdf].update(disabled=not only_docx)
        window[Event.compress].update(disabled=not only_images)
        window[Event.select_folder].update(disabled=not only_folder)
        window[Event.delete].update(disabled=len(selected_files) == 0)

        if event == Event.open_folder and open_folder != '':
            folder = open_folder
            files = os.listdir(folder)
            window[Event.files].update(files)
        elif event == Event.parent_folder:
            folder = os.path.abspath(os.path.join(folder, os.pardir))
            files = os.listdir(folder)
            window[Event.files].update(files)
        elif event == Event.pdfToDocx:
            for fileName in selected_files:
                pdf = os.path.join(folder, fileName)
                docx = os.path.join(folder, fileName[:-4] + ".docx")
                cv = Converter(pdf)
                cv.convert(docx)
                cv.close()
        elif event == Event.docxToPdf:
            for fileName in selected_files:
                docx = os.path.join(folder, fileName)
                pdf = os.path.join(folder, fileName[:-5] + ".pdf")
                print(docx, pdf)
                convert(docx, pdf)
        elif event == Event.compress:
            for fileName in selected_files:
                path = os.path.join(folder, fileName)
                image = Image.open(path)
                image.save(path, quality=1)
        elif event == Event.select_folder:
            folder = os.path.join(folder, selected_files[0])
            files = os.listdir(folder)
            window[Event.files].update(files)
        elif event == Event.delete:
            for selected_file in selected_files:
                os.remove(os.path.join(folder, selected_file))
            files = os.listdir(folder)
            window[Event.files].update(files)
        elif event == Event.exit or event == gui.WIN_CLOSED:
            break

    window.close()

main()