#!usr/bin/python3

import fitz, shutil
from fitz.utils import getColor

from os import listdir
from os.path import isdir, isfile, join

import PySimpleGUI as sg


print = sg.EasyPrint

fonts = ["Helvetica", "Helvetica-Oblique", "Helvetica-Bold",
         "Helvetica-BoldOblique", "Courier", "Courier-Oblique",
         "Courier-Bold", "Courier-BoldOblique", "Times-Roman",
         "Times-Italic", "Times-Bold", "Times-BoldItalic"]
colors = {"Black": getColor("black"), "White": getColor("white"),
         "Red": getColor("red"), "Blue": getColor("blue")}

layout = [
    [sg.Text('For stamps with number: Make sure you have numbered the original documents.\n\t\t\
            Stamp numbers will correspond to the first three letters of the document name.\n\t\t\
            Only put the label for numbers in Line 3.')],
    [sg.Text('Folder to be stamped',
             size=(28, 1)), sg.InputText(), sg.FolderBrowse(size=(10, 1))],
    [sg.Text('Output folder:', size=(28, 1)),
            sg.InputText(), sg.FolderBrowse(size=(10, 1))],
    [sg.Text('1st line of stamp:', size=(28, 2)), sg.InputText()],
    [sg.Text('2nd line of stamp:', size=(28, 2)), sg.InputText()],
    [sg.Text('3rd line or Label for numbering:', size=(28, 2)), sg.InputText()],
    [sg.Frame(layout=[
        [sg.Radio('Stamp all pages    OR', "RADIO1", default=True, size=(18, 1)),
        sg.Radio('Stamp only page', "RADIO1"), sg.InputText(size=(3, 1))],
        [sg.Checkbox('Stamp with numbers', size=(18, 1), default=False),
        sg.Text('Select font:'), sg.InputCombo((fonts), size=(18, 1)),
        sg.Text('Select Color:'), sg.InputCombo((list(colors.keys())), size=(12, 1))] ],
        title='Options', title_color='red', relief=sg.RELIEF_SUNKEN, tooltip='Use these to set options')],
    [sg.Submit("Stamp"), sg.Cancel()]
    ]

jobNo = 1
window = sg.Window('PDF Stamper with or without Numbering', layout)

while True:
    event, value_list = window.Read()
    if event is None or event == 'Cancel':
        break
    
    print("Job No:", jobNo)

    input_path = value_list[0]
    output_path = value_list[1]
    line_one = value_list[2]
    line_two = value_list[3]
    line_three = value_list[4]
    stampAll = value_list[5]
    stampPage = value_list[6]
    pageNo = value_list[7]
    stampNumbers = value_list[8]
    font = value_list[9]
    color = colors[value_list[10]]
    border = True
    overlay = True

    maxstring = 4 + max(len(line_one), len(line_two), len(line_three))
    leftwidth = 26 + (7 * maxstring)
    formatting = [color, font, overlay]

    if stampAll:
        print("Stamping all pages.")
    else:
        print("Stamping only page number:", f'{pageNo}')
    if stampNumbers:
        print("Stamping with numbers.")
    else:
        print('Stamping without numbers.')
    print("To be stamped folder: ", input_path,
        "\nOutput folder: ", output_path)

    input_files = [f for f in listdir(input_path)
                    if isfile(join(input_path, f))]
    output_files = []

    print('\n', "Input Files: ")
    for i in range(len(input_files)):
        print(i + 1, input_files[i])


    def draw(page):
        '''This function does the stamping.'''
        box = fitz.Rect(page.rect.width - leftwidth,
                        page.rect.height - 65,
                        page.rect.width - 25,
                        page.rect.height - 20)
        page.drawRect(box, color=colors["Black"], fill=colors["White"], overlay=True)
        page.insertTextbox(
            box, text, color=color, align=1, fontname=font, border_width=2)

    def output(output_path):
        '''This function moves the files to output folder.'''
        for f in output_files:
            shutil.move(f, output_path)


    for f in input_files:
        doc = fitz.open(f"{input_path}/{f}")
        stamp_numbers = stampNumbers and f[:3].isdigit()
        text = [f"{line_one}", f"{line_two}", f"{line_three}"]

        if stamp_numbers:
            text[2] = f"{line_three} " f"{f[:3]}"
        if stampAll:
            for page in doc:
                draw(page)
        else:
            draw(doc[int(pageNo) - 1])
        
        doc.save(f"{f}")
        output_files.append(f)

    output(output_path)

    print('\n', "Stamped files:")
    for i in range(len(output_files)):
        print(i + 1, output_files[i])

    print('\n', "You can find them in the folder:", f"{output_path}", '\n\n')
    jobNo += 1

window.Close()
