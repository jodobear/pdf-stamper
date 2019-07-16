#!/usr/bin/python3

import fitz, shutil
from fitz.utils import getColor

from os import listdir
from os.path import isdir, isfile, join

import PySimpleGUI as sg

print = sg.EasyPrint

help_lisc_input = "This software is liscenced under GPL v2, (c) Jodobear 2019.\n\nNOTE: Input path to the folder containing only pdf files to be stamped. Do not provide path to the file itself."
help_num = "FOR STAMPS WITH NUMBERS: Stamp numbers correspond to first three characters of the document name.\nFor example, a document named <<A14-Defence>> will be stamped with number <<A14>>.\nOptionally, input label (e.g. 'Exhibit No.') in Line 3 to get <<Exhibit No. A14>>.\n"
help_thanks = "Thank you for using this software. If you have any issues or want to contribute to the development, visit\nhttps://github.com/jodobear/pdf-stamper"
help_donation = "Please consider buying us a coffee at https://tallyco.in/jodobear \n"

fonts = ["Helvetica", "Helvetica-Oblique", "Helvetica-Bold",
         "Helvetica-BoldOblique", "Courier", "Courier-Oblique",
         "Courier-Bold", "Courier-BoldOblique", "Times-Roman",
         "Times-Italic", "Times-Bold", "Times-BoldItalic"]
colors = {"Black": getColor("black"), "White": getColor("white"),
         "Red": getColor("red"), "Blue": getColor("blue")}

layout = [
    [sg.T("Please click 'Help' for instructions.")],
    [sg.T('Folder to be stamped:',
             size=(18, 1)), sg.In(), sg.FolderBrowse(size=(10, 1))],
    [sg.T('Output folder:', size=(18, 1)),
     sg.InputText(), sg.FolderBrowse(size=(10, 1))],
    [sg.T('1st line of stamp:', size=(18, 2)), sg.In()],
    [sg.T('2nd line of stamp:', size=(18, 2)), sg.In()],
    [sg.T('3rd line of stamp:', size=(18, 2)), sg.In()],
    [sg.Frame(layout=[
        [sg.Radio('Stamp all pages    OR', "RADIO1", default=True, size=(18, 1)),
         sg.Radio('Stamp only page', "RADIO1"), sg.In(size=(3, 1))],
        [sg.Checkbox('Stamp with numbers', size=(18, 1), default=False),
         sg.T('Font:'), sg.InputCombo((fonts), size=(18, 1)),
         sg.T('Color:'), sg.InputCombo((list(colors.keys())), size=(6, 1))]],
        title='Options', title_color='red', relief=sg.RELIEF_SUNKEN, tooltip='Use these to set options')],
    [sg.Submit("Stamp", size=(8, 1)),
    sg.Cancel(size=(8, 1)), sg.T(' ' * 66),
    sg.Help(size=(10, 1), button_color=('black', 'orange'))]
    ]

window = sg.Window('PDF Stamper: with or without numbering', layout)

jobNo = 1

while True:
    event, value_list = window.Read()
    if event is None or event == 'Cancel':
        break

    if event == 'Help':
        sg.Popup('Help for pdfstamper_win0.2-en',
                 help_lisc_input, help_num, help_thanks, help_donation)
        continue

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

    maxstring = 4 + max(len(line_one), len(line_two), len(line_three))
    leftwidth = 26 + (7 * maxstring)

    # debug output of the selected job.
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

    # debug output of input files.
    print('\n', "Input Files: ")
    for i in range(len(input_files)):
        print(i + 1, input_files[i])


    def draw(page):
        '''This function draws the stamp.'''
        box = fitz.Rect(page.rect.width - leftwidth,
                        page.rect.height - 65,
                        page.rect.width - 25,
                        page.rect.height - 20)
        page.drawRect(box, color=colors["Black"], fill=colors["White"], overlay=True)
        page.insertTextbox(
            box, text, color=color, align=1, fontname=font, border_width=2)

    # calling draw and stamping.
    for f in input_files:
        doc = fitz.open(f"{input_path}/{f}")
        text = [f"{line_one}", f"{line_two}", f"{line_three}"]

        if stampNumbers:
            text[2] = f"{line_three} " f"{f[:3]}"
        if stampAll:
            for page in doc:
                draw(page)
        else:
            draw(doc[int(pageNo) - 1])
        doc.save(f"{f}")
        output_files.append(f)


    def output(output_path):
        '''This function moves the files to output folder.'''
        for f in output_files:
            shutil.move(f, output_path)

    output(output_path)

    # debug output of processed files and output folder.
    print('\n', "Stamped files:")
    for i in range(len(output_files)):
        print(i + 1, output_files[i])
    print('\n', "You can find them in the folder:", f"{output_path}", '\n\n')
    jobNo += 1

window.Close()
