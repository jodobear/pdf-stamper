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
colors = {"Noir": getColor("black"), "Blanc": getColor("white"),
         "Rouge": getColor("red"), "Bleu": getColor("blue")}

layout = [
    [sg.Text('Pour les tampons numérotés : le nom des fichiers originaux doit commencer par 001, 002,..\nLe numéro de tampon sera identique aux trois premiers caractères du nom de fichier.\nPar exemple, un fichier nommé A14-Défense aura pour numéro de tampon les caractères <<A14>>\nMettre <<Pièce n°>> en 3e ligne le cas échéant.')],
    [sg.Text('Dossier des fichiers pdf à tamponner:',
             size=(35, 2)), sg.InputText(), sg.FolderBrowse("Parcourir", size=(10, 1))],
    [sg.Text('Dossier dans lequel les fichiers tamponnées seront enregistrés:', size=(35, 2)),
            sg.InputText(), sg.FolderBrowse("Parcourir", size=(10, 1))],
    [sg.Text('1re ligne du tampon:', size=(35, 2)), sg.InputText()],
    [sg.Text('2e ligne du tampon:', size=(35, 2)), sg.InputText()],
    [sg.Text('3e ligne du tampon (<<Pièce n°>>)', size=(35, 2)), sg.InputText()],
    [sg.Frame(layout=[
        [sg.Radio('Tamponner toutes les pages', "RADIO1", default=True, size=(35, 1)),
        sg.Radio('Tamponner seulement la page n°', "RADIO1"), sg.InputText(size=(3, 1))],
        [sg.Checkbox('Numéroter les tampons (001, 002..)', size=(35, 1), default=False),
        sg.Text('Police:'), sg.InputCombo((fonts), size=(18, 1)),
        sg.Text('Couleur:'), sg.InputCombo((list(colors.keys())), size=(12, 1))] ],
        title='Options', title_color='red', relief=sg.RELIEF_SUNKEN, tooltip='Use these to set options')],
    [sg.Submit("Tamponner"), sg.Cancel("Annuler")]
    ]

jobNo = 1
window = sg.Window('Tampons PDFs numérotées ou non numérotés', layout)

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

    maxstring = max(len(line_one), len(line_two), (len(line_three)+4))
    leftwidth = 27 + (7 * maxstring)
    formatting = [color, font, overlay]

    if stampAll:
        print("Tamponnage de toutes les pages.")
    else:
        print("Tamponnage de la page n°", f'{pageNo}')
    if stampNumbers:
        print("Tampons numérotés")
    else:
        print('Tampons non numérotés.')
    print("Dossier des fichiers pdf à tamponner", input_path,
        "\nDossier dans lequel les fichiers tamponnées seront enregistrés:", output_path)

    input_files = [f for f in listdir(input_path)
                    if isfile(join(input_path, f))]
    output_files = []

    print('\n', "Fichiers à tamponner: ")
    for i in range(len(input_files)):
        print(i + 1, input_files[i])


    def draw(page):
        '''This function does the stamping.'''
        box = fitz.Rect(page.rect.width - leftwidth,
                        page.rect.height - 65,
                        page.rect.width - 25,
                        page.rect.height - 20)
        page.drawRect(box, color=colors["Noir"], fill=colors["Blanc"], overlay=True)
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

    print('\n', "Fichiers tamponnés:")
    for i in range(len(output_files)):
        print(i + 1, output_files[i])

    print('\n', "Vous pouvez les retrouver dans le dossier:", f"{output_path}", '\n\n')
    jobNo += 1

window.Close()
