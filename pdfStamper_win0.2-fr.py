#!/usr/bin/python3

import fitz, shutil
from fitz.utils import getColor

from os import listdir
from os.path import isdir, isfile, join

import PySimpleGUI as sg

print = sg.EasyPrint

help_lisc_input = "Ce logiciel est soumis à la licence GPL v2, (c) Jodobear 2019.\n\nMentionner uniquement le dossier comprenant les fichiers pdf."
help_num = "Pour tampons numérotés: Le numéro de tampon correspond au trois premiers caractères du nom de fichier.\nPar exemple, un fichier nommé <<A14-Défense>> aura pour tampon <<A14>>.\nMettre éventuellement <<Pièce n°>> en 3e ligne pour obtenir <<Pièce n° A14>>.\n"
help_thanks = "Merci d'utiliser ce logiciel. Si vous avez des problèmes ou voulez contribuer à son développement, voir\nhttps://github.com/jodobear/pdf-stamper"
help_donation = "Pour faire un don au développeur du projet, voir https://tallyco.in/jodobear \n"

fonts = ["Helvetica", "Helvetica-Oblique", "Helvetica-Bold",
         "Helvetica-BoldOblique", "Courier", "Courier-Oblique",
         "Courier-Bold", "Courier-BoldOblique", "Times-Roman",
         "Times-Italic", "Times-Bold", "Times-BoldItalic"]
colors = {"Noir": getColor("black"), "Blanc": getColor("white"),
         "Rouge": getColor("red"), "Bleu": getColor("blue")}

layout = [
    [sg.T("Appuyer sur 'Aide' pour instructions.")],
    [sg.T('Dossier à tamponner:',
             size=(18, 1)), sg.In(), sg.FolderBrowse("Parcourir", size=(10, 1))],
    [sg.T("Dossier d'enregistrement:", size=(18, 1)),
     sg.InputText(), sg.FolderBrowse("Parcourir", size=(10, 1))],
    [sg.T('1e ligne du tampon:', size=(18, 2)), sg.In()],
    [sg.T('2e ligne du tampon:', size=(18, 2)), sg.In()],
    [sg.T('3e ligne du tampon:', size=(18, 2)), sg.In()],
    [sg.Frame(layout=[
        [sg.Radio('Tamponner toutes les pages', "RADIO1", default=True, size=(24, 1)),
         sg.Radio('Tamponner la page n°', "RADIO1"), sg.In(size=(3, 1))],
        [sg.Checkbox('Numéroter les tampons', size=(18, 1), default=False),
         sg.T('Police:'), sg.InputCombo((fonts), size=(18, 1)),
         sg.T('Couleur:'), sg.InputCombo((list(colors.keys())), size=(6, 1))]],
        title='Options', title_color='red', relief=sg.RELIEF_SUNKEN, tooltip='Utiliser ceci pour régler les options')],
    [sg.Submit("Tamponner"),
    sg.Cancel("Annuler"), sg.T(' ' * 76),
    sg.Help("Aide", size=(10, 1), button_color=('black', 'orange'))]
    ]

window = sg.Window('Tampon 2 PDF : tampons numériques numérotés et non numérotés', layout)

jobNo = 1

while True:
    event, value_list = window.Read()
    if event is None or event == 'Annuler':
        break

    if event == 'Aide':
        sg.Popup('Aide pour pdfstamper_win0.2-fr',
                 help_lisc_input, help_num, help_thanks, help_donation)
        continue

    print("Numéro de tache:", jobNo)

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
        print("Tamponnage de toutes les pages.")
    else:
        print("Tamponnage de la page n°:", f'{pageNo}')
    if stampNumbers:
        print("Tampons numérotés.")
    else:
        print('Tampons non numérotés.')
    print("Dossier à tamponner: ", input_path,
        "\nDossier d'enregistrement: ", output_path)

    input_files = [f for f in listdir(input_path)
                    if isfile(join(input_path, f))]
    output_files = []

    # debug output of input files.
    print('\n', "Fichiers à tamponner: ")
    for i in range(len(input_files)):
        print(i + 1, input_files[i])


    def draw(page):
        '''This function draws the stamp.'''
        box = fitz.Rect(page.rect.width - leftwidth,
                        page.rect.height - 65,
                        page.rect.width - 25,
                        page.rect.height - 20)
        page.drawRect(box, color=colors["Noir"], fill=colors["Blanc"], overlay=True)
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
    print('\n', "Fichiers tamponnés: ")
    for i in range(len(output_files)):
        print(i + 1, output_files[i])
    print('\n', "Les fichiers tamponnées se trouvent ici: ", f"{output_path}", '\n\n')
    jobNo += 1

window.Close()
