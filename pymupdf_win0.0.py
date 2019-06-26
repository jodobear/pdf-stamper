
import fitz, shutil
from fitz.utils import getColor

from os import listdir
from os.path import isdir, isfile, join

import PySimpleGUI as sg

print = sg.EasyPrint

layout = [[sg.Text('For stmps with number: Make sure you have numbered the original docs. \n \
    Stamp numbers will correspond to the first three letters of the doc name. \n \
    Only put "Exhibit" in Line 3.')],
          [sg.Text('Folder to be stamped (docs numbered 001...pdf, 002...pdf, ...)',
                   size=(30, 2)), sg.InputText(), sg.FolderBrowse()],
          [sg.Text('Output folder:', size=(30, 2)),
           sg.InputText(), sg.FolderBrowse()],
          [sg.Text('1st line of stamp:', size=(30, 2)), sg.InputText()],
          [sg.Text('2nd line of stamp:', size=(30, 2)), sg.InputText()],
          [sg.Text('3rd line of stamp:', size=(30, 2)), sg.InputText()],
          [sg.Checkbox('Stamp all pages of pdf', size=(25, 1), default=False)],
          [sg.Submit(), sg.Cancel()]]

window = sg.Window('PDF Stamper with or without Counter', layout)

event, value_list = window.Read()

input_path = value_list[0]
output_path = value_list[1]
line_one = value_list[2]
line_two = value_list[3]
line_three = value_list[4]
stampAll = value_list[5]

maxstring = max((len(line_one)), (len(line_two)), (len(line_three) + 4))
leftwidth = maxstring*7+26

black = getColor("black")
white = getColor("white")
red = getColor("red")

print("Stamp all pages selected:", stampAll)
print("To be stamped folder: ", input_path, '\n' \
      "Output folder: ", output_path, '\n')

input_files = [f for f in listdir(input_path) if isfile(join(input_path, f))]

print("Input Files: ",'\n')
for i in range(len(input_files)):
    print(i + 1, input_files[i])

output_files = []

def draw(doc):
    if not stampAll:
        doc = doc[0]
    for page in doc:
        box = fitz.Rect(page.rect.width - leftwidth,
                        page.rect.height - 65,
                        page.rect.width - 25,
                        page.rect.height - 20)
        page.drawRect(box, color=black, fill=white, overlay=True)
        page.insertTextbox(
            box, text, color=black, align=1, fontname="Courier", border_width=2)

for f in input_files:
    text = [f"{line_one}", f"{line_two}", f"{line_three} {f[:3]}"]
    doc = fitz.open(f"{input_path}/{f}")
    draw(doc)
    doc.save(f"{f}")
    output_files.append(f)

def output(output_path):
    for f in output_files:
        if f.endswith(".pdf"):
            shutil.move(f, output_path)

output(output_path)

print("Stamped files:")
for i in range(len(output_files)):
    print(i + 1, output_files[i])


while True:
    event, values = window.Read()
    if event is None or event == 'Cancel':
        break
