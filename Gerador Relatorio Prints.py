

import glob
import os
from tkinter import font # Seleciona fonte
from docx import Document # Usado para criar e salvar docx
from docx.shared import Inches # Seleciona o tamanho das imagens em polegadas
from docx.enum.text import WD_ALIGN_PARAGRAPH # Alinhar textos
from docx.shared import Pt # Seleciona o tamanho da fonte
from PIL import Image

#################### Funções ####################




#################################################

# Cria um documento em DOCX na memoria
document = Document('Template de arquivo/Papel Timbrado Linea News.docx')

# Seleciona fonte e seu tamanho
style = document.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(11)

print('Iniciando\n')

# Adiciona as imagens na lista files
files = []
images = glob.iglob('prints/*.png', recursive=True)

for filepath in images:
    filename = os.path.basename(filepath)
    print('Carregando o arquivo ' + filename)
    files.append(filename)


for images in files:
    image = Image.open("prints/"+images)
    image.rotate(90, expand=True).save("prints/"+"rotacionado_"+images)


# Gera o relatorio
imgCounter = 0
numImgs = len(files)

while(imgCounter < numImgs):
    document.add_picture('prints/'+"rotacionado_"+files[imgCounter], width=Inches(5.5))
    imgCounter += 1

# Salva documento no PC
document.save('Relatorio Prints.docx')

print("\nDocumento docx gerado.\n")

input("Pressione ENTER para fechar")
















































































