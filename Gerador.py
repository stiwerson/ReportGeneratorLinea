import glob
from tkinter import font # Usado para listar imagens e adicionar todas no docx
from docx import Document # Usado para criar e salvar docx
from docx.shared import Inches # Seleciona o tamanho das imagens em polegadas
from docx.enum.text import WD_ALIGN_PARAGRAPH # Alinhar textos
from docx.shared import Pt # Seleciona o tamanho da fonte

print('Cobrinha marota ta trabalhando agora...\n')

# Cria um documento em DOCX na memoria
document = Document()

# Seleciona fonte e seu tamanho
style = document.styles['Normal']
font = style.font
font.name = 'Arial'
font.size = Pt(12)

# Coloca titulo no relatório
p1 = document.add_paragraph('''Relatório fotográfico dos pontos, referente sistema de informação TV Prefeitura
Nota fiscal número: XXX  emissão XXX Empenho nº XXX
Contrato nº XXX – Processo licitatório nº XXX\n
''')
p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
p1.style = document.styles['Normal']


# Adiciona as imagens e notifica quais foram adicionadas no console
for filename in glob.iglob('imagens/*.jpeg', recursive=True):
    print('Adicionado '+filename)
    document.add_picture(filename, width=Inches(1.9))

# Salva documento no PC
document.save('Relatorio.docx')

print("\nDocumento docx gerado.")

input("Pressione ENTER para fechar")