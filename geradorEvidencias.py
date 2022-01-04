from docx import Document
from docx.shared import Inches
import yaml


with open('data/dados.yaml') as arquivo:
    dados = yaml.load(arquivo, Loader=yaml.FullLoader)

#Deescobrindo quantos cenários foram executados
lista_cen = len(dados)
print(lista_cen)

#Percorrer a lista de cenários
count = 1
while count <= lista_cen:
    #Descobrindo quantos campos tem em cada cenário
    cen = 'cenario_'+str(count)
    lista_dados = len(dados[cen])
    print(lista_dados)

    cen_nome = dados[cen]['cen_nome']
    cen_desc = dados[cen]['cen_desc']
    cen_pre_requisitos = dados[cen]['cen_pre_requisitos']
    pasta_evidencias = dados[cen]['pasta_evidencias']

    print(str(cen_nome))
    print(str(cen_desc))
    print(str(cen_pre_requisitos))
    print(str(pasta_evidencias))

    doc = Document()
    doc.add_heading(cen_nome, 0)

    p = doc.add_paragraph('Descrição: ')
    p = doc.add_paragraph(str(cen_desc))
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.space_after = Inches(0.35)

    p = doc.add_paragraph('Pré-requisios: ')
    p = doc.add_paragraph(str(cen_pre_requisitos))
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.space_after = Inches(0.35)

    #percorrer os campos de cada cenário
    passos = (lista_dados - 4)/2
    count1 = 1
    print(int(passos))
    while count1 <= int(passos):
        cen_passo = 'passo_'+str(count1)
        cen_resultado = 'resultado_'+str(count1)

        evidencia = pasta_evidencias+'/'+cen_passo+'.png'

        txt_passo = '[Passo_'+str(count1)+']: '

        p = doc.add_paragraph(txt_passo).bold = True
        p = doc.add_paragraph((str(cen_passo)))
        p.paragraph_format.left_indent = Inches(0.5)


        p = doc.add_paragraph('Resultado esperado: ').bold = True
        p = doc.add_paragraph((str(cen_resultado)))
        p.paragraph_format.left_indent = Inches(0.5)

        doc.add_picture(evidencia, width=Inches(6.25))

        print(dados[cen][cen_passo])
        print(dados[cen][cen_resultado])
        count1 = count1+1


    path_doc = str(dados[cen]['pasta_evidencias'])+'/teste-evidencia.docx'
    doc.save(path_doc)

    count = count+1

