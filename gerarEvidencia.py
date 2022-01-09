import docx
from docx.text.run import *
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import yaml
import os
from support.actions import SupportActions

actions = SupportActions()

dados_evidencia  = actions.ler_dados_arquivo_yaml()

#Deescobrindo quantos cenários foram executados
lista_cen = len(dados_evidencia)

#Percorrer a lista de cenários
count = 1
while count <= lista_cen:
    #Descobrindo quantos campos tem em cada cenário
    cen = 'cenario_'+str(count)
    lista_dados = len(dados_evidencia[cen])
    pasta_evidencias = dados_evidencia[cen]['pasta_evidencias']

    actions.formatacao_arquivo_evidencia()

    #actions.inserir_cabecalho_arquivo_evidencia(0, str(dados_evidencia[cen]['cen_nome']), "simples")
    actions.inserir_titulo_arquivo_evidencia(str(dados_evidencia[cen]['cen_nome']))

    actions.inserir_informacoes_arquivo_evidencia("Descrição: ", str(dados_evidencia[cen]['cen_desc']))
    actions.inserir_informacoes_arquivo_evidencia("Pré-requisitos: ", str(dados_evidencia[cen]['cen_pre_requisitos']))
    actions.inserir_informacoes_arquivo_evidencia("Resultado esperado: ", str(dados_evidencia[cen]['resultado_esperado']))
    actions.inserir_informacoes_arquivo_evidencia("Plataforma: ", str(dados_evidencia[cen]['cen_plataforma']))
    actions.inserir_informacoes_arquivo_evidencia("Disposiitivo uilizado: ", str(dados_evidencia[cen]['cen_dispositivo']))
    actions.inserir_informacoes_arquivo_evidencia("Versão do Software: ", str(dados_evidencia[cen]['cen_versao_software']))
    actions.inserir_informacoes_arquivo_evidencia("Versão do App Next: ", str(dados_evidencia[cen]['cen_versao_app']))
    actions.inserir_informacoes_arquivo_evidencia("Executado por: ", str(dados_evidencia[cen]['cen_executor']))
    actions.inserir_informacoes_arquivo_evidencia("Massa uilizada: ", str(dados_evidencia[cen]['cen_massa']))
    actions.inserir_informacoes_arquivo_evidencia("Data da execução: ", str(dados_evidencia[cen]['cen_data_execucao']))
    actions.inserir_informacoes_arquivo_evidencia("Status execução: ", str(dados_evidencia[cen]['cen_status_execucao']))

    actions.inserir_quebra_de_pagina()

    #percorrer os campos de cada cenário
    passos = (lista_dados - 13)
    count1 = 1
    while count1 <= int(passos):
        cen_passo = 'passo_'+str(count1)
        evidencia = pasta_evidencias+'/'+cen_passo+'.png'
        txt_passo = '[Passo_'+str(count1)+']: '
                
        actions.inserir_espaco_antes_paragrafo(2)
        actions.inserir_informacoes_arquivo_evidencia(txt_passo, str(dados_evidencia[cen][cen_passo]))
        actions.inserir_informacoes_arquivo_evidencia("Resultado:", " ")
        actions.inserir_espaco_apos_paragrafo(4)
        actions.inserir_imagem_arquivo_evidencia(evidencia)

        if count1 < int(passos):
            actions.inserir_quebra_de_pagina()

        count1 = count1+1

    nome_cenario = str(dados_evidencia[cen]['cen_nome'])
    path_doc = str(dados_evidencia[cen]['pasta_evidencias'])+'/'+nome_cenario[0:21]+' - '+str(dados_evidencia[cen]['cen_plataforma'])+'.docx'
    actions.salvar_arquivo_evidencia(str(nome_cenario[0:21]), str(path_doc))
    count = count+1