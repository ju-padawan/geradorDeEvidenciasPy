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

    cenario = str(dados_evidencia[cen]['cen_nome'])
    actions.exibir_informacao_console("==================================================================================")
    actions.exibir_informacao_console("Gerando evidência paro o cenário de teste >> "+cenario[0:21])

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
    actions.inserir_status_colorido("Status execução: ", str(dados_evidencia[cen]['cen_status_execucao']))

    actions.inserir_quebra_de_pagina()

    actions.inserir_informacoes_arquivo_evidencia("Selecionando a Esteira", " ")
    evidencia_esteira = pasta_evidencias+"/esteira.png"
    actions.inserir_imagem_arquivo_evidencia(evidencia_esteira, "\nNão foi possível inserir a evidência da seleção de esteira...", "final")
    
    actions.inserir_informacoes_arquivo_evidencia("Tela inicial do APP", " ")
    evidencia_tela_inico_app = pasta_evidencias+"/inicio_app.png"
    actions.inserir_imagem_arquivo_evidencia(evidencia_tela_inico_app, "\nNão foi possível inserir evidência da tela inicial do APP...", "final")
    
    actions.inserir_informacoes_arquivo_evidencia("Tela de Login do APP", " ")
    evidencia_tela_login = pasta_evidencias+"/login_app.png"
    actions.inserir_imagem_arquivo_evidencia(evidencia_tela_login, "\nNão foi possível inserir evidência da tela de login do APP...", "final")

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

        actions.inserir_imagem_arquivo_evidencia(evidencia, "\nNão foi possível inserir evidência para o "+cen_passo+"...")

        if count1 < int(passos):
            actions.inserir_quebra_de_pagina()

        count1 = count1+1

    nome_cenario = str(dados_evidencia[cen]['cen_nome'])
    actions.salvar_arquivo_evidencia(str(nome_cenario[0:21]), str(dados_evidencia[cen]['pasta_evidencias']), str(dados_evidencia[cen]['cen_status_execucao']), str(dados_evidencia[cen]['cen_plataforma']))
    count = count+1