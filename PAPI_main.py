#8h0min

#TO DO
#Pular Linha Quando Input de Outro Ano
#Calcular IR
#Criar e Formatar Abas Padrão (INICIO - COM PORCENTAGEM)


import os
os.chdir(r"C:\Users\lucas\OneDrive\Área de Trabalho\PJT_Planilha FII")      #Alterar Diretório de Execução do Script Python

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter                                #Biblioteca para Função de Converte para Letras Índice da Coluna           
from openpyxl.styles import Alignment, Font, Border, Side, numbers          #Biblioteca Para Formatação das Células
from datetime import datetime                                               #Biblioteca Converte string para datetime
#Alignment - Alinhamento
#Font - Font
#Border e Side- Borda
#Numbers - Moeda


#Variavel Global
borda_padrao = Side(border_style="thin", color="000000")
borda_externa = Side(border_style="medium", color="000000")

##### FUNÇÕES PARA CRIAÇÃO E FORMATAÇÃO DE ESTILO #####

#Criar Novo Arquivo e Formatar Abas Padrão
def criar_arquivo(arquivo):
    aba_inicio = arquivo.active                             #Seleciona Aba Ativa - Criada Automaticamente
    aba_inicio.title = "INICIO"                             #Nomeia Aba Criada Automaticamente
    ocultar_grades(aba_inicio)
    
    
    nova_aba = arquivo.create_sheet(title="TOTAL_APORTES")  #Nova Aba para Conter Todos os Aportes
    ocultar_grades(nova_aba)

    #Alterar Largura das Colunas
    nova_aba.column_dimensions['A'].width = 8.8
    nova_aba.column_dimensions['B'].width = 8.8
    alterar_largura_colunas_sequencia(nova_aba, 2, 6, 12.8)

    alterar_altura_linha(nova_aba, 1)

    #Preencher Cabeçalho da Planilha
    nova_aba.cell(row=1, column=1).value = "Ativos"
    nova_aba.cell(row=1, column=2).value = "Cotas"
    nova_aba.cell(row=1, column=3).value = "Preço"
    nova_aba.cell(row=1, column=4).value = "Investido"
    nova_aba.cell(row=1, column=5).value = "Data"

    #Alterção do Estilo e Borda da Célula em Sequência com o Mesmo Padrão
    for linha in range(1,4):
        for celula in nova_aba[linha]:
            alterar_estilo_celula(celula)
            alterar_borda(celula, borda_padrao, borda_padrao, borda_padrao, borda_padrao)


#Criar Aba Nova para FII e Formatar
def criar_aba(arquivo, nome_ativo):
    nova_aba = arquivo.create_sheet(title=nome_ativo)       #Cria Nova Aba

    ocultar_grades(nova_aba)

    #Preencher Cabeçalho da Planilha
    nova_aba.merge_cells('B2:E2')                           #Mesclar Celulas
    nova_aba.cell(row=2, column=2).value = nome_ativo
    nova_aba.cell(row=3, column=2).value = "Cotas"
    nova_aba.cell(row=3, column=3).value = "Preço"
    nova_aba.cell(row=3, column=4).value = "Investido"
    nova_aba.cell(row=3, column=5).value = "Data"

    nova_aba.merge_cells('G2:J2')
    nova_aba.cell(row=2, column=7).value = "IR"
    nova_aba.cell(row=3, column=7).value = "Cotas"
    nova_aba.cell(row=3, column=8).value = "Preço Médio"
    nova_aba.cell(row=3, column=9).value = "Total"
    nova_aba.cell(row=3, column=10).value = "Ano"

    #Alterar Altura das Linhas 1 a 3
    for linha in range(1, 4):
        alterar_altura_linha(nova_aba, linha)

    #Alterar Largura das Colunas
    nova_aba.column_dimensions['A'].width = 2.8
    nova_aba.column_dimensions['B'].width = 8.8
    nova_aba.column_dimensions['F'].width = 4.8

    alterar_largura_colunas_sequencia(nova_aba, 2, 6, 12.8)
    alterar_largura_colunas_sequencia(nova_aba, 7, 11, 14.8)

    #Alterção do Estilo da Célula em Sequência com o Mesmo Padrão
    for linha in range(1,4):
        for celula in nova_aba[linha]:
            alterar_estilo_celula(celula)
    
    for linha in range(7,11):
        for celula in nova_aba[linha]:
            alterar_estilo_celula(celula)

    #Alterção da Borda da Célula Mesclada Para Titulo
    for coluna in range(2, 6):
        coluna_letra = get_column_letter(coluna)
        celula = nova_aba.cell(row=2, column=coluna)
        alterar_borda(celula, borda_externa, borda_externa, borda_externa, borda_externa)
    for coluna in range(7, 11):
        coluna_letra = get_column_letter(coluna)
        celula = nova_aba.cell(row=2, column=coluna)
        alterar_borda(celula, borda_externa, borda_externa, borda_externa, borda_externa)

    #Alterção da Borda da Célula Para SubTitulo
    alterar_borda(nova_aba.cell(row=3, column=2), borda_externa, borda_padrao, borda_externa, borda_externa)
    alterar_borda(nova_aba.cell(row=3, column=3), borda_padrao, borda_padrao, borda_externa, borda_externa)
    alterar_borda(nova_aba.cell(row=3, column=4), borda_padrao, borda_padrao, borda_externa, borda_externa)
    alterar_borda(nova_aba.cell(row=3, column=5), borda_padrao, borda_externa, borda_externa, borda_externa)
    alterar_borda(nova_aba.cell(row=3, column=7), borda_externa, borda_padrao, borda_externa, borda_externa)
    alterar_borda(nova_aba.cell(row=3, column=8), borda_padrao, borda_padrao, borda_externa, borda_externa)
    alterar_borda(nova_aba.cell(row=3, column=9), borda_padrao, borda_padrao, borda_externa, borda_externa)
    alterar_borda(nova_aba.cell(row=3, column=10), borda_padrao, borda_externa, borda_externa, borda_externa)


#Oculta as Linhas de Grade da Aba
def ocultar_grades(aba):
    aba.sheet_view.showGridLines = False


#Alterar Altura da Linha
def alterar_altura_linha(aba, linha):
    aba.row_dimensions[linha].height = 14


#Alterar Largura de Colunas em Sequência com o Mesmo Valor - Reutilização de Código
def alterar_largura_colunas_sequencia(nova_aba, coluna_inicio, coluna_final, tamanho):
    for coluna in range(coluna_inicio, coluna_final):  
        letra_coluna = get_column_letter(coluna)                    #Converte para Letras
        nova_aba.column_dimensions[letra_coluna].width = tamanho    #Alterar Largura de Coluna


#Alterar Estilo da Célula com o Mesmo Padrão - Reutilização de Código
# Para Criação e Adição??
def alterar_estilo_celula(celula):
    celula.alignment = Alignment(horizontal='center', vertical='center')  #Centralizar Texto
    celula.font = Font(name="Arial", size=12, bold=False, italic=False, color="000000")
    #Nome da Fonte; Tamanho; Negrito; Italico; Cor


#Boras
def alterar_borda(celula, left, right, top, bottom):
    celula.border = Border(left=left, right=right, top=top, bottom=bottom)


##### FUNÇÕES PARA ALIMENTAR TABELAS #####

#Obter Linha Vazia
#Vai até ultima que sofreu modificação (max_row)
def ultima_linha(aba):
    for row in range(2, aba.max_row+1):       
        if aba.cell(row=row, column=2).value is None:
            return row
    return aba.max_row + 1 #Se for 0 retorna 1


#Função para Adicionar Dados na Aba       
def adicionar_aporte(planilha, nome_fundo, cotas, valor, data, borda, aba_selecionada):
    #Seleciona Aba Especifica do Ativo ou Aba Geral
    aba = planilha[aba_selecionada]

    #Obter Linha
    linha = ultima_linha(aba)
    print(f'Linha: {linha}')

    #Verificar se é Primeiro Input ou Não e Se Linha Anterior Não É Vazia
    if linha <= 4:
        ano_anterior = "1000"       #Adiciona um Valor Para Comparar Porque é Primeiro Input
        print("1 Input")
    elif aba.cell(row=linha-1, column=2).value is not None:
        #Como Pula então talvez depois ele pega a linha vazia
        #Se conteudo tiver vazio
        print("aqUI")
        linha_anterior = linha - 1
        ano_anterior = aba.cell(row=linha_anterior, column = 5).value       #Pega Ultimo Ano Digitado formato datetime
        ano_anterior = str(ano_anterior)
        ano_anterior = ano_anterior[:4]         #Fatiamento Para Pegar Somente Ano
    else:
        print("oi")
        linha_anterior = linha
        ano_anterior = aba.cell(row=linha_anterior, column = 5).value       #Pega Ultimo Ano Digitado formato datetime
        ano_anterior = str(ano_anterior)
        ano_anterior = ano_anterior[:4]         #Fatiamento Para Pegar Somente Ano
    print(ano_anterior)

    if ano_anterior not in str(data.year):
        print("Ano Diferente")
    else:
        print("Ano Igual")

    # Se Não é Aba Geral e Se É Ano Diferente (Se Ano Digitado é Diferente do Ultimo Adicionado) - Pular uma Linha
    if aba_selecionada != "TOTAL_APORTES" and ano_anterior not in str(data.year) and linha > 4:
        linha += 1
        borda = borda_externa
        borda_topo = borda_externa
        #Formatar Borda do Fim do Ano Anterior
        alterar_borda(aba.cell(row=linha_anterior, column = 2),borda, borda_padrao, borda_padrao, borda_externa)
        alterar_borda(aba.cell(row=linha_anterior, column = 3),borda_padrao, borda_padrao, borda_padrao, borda_externa)
        alterar_borda(aba.cell(row=linha_anterior, column = 4),borda_padrao, borda_padrao, borda_padrao, borda_externa)
        alterar_borda(aba.cell(row=linha_anterior, column = 5),borda_padrao, borda, borda_padrao, borda_externa)    
    else:
        borda_topo = borda_padrao

    #Adicionar Conteudo na Planilha
    if aba_selecionada == "TOTAL_APORTES":
        aba.cell(row=linha, column = 1).value = nome_fundo
        alterar_borda(aba.cell(row=linha, column = 1),borda, borda_padrao, borda_padrao, borda_padrao)
        alterar_estilo_celula(aba.cell(row=linha, column = 1))
    aba.cell(row=linha, column = 2).value = cotas
    aba.cell(row=linha, column = 3).value = valor
    aba.cell(row=linha, column = 3).number_format = 'R$ #,##0.00'
    aba.cell(row=linha, column = 4).value = cotas * valor
    aba.cell(row=linha, column = 4).number_format = 'R$ #,##0.00'
    aba.cell(row=linha, column = 5).value = data
    aba.cell(row=linha, column = 5).number_format = "DD/MM/YYYY"

    #Formatar Estilo após Input
    for coluna in range(2, 6):
        alterar_estilo_celula(aba.cell(row=linha, column = coluna))
        alterar_altura_linha(aba, linha)
    alterar_borda(aba.cell(row=linha, column = 2),borda, borda_padrao, borda_topo, borda_padrao)
    alterar_borda(aba.cell(row=linha, column = 3),borda_padrao, borda_padrao, borda_topo, borda_padrao)
    alterar_borda(aba.cell(row=linha, column = 4),borda_padrao, borda_padrao, borda_topo, borda_padrao)
    alterar_borda(aba.cell(row=linha, column = 5),borda_padrao, borda, borda_topo, borda_padrao)    


######################################################

nome_arquivo = input("Nome da Planilha: ")+ ".xlsx"

#Verificação se Arquivo Existe
if os.path.exists(nome_arquivo):
    arquivo = load_workbook(nome_arquivo) #Carregar Arquivo - sem Macro
    # Como tá na mesma página, não precisa indicar caminho'''
else:
   arquivo = Workbook()                    #Cria Novo Arquivo
   criar_arquivo(arquivo)

nome_ativo = input("ATIVO: ")
nome_ativo = nome_ativo.replace('1',"")     #Retirar 11 se Usuário Informar
nome_ativo = nome_ativo.upper()             #Converter Toda String Para Maiusculo
cotas = 8
valor = 100
data = "20/12/2029"
data = datetime.strptime(data, "%d/%m/%Y")  # Converte String para datetime
if nome_ativo not in arquivo.sheetnames:    #Se Fundo Informado Não Tem Aba
    criar_aba(arquivo, nome_ativo)
adicionar_aporte(arquivo, nome_ativo, cotas, valor, data, borda_externa, nome_ativo)
#adicionar_aporte(arquivo, nome_ativo, cotas, valor, data, borda_padrao, "TOTAL_APORTES")

arquivo.save(nome_arquivo)    #Sobrescreve o Arquivo/ Outro Nome Gera Outro Arquivo