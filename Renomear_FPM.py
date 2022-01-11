from genericpath import exists
import os
import shutil
from openpyxl import Workbook, load_workbook
import time

def main():

    #Coletar informações dos colaboradores na base de dados
    Planilha_base_de_dados = load_workbook("C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Colaboradores\\Colaboradores.xlsx")
    Aba = Planilha_base_de_dados.active

    funcionários = []
    matriculas = []
    funções = []
    equipes= []

    
    for celula_nome in Aba['B']:  
        linha_nome = celula_nome.row
        nome = str(Aba["B{}".format(linha_nome)].value)
        if nome == "Funcionário":
            time.sleep(0.00001)
        else:
            funcionários.append(nome)

    for celula_matricula in Aba['A']:  
        linha_matricula = celula_matricula.row
        matricula = str(Aba["A{}".format(linha_matricula)].value)
        if matricula == "Matricula":
            time.sleep(0.00001)
        else:
            matriculas.append(matricula)

    for celula_função in Aba['C']:  
        linha_função= celula_função.row
        função = str(Aba["C{}".format(linha_função)].value)
        if função == "Função":
            time.sleep(0.00001)
        else:
            funções.append(função)

    for celula_equipe in Aba['D']:  

        linha_equipe = celula_equipe.row
        equipe = str(Aba["D{}".format(linha_equipe)].value)
        if equipe == "Equipe":
            time.sleep(0.00001)
        else:
          equipes.append(equipe)

    FPM = ['180020','180007','180006','180112','180107', '180108', '180113', '180114', '180115', '180116', '180117', '180118', '180128', '180129', '180131', '180133', '180136', '180137', '180138', '180139', '180141']

    Paginas = str(input("Qual o nº de Páginas?\n"))
    a = 0
    b = 0
    c = 0
    j = 0
    contador = ""

    while contador != Paginas:
        for i in range (10):
            n = str('{}{}{}'.format(a,b,c))
            c = c+1
            j = int(n)-1
            cont = j
            if cont < 10:
                contador = str('{}{}{}'.format(a,b,cont))
                if contador == Paginas:
                    break
                elif n == '000':
                    time.sleep(0.0000001)
                else:
                    old_file = "C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Espelho de Ponto\\2021\\DEZ2021 - 16a31\\FMP 16 – 31 – DEZ –_{}.pdf".format(n)
                    new = "C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Espelho de Ponto\\2021\\DEZ2021 - 16a31\\FMP 16 – 31 – DEZ – {}.pdf".format(funcionários[matriculas.index(FPM[j])])
                    os.rename(old_file, new)

            elif (cont >= 10) and (cont < 100):
                contador = str('{}{}'.format(a,cont))
                if contador == Paginas:
                    break
                else:
                    old_file = "C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Espelho de Ponto\\2021\\DEZ2021 - 16a31\\FMP 16 – 31 – DEZ –_{}.pdf".format(n)
                    new = "C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Espelho de Ponto\\2021\\DEZ2021 - 16a31\\FMP 16 – 31 – DEZ – {}.pdf".format(funcionários[matriculas.index(FPM[j])])
                    os.rename(old_file, new)
        c = 0    
        b = b+1

main()