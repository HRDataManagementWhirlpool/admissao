from tkinter import filedialog

from src.models.sheets import Sheets
from src.controllers.sheets import SheetsController

def main():
    folder_path = filedialog.askdirectory()
    conferencia = Sheets(folder_path, 'Conferência').load_sheet()
    contas = Sheets(folder_path, 'Contas').load_sheet()
    dependentes = Sheets(folder_path, 'Dependente').load_sheet()
    eSocial = Sheets(folder_path, 'eSocial').load_sheet()
    workForce = Sheets(folder_path, 'WorkForce').load_sheet()
    checkList, check = Sheets(folder_path, 'Check').clone_sheet('Conferência')

    if not all([conferencia, contas, dependentes, eSocial,workForce, checkList]):
        print('Algumas planilhas não foram encontradas. Verifique os nomes dos arquivos.')
        return
    
    resConferidos = [] # Armazena o RE para evitar duplicidade
    for re in checkList['B'][2:]: # Para linha na Coluna B a partir da terceira linha:
        nomesConferidos = [] # Armazena os nomes dos dependentes do RE que está sendo conferido para evitar duplicidade
        if re.value is not None: # Caso o valor do RE não seja vazio
            for celula in conferencia['A'][1:]: # Verifica se o valor do RE a ser conferido é igual ao RE dentro da tabela
                if celula.value == re.value: # Caso seja igual
                    if re.value not in resConferidos: # E não esteja duplicado
                        resConferidos.append(re.value)
                        array = SheetsController.get_data(celula.row, conferencia, contas, dependentes, eSocial, workForce, re.value, nomesConferidos)
                        SheetsController.fill_sheet(checkList, re.row, array)
    SheetsController.savefile(folder_path, check)

if __name__ == "__main__":
    main()
