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

    for re in checkList['B'][2:]:
        resConferidos = []
        nomesConferidos = []
        if re.value is not None:
            for celula in conferencia['A'][1:]:
                if celula.value == re.value:
                    if re.value not in resConferidos:
                        resConferidos.append(re.value)
                        array = SheetsController.get_data(celula.row, conferencia, contas, dependentes, eSocial, workForce, re.value, nomesConferidos)
                        SheetsController.fill_sheet(checkList, re.row, array)
    SheetsController.savefile(folder_path, check)

if __name__ == "__main__":
    main()
