import sys
from tkinter import filedialog
sys.path.append(".")
from src.models.sheets import Sheets
from src.controllers.sheets import SheetsController

def test_models_load_sheet_function():
    test_pass = True
    folder_path = filedialog.askdirectory()
    conferencia = Sheets(folder_path, 'Conferência').load_sheet()
    contas = Sheets(folder_path, 'Contas').load_sheet()
    dependentes = Sheets(folder_path, 'Dependente').load_sheet()
    eSocial = Sheets(folder_path, 'eSocial').load_sheet()
    workForce = Sheets(folder_path, 'WorkForce').load_sheet()
    checkList, check = Sheets(folder_path, 'Check').clone_sheet('Conferência')

    if not all([conferencia, contas, dependentes, eSocial,workForce, checkList]):
        test_pass = False
        return
        
    assert test_pass == True