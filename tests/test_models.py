from src.models.sheets import Sheets
import openpyxl

folder_path = 'C:/Users/DESOUR10/Downloads/Conferência'
conferencia = Sheets(folder_path, 'Conferência')
contas = Sheets(folder_path, 'Contas')
dependentes = Sheets(folder_path, 'Dependente')
eSocial = Sheets(folder_path, 'eSocial')
workForce = Sheets(folder_path, 'WorkForce')
checkList = Sheets(folder_path, 'Check')

def test_models_load_sheet_function():
    test_pass = True
    conferencia.load_sheet()
    contas.load_sheet()
    dependentes.load_sheet()
    eSocial.load_sheet()
    workForce.load_sheet()
    checkList.load_sheet()

    if not all([conferencia, contas, dependentes, eSocial,workForce, checkList]):
        test_pass = False
        
    assert test_pass == True
    
def test_models_clone_sheet_function():
    test_pass = True
    test_check = True
    test_checkList = True
    
    checkList, check = Sheets(folder_path, 'Check').clone_sheet('Conferência')
    
    if type(check) is not openpyxl.workbook.workbook.Workbook:
        test_check == False
    if type(checkList) is openpyxl.worksheet.worksheet.Worksheet:
        test_checkList == False
    
    assert test_pass == (test_check and test_checkList)