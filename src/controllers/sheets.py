import os
import datetime

class SheetsController:
    
    def format_date(date_string):
        return datetime.date(int(date_string[:4]), int(date_string[5:7]), int(date_string[8:10]))
    
    def fill_sheet(worksheet, row, array):
        for col in worksheet[f'AC{row}':f'AZ{row}']:
            for i, cell in enumerate(col):
                cell.value = array[i]
                
    def savefile(folder_path, workbook):
        now = datetime.datetime.now().strftime("%d-%m")
        filename = f'Att-{now}.xlsx'
        savefile = os.path.join(folder_path, filename)
        workbook.save(savefile)
        print('Processo concluído. Arquivo salvo com sucesso.')
        
    def workbook_data(workbook, init, reValue, col):
        for cell in workbook[init][1:]:
            row = cell.row
            if reValue in str(cell.value):
                data = str(workbook[f'{col}{row}'].value)
                return data
            
    
    def get_data(linha, conferencia, contas, dependentes, eSocial, workForce, reValue, nomesConferidos):
        dtInicio = SheetsController.format_date(str(conferencia[f'C{linha}'].value))
        vinculo = conferencia[f'AF{linha}'].value
        codCategoria = conferencia[f'AL{linha}'].value
        codVinculo = conferencia[f'AM{linha}'].value
        codExposicao = conferencia[f'AN{linha}'].value
        agencia = SheetsController.workbook_data(contas, 'A', reValue, 'D')[-4:]
        conta = SheetsController.workbook_data(contas, 'A', reValue, 'E')
        dependente = 0
        for cell in dependentes['A'][1:]:
            row = cell.row
            if cell.value == reValue:
                nome = dependentes[f'L{row}'].value
                if nome not in nomesConferidos:
                    dependente+=1
                    nomesConferidos.append(nome)
        saude = conferencia[f'P{linha}'].value
        saudeDt = SheetsController.format_date(str(conferencia[f'S{linha}'].value))
        vida = conferencia[f'V{linha}'].value
        vidaDt = SheetsController.format_date(str(conferencia[f'X{linha}'].value))
        fgts = conferencia[f'AG{linha}'].value
        pFgts = conferencia[f'AH{linha}'].value
        infotipo = '-'
        turno = conferencia[f'K{linha}'].value
        cargaHoraria = str(conferencia[f'J{linha}'].value)
        cargaHoraria = f'{cargaHoraria[:-2]} hs'
        salario = conferencia[f'AE{linha}'].value
        sindicato = conferencia[f'AB{linha}'].value
        gremio='-'
        if conferencia[f'A{linha}'].value == conferencia[f'A{linha+1}'].value:
            gremio = conferencia[f'AB{linha+1}'].value
        periculosidade = conferencia[f'AK{linha}'].value

        es = SheetsController.workbook_data(eSocial, 'J', reValue, 'G')
        wf = SheetsController.workbook_data(workForce, 'BD', reValue, 'CB')
        cargo = SheetsController.workbook_data(workForce, 'BD', reValue, 'BJ')
        
        varArray = [
                            dtInicio,
                            cargo,
                            vinculo,
                            codCategoria,
                            codVinculo,
                            codExposicao,
                            agencia,
                            conta,
                            dependente,
                            saude,
                            saudeDt,
                            vida,
                            vidaDt,
                            fgts,
                            pFgts,
                            infotipo,
                            turno,
                            cargaHoraria,
                            salario,
                            gremio,
                            sindicato,
                            periculosidade,
                            es,
                            wf
                        ]
        return varArray