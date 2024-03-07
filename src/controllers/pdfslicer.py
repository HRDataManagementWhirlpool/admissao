from pikepdf import Pdf
import os

class PdfController:
    def slice_and_rename_from_data(file, mensalists):
        with Pdf.open(file) as pdf:
            if (len(pdf.pages)) % (len(mensalists)) == 0: # Verificação para que o número de contratos seja compatível com o número de mensalistas
                div = int(len(pdf.pages) / len(mensalists)) # Armazena o número de páginas por contrato
                for i in range(0, len(pdf.pages), div):
                    new = Pdf.new() # Cria um novo arquivo pdf
                    for j in range(div): # Insere as páginas do contrato nesse novo arquivo pdf
                        new.pages.append(pdf.pages[i+j])
                    title = (f'CONTRATO DE TRABALHO - {mensalists[int(i/div)]['nome']} - {mensalists[int(i/div)]['re']}.pdf') # Renomeando o arquivo no padrão solicitado
                    savefile = os.path.join(r'files', title) # Armazenando o arquivo na pasta especificada
                    new.save(savefile) # Salva o arquivo e finaliza o processo
            else:
                return "Número de colaboradores incompatível com o número de contratos!"