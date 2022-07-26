from PyPDF2 import PdfFileReader
import fitz
import pyautogui
from tkinter import *
from tkinter.filedialog import *
import os
from tqdm import tqdm

#funções
def ExportarPDF():
    nomeSalvarPDF = asksaveasfilename(filetypes=[('PDF Files', '*.pdf')])
    tipo = '.pdf'
    while tipo not in nomeSalvarPDF:
        pyautogui.alert(text="Prezado humano, extensão .pdf não informada.", title='Aviso', button="OK")
        nomeSalvarPDF = asksaveasfilename(filetypes=[('PDF Files', '*.pdf')])
    return nomeSalvarPDF

def ListarPDF():
    diretorioComPDF = askdirectory()
    while diretorioComPDF == '':
        pyautogui.alert(text="Prezado humano, nenhuma pasta foi selecionada.", title='Aviso', button="OK")
        diretorioComPDF = askdirectory()
    return diretorioComPDF

pyautogui.alert(text="Prezado humano, selecione a pasta com um ou mais arquivos em PDF, com páginas frente e verso.", title='Aviso', button="OK")

#Palavra chave pra identificar se é a FRENTE do documento
palavraChave = "DATA DE EMISSÃO" 

pastaSelecionada = ListarPDF()
caminhos = [os.path.join(pastaSelecionada, nome) for nome in os.listdir(pastaSelecionada)]
arquivos = [arq for arq in caminhos if os.path.isfile(arq)]
pdf = [arq for arq in arquivos if arq.lower().endswith(".pdf")]

#Inicio do programa
if __name__ == "__main__":
    #verificando se possui arquivos pdf no diretótio
    if pdf != []:
        print('-----------------------------------------------------------------------')
        for i in pdf:
            caminhoDoPDFEntrada = i

            #Abre o arquivo PDF
            arquivoPDFOriginal = open(caminhoDoPDFEntrada, mode="rb") 

            #Leitura do PDF
            lerPDF = PdfFileReader(arquivoPDFOriginal)

            totalPaginasPDF = lerPDF.numPages
            print('Lendo PDF {}, total páginas = {}'.format(caminhoDoPDFEntrada, totalPaginasPDF))
            print('-----------------------------------------------------------------------')

            PDFEstadoPaginas = list()

            #Loop em cada pagina do PDF
            for i in tqdm(range(0, totalPaginasPDF)):
                paginaEmObjeto = lerPDF.getPage(i)
                textoNaPagina = paginaEmObjeto.extractText()
                if palavraChave in textoNaPagina:
                    PDFEstadoPaginas.append("FRENTE")
                else:
                    PDFEstadoPaginas.append("VERSO")

            listaDePaginas = list(range(0, totalPaginasPDF))
            selecaoDePaginas = list()

            estadoAnterior = None
            pagina = 0
            for estado in PDFEstadoPaginas:
                estadoAtual = estado
                
                if(estadoAtual != estadoAnterior):
                    selecaoDePaginas.append(pagina)

                estadoAnterior = estadoAtual
                pagina = pagina + 1


            if selecaoDePaginas == listaDePaginas:
                arquivoPDFOriginal.close()
                
            else:
                #Caminho do PDF de saída, para salvar o novo arquivo gerado
                caminhoDoPDFSaida = ExportarPDF()
                novoArquivoPDF = fitz.open(caminhoDoPDFEntrada)
                novoArquivoPDF.select(selecaoDePaginas)
                novoArquivoPDF.save(caminhoDoPDFSaida)
                novoArquivoPDF.close()
                arquivoPDFOriginal.close()
    else:
        pyautogui.alert(text='Não foi localizado nenhum arquivo pdf nesse diretório...', title='Aviso', button="Fechar")
        quit()
#fim do programa
pyautogui.alert(text='Finalizado!', title='Aviso', button="Fechar")