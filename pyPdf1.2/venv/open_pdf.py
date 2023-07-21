import PyPDF2
import os
from PIL import Image, ImageTk, ImageDraw, ImageFont
import tkinter as tk
from pdf2image import convert_from_path
import cv2
import pytesseract


def script():

    directory = "C:/Users/Serget/Documents/Dev/code.tester/"

    #Verifica diretório
    def list_files(directory):
        return [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]

    files = list_files(directory)

    for f in files:
        # checa se o arquivo contém enel radar e/ou finaliza em .pdf (é dot tipo pdf)
        checkString = "enel radar"
        ext = ".pdf"
        count = 0
        if checkString.upper() in f.upper():
            count += 1
        else:
            if ext.upper() in f.upper():
                ###f = '64300545810' + '.pdf'
                ###f = '542306410967' + '.pdf'
                f = 'A.pdf'
                # Abre o arquivo PDF
                pdf_path = directory+f

                pdf_file = open(pdf_path, "rb")
                pdf_reader = PyPDF2.PdfReader(pdf_file)

                def pdf_to_images(pdf_path):
                    images = convert_from_path(pdf_path)
                    return images
                img = pdf_to_images(pdf_path)
                def process_images(images):
                    for i, image in enumerate(images):
                        # Salvar a imagem em um arquivo temporário
                        image_path = f"image_{i}.jpg"
                        image.save(image_path, "JPEG")

                        # Extrair texto da imagem usando Tesseract OCR
                        recognized_text = pytesseract.image_to_string(image_path)

                        # Mostrar as palavras reconhecidas no console
                        words = recognized_text.split()
                        for word in words:
                            print(word)

                        # Remover o arquivo temporário
                        os.remove(image_path)

                process_images(img)
                print(f)


                # Obtem a quantidade de páginas no PDF
                num_pages = len(pdf_reader.pages)

                # Inicialize uma matriz para armazenar os dados
                data = []

                # Loop através de todas as páginas do PDF
                for page_num in range(num_pages):
                    page = pdf_reader.pages[page_num]
                    print(page)
                    text = page.extract_text()
                    print(text)
                    data.append(text)

                splitdata0 = data[0].split()
                #splitdata1 = data[1].split()
                print("splitdata0: ")
                for n, x in enumerate(splitdata0):
                    print("[",n,"] =", x, end =" "),
                print(" ")
                #print("splitdata1: ")
                #for n, x in enumerate(splitdata1):
                #    print("[",n,"] =", x, end =" "),



                #Nº Nota Fiscal
                #NF = splitdata0[34]

                def ancorar(matriz, ancoragem):
                    for indice, numero in enumerate(matriz):
                        if numero == ancoragem:
                            encontrado = True
                            chave = indice
                            return chave
                            break
                print(" ")
                #print("valor: " + splitdata1[25])



                #numeroConta = splitdata1[ancorar(splitdata1, "156") + 1] if ancorar(splitdata1, "156") is not None else 0
                # checar caso ENEL RADAR - 497764384 - R$ 69,57 - 210043 - 02.05.2023.pdf
                #chavevalor = splitdata0[ancorar(splitdata0, "PAULO/SP") + 1]
                #chavedia = splitdata0[ancorar(splitdata0, "PAULO/SP") + 2]
                #chavemes = ancorar(splitdata0, "PAULO/SP") + 3
                #chaveano = splitdata0[ancorar(splitdata0, "PAULO/SP") + 4]

               # print("chave-valor: " + chavevalor)
                #if ancorar(splitdata1, "156")==None:
                #    valor = chavevalor[9:]
                #    print(valor)
                #else:
                #    valor = splitdata1[ancorar(splitdata1, "156")+10]
                #    print("ancorar(splitdata1, 156)+10: ", splitdata1[ancorar(splitdata1, "156")+10])

                #print("chave-dia: " + chavedia)

                ##print("ancorar(splitdata1, 156): " + ancorar(splitdata1, "156"))
                #print(ancorar(splitdata1, "156") == None)

                mesAbrev = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
                mesNum = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

                def get_key(vector, target_string):
                    for i, string in enumerate(vector):
                        if string == target_string:
                            return i
                    return -1

                #ninscricao = f.replace(".pdf", "")

                #nf = splitdata0[ancorar(splitdata0, "Venda") - 6]
                #print("nf: " + nf)

                #target_string = splitdata0[chavemes]
                #key = get_key(mesAbrev, target_string)

                #print("mesNum[key]: " + mesNum[key])

                #nomeclatura = "ENEL RADAR - " + nf + " - R$ " + valor + " - 210043 - " + chavedia + "." + mesNum[key] + "." + "" + "2023"

                # Feche o arquivo PDF
                pdf_file.close()

                #os.rename(pdf_path, directory+nomeclatura+".pdf")
                ###break
#script()
