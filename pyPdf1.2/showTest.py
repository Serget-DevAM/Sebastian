from __future__ import print_function
import PyPDF2
import os
from PIL import Image, ImageTk, ImageDraw, ImageFont
import tkinter as tk

import gspread
from gspread.utils import rowcol_to_a1
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Defina as credenciais de acesso do Google Sheets
scope = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
spreadsheet_id = '1MctP90zSnUFq-BjiPUkc2GFs7RwNZXoFsezYfKpkwmo'
range1 = 'Página1!A1:Q136'

def runSheets(localizaNaLinha, localizaCol, Valores):
    print("runSheets")

    #Autentificação
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scope)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'C:/Users/Serget/Downloads/credentials.json', scope)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()
        #Ler
        result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                    range=range1).execute()

        values = result.get('values', [])

        for indice, l in enumerate(localizaNaLinha):
            print(f"localizaNaLinha index {indice} l {l} col {localizaCol[indice]}")
            for row, row_values in enumerate(values):
                for col, cell_value in enumerate(row_values):
                    #print("row ", row,"col ", col)
                    #print("cell_value", cell_value)
                    if cell_value == l:
                        cell_ref = rowcol_to_a1(row + 1, col + 1)
                        print(f"Célula encontrada na linha {row + 1}, coluna {col + 1}")
                        print("cell_value", cell_value) 
                        print("cell_ref", cell_ref)
                        print("linha", row+1)
                        print("cell_ref[0:1]", cell_ref[0:1])
                    if cell_value == l:
                        print()
            #Atualizar
            valueAdd = [["teste"],]
            print("valueAdd", valueAdd)

        #print(linha)
        '''result2 = sheet.values().update(spreadsheetId=spreadsheet_id,
                                    range="Página1!B2", valueInputOption="USER_ENTERED", body={'values': valueAdd}).execute()'''
        values = result.get('values', [])

        '''values2 = result2.set('values', ["2"])'''

        if not values:
            print('No data found.')
            return
        '''print(result)'''
        valores = result['values']
        print("result['values']", result['values'])
        '''print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[1]))'''
    except HttpError as err:
        print("err", err)


    def busca_coord(spreadsheet_id, range_name, id_value):
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name).execute()

        rows = result.get('values', [])

        for i, row in enumerate(rows):
            if row and row[0] == id_value:
                row_index = i + 1  # adiciona 1 ao índice para considerar o cabeçalho da tabela
                print("row_index ", row_index)
                break
        else:
            row_index = None  # se o ID não for encontrado, retorna None
    print("busca_coord", spreadsheet_id, range1)
    busca_coord(spreadsheet_id, range1, "teste")






from PIL import Image, ImageTk, ImageDraw, ImageFont
import tkinter as tk

# Abre a imagem usando o Pillow
# imagem = Image.open('C:/Users/Serget/Pictures/logo-serget-text-left(1).png')
# imagem = Image.open('H:/Meu Drive/Anexos/AutoCatch/logo-serget-text-left.png')

# Redimensiona a imagem para ter largura de 500 pixels e altura proporcional
'''largura = 500
altura = int(largura / imagem.size[0] * imagem.size[1])
imagem = imagem.resize((largura, altura))'''

# Cria uma janela do Tkinter e exibe a imagem
'''janela = tk.Tk()
tkimagem = ImageTk.PhotoImage(imagem)
label = tk.Label(janela, image=tkimagem)
label.pack()'''

# Define a fonte para a legenda
# fonte = ImageFont.truetype('C:/Users/Serget/Documents/anexos/Adrianna-Bold.ttf', 32)
'''fonte = ImageFont.truetype('H:/Meu Drive/Anexos/AutoCatch/Adrianna-Bold.ttf', 32)
fonte = ("Arial", 32)
'''
# Define o texto da legenda
# texto = "AutoCacth Serget - Dev AM || Processando..."
'''label_texto = tk.Label(janela, text=texto, font=fonte)
label_texto.pack(side="bottom")'''

'''
# Cria uma nova imagem que combina a imagem original e a legenda
nova_imagem = Image.new('RGB', (imagem.width, imagem.height + fonte.getsize(texto)[1]), color='white')
nova_imagem.paste(imagem, (0, 0))
desenho = ImageDraw.Draw(nova_imagem)
desenho.text((0, imagem.height), texto, font=fonte, fill=(0, 0, 0))

# Exibe a nova imagem
nova_imagem.show()

'''


def script():

    # directory = "./"  # Listar arquivos no diretório atual
    #   directory = "C:/Users/Serget/Documents/anexos/Enel/"
    #   dir = "C:/Users/Serget/Documents/anexos/Enel/"
    nInscricao = []
    mesesInsercao = []
    valores = []
    registrador_nInscricao = 0
    registrador_mesesInsercao = 0

    directory = "H:/Meu Drive/Anexos/Enel/"
    dir = "H:/Meu Drive/Anexos/Enel/"

    # Verifica diretório
    def list_files(directory):
        return [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]

    files = list_files(directory)

    files2 = list_files(dir)
    print(dir)
    print(files2)

    for f in files2:
        # checa se o arquivo contém enel radar e/ou finaliza em .pdf (é dot tipo pdf)
        checkString = "enel radar"
        ext = ".pdf"
        count = 0
        print(f)
        #print("checkString.upper() in f.upper(): ", checkString.upper() in f.upper())

        #if checkString.upper() in f.upper():
        if f.upper() in f.upper():
            # print("enel radar encontrado")
            count += 1
        #else:
            #print("ext.upper() in f.upper(): ", ext.upper() in f.upper()),
            if ext.upper() in f.upper():
            #if f.upper() in f.upper() :
                # print("é pdf")
                # print(directory + f)
                '''if files:
                    print("Arquivos no diretório:")
                    for file in files:
                        print(file)
                else:
                    print("Não há arquivos no diretório")'''

                # Abra o arquivo PDF
                # pdf_path = 'C:\\Users\\Serget\\Downloads\\658205047373.pdf'
                pdf_path = directory + f

                pdf_file = open(pdf_path, "rb")
                pdf_reader = PyPDF2.PdfReader(pdf_file)

                # Obtenha a quantidade de páginas no PDF
                num_pages = len(pdf_reader.pages)
                # Inicialize uma matriz para armazenar os dados
                data = []

                # Loop através de todas as páginas do PDF
                for page_num in range(num_pages):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    data.append(text)
                    # print(text)

                # print("sequencia")
                # print("data1: ", data[1])
                splitdata0 = []
                splitdata0 = data[0].split()
                print("splitdata0: ", splitdata0)
                splitdata1 = []
                splitdata1 = []
                splitdata1 = data[1].split()
                print("splitdata1: ", splitdata1)


                '''it = 0
                print(it)
                it +=1
                print(splitdata0[34])
                print(it)
                it += 1
                #Nº Nota Fiscal
                print(splitdata0[34])
                print(it)
                it += 1
                #Data Vencimento
                print(splitdata0[22])
                print(splitdata0[23])
                print(":")
                print(it)
                it += 1
                print(splitdata1[20])
                #Total a Pagar
                print(splitdata1[25])
                print(it)
                it += 1

                #CNPJ Contratante
                print(splitdata0[1])
                print(it)
                it += 1'''

                mesAbrev = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
                mesNum = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

                def get_key(vector, target_string):
                    for i, string in enumerate(vector):
                        if string == target_string:
                            return i
                    return -1

                # Nº Nota Fiscal
                nf = splitdata0[34]
                print("NF: ", nf)
                print("F: ", f.replace(".pdf", ""))
                nf = f.replace(".pdf", "")

                # print("splitdata0")
                # after 'PAULO/SP'
                # print(splitdata1[29])
                # print(splitdata0[22])
                # print(splitdata0)

                def ancorar(matriz, ancoragem):
                    for indice, numero in enumerate(matriz):
                        if numero == ancoragem:
                            #print(ancoragem)
                            encontrado = True
                            chave = indice
                            #print(chave)
                            # print(chavevalor)
                            # print(splitdata0[chavevalor])
                            # print(splitdata0[chavevalor][9:])
                            #localizado = matriz[chave][9:]
                            #print("localizado: ", matriz[chave+1])
                            return chave
                            break
                    else:
                        print("não encontrado")

                numeroConta = splitdata1[ancorar(splitdata1, "156")+1] if ancorar(splitdata1, "156") is not None else 0
                numeroInstalacao = splitdata0[ancorar(splitdata0, "PAULO/SP")+1][0:9]
                chavevalor = splitdata0[ancorar(splitdata0, "PAULO/SP")+1]
                chavedia = splitdata0[ancorar(splitdata0, "PAULO/SP")+2]
                # 696000481007
                chavemes = splitdata0[ancorar(splitdata0, "PAULO/SP")+3]
                chaveano = splitdata0[ancorar(splitdata0, "PAULO/SP")+4]
                nf = splitdata0[ancorar(splitdata0, "Venda") - 6]
                print("numeroInstalacao ", numeroInstalacao)
                nInscricao.append(numeroInstalacao)
                for indice, numero in enumerate(splitdata0):
                    if numero == 'PAULO/SP':
                        encontrado = True
                        chavevalor = indice + 1
                        chavedia = indice + 2
                        chavemes = indice + 3
                        chaveano = indice + 4
                        '''print(chavevalor)
                        print(splitdata0[chavevalor])
                        print(splitdata0[chavevalor][9:])'''
                        valor = splitdata0[chavevalor][9:]
                        break


                target_string = splitdata0[chavemes]
                mesesInsercao.append(target_string)
                key = get_key(mesAbrev, target_string)
                valores.append(splitdata1[25])
                nomeclatura = "ENEL RADAR - " + nf + " - R$ " + splitdata1[25] + " - 210043 - " + splitdata0[chavedia] + "." + mesNum[key] + "." + "" + "2023"  # splitdata1[chaveano]
                print(nomeclatura)
                '''
                Chamar func pra buscar n de instalação
                Se encontrado retornar coordenada da linha
                senão, try inserir em nova linha
                    chamar func pra identificar o mês
                    Se encontrado, retornar coordenada da Coluna
                        Chamar func pra inserir                 
                '''


                #if __name__ == '__main__':
                #    runSheets()
                # Feche o arquivo PDF
                pdf_file.close()
                print("\n")
                '''os.rename(pdf_path, directory + nomeclatura + ".pdf")'''

                # abre a planilha pelo ID
                '''gc = gspread.service_account(filename='C:/Users/Serget/Downloads/credentials.json')

                sheet = gc.open_by_key('1MctP90zSnUFq-BjiPUkc2GFs7RwNZXoFsezYfKpkwmo')
                worksheet = sheet.worksheet('Página1')
                cell = worksheet.cell(1, 1)
                cell.value = 'Olá, mundo!'''
                # Carregue a planilha para um objeto do Google Sheets
                # = client.open('ENEL').worksheet('Página1')

                # Defina o ID que será usado para localizar o registro que deve ser atualizado
                '''id_referencia = "595505785696"'''

                # Localize a linha que contém o ID de referência
                '''linha = sheet.find(str(id_referencia)).row'''

                # Atualize os valores das colunas desejadas na linha especificada
                '''sheet.update_cell(linha, 2, 'Novo Valor 1')
                sheet.update_cell(linha, 3, 'Novo Valor 2')
                sheet.update_cell(linha, 4, 'Novo Valor 3')'''

                # Exiba uma mensagem de confirmação
                print('Dados atualizados com sucesso.')

    # Fecha a janela
    print(len(nInscricao))
    print("ninscricao ", nInscricao)
    print(len(mesesInsercao))
    print("mesesInsercao ", mesesInsercao)
    if __name__ == '__main__':
        runSheets(nInscricao, mesesInsercao, valores)
script()

'''    janela.destroy()


# Aguarda 3 segundos antes de fechar a janela
janela.after(5000, script)

# Aguarda a interação do usuário
janela.mainloop()'''


'''
compilar prompt terminal
pyinstaller --noconsole --onefile --icon=LogoSERGET.ico main.py
pyinstaller --onefile --icon=LogoSERGET.ico main.py
pyinstaller --onefile --icon=LogoSERGET.ico -F main.py
pyinstaller --onefile --icon=C:\Users\Serget\Downloads\LogoSERGET.ico -F main.py
pyinstaller --onefile --icon=C:\Users\Serget\Downloads\LogoSERGET.ico main.py 
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
python.exe -m pip install --upgrade pip
pyinstaller -F main.py
Set-ExecutionPolicy RemoteSigned
pip install pyinstaller 
python pypdf2.py build  
pip install cx_freeze
pip install -U pyinstaller

'''