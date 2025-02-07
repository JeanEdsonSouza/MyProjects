from distutils.log import ERROR
import openpyxl as op
import sqlite3
from turtle import window_width
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A3
from curses import window
from optparse import Values
import back
from colorama import Cursor
from PySimpleGUI import PySimpleGUI as sg

sg.theme('Dark Purple3')
def janela_login():
    layout =[
        [sg.Text('Bem-Vindo(a)')],
        [sg.Image(r'/home/jean/Desktop/Art_Doce 1.2/Art_Doce.png')],
        [sg.Text('usuario'), sg.Input(key = 'usuario',size=(15,1))],
        [sg.Text('senha'), sg.Input(key = 'senha', password_char = '*',size=(15,1))],
        [sg.Button('entrar')], [sg.Button('cadastrar novo usuario')],
    ]
    return sg.Window('login', layout=layout ,finalize=True)

def janela_novo():
    layout =[
        [sg.Text('Novo Usuario')],
        [sg.Image(r'/home/jean/Desktop/Art_Doce 1.2/Art_Doce.png')],
        [sg.Text("nome"), sg.Input(key = 'Nome', size =(15,1))],
        [sg.Text("senha"), sg.Input(key = 'Senha', size = (15,1))],
        [sg.Button('Criar', key = 'Criar')],
        [sg.Button('voltar')]
    ]
    return sg.Window('Criar', layout = layout,resizable=True  ,finalize=True)

resultado = back.read_task()

def janela_consulta():
    layout =[
        [sg.Text('Consulta')],
        [sg.Text('Produtos')],
        [sg.Listbox(resultado, size=(125,20), key= '-BOX-')],
        [sg.Button('Atualizar'), sg.Button('Enviar p/PDF', key = 'pdf') ,sg.Button('voltar'),sg.Button('pagina cadastro')],
    ]
    return sg.Window('Consulta', layout=layout,resizable=True  ,finalize=True)


def janela_cadastro():
    layout =[
        [sg.Text('Pagina de Cadrasto')],
        [sg.Image(r'/home/jean/Desktop/Art_Doce 1.2/Art_Doce.png')],
        [sg.Text('nome do produto'),sg.Input(key ='nome', size = (15,1))],
        [sg.Text('codigo do produto'), sg.Input(key='codigo',size=(15,1))],
        [sg.Text('preço R$:'), sg.Input(key ='preço', size = (15,1))],
        [sg.Text('quantidade'), sg.Input(key = 'quantidade', size = (3,1))],
        [sg.Text('validade :'), sg.Input(key = 'validade', size = (10,1))],
        [sg.Button('Salvar')],
        [sg.Button('voltar')], [sg.Button('Pagina Consulta')]    
    ]
    return sg.Window('Cadastrar', layout = layout,resizable=True ,finalize=True)

janela1,janelac,janela2,janela3= janela_login(),None,None,None

while True:
    window,event,values = sg.read_all_windows()
    if window == janela1 and event == sg.WINDOW_CLOSED:
        break
    if window == janela1 and event == 'cadastrar novo usuario':
        janelac = janela_novo()
        janela1.hide()
    if window == janelac and event == sg.WINDOW_CLOSED:
        break
    if window == janelac and event == 'Criar':
        banco = sqlite3.connect('login.db')
        cursor = banco.cursor()
        nome = values ['Nome']
        senha = values ['Senha']
        #cursor.execute("CREATE TABLE IF NOT EXISTS login1(nome text, senha text)")
        cursor.execute("INSERT INTO login VALUES('"+nome+"','"+senha+"')")
        banco.commit()
        banco.close()
    if window == janelac and event == 'voltar':
        janelac.hide()
        janela1.un_hide()
    if window == janela1 and event == 'entrar':
        nome = values['usuario']
        senha = values ['senha']
        banco = sqlite3.connect('login.db')
        c = banco.cursor()
        try:
            c.execute("SELECT senha FROM login WHERE nome ='{}'".format(nome))
            senha_bd = c.fetchall()
            banco.close() 
        except:
            sg.popup("ERRO AO FAZER LOGIN!")
        if senha == senha_bd[0][0] :
             janela1.hide()
             window.find_element('senha').Update('')
             janela2=janela_cadrasto()
        else:
            sg.popup("Senha ou Login Invalidos!")
    if window == janela2 and event == 'voltar':
        janela1.un_hide()
        janela2.hide() 
    if window == janela2 and event == sg.WINDOW_CLOSED:
         break
    if window == janela2 and event == 'Pagina Consulta':
        janela2.hide()
        janela3 = janela_consulta()
    if window == janela3 and event == 'pagina cadastro':
        janela3.hide() 
        janela2=janela_cadastro()
    if window == janela3 and event == 'Atualizar':
        resultado = back.read_task()
        window.find_element('-BOX-').Update(resultado)
    if window == janela3 and event == 'pdf':
        resultado = back.read_task()
        y = 0
        pdf = canvas.Canvas("Art_Doce.pdf")
        pdf.setFont("Times-Bold",25)
        pdf.drawString(200,800, "Produtos Cadastrados:")
        pdf.setFont("Times-Bold",18)

        pdf.drawString(10,770, "NOME")
        pdf.drawString(110,770, "CODIGO")
        pdf.drawString(210,770, "PREÇO")
        pdf.drawString(310,770, "QUANT")
        pdf.drawString(410,770, "VALIDADE")

        for i in range(0,len(resultado)):
            y = y - 50
            pdf.drawString(10,270 - y, str(resultado[i][0]))
            pdf.drawString(110,270- y, str(resultado[i][1]))
            pdf.drawString(210,270 - y, str(resultado[i][2]))
            pdf.drawString(310,270 - y, str(resultado[i][3]))
            pdf.drawString(410,270 - y, str(resultado[i][4]))

        pdf.save()
        sg.popup("PDF Salvo com sucesso")
    if window == janela3 and event =='voltar':

        janela3.hide()
        janela2.un_hide()
    if window == janela2 and event == 'pagina cadastro':
        janela2 = janela_cadastro()
        janela3.hide()
    if window == janela3 and event == sg.WINDOW_CLOSED:
        break
    if window == janela2 and event == 'Salvar':
        banco = sqlite3.connect('Art_Doce.db')

        cursor = banco.cursor()

        #cursor.execute("CREATE TABLE Art_Doce(nome text, codigo integer, preço numeric,quantidade integer, validade text)")
        nome = values  ['nome']
        codigo = values  ['codigo']
        preço = values  ['preço']
        quant = values ['quantidade']
        validade = values ['validade']
        cursor.execute("INSERT INTO Art_Doce VALUES('"+nome+"', '"+codigo+"', '"+preço+"','"+quant+"','"+validade+"')")
        banco.commit()
        banco.close()
        #--------------------#
        nome = values ['nome']
        codigo = values['codigo']
        preço = values ['preço']
        quantidade = values ['quantidade']
        validade = values ['validade']
        book = op.load_workbook('Art_Doce.xlsx')
        art_doce = book['Art_Doce']
        #art_doce.append(['Nome','Codigo','Preço','Quantidade','Validade'])      
        art_doce.append([nome,codigo,preço,quantidade,validade])      
        
        for rows in art_doce.iter_rows(min_row=2,max_row=30):
            for cell in rows:
                if cell.value == '':
                    cell.value==resultado
        book.save('Art_Doce.xlsx')
        sg.popup('Excell Salvo com Sucesso!')
        sg.popup("Salvo com Sucesso!")
        window.find_element('nome').Update('')
        window.find_element('codigo').Update('')
        window.find_element('preço').Update('')
        window.find_element('quantidade').Update('')
        window.find_element('validade').Update('')
