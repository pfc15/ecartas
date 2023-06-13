# autor: Pedro Cruz (pfc15)
# contato: pedrofcrux@gmail.com

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
#from email.mime.base import MIMEBase
#from email import encoders
#from email.mime.image import MIMEImage
import datetime
import pandas as pd
from tkinter import *
import tkinter.filedialog as fd


'''
Ainda não comentei tudo, se vc quiser decifra isso, mas quando tiver a versão completa vou comentar!!!!
'''
server = ''
email_de = assunto = corpo = ''
def res_pag1():
    global email_de
    email_de = e1.get()
    login(email_de, str(e_senha.get()))
    print(email_de)
    #(e_senha.get())
    pag2()

def envia(corpo, dic, email_de, assunto):
    lista = []
    chaves = list(dic.keys())
    temp = {}
    for e in range(len(dic[chaves[0]])):
        for k, v in dic.items():
            temp[k] = v[e]
        lista.append(temp.copy())
        temp.clear()
    for d in lista:
        c = corpo.format(**d)
        assun = assunto.format(**d)
        email(email_de, d['email'], assunto=assun, corpo=c)

def res_pag2():
    #lista de destinatarios
    caminho = e1.get()
    arquivo = pd.read_excel(caminho)
    envio = arquivo.to_dict(orient='list')
    # corpo
    caminho = e2.get()
    corpo = ''
    with open(caminho, 'r', encoding='utf8') as file:
        corpo = file.read()
    # assunto
    assunto = e3.get()

    envia(corpo, envio, email_de, assunto)
    
    
    #for end in envio['email']:
    #    email(email_de, end, assunto=assunto, corpo=corpo)

def pag2():
    global l_exp_texto, l1_texto, l2_texto
    l_exp_texto.set('agora coloque o caminho dos arquivos necessários:')
    l1_texto.set('lista de destinatários:')
    l2_texto.set('corpo do email:')
    e1.delete(0, END)
    
    e2.grid(row=4, column=0)
    b1.grid(row=2, column=1)
    b2.grid(row=4, column=1)
    l3.grid(row=5, column=0)
    e3.grid(row=6, column=0)
    b3.grid(row=10, column=1)


    submit.config(command=res_pag2, text='enviar emails')
    


def teste():
    global email_de
    assunto = e3.get()
    caminho = e2.get()
    corpo = ''
    with open(caminho, 'r', encoding='utf8') as file:
        corpo = file.read()
    email(email_de, email_de, assunto='teste: '+assunto, corpo=corpo)

def browse1():
    e1.delete(0, END)
    filename = fd.askopenfilename(filetypes=(("xlsx","*.xlsx"),("All files","*.*")))
    e1.insert(END, filename) 

def browse2():
    e2.delete(0, END)
    filename = fd.askopenfilename(filetypes=(("txt","*.txt"),("All files","*.*")))
    e2.insert(END, filename) 


def login(email_de, senha):
    global server
    #senha = 'tjgwguzdxllfledb'
    #senha = 'qbmoxdicvsqyeqxq'
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_de, senha)


def email(email_de ,email_para , assunto = '', corpo=''):
    print(datetime.datetime.now().time())
    global server
     
    # configurando o cabeçalho do email
    msg = MIMEMultipart()
    msg['From'] = email_de
    msg['to'] = email_para
    msg['Subject'] = assunto

    # corpo do email
    msg.attach(MIMEText(corpo, 'plain'))
    #fp = open('img.jpeg', 'rb')
    #img = MIMEImage(fp.read())
    #fp.close()
    html = """

    <h1>Email incrivel!</h1>
    <p>estou usando um html baum zaum</p>
    <img src="img.jpeg" alt="">

    """
    #msgText = MIMEText(html, 'html')
    #msg.attach(img)
    #nome_arquivo = 'manifesto_comunista.pdf'
    #anexo = open(nome_arquivo, 'rb')
    #msg.attach(msgText)
    
    text = msg.as_string()
    # enviando o email
    server.sendmail(email_de, email_para, text)
    print(f'email enviado com sucesso!!\nde:{email_de}\npara:{email_para}\nassunto:{assunto}')
    print('-='*30)



if __name__ == '__main__':
    root = Tk()
    # elementos
    l_exp_texto = StringVar()
    l1_texto = StringVar()
    l2_texto = StringVar()
    l_exp_texto.set('Bem vindo ao enviador de emails em masssa!')
    l1_texto.set('email de: ')
    l2_texto.set('sua senha: ')
    #root.geometry('600x250')
    l_exp = Label(root, textvariable=l_exp_texto).grid(row=0, column=0)
    l1 = Label(root, textvariable=l1_texto).grid(row=1, column=0)
    e1= Entry(root, width=20)
    e1.grid(row=2, column=0)
    l2 = Label(root, textvariable=l2_texto).grid(row=3, column=0)
    e_senha = Entry(root, show='*', width=20)
    e_senha.grid(row=4, column=0)
    e2 = Entry(root)
    e3 = Entry(root)
    l3 = Label(root, text='Assunto: ')
    submit = Button(root, text='login', command=res_pag1)
    submit.grid(row=10, column=0)
    b1  = Button(root, text='browse', command=browse1)
    b2 = Button(root, text='browse', command=browse2)
    b3 = Button(root, text='teste', command=teste)
    #global senha, email_de
    #print(senha, email_de)
    root.mainloop()
if server:
    server.quit()
print('tchau')
