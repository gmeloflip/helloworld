import re
import os
import smtplib
import openpyxl
import email
from email.message import EmailMessage
from time import sleep
import pandas as pd
import pywhatkit as kit
import pyautogui

class ToDo:

    def iniciar(self):
        self.lista_tarefas = []
        self.email_destino()
        self.menu()
        self.criar_planilha()
        sleep(1)
        self.enviar_email()
        self.enviar_Wwats()

    def email_destino (self):
        while True:
            self.email = str(input('Email de destino: ')).lower()
            
            padrao_email = re.search(
            '^[a-z0-9._]+@[a-z0-9]+.[a-z]+(.[a-z]+)?$',self.email
            )
            
            if padrao_email:
                print('Email valido!!!')
                break
            else:
                print('Tente outro email, este é invalido')

    def menu(self):
        while True:
            menu_principal = int(input("""
            MENU PRINCIPAL
            [1] CADASTRAR
            [2] VISUALIZAR
            [3] SAIR
            OPÇÃO: """))
    
            match menu_principal:
                case 1: self.cadastrar()
                case 2: self.visualizar()
                case 3: break
                case _: print('\nOpção invalida!')

    def cadastrar(self):
        while True:
            self.tarefa = str(input('Digite uma tarefa ou [S] para sair: ')).capitalize()

            if self.tarefa == 'S': 
                break
            else:
                self.lista_tarefas.append(self.tarefa) #usa o self para acessar coisas dentro de outros metodos
                try:
                    with open('./src/Tarefas/historico-tarefas.txt', 'a', encoding='utf8') as arquivo: #precisa se atentar a deixar a barra pro outro lado / (o encoding='utf8' é para deixar o texto com os caracteres em portugues)
                        arquivo.write(f'{self.tarefa}\n')
                
                except FileExistsError as e:
                    print(f'Erro: {e}')
    
    def visualizar(self):
        try:
            with open ('./src/Tarefas/historico-tarefas.txt', 'r', encoding= 'utf8') as arquivo:
                print(arquivo.read())

        except FileExistsError as e:
            print(f'Erro: {e}')

    def criar_planilha(self):
        if len(self.lista_tarefas) > 0:
            try:
                df = pd.DataFrame({"Tarefas": self.lista_tarefas})    #data frame means tabela (python converte para qualquer coisa)
                self.nome_arquivo = str(input('Nome do arquivo: '))

                if self.nome_arquivo[-5:] == '.xlsx':
                    df.to_excel(f'./src/Tarefas/{self.nome_arquivo}', index=False) #index false é para tirar a primeira coluna com os numeros
                
                else:
                    df.to_excel(f'./src/Tarefas/{self.nome_arquivo}.xlsx', index=False)

                print('\nPlanilha criada com sucesso!')

            except Exception as e:
                print(f'Erro: {e}')
       

        else:
            print('\nLista de tarefas vazias.')

    def enviar_email(self):
        endereco = 'gmeloflip@gmail.com'
        senha = ''

        with open('./src/Senha/Senha.txt','r', encoding='utf8') as arquivo:
            s = arquivo.readlines() #epara ler linhas especificas
        
        senha = s[0] #serve para localizar qual a linha contem a senha

        msg = EmailMessage()
        msg['From'] = endereco
        msg ['To'] = self.email
        msg ['Subject'] = 'Chegou a planilha de tarefas!'
        msg.set_content(
            'Planilha em anexo.'
        )
        arquivos = [f'./src/Tarefas/{self.nome_arquivo}.xlsx']

        for arquivo in arquivos:
            with open(arquivo, 'rb') as arq: #rb é pra que a maquina leia do jeito certo o anexo
                dados = arq.read()
                nomes_arquivo = arq.name

            msg.add_attachment(
                dados,
                maintype='application',
                subtype='octet-stream',
                filename=nomes_arquivo
            ) #padrão para anexar arquivos

        server = smtplib.SMTP_SSL('smtp.gmail.com', 465) #transações de email da internet
        server.login(endereco, senha, initial_response_ok=True)
        server.send_message(msg)
        print ('Email enviado com sucesso.')

#https://myaccount.google.com/apppasswords para configurar o gmail a mandar emails
        

    def enviar_Wwats(self):
        try:
            numero_destino = '+5511964451910'
            mensagem = 'Teste de mensagem\nteste'

            kit.sendwhatmsg_instantly(numero_destino, mensagem, wait_time=15)

        except Exception as e:
            print(f'Erro: {e}')
















start = ToDo()
start.iniciar()







