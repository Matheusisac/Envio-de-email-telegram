from email import message
import matplotlib.pyplot as plt
import time
import telebot
import win32com.client
import pythoncom
import os

email = " "
CHAVE_API = "TOKEN"
bot = telebot.TeleBot(CHAVE_API)
Problema = " "

@bot.message_handler(commands=["clear"])
def clean(mensagem):
    global email
    global Problema
    email = " "
    Problema = " "
    texto = (""" Informações inseridas apagadas
    Escolha a opção:
    Abrir chamado /NovoChamado """) 
    bot.send_message(mensagem.chat.id, texto)

def envio(mensagem):
    if((Problema == "Internet") or (Problema == "Computador") or (Problema == "Impressora") or (Problema == "SEI")):
        if((mensagem.text != "/start") and (mensagem.text != "/NovoChamado") and (mensagem.text != "/Impressora") and (mensagem.text != "/Computador") and (mensagem.text != "/Internet") and (mensagem.text != "/SEI")and (mensagem.text != "/GAAL")and (mensagem.text != "/Movimentacao")and (mensagem.text != "/GT")and (mensagem.text != "/enviar")):
            return True
    else: 
        return False

@bot.message_handler(func=envio)
def enviar(mensagem):
    mail = win32com.client.Dispatch("outlook.application",pythoncom.CoInitialize())
    email = mail.CreateItem(0)
    email.To = email
    email.Subject = "O Usuário "+ mensagem.chat.first_name + " " + mensagem.chat.last_name +" está com problemas com " + Problema
    email.HTMLBody = "<p>" + mensagem.text + "<p>"
    email.Send()
    texto = "Chamado realizado com sucesso"
    bot.send_message(mensagem.chat.id,texto)


@bot.message_handler(commands=["Computador"])
def pc(mensagem):
    global Problema
    Problema = "Computador"
    texto = "Escreva o corpo da mensagem"
    bot.send_message(mensagem.chat.id,texto)

@bot.message_handler(commands=["SEI"])
def sei(mensagem):
    global Problema
    Problema = "SEI"
    texto = "Escreva o corpo da mensagem"
    bot.send_message(mensagem.chat.id,texto)


@bot.message_handler(commands=["Impressora"])
def impressora(mensagem):
    global Problema
    Problema = "Impressora"
    texto = "Escreva o corpo da mensagem"
    bot.send_message(mensagem.chat.id,texto)

@bot.message_handler(commands=["Nudes"])
def envi(mensagem):
    bot.send_photo(mensagem.chat.id,"http://www.planetscott.com/img/3498/large/ecuadorian-ground-dove-(columbina-buckleyi).jpg")
    bot.send_message(mensagem.chat.id,"Essa é minha rolinha")

@bot.message_handler(commands=["GT"])
def Chamado(mensagem):
    if(mensagem.chat.type == "private"):
        global email
        email = "SEU EMAIL"
        texto = """Clique no problema: 
        - Problema de conexão com a internet /Internet
        - Problemas ao ligar o Computador /Computador
        - Problemas no SEI /SEI
        - Problemas com impressora /Impressora
        """
    else:
        texto = "Atendimento apenas no privado"

    bot.send_message(mensagem.chat.id,texto)

@bot.message_handler(commands=["NovoChamado"])
def Chamado(mensagem):
    if(mensagem.chat.type == "private"):
        texto = """Clique no problema: 
        - Para Gerencia de Tecnologia /GT
        """
    else:
        texto = "Atendimento apenas no privado"

    bot.send_message(mensagem.chat.id,texto)



@bot.message_handler(commands=["start"])
def responder(mensagem):
    if(mensagem.chat.type == "private"):
        texto = ("Ola " + mensagem.from_user.first_name + " " + mensagem.chat.last_name +"""
        Escolha a opção:
        Abrir chamado /NovoChamado """) 
    else:
        texto = "Atendimentos apenas no privado"
    print(mensagem)
    bot.send_message(mensagem.chat.id, texto)

try:
    bot.polling(none_stop=True)
except:
    time.sleep(3)
