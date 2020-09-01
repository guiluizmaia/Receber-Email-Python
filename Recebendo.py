#Importações necessárias
import imaplib
import email
from email.header import decode_header

EMAIL = 'SeuEmail@gmail.com'
SENHA = 'SuaSenha'
#Aqui é o servidor do gmail, procure o
#do seu email na internet
SERVER = 'imap.gmail.com'

#Abrindo conexão com o servidor do email
eM = imaplib.IMAP4_SSL(SERVER)
eM.login(EMAIL, SENHA)
#Aqui selecionamos a caixa que queremos extrair
#E receberemos todos os status dos emails
#Seu status e seus ids
#O data é a lista com os ids
status, messages =eM.select('inbox')

#Número de emails que você quer pegar de cima para
#baixo
N = 3
contador = 1
#Total do número de emails
messages = int(messages[0])
N -= 1
#Vamos fazer uma varredura de cima pra baixo
for i in range(messages, messages-N, -1):
    print("-------------------------------------------------------------------------")
    contador += 1
    print("Email: ", contador)
    #Buscar no email pelo ID
    res, msg = eM.fetch(str(i), "(RFC822)")
    #O id vem em uma tupla(Parecido com lista porem
    # é imutavel)
    for resposta in msg:
        #O isinstance chega se é o tipo de dado
        #que entra como parâmetro
        if isinstance(resposta, tuple):
            #Transforma os bytes do email 
            #para um objeto mensagem
            #Colocamos o 1 pois a mensagem está na segunda
            #seção da tupla, a primeira é o cabeçalho
            msg = email.message_from_bytes(resposta[1])
            #Usaremos o decode_header para decodificar o email
            tema = decode_header(msg["Subject"])[0][0]
            if isinstance(tema, bytes):
                #Se for byte transformar em String
                tema = tema.decode()
            #Remetente
            de = msg.get("From")
            print(f"De: {de}")
            print(f"Tema: {tema}")

            #Agora veremos se o email tem anexo
            #ou seja, se é Multipart
            if msg.is_multipart():
                #Andaremos por todas as parte do email
                for parte in msg.walk():
                    #Extrair o conteudo e o tipo do email
                    tipo_cont = parte.get_content_type()
                    cont = str(parte.get("Content-Disposition"))
                    #Colocaremos um try apenas para evitar
                    #eventuais erros
                    try:
                        #Pegar o corpo do email
                        #Ativa o decode para decodificar
                        corpo = parte.get_payload(decode=True).decode()
                    except:
                        pass
                    #Conferindo se o tipo do conteudo é text
                    # e se não tem anexo
                    if tipo_cont == "text/plain" and "attachment" not in cont:
                        print(corpo)
                    #Se tiver anexo
        print("-------------------------------------------------------------------------")
#Fechando o email
eM.close()
eM.logout() 
                    
