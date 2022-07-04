import json
import win32com.client as win32


with open(f'{input(r"Informe o caminho do arquivo: ")}\{input("Informe o nome do arquivo: ")}.json', 'r') as arq:
    arq_json = json.loads(arq.read())


name = arq_json["full_name"].split(" ")
date = arq_json["date"]
date = date.split(" ")
hour = date[1]
date = date[0]
date = date.split("-")
date = f'{date[2]}/{date[1]}/{date[0]}'

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
print('\nConta logada ao Outlook selecionada para o envio da mensagem!')

email.To = input("\nInforme o email para enviar a mensagem: ")
email.Subject = 'Confirmação da compra'
email.Body = f"""
Ola {name[0].capitalize()}, a sua compra no valor de {arq_json["amount"]} foi confirmada dia {date} as {hour}.

Informações completas da compra:

ID: {arq_json["id"]}                 
CPF: {arq_json["user_cpf"]}                 
Nome Completo: {arq_json["full_name"]}                
Valor: {arq_json["amount"]}                
E-mail: {arq_json["e-mail"]}               
Data: {date}              
Hora: {hour}             
"""

email.Send()

print('\nfim')
