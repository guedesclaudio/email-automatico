"""
Created on Tue May 10 17:39:58 2022
@author: claud
"""

import smtplib
import email.message
import pandas as pd


def EnviaEmail():
    
    diretorio = input('Digite o caminho da sua pasta: ') 
    tabela = pd.read_excel(diretorio)
    assuntoEmail = input('Digite o assunto desejado: ')
    meuEmail = input('Digite seu email: ')
    password = input('Digite a sua senha: ')
    
    
    
    x = 0
        
    while x < len(tabela):
        cliente = tabela['NOME'][x]
        emailCliente = tabela['EMAIL'][x]
             
            
        corpo_email = f"""
        <p>Prezado, {cliente}.</p>
        <p>Tudo bem?</p>
        """
        
        msg = email.message.Message()
        msg['Subject'] = assuntoEmail
        msg['From'] = meuEmail
        msg['To'] = emailCliente
        password = password
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )
        
        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('')
        print(f'Email enviado para {cliente} com sucesso')
        print('-----------------------------------------')
            
        x+=1


EnviaEmail()
