import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurar detalhes do servidor de e-mail (lembrando que o email não deve ter autenticador 2 fatores)
email_de_envio = 'seu email'
senha = 'sua senha'
servidor_smtp = 'smtp.outlook.com' # Verifique o servidor smtp do seu email
porta_smtp = 587

# Configurar o corpo do e-mail
assunto = 'Assunto do Email'
texto_do_email = '''
Um belo texto para você fazer e enviar aos envolvidos :)
'''

link = 'anexado no final do texto' #Opcional, caso queira anexar algum link de Forms ou direcionar a pessoa para algum site, etc.
corpo_do_email = texto_do_email + link

# Carregar dados do Excel
df = pd.read_excel('caminho/do/seu/arquivo', sheet_name= 0)  # Caso sua planilha não tenha um nome, atribua o sheet_name para 0 (inicio da coluna). Caso tenha nome, substitua o '0' pelo nome real da sua planilha

# Iterar sobre os destinatários
for index, row in df.iterrows():
    destinatario = row['email']  # Supondo que a coluna com os e-mails seja chamada 'email'

    # Configurar e-mail
    msg = MIMEMultipart()
    msg['From'] = email_de_envio
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Adicionar corpo do e-mail
    msg.attach(MIMEText(corpo_do_email, 'plain'))

    # Configurar servidor SMTP
    servidor = smtplib.SMTP(servidor_smtp, porta_smtp)
    servidor.starttls()

    # Fazer login e enviar e-mail
    servidor.login(email_de_envio, senha)
    servidor.sendmail(email_de_envio, destinatario, msg.as_string())
    servidor.quit()

    print(f'E-mail enviado para {destinatario}')