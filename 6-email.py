import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

# Dados do email
password = open('senha', 'r').read()
from_email = 'tatianalimongi@gmail.com'
to_email = 'tatianalimongi@gmail.com'
subject = 'Automação da Planilha'
body = '''
Olá. Segue em anexo a automação da planilha para a empresa X.

Qualquer dúvida estou a disposição!
'''

# Montar a estrutura do email
message = EmailMessage()
message['From'] = from_email
message['To'] = to_email
message['Subject'] = subject

message.set_content(body)
safe = ssl.create_default_context()

# Adicionar anexo
anexo = 'teste.xlsx'
mime_type, mime_subtype = mimetypes.guess_type(anexo)[0].split('/')
with open(anexo, 'rb') as a:
    message.add_attachment(
        a.read(),
        maintype = mime_type,
        subtype = mime_subtype,
        filename = anexo
    )

# Enviar email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context = safe) as smtp:
        smtp.login(from_email, password)
        smtp.sendmail(
            from_email,
            to_email,
            message.as_string()
        )
