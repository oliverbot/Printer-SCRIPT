# -------------------------------------------

# Importando dependências necessárias.

# Manipulação de CSV.
import csv

# Manipulação das pastas e arquivos gerados pelo script.
import os
import shutil

# Módulos de data e hora.
import time
import datetime
from datetime import date

# Manipulação dos dados.
import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter

# Dependências para envio do e-mail.
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# -------------------------------------------

# Variáveis de data.
timestr0 = time.strftime("%d-%m-%Y")
timestr1 = time.strftime("%Y-%m-%d")

# Abre o arquivo de log das impressões.
# Remove caracteres desnecessários e realiza limpeza da formatação.
# Reescreve em novo arquivo.
with open('papercut-print-log-'+timestr1+'.csv') as log:
    with open('LOG PRINTER '+timestr0+'.csv', 'w') as newLog:
        next(log)
        for line in log:
            line = line.replace('"2', '2')
            line = line.replace(',"', ',')
            line = line.replace('",', ',')
            newLog.write(line)

# Abre o novo arquivo e formata tudo direitinho.
a = open('LOG PRINTER '+timestr0+'.csv', encoding='utf-8')
csvArq = csv.DictReader(a, delimiter=',', quotechar='"')

# Inicia contadores.
usr = Counter()
prt = Counter()
gsc = Counter()
ngsc = Counter()

# Conta as variáveis e obtém os dados necessários.
for row in csvArq:
    pg = int(row['Pages'])
    cp = int(row['Copies'])
    gs = row['Grayscale']
    usr[row['User']] += pg * cp
    prt[row['Printer']] += pg * cp
    if gs == 'GRAYSCALE':
        gsc[row['User']] += pg * cp
    if gs == 'NOT GRAYSCALE':
        ngsc[row['User']] += pg * cp

# Organiza os dados em variáveis.
pages_max = sum(usr.values())
users_by_pages = usr.most_common()
users_top5 = usr.most_common(5)
most_used_printers = prt.most_common()
users_if_grayscale = gsc.most_common()
users_if_color = ngsc.most_common()
maxx = str(pages_max)

# Cria a pasta /img/
os.mkdir('img')

# Gera os gráficos e tabelas HTML através do Pandas e do MatPlotLib.
plt.style.use('ggplot')
users_pages = pd.DataFrame(users_by_pages, columns = ['Usuário', 'N. de Páginas'])
users_pages.plot(kind='barh',x='Usuário',y='N. de Páginas',figsize=(10,5))
plt.savefig('img/users_pages.png', bbox_inches='tight')
users_pages.to_html('users_pages.html', index=False, border=0, justify="left")

plt.style.use('ggplot')
topPrinters = pd.DataFrame(most_used_printers, columns = ['Impressora', 'N. de Páginas'])
topPrinters.plot(kind='barh',x='Impressora',y='N. de Páginas',figsize=(10,5))
plt.savefig('img/printers.png', bbox_inches='tight')
topPrinters.to_html('printers.html', index=False, border=0, justify="left")

grey = pd.DataFrame(users_if_grayscale, columns = ['Usuário', 'N. de Páginas P&B'])
notGrey = pd.DataFrame(users_if_color, columns = ['Usuário', 'N. de Páginas Coloridas'])

GreyOrNot = pd.merge(notGrey, grey, on='Usuário', how='outer')

plt.style.use('ggplot')
pd.concat(
    [GreyOrNot['Usuário'], GreyOrNot['N. de Páginas Coloridas'], 
     GreyOrNot['N. de Páginas P&B']],
    axis=1).plot(kind='barh', x='Usuário',figsize=(10,6))
plt.savefig('img/greynot.png', bbox_inches='tight')

top5 = pd.DataFrame(users_top5, columns = ['Usuário', 'N. de Páginas'])
top5.to_html('top5.html', index=False, border=0, justify="left")

# Configura as respectivas variáveis pra serem chamadas no arquivo HTML que será usado como corpo do E-Mail.
usersHTML = open('users_pages.html','r')
usersHTML = usersHTML.read()

printsHTML = open('printers.html', 'r')
printsHTML = printsHTML.read()

top5HTML = open('top5.html', 'r')
top5HTML = top5HTML.read()

today = date.today()
todayHTML = today.strftime("%d/%m/%Y")

css = ('.footer,.header{background-color:#e24a33;color:#fff;border-radius:4px}.header{text-align:center;padding:40px 0 30px 0;font-size:2vw}.footer{padding:15px 10px 15px 10px}.content{padding:15px 10px 15px 10px}.dataframe{width:95%}.dataframe td{border-bottom:1px solid #ddd;text-transform:uppercase;padding:0 10px 0 10px}.dataframe th{border-bottom:1px solid #ddd;background-color:#e24a33;border-radius:3px;color:#fff;text-transform:uppercase;padding:3px 3px}body{background-color:white;font-family:verdana;margin:0;padding:0}.stat-head{padding:0 10px 0 10px}td img{display:block;margin-left:auto;margin-right:auto}.main{background-color:#fff;}')

# Fecha o arquivo CSV.
a.close()


# Configura o envio do E-Mail.
# -------------------------------------------

# Informação básica.
TO = ' ??? '
FROM =' ??? '
SUBJECT = 'Relatório do uso das impressoras -- {time}'.format(time=todayHTML)

# Senha para a autenticação de E-Mail.
password = ' ??? '

# Corpo HTML do E-Mail.
email_content = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"> <html xmlns="http://www.w3.org/1999/xhtml"> <head> <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /> <title>Relatório do Uso das Impressoras - VECTRA</title> <meta name="viewport" content="width=device-width, initial-scale=1.0" /> <style> {cs} </style> </head> <body> <table border="0" cellpadding="0" cellspacing="0" width="100%"> <!-- Main table --> <tr> <td> <table class="main" align="center" border="" cellpadding="0" cellspacing="0" width="650" style="border-collapse: collapse;max-width:650px;"> <!-- Three main rows --> <tr> <td class="header"> <!-- Row 1 --> RELATÓRIO DO USO DAS IMPRESSORAS - VECTRA </td> </tr> <tr> <td class="content"> <!-- Row 2 --> <table border="0" cellpadding="0" cellspacing="0" width="100%"> <!-- Rwo 2 CONTENT --> <tr> <td class="stat-head"> <h3>USUÁRIOS / Nº DE IMPRESSÕES</h3><hr> </td> </tr> <tr> <td> <img src="cid:image1" width="98%" /> </td> </tr> <tr> <td align="center"> {us} </td> </tr> <tr> <td class="stat-head"> <br /><h3>IMPRESSORAS / Nº DE IMPRESSÕES</h3><hr> </td> </tr> <tr> <td> <img src="cid:image2" width="98%"/> </td> </tr> <tr> <td align="center"> {im} </td> </tr><br /> <tr> <td> <table border="0" cellpadding="0" cellspacing="0" width="100%"> <tr> <td class="stat-head" width="260" valign="top"> <h4> NÚMERO TOTAL DE IMPRESSÕES </h4><hr></td> <td class="stat-head" width="260" valign="top"> <h4> TOP 5 USUÁRIOS QUE MAIS IMPRIMIRAM </h4><hr></td> </tr> <tr> <td width="260" valign="top"> <h1 align="center" style="font-size:6vw;"> {ma} </h1> </td> <td align="center" width="260" valign="top"> {tt} </td> </tr> </table> </td> </tr> <tr> <td class="stat-head"> <br /> <h3>USUÁRIOS / Nº DE IMPRESSÕES COLORIDAS OU P&B </h3><hr> </td> </tr> <tr> <td> <img src="cid:image3" width="98%" /> </td> </tr> </table> </td> </tr> <tr> <td class="footer"> <!-- Row 3 --> Este relatório é gerado automaticamente. </br> {td}. </td> </tr> </table> </td> </tr> </table> </body> </html>'.format(cs=css, us=usersHTML, im=printsHTML, ma=maxx, tt=top5HTML, td=todayHTML)

# Configura as informações básicas ao E-Mail. 
MESSAGE = MIMEMultipart('related')
MESSAGE['subject'] = SUBJECT
MESSAGE['To'] = TO
MESSAGE['From'] = FROM
MESSAGE.preamble = 'Your mail reader does not support this format.'

# Configura o corpo HTML como parte do E-Mail.
HTML_BODY = MIMEText(email_content, 'html')
MESSAGE.attach(HTML_BODY)

# Atrela os gráficos em .PNG já previamente gerados ao E-Mail como anexos.
im1 = open('img/users_pages.png', 'rb')
msgImage1 = MIMEImage(im1.read())
im1.close()
im2 = open('img/printers.png', 'rb')
msgImage2 = MIMEImage(im2.read())
im2.close()
im3 = open('img/greynot.png', 'rb')
msgImage3 = MIMEImage(im3.read())
im3.close()

# Atrela cada imagem anexada um CID (Contend ID).
msgImage1.add_header('Content-ID', '<image1>')
MESSAGE.attach(msgImage1)
msgImage2.add_header('Content-ID', '<image2>')
MESSAGE.attach(msgImage2)
msgImage3.add_header('Content-ID', '<image3>')
MESSAGE.attach(msgImage3)

# Configurando servidor e enviando E-Mail.
server = smtplib.SMTP('smtp.gmail.com:587')
server.starttls()
server.login(FROM,password)
server.sendmail(FROM, [TO], MESSAGE.as_string())
server.quit()

# -------------------------------------------

# Deletando todos os arquivos previamente gerados e limpando a pasta.
shutil.rmtree('img')
os.remove('printers.html')
os.remove('top5.html')
os.remove('users_pages.html')
os.remove('LOG PRINTER '+timestr0+'.csv')
os.remove('papercut-print-log-'+timestr1+'.csv')