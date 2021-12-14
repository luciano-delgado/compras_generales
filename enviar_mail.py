from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os, smtplib



def enviarmail(clave,cuenta,destino,cc,asunto,texto=''):



    # create message object instance
    msg = MIMEMultipart()
    password = clave
    msg['From'] = cuenta
    msg['To'] = destino

    #msg['To'] = ','.join(destino) ### Con varios destinos
    msg['Subject'] = asunto
    msg['Cc'] = ','.join(cc) ### Con varios CC
    #msg['Cc'] = cc2
    body = MIMEText(texto) # convert the body to a MIME compatible string
    msg.attach(body) # attach it to your main message

    # attach image to message body
    #fp=open(r"C:\Users\facucores\Desktop\descarga.jpg",'rb')
    #msgImage = MIMEImage(fp.read())
    #fp.close()
    #msg.attach(msgImage) 
    #Enviar mail con archivo no imagen por python
    #part = MIMEBase('application', "octet-stream")
    #part.set_payload(open(archivo, "rb").read())
    #encoders.encode_base64(part)
    #part.add_header('Content-Disposition', 'attachment; filename="%s"' % nombredelarchivo)
    #msg.attach(part)
    
    # create server
    server = smtplib.SMTP('smtp.gmail.com: 587')
    server.starttls()
    
    # Login Credentials for sending the mail
    try:
        server.login(msg['From'], password)
    except:
        print(f'=============================================')
        print(f'~~~ATENCION: Usuario o contraseña incorrectos')
        print(f'============================================')
    
    # send the message via the server.
    #server.sendmail(msg['From'], destino, msg.as_string()) ### Varios destinos
    server.sendmail(msg['From'], destino.split(","), msg.as_string())
    server.sendmail(msg['From'], cc, msg.as_string())
    
    server.quit()
#----------------------------------------------------------------------------
# contra='Chewie2019!'
# emisor='ldelgado@scienza.com.ar'

# cc = ['lucianodelgado92@gmail.com','oyp@scienza.com.ar']
# destino ='lucianodelgado92@hotmail.com'

# asunto = 'PRUEBA'
# cuerpo = 'hola hola hola'

# enviarmail(contra,emisor,destino,cc,asunto,cuerpo)


#-----
# Destino=['correo1@gmail.com','correo2@gmail.com','correo3a@gmail.com','correo4@gmail.com' ]
# msg['To'] = ','.join(Destino)

# server.sendmail(msg['From'], Destino, msg.as_string())
    