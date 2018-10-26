#============================PROCESO==================================



#=======================Mandar Mail=======================
library(mailR)
#Para enviar un correo electrónico con uno o más archivos adjuntos y 
#configurar el parámetro de depuración para ver un mensaje de registro 
#detallado:
#Carlos <betinpaovitto@gmail.com> 1712331212

send.mail(from = "sender@gmail.com",
          to = c("recipient1@gmail.com", "recipient2@gmail.com"),
          subject = "Subject of the email",
          body = "Body of the email",
          smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "gmail_username", passwd = "password", ssl = TRUE),
          authenticate = TRUE,
          send = TRUE,
          attach.files = c("./download.log", "upload.log", "https://dl.dropboxusercontent.com/u/5031586/How%20to%20use%20the%20Public%20folder.rtf"),
          file.names = c("Download log.log", "Upload log.log", "DropBox File.rtf"), # optional parameter
          file.descriptions = c("Description for download log", "Description for upload log", "DropBox File"), # optional parameter
          debug = TRUE)

#Para enviar un correo electrónico con formato HTML:
send.mail(from = "sender@gmail.com",
          to = c("recipient1@gmail.com", "recipient2@gmail.com"),
          subject = "Subject of the email",
          body = "<html>The apache logo - <img src=\"http://www.apache.org/images/asf_logo_wide.gif\"></html>", # can also point to local file (see next example)
          html = TRUE,
          smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "gmail_username", passwd = "password", ssl = TRUE),
          authenticate = TRUE,
          send = TRUE)



#To send a HTML formatted email with embedded inline images:

send.mail(from = "sender@gmail.com",
          to = c("recipient1@gmail.com", "recipient2@gmail.com"),
          subject = "Subject of the email",
          body = "path.to.local.html.file",
          html = TRUE,
          inline = TRUE,
          smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "gmail_username", passwd = "password", ssl = TRUE),
          authenticate = TRUE,
          send = TRUE)

#To send an email with utf-8 or other encoding:

email <- send.mail(from = "betinpaovitto@gmail.com",
                   to = "betinpaovitto@hotmail.com",
                   subject = "A quote from Gandhi",
                   body = "In Hindi :  थोडा सा अभ्यास बहुत सारे उपदेशों से बेहतर है।
                   English translation: An ounce of practice is worth more than tons of preaching.",
                   encoding = "utf-8",
                   smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "betinpaovitto", passwd = "1712331212", ssl = T),
                   authenticate = TRUE,
                   send = TRUE)
#email$send()


