# sheet_email_python

# Português

Projeto de Integração de Planilha e Envio de E-mails Automatizado
Este projeto visa automatizar a integração de dados de um arquivo de texto para uma planilha do Google Sheets, além de enviar um e-mail com o conteúdo atualizado da planilha. Utiliza Python para integração com a API do Google Sheets e envio de e-mails via SMTP.

Configuração das Credenciais do Google:
Crie um projeto no Google Cloud Console e habilite a API do Google Sheets.
Baixe as credenciais JSON e salve o arquivo no diretório do projeto com o nome credentials.json.

Configuração do Arquivo de Texto:
Prepare seu arquivo de texto com os dados a serem inseridos na planilha. Cada linha deve conter os dados separados por vírgula.

Configuração do E-mail:
Defina o endereço de e-mail remetente (fromaddr) e a senha correspondente (pwdEmail).
Especifique os endereços de e-mail dos destinatários (toaddr).

Execução do Script:
Execute o script Python main.py para integrar os dados na planilha e enviar o e-mail.

# English

Automated Spreadsheet Integration and Email Sending Project
This project aims to automate the integration of data from a text file into a Google Sheets spreadsheet, and subsequently send an email with the updated spreadsheet content. It utilizes Python for integration with the Google Sheets API and for sending emails via SMTP.

Configure Google Credentials:
Create a project in Google Cloud Console and enable the Google Sheets API.
Download the JSON credentials file and save it in the project directory as credentials.json.

Configure the Text File:
Prepare your text file with the data to be inserted into the spreadsheet. Each line should contain comma-separated values.

Configure Email Settings:
Set the sender email address (fromaddr) and corresponding password (pwdEmail).
Specify the recipient email address (toaddr).

Run the Script:
Execute the Python script main.py to integrate the data into the spreadsheet and send the email.