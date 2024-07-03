from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

class Spreadsheet:
    def __init__(self, sheet_url, sheet_name, credentials_path):
        spreadsheet_id = sheet_url.split('/')[-2]
        self.creds = self.authenticate(credentials_path)
        self.SPREADSHEET_ID = spreadsheet_id
        service = build('sheets', 'v4', credentials=self.creds, cache_discovery=False)
        self.sheet = service.spreadsheets()
        self.sheet_name = sheet_name
        self.values = []
        self.__reload_values()

    def authenticate(self, credentials_path):
        flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
        creds = flow.run_local_server(port=0)
        return creds

    def __reload_values(self):
        try:
            self.values = self.sheet.values().get(
                spreadsheetId=self.SPREADSHEET_ID,
                range=f'{self.sheet_name}!A:AZ').execute().get('values', [])
        except KeyError:
            pass

    def write_values(self, first_cell: str, values: list):
        request = self.sheet.values().update(
            spreadsheetId=self.SPREADSHEET_ID,
            range=f"{self.sheet_name}!{first_cell}",
            valueInputOption="RAW",
            body={"values": values})
        request.execute()

    def get_number_of_lines_filled(self):
        self.__reload_values()
        return len(self.values)

def read_txt_file(file_path):
    with open(file_path, 'r') as file:
        content = file.readlines()
    return [line.strip().split(',') for line in content]

def send_email(content, subject, toaddr, fromaddr, pwdEmail):
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = ', '.join(toaddr)
    msg['Subject'] = subject

    html = f"""\
    <html>
    <head></head>
    <body>
        <p>{content}</p>
    </body>
    </html>
    """
    msg.attach(MIMEText(html, 'html'))

    send = smtplib.SMTP('smtp-mail.outlook.com:587')
    send.starttls()
    send.login(fromaddr, pwdEmail)
    text = msg.as_string().encode('utf-8')
    send.sendmail(fromaddr, toaddr, text)
    send.quit()
    print('[E-mail] Report enviado!')

def main():
    sheet_url = 'link da planilha'
    sheet_name = 'teste de integração'
    txt_file_path = 'path\file.txt'
    credentials_path = 'path\credentials.json'
    fromaddr = 'remetente@hotmail.com'
    pwdEmail = 'senha'
    toaddr = ['destinatario@gmail.com']

    # Inicializa a classe Spreadsheet
    spreadsheet_class = Spreadsheet(sheet_url, sheet_name, credentials_path)
    # Lê os dados do arquivo .txt
    data = read_txt_file(txt_file_path)

    # Escreve os dados na planilha
    new_line = spreadsheet_class.get_number_of_lines_filled() + 1
    spreadsheet_class.write_values(f'A{new_line}', data)

    # Envia e-mail com o conteúdo escrito na planilha
    content = pd.DataFrame(data).to_html(index=False)
    send_email(content, "Atualização da Planilha", toaddr, fromaddr, pwdEmail)

if __name__ == "__main__":
    main()