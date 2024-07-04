import mysql.connector
import pandas as pd
from typing import List
from tqdm import tqdm

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path

class Connection:
    '''
    Manipulação de API do Google Sheets.
    '''
    def __init__(self, spreadsheet_id: str = None, worksheet: str = None, creds_path: str = 'credentials.json') -> None:
        self.spreadsheet_id = spreadsheet_id
        self.worksheet = worksheet

        # Initialize the Google Sheets API
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

        creds = None
        token_pickle = 'token.pickle'

        # Load token if it exists
        if os.path.exists(token_pickle):
            with open(token_pickle, 'rb') as token:
                creds = pickle.load(token)

        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
                creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open(token_pickle, 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

        # Call the Sheets API
        self.sheet = service.spreadsheets()

    def create_worksheet(self, title):
        body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': title,
                    }
                }
            }]
        }

        result = self.sheet.batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body=body
        ).execute()

        return result

    def insert_data(self, worksheet_title: str, values: List[List[str]]):
        values_body = {
            'valueInputOption': "USER_ENTERED",
            'data': [
                {
                    "range": f"{worksheet_title}!A1",
                    "values": values
                }
            ]
        }
        result = self.sheet.values().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body=values_body
        ).execute()
        return result

    def delete_worksheet(self, title: str):
        sheets = self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()
        sheets = sheets['sheets']
        sheetId = False
        for sheet in sheets:
            if sheet['properties']['title'] == title:
                sheetId = sheet['properties']['sheetId']
        
        if not sheetId:
            return 'Planilha não existe'

        body = {
            "requests": [
                {
                "deleteSheet": {
                    "sheetId": sheetId
                }
                }
            ]
        }
        result = self.sheet.batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body=body
        ).execute()
        return result

    def verify_worksheet(self, worksheet, results_list):
        print(f'Operando em {worksheet}')
        planilha = Connection(
            spreadsheet_id='1om2CebC9eH3brMJZE9A7RaUMXlvOc6f2jyBaj-a9Pfk',
            worksheet=worksheet
        )
        try:
            print('Deletando aba para recriar')
            planilha.delete_worksheet(worksheet)
        except Exception:
            print('Aba deletada, recriando')
        try:
            planilha.create_worksheet(worksheet)
        except Exception:
            print('Aba já criada, pulando.')

        self.all_values = [list(results_list[0].keys())]  # Cria a lista de cabeçalho
        for occurrence in tqdm(results_list, desc=f'Saving values from {worksheet}'):
            self.all_values.append([occurrence[key] for key in occurrence.keys()])

        planilha.insert_data(
            worksheet_title=worksheet,
            values=self.all_values
        )

    def connect_db(self, user, password, db, host='localhost', port='sua porta'):
        connection = False
        try:
            connection = mysql.connector.connect(
                host=host,
                port=port,
                database=db,
                user=user,
                passwd=password
            )
        except Exception as e:
            print(f"Erro ao se conectar com Banco de Dados: {str(e)}")
        return connection

    def execute_query(self, connection, query, params=None):
        cursor = connection.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            result = cursor.fetchall()
            return result
        except Exception as e:
            print(f"Erro executando a SQL query: {str(e)}")
            return None
        finally:
            cursor.close()

    def main(self):
        connection = Connection.connect_db(self, user='seu user', password='sua senha', db='db', port='sua porta')
        if not connection:
            print("Error connecting to the database.")
            exit(0)

        query = '''
        SELECT
            *
        FROM
            vagas
    '''

        result = script.execute_query(connection, query)

        results_list = []

        if result:
            for row in result:
                row_dict = {
                    'id': row[0],
                    'nome': row[1],
                    'idade': row[2],
                    'cidade': row[3],
                }
                results_list.append(row_dict)
        script.verify_worksheet("teste", results_list)

if __name__ == '__main__':
    script = Connection()
    script.main()