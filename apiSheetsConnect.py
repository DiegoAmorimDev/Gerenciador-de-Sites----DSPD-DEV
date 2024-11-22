import os.path
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Se deixar .readonly no final, o escopo fica limitado a somente ler os dados
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# O ID da planilha e o intervalo de células da página.
SAMPLE_SPREADSHEET_ID = "1pvPLtqu06VXF2PwSFE-jqqWRq586vnOLSIin95QaL7U"
SAMPLE_RANGE_NAME = "dados_gerais!A1:F1015"

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # Verifica se já existe o token
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Gera um token
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        # Constrói o serviço para a API do Google Sheets
        service = build("sheets", "v4", credentials=creds)

        # Lê os valores da planilha
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=SAMPLE_RANGE_NAME
        ).execute()
        values = result.get('values', [])

        if not values:
            print("Nenhum dado encontrado.")
            return

        def mostrarPlanilha(result):
            df = pd.DataFrame(values[1:], columns=values[0])  # Usa a primeira linha como cabeçalho
            print(df)

        #leitura
        mostrarPlanilha(result)

        #criar
        # def adicionarDado(result):
        #     result.update_cell(row=191, col=1, value="testeconecct")
        #     mostrarPlanilha(result)

        # #adicionar
        # adicionarDado(result)

    except HttpError as err:
        print(f"Ocorreu um erro: {err}")

if __name__ == "__main__":
    main()
