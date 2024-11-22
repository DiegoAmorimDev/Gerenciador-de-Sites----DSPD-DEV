import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Se deixar .readonly no final, o escopo fica limitado a somente ler os dados
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

#O id da planilha e o intervalo de celulas da pagina.
SAMPLE_SPREADSHEET_ID = "1pvPLtqu06VXF2PwSFE-jqqWRq586vnOLSIin95QaL7U"
SAMPLE_RANGE_NAME = "dados_gerais!A1:F1015"


def main():
  """Shows basic usage of the Sheets API.
  Prints values from a sample spreadsheet.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  #verifica se ja existe o token
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

#usar a planilha
  try:
    service = build("sheets", "v4", credentials=creds)

    # (ler informações do google sheets)
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
        .execute()
    )

    # adicionar informações:
    # valores_adicionar = [
    #   ["Testestes", 'Testando'],
    #   ["Testandos", 'Teste20'],
    # ]
    # result = (
    #     sheet.values()
    #     .update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="A189", valueInputOption="RAW", body={'values': valores_adicionar})
    #     .execute()
    # )


    valores = result['values']

    valores_adicionar = [
      ["imposto"],
    ]

    for i, linha in enumerate(valores):
      if i > 0:
        print(linha, i)
        
      

    result = (
        sheet.values()
        .clear(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="A189")
        .execute()
    )

#listas, manipulando como listas
    # valores = result['values']
    # print(valores)
    
   
  except HttpError as err:
    print(err)


if __name__ == "__main__":
  main()