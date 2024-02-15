import gspread
from oauth2client.service_account import ServiceAccountCredentials
from math import ceil

# Após a 5~10 linha o programa se encerra devido ao número máximo de requisições por minuto por usuário;

class GoogleSheetProcessor:
    def __init__(self, json_credentials_path, spreadsheet_name, worksheet_name):
        self.scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(json_credentials_path, self.scope)
        self.client = gspread.authorize(self.creds)
        self.spreadsheet = self.client.open(spreadsheet_name)
        self.worksheet = self.spreadsheet.worksheet(worksheet_name)

    def get_column_values(self, column_name):
        column_index = self.get_column_index(column_name)
        return self.worksheet.col_values(column_index)

    def get_column_index(self, column_name):
        headers = self.worksheet.row_values(1)
        return headers.index(column_name) + 1

    def calculate_average(self, p1, p2, p3):
        return (p1 + p2 + p3) / 3

    def process_students(self):
        num_rows = len(self.worksheet.get_all_values())
        total_classes = 60

        for row in range(2, num_rows + 1):
            try:
                # Converção
                matricula = int(self.worksheet.cell(row, 1).value)
            except ValueError:
                continue

            aluno = self.worksheet.cell(row, 2).value
            faltas = int(self.worksheet.cell(row, 3).value)
            p1 = float(self.worksheet.cell(row, 4).value)
            p2 = float(self.worksheet.cell(row, 5).value)
            p3 = float(self.worksheet.cell(row, 6).value)

            # Média
            average = self.calculate_average(p1, p2, p3)
            # Reprovado ou não
            if faltas > total_classes * 0.25:
                situacao = "Reprovado por Falta"
                naf = 0
            else:    
                if average < 50:
                    situacao = "Reprovado por Nota"
                    naf = 0
                if average >= 50 and average < 70:
                        situacao = "Exame Final"
                        #Nota para Aprovação Final
                        naf = ceil(100 - average)
                if average >= 70:
                        situacao = "Aprovado"
                        naf = 0

            # Update 
            self.worksheet.update_cell(row, 7, situacao)
            self.worksheet.update_cell(row, 8, naf)

            
            print(f"Processed student {matricula} ({aluno}): Average={average}, Situacao={situacao}, NAF={naf}")


json_credentials_path = 'teste.json'
spreadsheet_name = 'Engenharia de Software - DesafioRafaelYudiShirakura'
worksheet_name = 'engenharia_de_software'

sheet_processor = GoogleSheetProcessor(json_credentials_path, spreadsheet_name, worksheet_name)
sheet_processor.process_students()
