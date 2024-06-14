# -*- coding: utf-8 -*- 

import os
import gspread
from gspread.worksheet import Worksheet
from gspread.spreadsheet import Spreadsheet
from google.oauth2.service_account import Credentials
import pandas as pd
import json
from time import sleep
SCOPES = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

def __get_id_by_name(gc, name:str):
    files = gc.list_spreadsheet_files()
    sheet_file_id = None
    for file in files:
        if file['name'] == name:
            sheet_file_id = file['id']
            break
    if sheet_file_id:
        return sheet_file_id
    else:
        return None

def send_sheets():
    def login():
        credentials = Credentials.from_service_account_file('filename.json', scopes=SCOPES)
        gc = gspread.auth.authorize(credentials)
        return gc

    def get_sheet_file(gc, file_path):
        file = open('settings.json', 'r+', encoding='utf-8')
        settings = json.load(file)
        filename = f'{os.path.basename(file_path).replace(".xlsx","")}'
        try:
            # id_sheet = __get_id_by_name(gc, settings[filename])
            # spreadsheet = gc.open_by_key(id_sheet)
            spreadsheet = gc.open_by_key(settings[filename])
        except KeyError:
            spreadsheet = gc.create(filename)
            settings[filename] = spreadsheet.id
            file.truncate(0)
            file.seek(0)
            json.dump(settings, file, indent=4, ensure_ascii=False)
        return spreadsheet
    
    def convert_excel_to_csv(file_path):
        def create_folder_csv():
            folder_path = f'planilhas/{os.path.basename(file_path).replace(".xlsx", "")}/'
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            return folder_path

        data_excel = pd.ExcelFile(file_path)

        for sheet_name in data_excel.sheet_names:
            data_sheet = pd.read_excel(data_excel, sheet_name=sheet_name)
            folder_path = create_folder_csv()
            csv_file_path = os.path.join(folder_path, f'{sheet_name}.csv')
            data_sheet.to_csv(csv_file_path, index=False)
        return f'planilhas/{os.path.basename(file_path).replace(".xlsx", "")}/'
    
    def push_worksheets(spreadsheet:Spreadsheet, folder_path):
        print('atualizando planilhas')
        def get_or_create_worksheet(spreadsheet, sheet_title, num_rows, num_cols):
            try:
                worksheet = spreadsheet.worksheet(sheet_title)
            except gspread.exceptions.WorksheetNotFound:
                print(sheet_title)
                worksheet = spreadsheet.add_worksheet(title=sheet_title[:100], rows=num_rows, cols=num_cols)
            return worksheet
        
        def update_worksheets():
            def check_duplicates(ws:Worksheet):
                novos_dados = ws.get_all_values()
                df_novos_dados = pd.DataFrame(novos_dados[1:], columns=novos_dados[0])
                duplicatas = df_novos_dados[df_novos_dados.duplicated(keep=False)]
                if not duplicatas.empty:
                    print("Duplicatas encontradas na planilha:")
                    print(duplicatas)
                else:
                    print("Não há duplicatas na planilha.")

            def remove_duplicates(ws:Worksheet):
                data = ws.get_all_values()
                unique_lines = []
                lines_already_seen = set()
                for linha in data:
                    chave = tuple(linha)
                    if chave not in lines_already_seen:
                        unique_lines.append(linha)
                        lines_already_seen.add(chave)
                ws.clear()
                ws.update(unique_lines)

            files = os.listdir(folder_path)
            for file in files:
                with open(f'{folder_path}/{file}', 'r') as file:
                    csv_contents = file.read().splitlines()
                    have_text = False
                    for line in csv_contents:
                        if not line.isspace() and have_text is False:
                            have_text = True
                            continue
                        else:
                            break
                    if have_text is False:
                        continue

                    with pd.ExcelFile(folder_path[:-1]+'.xlsx') as xls:
                        data_sheet = pd.read_excel(xls, os.path.basename(file.name).replace('.csv',''))
                        num_rows, num_cols = data_sheet.shape
                ws = get_or_create_worksheet(spreadsheet, os.path.basename(file.name).replace('.csv',''), num_rows, num_cols)
                ws:Worksheet
                # update de valores
                try:
                    ws.update([line.split(',') for line in csv_contents])
                except Exception as e:
                    print(f'{e}: \n\n{csv_contents}\n\n')
                    continue
                # removendo qualquer row duplicata
                remove_duplicates(ws)
                # final check
                # check_duplicates(ws)
            try:
                sheet_origin = spreadsheet.get_worksheet(0)
                if sheet_origin.title == 'Sheet1' or sheet_origin.title == 'Página1':
                    spreadsheet.del_worksheet(sheet_origin)
            except Exception as e:
                print(f'remove sheet problem: {e}')

        def share_with_users():
            emails_to_share = ['bla bla bla']
            users_with_access = spreadsheet.list_permissions()
            users_to_share = []
            for email in emails_to_share:
                has_access = False
                for user in users_with_access:
                    if email == user['emailAddress']:
                        has_access = True
                        break
                if not has_access:
                    users_to_share.append(email)

            for user in users_to_share:
                spreadsheet.share(user, perm_type='user', role='writer')
        update_worksheets()
        share_with_users()
        print(f"Arquivo {spreadsheet.title} carregado com sucesso para o Google Planilhas!")

    gc = login()
    sheet_files = os.listdir('planilhas')
    sheet_files = [file for file in sheet_files if file.endswith('.xlsx')]
    for sheet_file in sheet_files:
        file_path = f'planilhas/{sheet_file}'
        spreadsheet = get_sheet_file(gc, file_path)
        folder_path = convert_excel_to_csv(file_path)
        push_worksheets(spreadsheet, folder_path)

def change_values():
    def login():
        credentials = Credentials.from_service_account_file('salao-jovem-avec.json', scopes=SCOPES)
        gc = gspread.auth.authorize(credentials)
        return gc

    def get_sheet_files():
        file = open('settings.json', 'r+', encoding='utf-8')
        settings = json.load(file)
        sheet_files = []

        no_priotity_list = ['last_date_updated','Campanhas', 'Profissionais', 'Agenda', 'Auditoria', 'Clientes']
        # fazendo os prioritarios antes
        for key, value in settings.items():
            if key in no_priotity_list:
                continue
            print(f'\n\n{key}\n\n')
            sheet_files.append(value)
        # fazendo os demais
        for key, value in settings.items():
            if key in no_priotity_list and key != 'last_date_updated':
                print(f'\n\n{key}\n\n')
                sheet_files.append(value)
        return sheet_files

    def alter_values(sheet_files):

        def col_idx_to_str(col_idx):
            """Converts a column index to a column letter (e.g., 1 -> 'A', 27 -> 'AA')."""
            col_str = ''
            while col_idx > 0:
                col_idx, remainder = divmod(col_idx - 1, 26)
                col_str = chr(65 + remainder) + col_str
            return col_str

        for sheet_file in sheet_files:
            gc = login()
            spreadsheet = gc.open_by_key(sheet_file)
            worksheets = spreadsheet.worksheets()
            
            for worksheet in worksheets:
                if worksheet.title == 'Sheet1' or worksheet.title == 'Página1':
                    continue
                
                all_cells = worksheet.get_all_values()
                modified_rows = []

                for row in all_cells:
                    new_row = []
                    for cell in row:
                        if cell.count('"') == 1:
                            new_value = cell.replace('"', '')
                        elif cell.endswith('.0') and cell.count('.0') == 1:
                            new_value = int(cell.replace('.0', ''))
                        elif cell.isdigit():
                            new_value = int(cell)
                        else:
                            new_value = cell
                        new_row.append(new_value)
                    modified_rows.append(new_row)

                # Determinar a faixa completa para a atualização
                num_rows = len(modified_rows)
                num_cols = len(modified_rows[0]) if num_rows > 0 else 0
                last_col_letter = col_idx_to_str(num_cols)
                range_name = f'A1:{last_col_letter}{num_rows}'

                try:
                    worksheet.update(values=modified_rows, range_name=range_name)
                    print(f"Worksheet {worksheet.title} updated successfully.")
                    sleep(2)
                except Exception as e:
                    print(f"Failed to update worksheet {worksheet.title}: {e}")
                    sleep(10)
                    try:
                        worksheet.update(values=modified_rows, range_name=range_name)
                        print(f"Retrying: Worksheet {worksheet.title} updated successfully.")
                        sleep(2)
                    except Exception as e:
                        print(f"Failed to update worksheet {worksheet.title} on retry: {e}")
                        sleep(10)

                # try:
                #     worksheet.update(values=modified_rows, range_name=range_name)
                #     print(f"Worksheet {worksheet.title} updated successfully.")
                #     sleep(2)
                # except Exception as e:
                #     print(f"Failed to update worksheet {worksheet.title}: {e}")
                #     sleep(10)
                #     try:
                #         worksheet.update(values=modified_rows, range_name=range_name)
                #         print(f"Retrying: Worksheet {worksheet.title} updated successfully.")
                #         sleep(2)
                #     except Exception as e:
                #         print(f"Failed to update worksheet {worksheet.title}: {e}")
                #         sleep(10)
                #         try:
                #             worksheet.update(values=modified_rows, range_name=range_name)
                #             print(f"Retrying: Worksheet {worksheet.title} updated successfully.")
                #             sleep(2)
                #         except:
                #             print(f"Failed to update worksheet {worksheet.title} on retry: {e}")
                #             sleep(10)
                #             continue

    sheet_files = get_sheet_files()
    alter_values(sheet_files)

if __name__ == '__main__':
    with open('settings.json', 'r') as file:
        settings = json.load(file)
    change_values()
