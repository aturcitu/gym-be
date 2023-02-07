import imaplib
import email
import yaml
import gspread
from gspread.cell import Cell
import json
from datetime import datetime

#path for external files
PATH_CREDENTIALS_MAIL= "credentials/credentials_logmail.yml"
PATH_CREDENTIALS_GSPREAD= "credentials/credentials_gsuit.json"
LOG_PATH = 'log.txt'
#sheet references
GSDOC_NAME = "Gimnasio - Schedule"
WKS_NAME = "Log"
USER_ID_WKS = "Usuarios"
EXE_INFO_WKS = "Ejercicios"

def log_info(data, tabs=0, log_path = LOG_PATH):
    with open(log_path, "a") as log_file:
        log_file.write("\t" * tabs + data + "\n")
        print("\t" * tabs + data)

def login_mail(path_credentials = PATH_CREDENTIALS_MAIL):
    with open(path_credentials) as f:
        content = f.read()
    #load credentials
    my_credentials = yaml.load(content, Loader=yaml.FullLoader)
    user, password = my_credentials["user"], my_credentials["password"]
    try:
        # Connection with GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL('imap.gmail.com')
        # Log in using your credentials
        my_mail.login(user, password)
        log_info("Login to mail OK", tabs=1)
    except:
        log_info("Loging to mail NOK", tabs=1)
        return False, None

    return my_mail, user

def load_worksheet(path_credentials=PATH_CREDENTIALS_GSPREAD, doc_name=GSDOC_NAME, wks_name=WKS_NAME):
    try:
        sa = gspread.service_account(path_credentials)
        sh = sa.open(doc_name)
        wks = sh.worksheet(wks_name)
        log_info("Worksheet loaded OK", tabs=1)
        return wks, sh 
    except:
        log_info("Worksheet loaded NOK", tabs=1)
        return None, None 

def load_user_id(sh):
    worksheet = sh.worksheet(USER_ID_WKS)
    user_ids = {}
    for index, user_id in enumerate(worksheet.col_values(1)[1:]):
        user_ids.update({int(user_id): worksheet.col_values(2)[1+index]})
    return user_ids

def load_exercises(sh):
    worksheet = sh.worksheet(EXE_INFO_WKS)
    exercies_info = {}
    ejercicio_data = worksheet.col_values(worksheet.find("ejercicio").col)[1:]
    tipo_data = worksheet.col_values(worksheet.find("tipo").col)[1:]
    for index, exer_id in enumerate(worksheet.col_values(worksheet.find("id").col)[1:]):
        exercies_info.update({
                int(exer_id): {
                    'ejercicio': ejercicio_data[index],
                    'tipoDato':  tipo_data[index] }
            })
    return exercies_info

def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return int(len(str_list)+1)

def clean_json(input_string):
    input_string = input_string.replace('=\r\n', '').replace('\r\n', '')
    start_index = input_string.find("{")
    end_index = input_string.rfind("}")
    if start_index != -1 and end_index != -1:
        return input_string[start_index:end_index + 1]
    return input_string

def trainning_form_mail(messages):
    training_msgs = {}
    for msg in messages[::-1]:
        for response_part in msg:
            if type(response_part) is tuple:
                my_msg=email.message_from_bytes((response_part[1]))
                for word_subject in my_msg['subject'].split():
                    #only keep messages with a "%d/%m/%Y" date on the subject
                    try:  
                        datetime.strptime(word_subject, "%d/%m/%Y")
                        for part in my_msg.walk():
                            #extract text body fro mmail 
                            if part.get_content_type() == 'text/plain':
                                training_msgs.update({word_subject: json.loads(clean_json(part.get_payload()))})
                                break
                    except:
                        next    
    return training_msgs

def json_to_cells(json_dic, COL_NAMES, free_row):
    rows = [] 
    for key in json_dic: 
        new_row = []
        for col_name in COL_NAMES:  
            try:
                new_row.append(json_dic[key][col_name])
            except:
                new_row.append("")
                #print("Col name not found:",col_name)  
        rows.append(new_row)
    cells = []
    for i, exercise in enumerate(rows):
        for j, cell_data in enumerate(exercise):
            cells.append(Cell(row=i+free_row, col=j+1, value = cell_data)) 
    return cells, rows

def upload_relative_load(wks):
    wks_dic = wks.get_all_records()
    # Get max load for each Exercise ID
    exer_records = {}
    for wks_row in wks_dic:
        row_id = wks_row['id']
        row_load = wks_row['carga']
        try:
            if row_load > exer_records[row_id]:
                exer_records[row_id] = row_load
        except KeyError:
            exer_records[row_id] = row_load
    # Compute relative load (normlizing to max -> 100)
    exer_rel_load = []
    for wks_row in wks_dic:
        row_id = wks_row['id']  
        row_load = wks_row['carga']
        try:
            rel_load = int((row_load/exer_records[row_id])*100)
        except:
            rel_load = 0
        exer_rel_load.append(rel_load)
    # Update values
    log_info("All relative load have been re-computed", tabs=1)
    cells = []
    wks_col = [item for item in wks.row_values(1) if item].index('cargaRelativa')+1
    for wks_row, cell_data in enumerate(exer_rel_load):
        cells.append(Cell(row=wks_row + 2, col=wks_col, value = cell_data)) 
    try:
        wks.update_cells(cells, value_input_option='USER_ENTERED')
        log_info("Spreadsheet updated OK", tabs=1)
    except:
        log_info("Spreadsheet updated NOK", tabs=1)
