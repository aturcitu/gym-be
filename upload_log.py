import imaplib
import email
import yaml
import gspread
from gspread.cell import Cell
import json
from datetime import datetime

PATH_CREDENTIALS_MAIL= "credentials/credentials_logmail.yml"
PATH_CREDENTIALS_GSPREAD= "credentials/credentials_gsuit.json"
#put data into external file
GSDOC_NAME = "Gimnasio - Schedule"
WKS_NAME = "Log"
USER_ID_WKS = "Usuarios"
EXE_INFO_WKS = "Ejercicios"
LOG_PATH = 'log.txt'


def log_info(data, tabs=0, log_path = LOG_PATH):
    with open(log_path, "a") as log_file:
        log_file.write("\t" * tabs + data + "\n")
        print("\t" * tabs + data)

def login_mail(path_credentials):
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

def load_worksheet(path_credentials, doc_name, wks_name):
    try:
        sa = gspread.service_account(path_credentials)
        sh = sa.open(doc_name)
        wks = sh.worksheet(wks_name)
        log_info("Worksheet loaded OK", tabs=1)
        return wks, sh 
    except:
        log_info("Worksheet loaded NOK", tabs=1)
        return None, None 

def laod_user_id(sh):
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
                    'tipoDato':  tipo_data[index]
                }
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



#--- PREPARE MAIL ACCOUNT AND GET INFO ---#
log_info("Execution started at " + str(datetime.now()))
#log into mail account
my_mail, mail_user = login_mail(PATH_CREDENTIALS_MAIL)
if mail_user:
    #check inbox mails from same account
    my_mail.select('Inbox')
    _, data = my_mail.search(None, "FROM", mail_user) 
    mail_id_list = data[0].split()  #IDs of all emails that we want to fetch 
    #Iterate through messages and extract data into the msgs list
    log_info(str(len(mail_id_list))+" mails found",tabs=1)
    msgs = [] 
    for num in mail_id_list:
        typ, data = my_mail.fetch(num, '(RFC822)') #RFC822 returns whole message 
        msgs.append(data)
    #filter trainnings from mails
    training_json_dic = trainning_form_mail(msgs)
else:
    None
    #break when in main
#--- LOAD EXCEL INFO ---#
wks, sh = load_worksheet(PATH_CREDENTIALS_GSPREAD, GSDOC_NAME, WKS_NAME)
if wks:
    user_ids = laod_user_id(sh)
    exercises_info = load_exercises(sh)
    last_date = wks.col_values(wks.find("fecha").col)[-1]
else:
    None
    #break when in main

# Load the JSON input and format data 
new_rows_dic = {}
num_new_rows = 0
timestamp_last_date = datetime.strptime(last_date, "%d/%m/%Y")
#chech date every valid json, keep only new ones
for sorted_training in sorted(training_json_dic.items(), key = lambda x:datetime.strptime(x[0], "%d/%m/%Y")):
    json_date = sorted_training[0]
    json_input = sorted_training[1]
    if datetime.strptime(json_date, "%d/%m/%Y") > timestamp_last_date: 
        log_info("Parsing data from: " + json_date, tabs=1)
        # valid data
        timestamp_init = datetime.fromtimestamp(json_input['fechaInicio']/1000)
        timestamp_end = datetime.fromtimestamp(json_input['fechaFin']/1000)
        # common data for whole json
        common_data = {
            "fecha":  timestamp_init.strftime("%d/%m/%Y"),
            "fechaInicio":  int(datetime.timestamp(timestamp_init)),
            "fechaFin":  int(datetime.timestamp(timestamp_end)),
            "tiempo":  int(datetime.timestamp(timestamp_end) - datetime.timestamp(timestamp_init)),
            "entrenamiento": json_input['entrenamiento'],
            "mesociclo": json_input['mesociclo'],
            "tipoEntrenamiento": json_input['tipoEntrenamiento'],
        }
        # common data for each exercise
        for exercise_dic in json_input["ejercicios"]:
            # Direct data from json all users
            exer_data = {}
            exer_data_names = ['orden','id','series','tempo','objetivo']
            for data in exer_data_names:
                exer_data.update({data: exercise_dic[data]})
            # Load repeticiones
            for idx, reps in enumerate(exercise_dic['repeticiones']):
                exer_data.update({'repeticiones'+str(idx+1): reps})
            #load non-direct exercise data
            exer_data.update(exercises_info[exercise_dic['id']])
            # user depending data
            for user_dic in exercise_dic["usuarios"]:
                user_data = {
                    'usuario': user_ids[user_dic['id']],
                    'sensacion': user_dic['sensacion']
                    }
                # Load repeticiones efectivas y pesos
                for idx, reps in enumerate(user_dic['repeticionesRealizadas']):
                    exer_data.update({'realizadas'+str(idx+1): reps})                
                for idx, peso in enumerate(user_dic['pesos']):
                    exer_data.update({'peso'+str(idx+1): peso})
                # compute total load
                load = 0
                for idx, reps in enumerate(user_dic['repeticionesRealizadas']):
                    load = load + reps * user_dic['pesos'][idx]
                exer_data.update({'carga': load})

                new_rows_dic.update({num_new_rows: common_data | exer_data | user_data })      
                num_new_rows = num_new_rows +1

COL_NAMES = [item for item in wks.row_values(1) if item]
free_row = next_available_row(wks)
new_cells, new_rows_sh = json_to_cells(new_rows_dic, COL_NAMES, free_row)


if new_cells:
    log_info(str(len(new_cells))+" new cells ready to be uploaded", tabs=1)
    try:
        wks.update_cells(new_cells, value_input_option='USER_ENTERED')
        new_free_row = next_available_row(wks)
        log_info("Spreadsheet upload OK", tabs=1)
        log_info("New data from row "+str(free_row)+" to row "+str(new_free_row-1), tabs=1)
    except:
        log_info("Spreadsheet upload NOK", tabs=1)
else:
    log_info("There is not new data to be loaded at the moment", tabs=1)

log_info("Execution finished at " + str(datetime.now()))
