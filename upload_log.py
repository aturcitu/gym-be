
from datetime import datetime
from gymdef import log_info, login_mail, load_worksheet, load_user_id, load_exercises, next_available_row, clean_json, trainning_form_mail, json_to_cells, upload_relative_load

#--- PREPARE MAIL ACCOUNT AND GET INFO ---#
log_info("Execution started at " + str(datetime.now()))
#log into mail account
my_mail, mail_user = login_mail()
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
wks, sh = load_worksheet()
if wks:
    user_ids = load_user_id(sh)
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
            "tipoEntrenamiento": json_input['tipoEntrenamiento']}
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
    # Using new data recompute relative load
    upload_relative_load(wks)
else:
    log_info("There is not new data to be loaded at the moment", tabs=1)

log_info("Execution finished at " + str(datetime.now()))
