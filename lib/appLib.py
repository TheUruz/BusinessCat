import os
import fitz # PyMuPDF
import json
import mysql
import smtplib
from threading import Thread
from operator import itemgetter

import pandas as pd

# PATHS
db_config_path = "../config_files/db/config.json"
email_config_path = "../config_files/email_service/config.json"
logo_path = "../config_files/imgs/BusinessCat.png"
icon_path = "../config_files/imgs/Cat.ico"

# COLORS
default_background = "#ffffff"
color_light_orange = "#fff6e5"
color_green = "#009922"
color_yellow = "#ffe01a"
color_red = "#ff3333"
color_orange = "#e59345"


''' UTILITIES '''

def read_page(pageObj, lookup_interval=[]):
    def check_match(lookup_interval, splitted):
        months = [
            "Gennaio",
            "Febbraio",
            "Marzo",
            "Aprile",
            "Maggio",
            "Giugno",
            "Luglio",
            "Agosto",
            "Settembre",
            "Ottobre",
            "Novembre",
            "Dicembre"
        ]

        matching = False
        name = None

        # per ogni cosa nell'intervallo di questa pagina
        for x in splitted:
            x = str(x)
            if (x not in months) and not ((x.strip()).isdigit()) and not (x.strip() == ""):

                # checko che non sia un nome ma comprenda anche i digit
                found_digit = False
                for char in x:
                    if char.isdigit():
                        found_digit = True
                        name = x

                # se non trovo un numero nella stringa passo al prossimo valore
                if not found_digit:
                    continue

                # controllo il matching col foglio precedente
                for y in lookup_interval:
                    if y == x:
                        matching = True

        print(name)

        if matching:
            return (True, name)
        else:
            return (False, name)

    # estratto il testo dall pagina
    text = pageObj.extractText()

    # recupero l'intervallo interessato di questa pagina
    splitted = text.split(" ")[420:470]

    # CHECK MATCH CON LA PAGINA PRECEDENTE
    matching, name = check_match(lookup_interval, splitted)

    # CONDITIONAL RETURN SULLA BASE DEL MATCH CON IL FOGLIO PRECEDENTE
    if matching:
        # print(f"pagine {page -1} - {page} : busta paga di -> {name}")
        return (lookup_interval, True, name)
    else:
        # print(f"pagina {page} : busta paga di  -> {name}")
        lookup_interval = splitted
        return (lookup_interval, False, name)

def CREATE_BUSTE(file_to_split, dirname):
    mesi = [
        "gennaio",
        "febbraio",
        "marzo",
        "aprile",
        "maggio",
        "giugno",
        "luglio",
        "agosto",
        "settembre",
        "ottobre",
        "novembre",
        "dicembre"
    ]
    inputpdf = fitz.open(file_to_split)
    check_name = ""

    # possible slots where the name is stored in the pdf {slot:array_split_at}
    lookup_range = {
    19:1,
    21:1,
    117:0
}

    # chech if file in directory
    if not os.path.exists(file_to_split):
        raise Exception(f"Non ho trovato il file {file_to_split} nella mia stessa directory!")

    # create dir if it does not exists
    if not os.path.exists(dirname):
        os.mkdir(dirname)

    # per ogni pagina
    for i in range(inputpdf.pageCount):
        page = inputpdf.loadPage(i)
        blocks = page.getText("blocks")
        blocks.sort(key=lambda block: block[1])  # sort vertically ascending

        paycheck_owner = None

        for index, b in enumerate(blocks):

            # per ogni range in cui è possibile trovare il nome
            for key in lookup_range:
                if index == key:

                    name = b[4]

                    if not name.startswith("<"):
                        name = name.split('\n')[lookup_range[key]]

                    if key != 117:
                        if name != "Codicesdipendente" \
                                and name != "CodicesFiscale" \
                                and not name[0].isdigit() \
                                and not name.startswith("<"):
                            name_ = name.split(" ")
                            new_name = ""
                            for word in name_:
                                new_name += word[0].upper() + word[1:].lower() + " "
                            paycheck_owner = new_name[:-1]
                    else:
                        name_ = (name[12:]).split()[:-1]
                        new_name = ""
                        for word in name_:
                            new_name += word[0].upper() + word[1:].lower() + " "
                        paycheck_owner = new_name[:-1]

        if not paycheck_owner:
            raise Exception(f"Errore non ho trovato il proprietario della busta numero {i}")
        else:
            paycheck = fitz.Document()

            # se il nome è lo stesso del foglio precedente salvo il pdf come due facciate
            if paycheck_owner == check_name:
                paycheck.insertPDF(inputpdf, from_page=i - 1, to_page=i)
            # altrimenti salvo solo la pagina corrente
            else:
                paycheck.insertPDF(inputpdf, from_page=i, to_page=i)

            check_name = paycheck_owner
            paycheck.save(dirname + "/" + paycheck_owner + ".pdf")

def CREATE_CARTELLINI(file_to_split, dirname):
    try:

        inputpdf = fitz.open(file_to_split)

        # chech if file in directory
        if not os.path.exists(file_to_split):
            raise Exception(f"Non ho trovato il file {file_to_split} nella mia stessa directory!")

        # create dir if it does not exists
        if not os.path.exists(dirname):
            os.mkdir(dirname)

        # per ogni pagina
        for i in range(inputpdf.pageCount):
            page = inputpdf.loadPage(i)
            obj = page.getTextBlocks()
            full_name = ""

            # debug
            # print(obj[37])

            # provo a cercare il nome in uno slot
            try:
                for index, x in enumerate(obj[37]):
                    if not str(x)[0].isdigit():
                        name_arr = (x.split("\n")[0]).split(" ")

                        for string in range(len(name_arr)):
                            if name_arr[string]:
                                full_name += (name_arr[string][0].upper() + name_arr[string][1:].lower() + " ")

                        full_name = (full_name[:-1])

                if not full_name:
                    raise

            # altrimenti lo cerco nell'altro
            except:
                for index, x in enumerate(obj[36]):
                    if not str(x)[0].isdigit():
                        name_arr = (x.split("\n")[0]).split(" ")

                        for string in range(len(name_arr)):
                            if name_arr[string]:
                                full_name += (name_arr[string][0].upper() + name_arr[string][1:].lower() + " ")

                        full_name = (full_name[:-1])

                if not full_name:
                    raise

            # print(full_name)

            # salvo il cartellino
            badge = fitz.Document()
            badge.insertPDF(inputpdf, from_page=i, to_page=i)
            badge.save(dirname + "/" + full_name + ".pdf")


    except:
        os.remove(dirname)
        raise Exception("#######################################\n"
                        "ERRORE nella generazione dei cartellini\n"
                        "#######################################\n")

def START_in_Thread(process):
    """prevent process to freeze the app upon launch"""
    Thread(target=process).start()

def check_paycheck_badges():

    done_paycheck = False
    done_badges = False

    if os.path.exists("BUSTE PAGA"):
        if len(os.listdir("BUSTE PAGA")) > 0:
            done_paycheck = True
    if os.path.exists("CARTELLINI"):
        if len(os.listdir("CARTELLINI")) > 0:
            done_badges = True

    check = (done_paycheck, done_badges)

    return check


''' MAILS '''

def load_email_server_config():
    try:
        with open(email_config_path, 'r') as f:
            config = json.load(f)
        return config
    except:
        raise Exception("Cannot find email server config file")

def connect_to_mail_server():
    config = load_email_server_config()
    smtp = smtplib.SMTP_SSL(config['server'], config['port'], timeout=10)
    smtp.ehlo()
    try:
        smtp.login(config['email'], config['password'])
        print(f"""CONNECTED SUCCESSFULLY:\n
        -> SERVER {config['server']}\n
        -> PORT {config['port']}\n
        -> USER {config['email']}\n
        -> PWD {"*"*len(str(config['password']))}
        """)
    except:
        raise Exception("Login al server non riuscito")

    return smtp


''' DB UTILS '''

def load_db_config():
    try:
        with open(db_config_path, "r") as f:
            config = json.load(f)
    except:
        raise Exception("Cannot find db config file")

    return config

def connect_to_db(config):

    db = mysql.connector.connect(
        host=config['host'],
        password=config['password'],
        username=config['username'],
        database=config['database'],
        port=config['port']
    )

    return db

def add_user(db, cursor, email, pwd, product_key, workstations):
    workstations = json.dumps(workstations)

    task = f"INSERT INTO users (email, pwd, product_key, workstations) VALUES ('{email}','{pwd}','{product_key}','{workstations}')"
    print(task)

    cursor.execute(task)
    print(f"-> user added")
    db.commit()
    return cursor

def check_registered(cursor, email):
    already_registered = False

    task = f"SELECT * FROM users WHERE email = '{email}'"
    cursor.execute(task)

    if cursor.fetchone():
        already_registered = True

    return already_registered







"""
PDF_file = '../test_data/controllo.pdf'
check_values = {
    "festivita' godute": -1,
    "ferie godute": -1,
    "permessi rol goduti": -1,
    "festivita' non godute": -1,
    "indennita' mulettista 2": -1,
    "buono pasto elettronico": -1,
    "assegno nucleo familiare": -1,
    "fesica-confsal": -1,
    "trattenuta sindacale": -1,
    "trattamento integrativo l.": -1,
    "quota t.f.r.": -1,
    "indennita' container": -1,
    "recupero acconto": -1,
    "malattia inps 50%": -1,
    "int. malattia": -1,
    "lordizzazione inps": -1,
    "cessione":-1,
    "cessione 1/5 stipendio": -1,
    "trattenuta p.": -1,
}
col_to_merge = {
    "Ferie/fest. pagate": ["festivita' godute", "festivita' non godute","permessi rol goduti", "ferie godute"],
    "T.F.R.": ["quota t.f.r", "quota t.f.r. a fondi"],
    "Trattenute sindacali": ["fesica-confsal", "trattenuta sindacale"],
    "Cessione 1/5": ["cessione","cessione 1/5 stipendio", "trattenuta p."]
}

def parse_block_content(array):
    key = ""
    value = []

    index = 0

    while index < len(array):

        item = array[index].strip()
        if item:
            try:
                item = item.replace('€','').strip()
                if ',' in item:
                    split_item = item.split(',')
                    if len(split_item)<2:
                        for i_ in split_item:
                            if not i_.isnumeric():
                                raise

                    item = item.replace('.','').replace(',','.')
                    item = float(item)
                else:
                    item = int(item)
                value.append(item)
            except:
                key += (str(item) + ' ')

        index += 1

    return {key[:-1]:value}

def page_extractor(PDF):
    inputpdf = fitz.open(PDF)
    total_content = {}

    # per ogni pagina
    for i in range(inputpdf.pageCount):
        page = inputpdf.loadPage(i)
        blocks = page.getText("blocks")
        blocks.sort(key=lambda block: block[1])  # sort vertically ascending


        ####### find name in page
        name = None
        for index, b in enumerate(blocks):
            for elem in b:
                elem = str(elem)
                if ('cognome' in elem and 'nome' in elem) or ('COGNOME' in elem and 'NOME' in elem):
                    name_array = (blocks[index+1][4]).split('\n')

                    # conditional fix array
                    if len(name_array) == 2:
                        for check_ in name_array:
                            if "cessat" in check_.lower():
                                name_array = (blocks[index+2][4]).split('\n')
                                break

                    #check if there's a name in the array wich was found in the block
                    for value in name_array:
                        valid = True
                        for char in value:
                            if char.isdigit():
                                valid = False
                                break

                        if valid and value:
                            if value != None and not value.isspace() and value[0].isalpha():
                                name = value
                                break

        #setup name in return object
        if name not in total_content.keys(): total_content[name] = []

        ####### get page content
        page_content = []
        for index, b in enumerate(blocks):
            content = b[4].split('\n')
            content = parse_block_content(content)
            page_content.append(content)

        total_content[name].append(page_content)








    return total_content

def parse_pages_data(total_content, check_values):

    parsed_content = {}

    for worker in total_content:
        if worker:
            parsed_content[worker] = {}
            worker_content = total_content[worker]

            for page in worker_content:

                for check_ in check_values:
                    val = None
                    for index, elem in enumerate(page):
                        for e in elem:
                            if check_ in e.lower():

                                # parse key
                                key_words = e.split(' ')
                                key = ""
                                for part in key_words:
                                    if part[0].isalpha() and part != "ORE" and part != "GG":
                                        key += (part + ' ')
                                key = key[:-1]


                                if val:
                                    if key in val:
                                        val[key] += elem[e][check_values[check_]]
                                else:
                                    val = {key:elem[e][check_values[check_]]}

                    if val:
                        parsed_content[worker].update(val)

    return parsed_content

def merge_columns(parsed_data, columns_to_merge):

    merged_data = {}


    for worker in parsed_data:
        worker_new_data = {}
        processed_keys = []

        for key in list(parsed_data[worker]):
            edited = False

            # se la chiave non è già stata processata
            if key not in processed_keys:
                for col_name in columns_to_merge:
                    for component in columns_to_merge[col_name]:
                        if (component.lower() in key.lower()) and (key.lower() not in processed_keys):

                            if col_name in worker_new_data:
                                worker_new_data[col_name] += parsed_data[worker][key]
                            else:
                                worker_new_data[col_name] = parsed_data[worker][key]

                            processed_keys.append(key.lower())
                            edited = True

            if not edited:
                worker_new_data[key] = parsed_data[worker][key]

        merged_data[worker] = worker_new_data

    return merged_data

def create_Excel(data,sorted=True):
    df = pd.DataFrame.from_dict(data).T

    if sorted:
        df = df.sort_index()

    with pd.ExcelWriter('test.xlsx') as writer:
        df.to_excel(writer, sheet_name='Foglio1')

parsed_content = parse_pages_data(page_extractor(PDF_file), check_values)
parsed_content = merge_columns(parsed_content, col_to_merge)
#create_Excel(parsed_content)
"""
