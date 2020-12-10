import os
import fitz # PyMuPDF
import json
import mysql
import smtplib
from threading import Thread

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


