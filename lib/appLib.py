import os
from PyPDF2 import PdfFileWriter, PdfFileReader  # (for paychecks)
import fitz # PyMuPDF (for badges)
import json
import mysql

# PATHS
config_path = "../config_files/db/config.json"
logo_path = "../config_files/imgs/BusinessCat.png"
icon_path = "../config_files/imgs/Cat.ico"

# COLORS
default_background = "#ffffff"
color_light_orange = "#fff6e5"
color_green = "#80ff80"
color_yellow = "#ffeb64"
color_red = "#ff4d4d"
color_orange = "#ff9632"


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
    try:

        inputpdf = PdfFileReader(open(file_to_split, "rb"))
        lookup_interval = []

        # chech if file in directory
        if not os.path.exists(file_to_split):
            raise Exception(f"Non ho trovato il file {file_to_split} nella mia stessa directory!")

        # create dir if it does not exists
        if not os.path.exists(dirname):
            os.mkdir(dirname)

            # iterate throught pages and split them
            for page in range(inputpdf.numPages):
                pageObj = inputpdf.getPage(page)

                # gestisco le iterazioni
                if lookup_interval:
                    lookup_interval, matching, name = read_page(pageObj, lookup_interval)
                else:
                    lookup_interval, matching, name = read_page(pageObj)

                '''
                lookup interval - (intervallo col nome appena letto da paragonare col prossimo foglio)
                matching - (un nome in caso di match affermativo o None in caso di foglio singolo)
                '''

                # genero il pdf
                output = PdfFileWriter()

                # se matchano aggiungo all'output prima la pagina precedente
                if matching:
                    output.addPage(inputpdf.getPage(page - 1))

                output.addPage(inputpdf.getPage(page))

                with open(f"BUSTE PAGA/{name}.pdf", "wb") as outputStream:
                    output.write(outputStream)

    except:
        os.remove(dirname)
        raise Exception("####################################\n"
                        "ERRORE nella generazione delle buste\n"
                        "####################################\n")

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



''' DB UTILS '''

def load_config():
    try:
        with open(config_path, "r") as f:
            config = json.load(f)
    except:
        raise Exception("Cannot find config file")

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
