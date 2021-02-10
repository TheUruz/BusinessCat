import os
import copy
import io
import pickle
import fitz # PyMuPDF
import json
import mysql
import smtplib
import openpyxl

from openpyxl.utils.cell import get_column_letter
from openpyxl.styles import PatternFill
from threading import Thread
from operator import itemgetter
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request, AuthorizedSession
from email.message import EmailMessage
from email.mime.text import MIMEText

import pandas as pd

# PATHS
db_config_path = "../config_files/db/config.json"
email_config_path = "../config_files/email_service/config.json"
google_config_path = "../config_files/google/google_userconfig.json"
logo_path = "../config_files/imgs/BusinessCat.png"
icon_path = "../config_files/imgs/Cat.ico"

# COLORS
default_background = "#ffffff"
color_light_orange = "#fff6e5"
color_green = "#009922"
color_yellow = "#ffe01a"
color_red = "#ff3333"
color_orange = "#e59345"

# GOOGLE VARIABLES
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = '../config_files/google/google_credentials.json'
creds = None
SCOPE = [
    'https://www.googleapis.com/auth/drive'
]


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

def get_sheetnames_from_bytes(bytes_):
    wb = openpyxl.load_workbook(bytes_)
    sheetnames = wb.sheetnames
    return sheetnames


''' GOOGLE API METHODS '''

def authenticate(func):
    def auth_wrapper(*args, **kwargs):

        global creds
        global SCOPE

        # load token.pickle if present
        if os.path.exists('../config_files/google/token.pickle'):
            with open('../config_files/google/token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            # ? if token needs to be refreshed it will be refreshed
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            # ? otherwise authenticate
            else:
                flow = InstalledAppFlow.from_client_secrets_file(os.environ['GOOGLE_APPLICATION_CREDENTIALS'], SCOPE)
                creds = flow.run_local_server(port=0)

            with open('../config_files/google/token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        # raise if user does not authenticate
        if not creds:
            raise Exception("You are not allowed to run this method >> Unauthorized")

        return func(*args, **kwargs)

    return auth_wrapper

@authenticate
def create_auth_session(credentials=creds):
    return AuthorizedSession(credentials)

@authenticate
def build_service(service, version="v3"):
    return build(service, version, credentials=creds)

def get_df_bytestream(ID):
    """
    get google sheet for comparison. return it as a bytestream
    """
    service = build_service("drive")
    conversion_table = service.about().get(fields="exportFormats").execute()
    request = service.files().export_media(fileId=ID,
                                           mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    fh = io.BytesIO()

    # download file bytestream
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()

    return fh

def get_comparison_df(bytestream, month):
    """
    parse bytestream as a pandas dataframe
    """

    # parse bytestream into df
    df = pd.read_excel(bytestream, sheet_name=month.upper(), index_col=0)
    return df

def get_sheetlist():
    """
    get google sheet list in google drive
    """
    service = build_service("drive")
    results = service.files().list(pageSize=10, fields="nextPageToken, files(id, name, mimeType)", q="mimeType='application/vnd.google-apps.spreadsheet'").execute()
    items = results.get('files', [])
    items = sorted(items, key=itemgetter('name'))
    return items


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


''' CLASSES '''

class PaycheckController():
    def __init__(self):
        """
        paychecks_to_check (str) -> path to multiple pages pdf containing all paychecks to check
        badges_to_check (str) -> path to folder containing all badges files as .pdf
        """
        """
        try:
            self.__validate_data(paychecks_to_check, badges_to_check)
        except Exception() as e:
            raise Exception(e)
        """

        self.badges_path = "" # path to CARTELLINI folder
        self.paychecks_to_check = "" # lul_controllo
        self.conversion_table_path = "../config_files/conversion_table.json"

        self.verify_filename = "Verifica.xlsx" #name of the output verification xlsx
        self.highlight_error = "FFFFFFFF"
        self.default_configuration = {
                "col_codes": {},
                "col_to_extract": [
                    "Z01100",
                    "Z00246",
                    "Z00255",
                    "Z01138",
                    "000279",
                    "001008",
                    "ZP0160",
                    "ZP0162",
                    "ZPS000",
                    "F02701",
                    "000282",
                    "003955",
                    "Z05031",
                    "ZP0001",
                    "ZP0030",
                    "003951",
                    "002099",
                    "002101",
                    "PLACEHOLDER1",
                    "PLACEHOLDER2",
                    "F09080",
                    "F09100",
                    "F09130",
                    "F09081",
                    "Z50022",
                    "Z50023",
                    "Z51000",
                    "Z51010",
                    "000085",
                    "000086",
                    "ZP8134",
                    "ZP8138",
                    "ZP8130",
                    "003450",
                    "000283",
                    "quota t.f.r.",
                    "quota t.f.r. a fondi"
                ],
                "merging_columns": {
                    "Ferie/fest. pagate": [
                        "Z01100",
                        "Z01138",
                        "Z00255",
                        "Z51010",
                        "Z00246",
                        "Z51000"
                    ],
                    "T.F.R.": [
                        "quota t.f.r.",
                        "quota t.f.r. a fondi",
                        "ZP8134",
                        "003450",
                        "fondo t.f.r. al 31/12"
                    ],
                    "Cessione 1/5": [
                        "002101",
                        "003951",
                        "002099"
                    ],
                    "Assegni Familiari": [
                        "ZP0160",
                        "ZP0162"
                    ]
                }
            }
        self.config = None

        self.__load_conversion_table()

    """ PRIVATE METHODS """
    def __load_conversion_table(self):
        """
        set PaycheckController configuration. with this configuration the program knows wich field
        extract from paychecks
        """
        if os.path.exists(self.conversion_table_path):
            with open(self.conversion_table_path, "r") as f:
                self.config = json.load(f)
                return
        else:
            return copy.deepcopy(self.default_configuration)


    """ PUBLIC METHODS """
    def create_config_from_csv(self, csv_path):
        """
        csv must have at least two columns 'Codice voce' and 'Descrizione voce'
        """

        with open(csv_path, 'rt')as f:
            df = pd.read_csv(f, sep=";")
            columns = {}

            for index, row_ in df.iterrows():
                row = dict(row_)

                # parse out nan values
                parsed_row = {}
                for val in row:
                    if not pd.isnull(row[val]):
                        parsed_row[val] = row[val]

                if parsed_row['Codice voce'] not in self.config:
                    columns[parsed_row['Codice voce']] = parsed_row
                    columns[parsed_row['Codice voce']]['col_value'] = -1

        # adding keys
        columns["quota t.f.r."] = {"Descrizione voce": "Quota T.F.R.", "col_value": -1}
        columns["quota t.f.r. a fondi"] = {"Descrizione voce": "Quota T.F.R. a Fondi", "col_value": -1}

        # setting self.config with parsed data
        self.config["col_codes"] = columns

        with open(self.conversion_table_path, "w") as f:
            f.write(json.dumps(self.config, indent=4, ensure_ascii=True))

        print(f"* * conversion_table.json created from this file >> {csv_path}")
        self.__load_conversion_table()

    def validate_data(self):
        if not self.paychecks_to_check:
            raise Exception("Error: paycheck to check not specified")
        if not os.path.exists(self.paychecks_to_check):
            raise Exception(f"Error: cannot find {self.paychecks_to_check}")
        if not os.path.exists(self.badges_path):
            raise Exception(f"Error: cannot find {self.badges_path}")
        if not os.path.isdir(self.badges_path):
            raise Exception(f"Error: {self.badges_path} is not a folder")
        if len(os.listdir(self.badges_path)) < 1:
            raise Exception(f"Error: {self.badges_path} is empty!")

        return True

    # setters
    def set_badges_path(self, path):
        """setter for path to badges folder. it should contain .pdf files from every badge (splitted from BusinessCat)"""
        self.badges_path = path

    def set_paychecks_to_check_path(self, path):
        """setter for path to paychecks to check. it should be a .pdf file"""
        self.paychecks_to_check = path

    # main functions
    def paycheck_verification(self, create_Excel=True):
        PDF_file = self.paychecks_to_check
        total_content = {}

        col_codes = self.config['col_codes']
        col_to_extract = self.config['col_to_extract']
        merging_columns = self.config['merging_columns']

        # if there are no columns to extract set in the configuration, raise exception
        if not col_codes:
            raise Exception("No columns specified for extraction in config file")

        inputpdf = fitz.open(PDF_file)

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

            if name:
                name = name.upper()

                # add name to total_content
                if name not in total_content:
                    total_content[name] = {}

                ####### get page content
                for index, b in enumerate(blocks):
                    content = b[4].split('\n')
                    content = [string for string in content if (string != "" and not string.startswith("   "))]

                    # find ordinary and overtime hours in paycheck
                    h_check = [h.lower().strip() for h in content]
                    if "ore ordinarie" and "ore straordinarie" in h_check:
                        w_hours = [x for x in blocks[index+2][4].split('\n') if "," in x]
                        if len(w_hours) == 1:
                            w_hours.append("0")
                        try:
                            total_content[name]['Ore ordinarie'] = float(w_hours[0].replace(",","."))
                            total_content[name]['Ore straordinarie'] = float(w_hours[1].replace(",","."))
                        except IndexError:
                            pass

                    # parsing content
                    for i, elem in enumerate(content):

                        # parse netto del mese
                        if "NETTOsDELsMESE" in elem:
                            netto = blocks[index+1][-3]
                            netto = netto.replace("€", "").strip().replace(",",".")
                            netto = netto.replace(".", "", (netto.count('.') - 1))
                            try:
                                netto = float(netto)
                                total_content[name]["Netto del Mese"] = netto
                            except:
                                pass

                        elem = (elem.replace("*", "")).strip()

                        for paycheck_code in col_codes:

                            if (paycheck_code.lower() == elem.split(" ")[0].lower()) \
                                    or (len(paycheck_code.split(" ")) > 1) and elem.lower().startswith(paycheck_code.lower()):

                                # check if value is to be extracted
                                if paycheck_code in col_to_extract:

                                    # check if it needs to be merged
                                    elem_colname = None
                                    for colname in merging_columns:
                                        for subname in merging_columns[colname]:
                                            if subname == paycheck_code:
                                                elem_colname = colname

                                    if not elem_colname:
                                        elem_colname =  col_codes[paycheck_code]['Descrizione voce']

                                    # get row value
                                    val = content[col_codes[paycheck_code]['col_value']].replace(",", ".")
                                    val = val.replace(".", "", (val.count('.') - 1))
                                    try:
                                        val = float(val)
                                    except:
                                        pass

                                    # report data in the proper column
                                    if elem_colname not in total_content[name]:
                                        total_content[name][elem_colname] = val
                                    else:
                                        total_content[name][elem_colname] += val

        if create_Excel:
            sheet_name = "Verifica Buste Paga"
            self.create_Excel(total_content, sheet_name)
            print(f"File {self.verify_filename} generato con successo, {sheet_name} aggiunto al suo interno")

    def badges_verification(self, create_Excel=True):

        def parse_decimal_time(decimal_time):

            #if 0 not present attach it
            if len(decimal_time) == 1:
                decimal_time = decimal_time + "0"

            hour_min = 60
            result = (100*int(decimal_time))/hour_min
            return str(int(result))

        total_content = {}

        #for each badge
        for file in os.listdir(self.badges_path):
            path_to_badge = "/".join((self.badges_path, file))
            full_name = None
            last_block = []

            inputpdf = fitz.open(path_to_badge)

            ####### find name
            page = inputpdf.loadPage(0)
            blocks = page.getText("blocks")
            blocks.sort(key=lambda block: block[1])  # sort vertically ascending

            for index, b in enumerate(blocks):
                if "cognome e nome" in b[4].split('\n')[0].lower():
                    full_name = blocks[index+1][-3].split('\n')[0]
                    full_name = " ".join(full_name.split())
                    break

            ###### find values
            if full_name:
                total_content[full_name] = {"Ore ordinarie": 0.0, "Ore straordinarie": 0.0}

                for index, b in enumerate(blocks):

                    ##### check if block is a day
                    if len(b[4].split()[0][:-1]) <= 2 and b[4].split()[0][-1].isalpha() and b[4].split()[0][:-1].isdigit():
                        day = b[4].split()
                        day, day_values_ = day[0], day[1:]

                        # check if data is valid
                        if len(day[1:])%2 != 0 and day[1:][0].isdigit():
                            raise Exception(f"Expected even pairs, got uneven pairs for worker {full_name} on day {day}")

                        day_values_ = list(zip(day_values_[0::2], day_values_[1::2]))
                        day_values = []

                        # parse tuples in day values_
                        for i_, tupla in enumerate(day_values_):
                            # remove errors and start/end of workshift
                            if len((tupla[0])) > 4 or (len(tupla[0]) == 4 and "," not in tupla[0]):
                                continue
                            if len(day_values) < 2:
                                day_values.append(tupla)

                        # if not hours as first value skip day
                        if day_values:
                            if len(day_values[0][0]) == 3:
                                pass
                            else:
                                # parse decimal time values to number values
                                value_to_add = day_values[0][0].replace(",", ".").split(".")
                                if len(value_to_add) == 2:
                                    if value_to_add[1] != "00" and value_to_add[1] != "0":
                                        value_to_add[1] = parse_decimal_time(value_to_add[1])
                                    value_to_add = float(".".join(value_to_add))
                                else:
                                    value_to_add = float(day_values[0][0])
                                # if there's more than a pair of values in the day
                                if len(day_values) > 1:
                                    # if hours are overtime skip them
                                    if day_values[1][0] == "306" or day_values[1][0] == "006":
                                        pass
                                    else:
                                        total_content[full_name]['Ore ordinarie'] += value_to_add
                                # otherwise assume they are ordinary hours
                                else:
                                    total_content[full_name]['Ore ordinarie'] += value_to_add

                    ##### check if block is last and parse its data
                    if b[4].split()[0].isdigit() and not b[4].split()[0][-1].isalpha() and len(b[4].split()[0])<=3:
                        last_block_values = b[4].split()

                        # parse values
                        for val in last_block_values:
                            is_number = True
                            for letter in val:
                                if letter.isalpha() and letter != ",":
                                    is_number = False
                            # append conditionally
                            if is_number:
                                if "," in val:
                                    last_block.append(float(val.replace(",",".")))
                                else:
                                    last_block.append(int(val))
                            else:
                                if type(last_block[-1]) == int or type(last_block[-1]) == float:
                                    last_block.append(val)
                                else:
                                    last_block[-1] = last_block[-1] + " " + val

                        # check if data is valid
                        if len(last_block)%3 != 0:
                            raise Exception(f"Expected triplets on badge footer values for worker {full_name}")

                ##### last fixes
                total_content[full_name]["Ore ordinarie"] = total_content[full_name]["Ore ordinarie"]
                last_block = list(zip(last_block[0::3], last_block[1::3], last_block[2::3]))
                for pair in last_block:

                    #parse pair[2] time value to number value
                    pair_value = str(pair[2]).replace(",", ".").split(".")
                    if len(pair_value) == 2:
                        if pair_value[1] != "00" and pair_value[1] != "0":
                            pair_value[1] = parse_decimal_time(pair_value[1])
                        pair_value = float(".".join(pair_value))

                    if pair[0] == 306 or pair[0] == 6:
                        total_content[full_name]["Ore straordinarie"] += pair_value
                    total_content[full_name][str(pair[0]) + " " + str(pair[1])] = pair_value

        if create_Excel:
            sheet_name = "Verifica Cartellini"
            self.create_Excel(total_content, sheet_name)
            print(f"File {self.verify_filename} generato con successo, {sheet_name} aggiunto al suo interno")

    def create_Excel(self, content, sheet_name, transposed=True):
        df = pd.DataFrame.from_dict(content)

        if transposed:
            df = df.T

        # sort alphabetically rows and columns
        df = df.sort_index()
        df = df.reindex(sorted(df.columns), axis=1)

        open_mode = "a" if os.path.exists(self.verify_filename) else "w"

        with pd.ExcelWriter(self.verify_filename, mode=open_mode) as writer:
            df.to_excel(writer, sheet_name=sheet_name)

    def compare_badges_to_paychecks(self, keep_refer_values=True):

        CHECK_SUFFIX = " PAYCHECK"
        badges_df = pd.read_excel(self.verify_filename, sheet_name="Verifica Cartellini", index_col=0).fillna(0)
        paychecks_df = pd.read_excel(self.verify_filename, sheet_name="Verifica Buste Paga", index_col=0).fillna(0)

        # set indexes name
        badges_df.index.name = "LAVORATORI"
        paychecks_df.index.name = "LAVORATORI"

        # uniform indexes
        badges_df.index = badges_df.index.str.upper()
        paychecks_df.index = paychecks_df.index.str.upper()

        # create df with all data
        common_columns = set(badges_df.columns.values).intersection(set(paychecks_df.columns.values))
        common_df = paychecks_df[list(common_columns)].copy()
        renaming = {key: key + CHECK_SUFFIX for key in common_df.columns.values}
        common_df = common_df.rename(columns=renaming)

        # fix wrong indexes in badges
        badges_df = badges_df.rename(index={'COBIANCHI MARCO': 'COBIANCHI MARCO GABRIELE',
                                         'GUZMAN URENA ALEXANDER': 'GUZMAN URENA ALEXANDER DE JESUS',
                                         'NUTU LOREDANA ADRIAN': 'NUTU LOREDANA ADRIANA'
                                         })

        data_df = badges_df.merge(common_df, left_index=True, right_index=True)
        self.create_Excel(data_df, sheet_name="temp", transposed=False)

        destination_workbook = openpyxl.load_workbook(self.verify_filename)
        ws = destination_workbook["temp"]

        # find column of columns to highlight
        headings = [row for row in ws.iter_rows()][0]
        headings = [x.value for x in headings]
        matching_ = dict.fromkeys(common_columns, 0)
        matching = {}
        for col in matching_:
            matching[col] = 0
            matching[col + CHECK_SUFFIX] = 0
        for col in headings:
            if col in matching.keys():
                matching[col] = headings.index(col)

        for index, row in enumerate(ws.iter_rows()):
            # headers exclueded
            if index != 0:
                row_values = [x.value for x in row]
                worker_check = {}
                worker_errors = []

                # gather worker data
                for val in matching:
                    worker_check[val] = row_values[matching[val]]

                # find worker errors
                for data in worker_check:
                    check_val = data + CHECK_SUFFIX
                    if CHECK_SUFFIX not in data and check_val in worker_check:
                        if worker_check[data] - worker_check[check_val] != 0:
                            worker_errors.append(data)
                            worker_errors.append(check_val)

                if worker_errors:
                    # parse errors to cells
                    highlight_row = index + 1
                    for _i, error in enumerate(worker_errors):
                        highlight_column = get_column_letter(matching[error]+1)
                        worker_errors[_i] = str(highlight_column) + str(highlight_row)

                    for c in worker_errors:
                        cell = ws[c]
                        cell.fill = PatternFill(start_color='FFEE1111', end_color='FFEE1111', fill_type='solid')

        # drop refer columns conditionally
        if not keep_refer_values:
            col_to_remove = []
            for val in matching:
                if CHECK_SUFFIX in val:
                    col_to_remove.append(matching[val]+1)
            for val in sorted(col_to_remove, reverse=True):
                ws.delete_cols(val)


        #replace old verification with edited one
        destination_workbook.remove(destination_workbook["Verifica Cartellini"])
        ws.title = "Verifica Cartellini"
        destination_workbook.save(self.verify_filename)

        print(f">> BADGES COMPARED WITH PAYCHECKS SUCCESSFULLY")

        """
        # funziona male. evidenzia giusto ma non tutto
        styled_df = data_df.style
        for column in common_columns:
            col_to_check = column + " PAYCHECK"
            styled_df = styled_df.apply(lambda i_: ["background-color:red" if data_df.iloc[i_][col_to_check] - data_df.iloc[i_][column] != 0 else "" for i_, row in enumerate(data_df.iterrows())], subset=[column], axis=0)
        styled_df.to_excel("test.xlsx", engine="openpyxl", index=True)
        """

    def compare_paychecks_to_drive(self, df_bytestream, sheet, keep_refer_values=True):

        CHECK_SUFFIX = " DRIVE"
        drive_df = get_comparison_df(df_bytestream, sheet).fillna(0)
        paychecks_df = pd.read_excel(self.verify_filename, sheet_name="Verifica Buste Paga", index_col=0).fillna(0)

        # set indexes name
        paychecks_df.index.name = "LAVORATORI"
        drive_df.index.name = "LAVORATORI"

        # uniform indexes
        drive_df.index = drive_df.index.str.upper()
        paychecks_df.index = paychecks_df.index.str.upper()

        # create df with all data
        common_columns = set(drive_df.columns.values).intersection(set(paychecks_df.columns.values))
        common_df = drive_df[list(common_columns)].copy()
        renaming = {key: key + CHECK_SUFFIX for key in common_df.columns.values}
        common_df = common_df.rename(columns=renaming)
        data_df = paychecks_df.merge(common_df, left_index=True, right_index=True)

        self.create_Excel(data_df, sheet_name="temp", transposed=False)

        destination_workbook = openpyxl.load_workbook(self.verify_filename)
        ws = destination_workbook["temp"]

        # find column of columns to highlight
        headings = [row for row in ws.iter_rows()][0]
        headings = [x.value for x in headings]
        matching_ = dict.fromkeys(common_columns, 0)
        matching = {}
        for col in matching_:
            matching[col] = 0
            matching[col + CHECK_SUFFIX] = 0
        for col in headings:
            if col in matching.keys():
                matching[col] = headings.index(col)

        for index, row in enumerate(ws.iter_rows()):
            # headers exclueded
            if index != 0:
                row_values = [x.value for x in row]
                worker_check = {}
                worker_errors = []

                # gather worker data
                for key in matching:
                    val = row_values[matching[key]]
                    if isinstance(val, str) and "€" in val:
                        val = val.replace("€", "").replace("-", "").replace(",", ".").strip()
                    worker_check[key] = float(val) if val else 0

                # find worker errors
                for data in worker_check:
                    check_val = data + CHECK_SUFFIX
                    if CHECK_SUFFIX not in data and check_val in worker_check:
                        try:
                            if worker_check[data] - worker_check[check_val] != 0:
                                worker_errors.append(data)
                                worker_errors.append(check_val)
                        except Exception as e:
                            print(e)

                if worker_errors:
                    # parse errors to cells
                    highlight_row = index + 1
                    for _i, error in enumerate(worker_errors):
                        highlight_column = get_column_letter(matching[error] + 1)
                        worker_errors[_i] = str(highlight_column) + str(highlight_row)

                    for c in worker_errors:
                        cell = ws[c]
                        cell.fill = PatternFill(start_color='FFEE1111', end_color='FFEE1111', fill_type='solid')

        # drop refer columns conditionally
        if not keep_refer_values:
            col_to_remove = []
            for val in matching:
                if CHECK_SUFFIX in val:
                    col_to_remove.append(matching[val] + 1)
            for val in sorted(col_to_remove, reverse=True):
                ws.delete_cols(val)

        # replace old verification with edited one
        destination_workbook.remove(destination_workbook["Verifica Buste Paga"])
        ws.title = "Verifica Buste Paga"
        destination_workbook.save(self.verify_filename)

        print(f">> PAYCHECKS COMPARED WITH DRIVE {sheet} VALUES SUCCESSFULLY")
