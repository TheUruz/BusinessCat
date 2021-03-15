import os
import copy
import io
import re
import pickle
import fitz # PyMuPDF
import json
import mysql
import smtplib
import openpyxl
import xlrd
import datetime
import holidays
import random
import pandas as pd
import numpy as np

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
color_grey = "#e3e3e3"

# GOOGLE VARIABLES
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = '../config_files/google/google_credentials.json'
creds = None
SCOPE = [
    'https://www.googleapis.com/auth/drive'
]


''' UTILITIES '''

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
    """ open bytestream as openpyxl workbooks and return an array of worksheet names """
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
                    "Z05075",
                    "ZP0001",
                    "ZP0003",
                    "ZP0030",
                    "003951",
                    "002099",
                    "002101",
                    "003802",
                    "002100",
                    "F09080",
                    "F09100",
                    "F09130",
                    "F09081",
                    "Z50022",
                    "Z50023",
                    "Z51000",
                    "Z51010",
                    "Z05004",
                    "Z05041",
                    "Z05065",
                    "Z05060",
                    "ZP0029",
                    "000085",
                    "000229",
                    "000086",
                    "ZP8134",
                    "ZP8138",
                    "ZP8130",
                    "003450",
                    "003411",
                    "000283",
                    "000031",
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
                        "Z51000",
                        "000031"
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
                        "002100",
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
                new_config = json.load(f)
                self.config = new_config
                return
        else:
            self.config = copy.deepcopy(self.default_configuration)


    """ PUBLIC METHODS """
    def create_config_from_csv(self, csv_path):
        """
        csv must have at least two columns 'Codice voce' and 'Descrizione voce'
        """
        new_config = copy.deepcopy(self.config)

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

                if parsed_row['Codice voce'] not in new_config["col_codes"]:
                    columns[parsed_row['Codice voce']] = parsed_row
                    columns[parsed_row['Codice voce']]['col_value'] = -1

        # adding keys
        columns["quota t.f.r."] = {"Descrizione voce": "Quota T.F.R.", "col_value": -1}
        columns["quota t.f.r. a fondi"] = {"Descrizione voce": "Quota T.F.R. a Fondi", "col_value": -1}

        # setting new_config with parsed data and dumping it
        new_config["col_codes"] = columns
        with open(self.conversion_table_path, "w") as f:
            f.write(json.dumps(new_config, indent=4, ensure_ascii=True))

        self.__load_conversion_table()
        print(f"* * conversion_table.json created from this file >> {csv_path}")

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
        badges_df = pd.read_excel(self.verify_filename, sheet_name="Verifica Cartellini", index_col=0)
        paychecks_df = pd.read_excel(self.verify_filename, sheet_name="Verifica Buste Paga", index_col=0)

        # set indexes name
        badges_df.index.name = "LAVORATORI"
        paychecks_df.index.name = "LAVORATORI"

        # fix wrong indexes in badges
        badges_df = badges_df.rename(index={
                                         'GUZMAN URENA ALEXANDER DE JE': 'GUZMAN URENA ALEXANDER DE JESUS',
                                         'NUTU LOREDANA ADRIAN': 'NUTU LOREDANA ADRIANA'
                                         })

        # uniform indexes
        badges_df.index = badges_df.index.str.upper()
        paychecks_df.index = paychecks_df.index.str.upper()

        # create df with all data
        common_columns = set(badges_df.columns.values).intersection(set(paychecks_df.columns.values))
        common_df = paychecks_df[list(common_columns)].copy()
        renaming = {key: key + CHECK_SUFFIX for key in common_df.columns.values}
        common_df = common_df.rename(columns=renaming)


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

                        if worker_check[data] == 0 and worker_check[check_val] == None:
                            continue
                        else:
                            if (worker_check[data]>0 and worker_check[check_val] == None) or (worker_check[data] - worker_check[check_val] != 0):
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

    def compare_paychecks_to_drive(self, df_bytestream, sheet, keep_refer_values=True, leave_blanks=False):

        CHECK_SUFFIX = " DRIVE"
        drive_df = get_comparison_df(df_bytestream, sheet)
        drive_df = drive_df[drive_df.index.notnull()]
        paychecks_df = pd.read_excel(self.verify_filename, sheet_name="Verifica Buste Paga", index_col=0)

        # set indexes name
        paychecks_df.index.name = "LAVORATORI"
        drive_df.index.name = "LAVORATORI"

        # uniform indexes
        drive_df.index = drive_df.index.str.upper()
        paychecks_df.index = paychecks_df.index.str.upper()

        # check divergences on dataframes
        problems = {"uncommon_indexes": [],"different_lenght": False, "error_string": ""}
        uncommon_indexes = list(set(drive_df.index.values) - set(paychecks_df.index.values))
        same_index_length = True if len(drive_df.index) - len(paychecks_df.index) == 0 else False
        # if there are not hired people in drive_df
        if (uncommon_indexes and not same_index_length):
            problems["uncommon_indexes"] = uncommon_indexes
            problems["error"] = "Non assunti sul Drive"
        # if there are typos in drive_df
        elif (uncommon_indexes and same_index_length):
            problems["uncommon_indexes"] = uncommon_indexes
            problems["error"] = "Errori di scrittura sul Drive"

        # merge dataframes
        common_columns = list(set(drive_df.columns.values).intersection(set(paychecks_df.columns.values)))
        common_df = drive_df[common_columns]
        common_df = common_df.rename(columns={key: key + CHECK_SUFFIX for key in common_df.columns.values})
        #data_df = pd.merge_ordered(left=common_df.reset_index(), right=paychecks_df.reset_index(), left_on="LAVORATORI", right_on="LAVORATORI", left_by="LAVORATORI").set_index("LAVORATORI").sort_index()
        data_df = paychecks_df.merge(common_df, left_index=True, right_index=True).sort_index()
        self.create_Excel(data_df, sheet_name="temp", transposed=False)


        destination_workbook = openpyxl.load_workbook(self.verify_filename)
        ws = destination_workbook["temp"]

        # create empty rows for uncommon_indexes (shouldnt be used)
        if leave_blanks:
            index_checkup = {k: v for k, v in enumerate(common_df.index.values)}
            # add blank lines based on index_checkup
            rows = list(enumerate(ws.iter_rows()))
            added_rows = 0
            for row in rows[::-1]:
                i_ = row[0]+1
                w_name = row[1][0].value
                try:
                    if index_checkup[i_-added_rows] != w_name:
                        ws.insert_rows(i_+added_rows)
                        added_rows += 1
                except:
                    pass

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

                # gather worker data from his row
                for key in matching:
                    val = row_values[matching[key]]
                    if isinstance(val, str) and "€" in val:
                        val = val.replace("€", "").replace("-", "").replace(",", ".").strip()
                        worker_check[key] = float(val) if val else 0
                    elif not val:
                        worker_check[key] = 0
                    elif isinstance(val, float) or isinstance(val, int):
                        worker_check[key] = val

                # find worker errors
                for data in worker_check:
                    if CHECK_SUFFIX not in data:
                        try:
                            if worker_check[data] - worker_check[data + CHECK_SUFFIX] != 0:
                                worker_errors.append(data)
                                worker_errors.append(data + CHECK_SUFFIX)
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
        return problems

class BillingManager():
    def __init__(self, month=datetime.datetime.now().month, year=datetime.datetime.now().year):
        self.bill_name = "Fattura di prova.xlsx"
        self.badges_path = None #badges_path
        self.regex_day_pattern = "([1-9]|[12]\d|3[01])[LMGVSF]"
        self.name_cell = "B5" # in che cella del badge_path si trova il nome
        self.pairing_schema = {
            "COD QTA": ["COD", "QTA"],
            "ENT USC": ["ENT", "USC"],
            "GIOR PROG": ["GIOR", "PROG"]
        }
        self.billing_schema = None
        self.total_content = None


        # config paths
        #self.__bills_path = "../config_files/BusinessCat billing/bills.json"
        self.__billing_profiles_path = "../config_files/BusinessCat billing/billing_profiles.json"
        self.__jobs_path = "../config_files/BusinessCat billing/jobs.json"

        # load data from config paths
        #self.__load_bills()
        self.__load_billing_profiles()
        self.__load_jobs()

        # set billing time
        self.set_billing_time(month, year)

        print(">> BillingManager Ready")


    """    PRIVATE METHODS    """
    def __get_engine(self):
        """get engine conditional based on extension of self.badges_path"""
        if not self.badges_path:
            raise Exception("badges_path missing!")
        elif self.badges_path.rsplit(".")[-1] == "xlsx":
            engine = "openpyxl"
        elif self.badges_path.rsplit(".")[-1] == "xls":
            engine = "xlrd"
        else:
            raise TypeError("self.badges_path is not an Excel!")
        return engine

    def __load_bills(self):
        """ read and load current billing_profiles file """
        with open(self.__bills_path,"r") as f:
            self.bills = json.load(f)
        print("** bills caricati")

    def __load_billing_profiles(self):
        """ read and load current billing_profiles file """
        with open(self.__billing_profiles_path,"r") as f:
            self.billing_profiles = json.load(f)
        print("** billing_profiles caricati")

    def __load_jobs(self):
        """ read and load current billing_profiles file """
        with open(self.__jobs_path,"r") as f:
            self.jobs = json.load(f)
            self.jobs_namelist = sorted([job["name"] for job in self.jobs])
            self.jobs_namelist.insert(0,"")
        print("** jobs caricati")

    def __load_Excel_badges(self):
        """Load excel data of badges file (must be .xls or .xlsx)"""
        engine = self.__get_engine()
        try:
            if engine == "openpyxl":
                xlsx_data = openpyxl.load_workbook(self.badges_path)
                sheet_names = xlsx_data.sheetnames
            elif engine == "xlrd":
                xlsx_data = xlrd.open_workbook(self.badges_path, on_demand=True, logfile=open(os.devnull, 'w'))
                sheet_names = xlsx_data.sheet_names()
            else:
                raise
            return xlsx_data, sheet_names, engine
        except Exception as e:
            raise Exception(f"Cannot load_Excel_badges. Error: {e}")

    def __manage_columns(self, df):
        """ private method to fix original column names"""
        fixed_columns = []
        prev_fixed = False
        for index, v in enumerate(df.columns.values):
            if prev_fixed:
                prev_fixed = False
                continue
            new_value = df.columns.values[index].split()
            new_value = " ".join(new_value).strip()
            if new_value.startswith("COD QTA"):
                new_value = new_value.split()
                if len(new_value[1]) > 3:
                    new_value[0] = new_value[0] + new_value[1][3:]
                fixed_columns.append(new_value[0])
                fixed_columns.append(new_value[1])
                prev_fixed = True
            else:
                fixed_columns.append(new_value)
        df.columns = fixed_columns

        to_remove = []
        for index, col in enumerate(df.columns.values):
            if col.startswith("Unnamed"):
                # if col not in fixed_columns:
                to_remove.append(index)
        return df.drop(df.columns[to_remove], axis=1)

    def __get_badge_name(self, sheet_obj):
        """get owner's name out of sheet"""
        engine = self.__get_engine()
        try:
            if engine == "openpyxl":
                badge_name = (sheet_obj[self.name_cell]).value
            elif engine == "xlrd":
                badge_name = sheet_obj.cell_value(int(self.name_cell[1:]) - 1, int(openpyxl.utils.cell.column_index_from_string(self.name_cell[0])) - 1)
            else:
                raise
            badge_name = " ".join(badge_name.split())
            return badge_name
        except Exception as e:
            raise Exception(f"Cannot get_badge_name. Error: {e}")

    def __minutes_to_int(self, decimal_time):
        """ decimal_time => (str) MM"""
        # if 0 not present attach it
        if len(decimal_time) == 1:
            decimal_time = decimal_time + "0"

        hour_min = 60
        result = (100 * int(decimal_time)) / hour_min
        return str(int(result))

    def __round_float(self, float_number, decimal_pos=2):
        try:
            float(float_number)
        except:
            raise TypeError("Cannot round: not a float number")
        rounded = str(float_number).split(".")
        rounded = float(rounded[0] + "." + rounded[1][:decimal_pos])
        return rounded

    def __smart_renamer(self, name):
        old_name = name.split(" ")
        new_name = ""

        for index, word in enumerate(old_name):
            if index < len(old_name):
                new_name += (old_name[index][0].upper() + old_name[index][1:].lower() + " ")
            else:
                new_name += (old_name[index][0].upper() + old_name[index][1:].lower())

        return new_name

    """    PUBLIC METHODS    """
    def set_badges_path(self, badges_path):
        if not os.path.exists(badges_path):
            raise ValueError("ERROR: Cannot find badges path")
        self.badges_path = badges_path
        print("** badges_path caricato")

    def set_billing_time(self, month, year):
        self.billing_year = int(year)
        self.billing_month = int(month)
        self._holidays = holidays.IT(years=[self.billing_year, self.billing_year - 1])

    def get_all_badges_names(self):
        """ return an array of all names found in excel file """
        xlsx_data, sheet_names, engine = self.__load_Excel_badges()
        names = []
        for sheetNo, sheet in enumerate(sheet_names):
            sheet_data = xlsx_data[sheet] if engine == "openpyxl" else xlsx_data.get_sheet(sheetNo)
            names.append(self.__get_badge_name(sheet_data))
        return names

    def parse_badges(self, names=[]):
        """
        read and fix the badges form, adjusting column names and preparing data to be read by other methods.
        returning a dict containing every worker as key and a subdict containing every of its workday as value
        """
        xlsx_data, sheet_names, engine = self.__load_Excel_badges()
        total_content = {}

        for sheetNo, sheet in enumerate(sheet_names):
            sheet_data = xlsx_data[sheet] if engine == "openpyxl" else xlsx_data.get_sheet(sheetNo)
            badge_name = self.__get_badge_name(sheet_data)

            if names:
                if not badge_name in names:
                    continue

            total_content[badge_name] = {}

            # getting df, fixing columns, removing empty columns
            df = pd.read_excel(xlsx_data, sheet_name=sheet, header=9, index_col=0, engine=engine)

            # set columns
            df = self.__manage_columns(df)

            # parse rows
            for row_ in df.iterrows():
                i = str(row_[0]).strip()
                row = dict(row_[1])

                try:
                    if re.search(self.regex_day_pattern, i):
                        row_dict = {}

                        # pairing values
                        already_parsed = []
                        for val in row:
                            paired = False
                            for key in self.pairing_schema:
                                number = ""
                                if val.split(".")[0] in self.pairing_schema[key]:
                                    in_parsing = list(filter(lambda x: x != "", val.split(".")))
                                    if len(in_parsing) > 1 and in_parsing[1]:
                                        number = f".{in_parsing[1]}"

                                    main_key = key + "." + number if number else key + number
                                    if main_key not in already_parsed:
                                        row_dict[main_key] = []
                                        for v_ in self.pairing_schema[key]:

                                            # find the correct refer_key
                                            for _i in range(3):
                                                refer_key = v_ + "."*_i + number
                                                if refer_key in row.keys():
                                                    break

                                            if not isinstance(row[refer_key], pd.Series):
                                                row_dict[main_key].append(str(row[refer_key]).strip())
                                            else:
                                                check_val = ""
                                                for index in list(row[refer_key].to_list()):
                                                    if not str(index).isspace():
                                                        check_val = str(index).strip()
                                                row_dict[main_key].append(check_val)

                                        already_parsed.append(main_key)
                                    paired = True

                            if not paired:
                                row_dict[val] = row[val]

                        total_content[badge_name][i] = row_dict

                except TypeError:
                    pass

        self.total_content = total_content
        return total_content

    def parse_days(self, total_content):
        """ Pointing out what is the type of the hours worked by the worker that day """

        to_return = {}
        for worker in total_content:
            to_return[worker] = {}
            for day in total_content[worker]:
                day_content = total_content[worker][day]
                parsed_day = {
                    "OR": 0.0,  # ordinarie
                    "ST": 0.0,  # straordinarie
                    "MN": 0.0,  # maggiorazione notturna
                    "OF": 0.0,  # ordinario festivo
                    "SF": 0.0,  # straordinario festivo
                    "SN": 0.0,  # straordinario notturno
                    "FN": 0.0  # festivo notturno
                }

                # se non ci sono ore ordinarie o straordinarie return empty day
                if not any(day_content["GIOR PROG"]) and not any(day_content["GIOR PROG..1"]):
                    to_return[worker][day] = parsed_day
                    continue

                # setting starting ordinary and overtime values
                if day_content["GIOR PROG"][0]:
                    val = day_content["GIOR PROG"][0] if len(day_content["GIOR PROG"][0]) >= 4 else "0" + day_content["GIOR PROG"][0]
                    if "." in val:
                        val = val.split(".")[0] + "." + self.__minutes_to_int(val.split(".")[1])
                    parsed_day["OR"] += float(val)
                if day_content["GIOR PROG..1"][0]:
                    val = day_content["GIOR PROG..1"][0] if len(day_content["GIOR PROG..1"][0]) >= 4 else "0" + day_content["GIOR PROG..1"][0]
                    if "." in val:
                        val = val.split(".")[0] + "." + self.__minutes_to_int(val.split(".")[1])
                    parsed_day["ST"] += float(val)

                # check every COD key for special hours
                for key in day_content:
                    if key.startswith("COD") and any(day_content[key]):

                        # night shifts
                        if day_content[key][0] == "MN":
                            hours = day_content[key][1] if len(day_content[key][1]) >= 4 else "0" + day_content[key][1]
                            if "." in hours:
                                hours = hours.split(".")[0] + "." + self.__minutes_to_int(hours.split(".")[1])
                            hours = float(hours)
                            parsed_day["OR"] -= hours
                            parsed_day["MN"] += hours

                        # overtime night shifts
                        elif day_content[key][0] == "SN":
                            hours = day_content[key][1] if len(day_content[key][1]) >= 4 else "0" + day_content[key][1]
                            if "." in hours:
                                hours = hours.split(".")[0] + "." + self.__minutes_to_int(hours.split(".")[1])
                            hours = float(hours)
                            parsed_day["ST"] -= hours
                            parsed_day["SN"] += hours

                # if day is holiday decrement ordinary to increase holiday values
                try:
                    check_day = f"{self.billing_month}/{day[:-1]}/{self.billing_year}"
                    if check_day in self._holidays:
                        parsed_day["OF"] += parsed_day["OR"]
                        parsed_day["OR"] -= parsed_day["OR"]
                        parsed_day["SF"] += parsed_day["ST"]
                        parsed_day["ST"] -= parsed_day["ST"]
                        parsed_day["FN"] += parsed_day["MN"]
                        parsed_day["MN"] -= parsed_day["MN"]
                except ValueError as e:
                    if str(e).startswith("Cannot parse date from string '2/29/"):
                        raise ValueError(f"ERRORE: {worker} ha lavorato il giorno 29 Febbraio di un anno non bisestile!")

                to_return[worker][day] = parsed_day
        return to_return

    def get_jobname(self, job_id):
        """ given id, gets job name """
        name = ""
        for job in self.jobs:
            if job["id"] == job_id:
                name = job["name"]
                break
        if not name and job_id:
            name = f"Job {job_id} non trovato"
        return name

    def get_billingprofile(self, job_id):
        """ return billing profile id of given job id """
        pass


    def old_parse_day(self,day, day_content):
        """ Pointing out what is the type of the hours worked by the worker that day """
        parsed_day = {
            "OR": 0.0,  # ordinarie
            "ST": 0.0,  # straordinarie
            "MN": 0.0,  # maggiorazione notturna
            "OF": 0.0,  # ordinario festivo
            "SF": 0.0,  # straordinario festivo
            "SN": 0.0,  # straordinario notturno
            "FN": 0.0  # festivo notturno
        }

        # se non ci sono ore ordinarie o straordinarie return empty day
        if not any(day_content["GIOR PROG"]) and not any(day_content["GIOR PROG..1"]):
            return parsed_day

        # setting starting ordinary and overtime values
        if day_content["GIOR PROG"][0]:
            val = day_content["GIOR PROG"][0] if len(day_content["GIOR PROG"][0]) >= 4 else "0" + day_content["GIOR PROG"][0]
            if "." in val:
                val = val.split(".")[0] + "." + self.__minutes_to_int(val.split(".")[1])
            parsed_day["OR"] += float(val)
        if day_content["GIOR PROG..1"][0]:
            val = day_content["GIOR PROG..1"][0] if len(day_content["GIOR PROG..1"][0]) >= 4 else "0" + day_content["GIOR PROG..1"][0]
            if "." in val:
                val = val.split(".")[0] + "." + self.__minutes_to_int(val.split(".")[1])
            parsed_day["ST"] += float(val)

        # check every COD key for special hours
        for key in day_content:
            if key.startswith("COD") and any(day_content[key]):

                # night shifts
                if day_content[key][0] == "MN":
                    hours = day_content[key][1] if len(day_content[key][1]) >= 4 else "0" + day_content[key][1]
                    if "." in hours:
                        hours = hours.split(".")[0] + "." + self.__minutes_to_int(hours.split(".")[1])
                    hours = float(hours)
                    parsed_day["OR"] -= hours
                    parsed_day["MN"] += hours

                # overtime night shifts
                elif day_content[key][0] == "SN":
                    hours = day_content[key][1] if len(day_content[key][1]) >= 4 else "0" + day_content[key][1]
                    if "." in hours:
                        hours = hours.split(".")[0] + "." + self.__minutes_to_int(hours.split(".")[1])
                    hours = float(hours)
                    parsed_day["ST"] -= hours
                    parsed_day["SN"] += hours

        # if day is holiday decrement ordinary to increase holiday values
        check_day = f"{self.billing_month}/{day[:-1]}/{self.billing_year}"
        if check_day in self._holidays:
            parsed_day["OF"] += parsed_day["OR"]
            parsed_day["OR"] -= parsed_day["OR"]
            parsed_day["SF"] += parsed_day["ST"]
            parsed_day["ST"] -= parsed_day["ST"]
            parsed_day["FN"] += parsed_day["MN"]
            parsed_day["MN"] -= parsed_day["MN"]

        return parsed_day

    def apply_billing_profile(self, hours_to_bill, billing_profile):
        """
        steps: 1. adding time, 2. apply pattern, 3. apply pricing
        """
        priced_hours = {}

        # check integrity and get tag to focus
        if hours_to_bill["OR"] and hours_to_bill["OF"]:
            raise ValueError("ERROR: there are both ordinary and holiday hours on a single day")
        else:
            if hours_to_bill["OF"]:
                tag = "OF"
            elif hours_to_bill["OR"]:
                tag = "OR"
            else:
                tag = None

        # adding time
        if tag:
            if billing_profile["time_to_add"] and hours_to_bill[tag]:
                if billing_profile["add_over_threshold"]:
                    if hours_to_bill[tag] >= billing_profile["threshold_hour"]:
                        hours_to_bill[tag] += billing_profile["time_to_add"]
                else:
                    hours_to_bill[tag] += billing_profile["time_to_add"]

            # apply pattern
            if billing_profile["pattern"] and hours_to_bill[tag]:
                new_amount = 0.0
                start_val = copy.deepcopy(hours_to_bill[tag])
                for i in range(len(billing_profile["pattern"])):
                    operation = billing_profile["pattern"][i]["perform"].strip()
                    amount = billing_profile["pattern"][i]["amount"]
                    if operation == "/":
                        start_val /= amount
                    elif operation =="-":
                        start_val -= amount
                    elif operation == "+":
                        start_val += amount
                    elif operation =="*":
                        start_val *= amount
                    else:
                        raise Exception("ERROR: invalid operator in pattern. operator must be one of + - * /")
                    if billing_profile["pattern"][i]["keep"]:
                        new_amount += start_val
                hours_to_bill[tag] = new_amount

        # apply pricing
        if billing_profile["pricelist"]:
            for hour_type in hours_to_bill:
                for p in billing_profile["pricelist"]:
                    if hour_type == p["tag"]:
                        priced_hours[hour_type] = hours_to_bill[hour_type]*p["price"]
                        priced_hours[hour_type] = self.__round_float(priced_hours[hour_type], decimal_pos=2)
        else:
            raise Exception("ERROR: No pricelist specified!")

        return priced_hours

    def create_billing_schema(self, total_content, random_=True):

        billing_schema = {}

        # assign random billing profile to workers
        if random_:
            for worker in total_content:
                profile = random.choice(list(self.billing_profiles.keys()))
                billing_schema[worker] = {k:profile for k in total_content[worker]}
        else:
            raise Exception("Not implemented yet")

        self.billing_schema = billing_schema
        return billing_schema

    def parse_total(self, data):
        """ return a tuple containing ({worker:total}, {total:total})  """
        total = {}
        new_data = {}

        for worker in data:
            new_data[worker] = {}
            for day in data[worker]:
                for hour_type in data[worker][day]:

                    # add to worker data
                    if hour_type in new_data[worker]:
                        new_data[worker][hour_type] += data[worker][day][hour_type]
                    else:
                        new_data[worker][hour_type] = data[worker][day][hour_type]

                    # add to total
                    if hour_type in total:
                        total[hour_type] += data[worker][day][hour_type]
                    else:
                        total[hour_type] = data[worker][day][hour_type]

            # round values
            for h in new_data[worker]:
                new_data[worker][h] = self.__round_float(new_data[worker][h], decimal_pos=2)

        # round values
        for h in total:
            total[h] = self.__round_float(total[h], decimal_pos=2)

        return (new_data, total)

    def bill(self, total_content, billing_schema, dump_values=False, dump_detailed=False):

        hours_data = {}
        billing_data = {}

        for worker in total_content:
            hours_data[worker] = {}
            billing_data[worker] = {}
            for day_ in total_content[worker]:

                # parsing hours of the day
                hours_to_bill = self.parse_day(day_, total_content[worker][day_])
                hours_data[worker][day_] = hours_to_bill

                # apply billing profile for the day
                billing_profile = self.billing_profiles[billing_schema[worker][day_]]
                billed_hours = self.apply_billing_profile(hours_to_bill, billing_profile)
                billing_data[worker][day_] = billed_hours

        new_billing_data, total_billing = self.parse_total(billing_data)
        new_hours_data, total_hours = self.parse_total(hours_data)

        # conditional dump values
        if dump_detailed:
            with open("DETAIL_ore_lavoratori.json", "w") as f:
                f.write(json.dumps(hours_data, indent=4, ensure_ascii=True))
            with open("DETAIL_valori_da_fatturare.json", "w") as f:
                f.write(json.dumps(billing_data, indent=4, ensure_ascii=True))

        if dump_values:
            with open("ore_lavoratori.json", "w") as f:
                f.write(json.dumps(new_hours_data, indent=4, ensure_ascii=True))
            with open("valori_da_fatturare.json", "w") as f:
                f.write(json.dumps(new_billing_data, indent=4, ensure_ascii=True))

        self.create_Excel(new_hours_data, total_billing, billing_schema)
        print(">> BillingManager Done")

    def create_Excel(self, content, total_billing, billing_schema, transposed=True):
        df = pd.DataFrame.from_dict(content)

        if transposed:
            df = df.T

        # sort alphabetically rows and columns
        df = df.sort_index()
        df.rename(index=lambda x: self.__smart_renamer(x), inplace=True)

        # add totals (rows total/column total)
        df.loc[">> ORE TOTALI <<"] = df.sum(axis=0)
        df.loc[">> € DA FATTURARE <<"] = total_billing
        df['TOTALI'] = df.sum(axis=1)

        # polish
        df.index.rename("LAVORATORI", inplace=True)
        df.replace(0, np.nan, inplace=True)

        #generating excel with df data
        with pd.ExcelWriter(self.bill_name, mode="w") as writer:
            sh_name = "Report Fatturazione"
            df.to_excel(writer, sheet_name=sh_name, na_rep="", float_format="%.2f")
            ws = writer.sheets[sh_name]
            footer_color = PatternFill(start_color="e6e6e6", end_color="e6e6e6", fill_type="solid")

            # style total hours
            for cell in ws[f"{len(df)}:{len(df)}"]:
                cell.font = openpyxl.styles.Font(bold=True)
                cell.fill = footer_color

            # style billing totals
            for cell in ws[f"{len(df)+1}:{len(df)+1}"]:
                cell.number_format = '#,##0.00€'
                cell.font = openpyxl.styles.Font(bold=True)
                cell.fill = footer_color

        print(f">> {self.bill_name} REDATTO CON SUCCESSO")

"""
path = "../test_data/cartellini.xlsx"
Fatturatore = BillingManager()

Fatturatore.set_badges_path(path)
all_badges = Fatturatore.parse_badges()
billing_schema = Fatturatore.create_billing_schema(all_badges, random_=True)

Fatturatore.bill(all_badges, billing_schema)
"""