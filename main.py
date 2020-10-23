import os
import sys
from PyPDF2 import PdfFileWriter, PdfFileReader  # (per le buste)
import fitz # PyMuPDF (per i cartellini)
import time

def boot():
    print_instruction()

    valid = False
    print("Vuoi dividere le buste e i cartellini?\n"
          "- Digita s (Si) e premi invio per avviare il programma\n"
          "- Digita n (No) e premi invio per terminare il programma\n")
    while not valid:
        user_input = input("[s/n] --> ")

        if user_input.lower() == "s" or user_input.lower() == "si":
            print("\n" * 5)
            print(("V" * 50) + " PROGRAMMA AVVIATO " + ("V" * 50))
            print("\n")
            valid = True
        elif user_input.lower() == "n" or user_input.lower() == "no":
            print("Chiusura programma . . .")
            time.sleep(1)
            sys.exit(0)
        else:
            print(f"Input errato: hai digitato '{user_input}'")

def print_instruction():

    def split_rows(stringa, row_length):
        stringa = stringa.split(" ")
        new_stringa = ""

        for word in stringa:
            actual_string = new_stringa.split("\n")[-1]
            if (len(actual_string) + len(word + " ")) > row_lenght:
                new_stringa += (word + "\n")
            else:
                new_stringa += (word + " ")

        return new_stringa

    row_lenght = 108
    char = "*"
    raw_instructions = "Grazie per utilizzare BusinessCat, questo è un programma Open Source che divide i tuoi cartellini e le tue buste paga per singolo lavoratore!" \
                       "\n\n" \
                       "COME FUNZIONA?\n" \
                       "- BUSTE: per funzionare il programma ha bisogno di un file rinominato 'LULDIP.pdf' nella sua stessa directory. " \
                       "Il programma dividerà le buste per lavoratore rispettando la lunghezza delle buste lunghe due facciate.\n\n" \
                       "- CARTELLINI: per funzionare il programma ha bisogno di un file rinominato 'cartellini.pdf' nella sua stessa directory. " \
                       "Il programma dividerà i cartellini per lavoratore. I cartellini a differenza delle buste vengono divisi pagina per pagina quindi non sono ammessi" \
                       " cartellini lunghi più di una facciata."

    raw_credits = "CREDITS: Fabio Magrotti @ CSI - Centro Servizi Industriali - Broni 27043 (PV)"

    print("\n" + char * row_lenght)
    print(char * row_lenght, "\n")

    print(split_rows(raw_instructions, row_lenght) + "\n\n" + (split_rows(raw_credits, row_lenght)).upper())

    print("\n" + char * row_lenght)
    print(char * row_lenght, "\n")

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

def read_page(pageObj, lookup_interval=[]):
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

            #debug
            #print(obj[37])

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


            #print(full_name)

            # salvo il cartellino
            badge = fitz.Document()
            badge.insertPDF(inputpdf, from_page=i, to_page=i)
            badge.save(dirname + "/" + full_name + ".pdf")


    except:
        raise Exception("#######################################\n"
                        "ERRORE nella generazione dei cartellini\n"
                        "#######################################\n")


if __name__ == "__main__":

        boot()

        # PARAMETRI BUSTE
        file_buste = "LULDIP.pdf"
        directory_buste = "BUSTE PAGA"

        # PARAMETRI CARTELLINI
        file_cartellini = "cartellini.pdf"
        directory_cartellini = "CARTELLINI"

        try:

            # create paycheck and badges
            CREATE_BUSTE(file_buste, directory_buste)
            CREATE_CARTELLINI(file_cartellini, directory_cartellini)

            print("\n##############################################\n"
                  "FATTO: buste e cartellini divisi!\n\n"
                  
                  "Premi un tasto per chiudere il programma . . .\n"
                  "##############################################\n")

            input()
            print("Grazie per aver usato BusinessCat!")
            time.sleep(1)
            sys.exit(0)


        except Exception as e:
            print(e)
