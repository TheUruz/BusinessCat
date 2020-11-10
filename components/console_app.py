import sys
import time
from lib.appLib import CREATE_CARTELLINI, CREATE_BUSTE

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
