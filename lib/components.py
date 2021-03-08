from lib import appLib
import PIL.Image, PIL.ImageTk
import os
import copy
import uuid
import json
import shutil
import fitz
import pandas as pd
from pandastable import Table, TableModel, config as tab_config
from email.message import EmailMessage
import mimetypes
import smtplib
from tkinter import *
from ttkwidgets.frames import Balloon
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email import encoders

import mysql.connector

class Custom_Toplevel(Toplevel):
    def __init__(self, master=None, width=100, height=100):
        super().__init__(master)

        self.width = width
        self.height = height
        self.__screen_width = master.winfo_screenwidth()
        self.__screen_height = master.winfo_screenheight()
        self.geometry(f"{self.width}x{self.height}+{int((self.__screen_width / 2 - self.width / 2))}+{int((self.__screen_height / 2 - self.height / 2))}")
        self.title("BusinessCat")
        self.iconbitmap(appLib.icon_path)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", lambda: self._close(master))
        self.prior_window = None

    def open_new_window(self, master, new_window):
        """ use this method to pass from one window to another"""
        self.destroy()
        new = new_window(master)
        new.prior_window = type(self)

    def apply_balloon_to_widget(self, widget, text):
        """ use this method to apply a balloon to a given widget"""
        self.widget_balloon = Balloon(master=widget, text=text, timeout=0.5, height=100)

    def _close(self, master):
        """ defining closing window behavior """
        if messagebox.askokcancel("Esci", "Vuoi davvero uscire da BusinessCat?"):
            master.destroy()

class Login_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 340
        self.height = 600
        super().__init__(master, self.width, self.height)

        self.title(self.title().split("-")[:1][0] + " - " + "Login")

        # define mainframe
        self.mainframe = Frame(self, width=self.width, height=self.height, bg=appLib.default_background) # change color here to fill spaces
        self.mainframe.pack(expand=True, fill="both", anchor="center")


        # define logo and appname
        self.image = PIL.Image.open(appLib.logo_path)
        self.image = PIL.ImageTk.PhotoImage(self.image)
        self.label = Label(self.mainframe, image=self.image, bg=appLib.default_background)
        self.label.pack(fill="both", anchor="center", pady=(50,0))

        self.name = Label(self.mainframe, text="BusinessCat", font=("Calibri", 22, "bold"), fg=appLib.color_orange, bg=appLib.default_background)
        self.name.pack(fill="both", anchor="center")


        # define email label and textbox
        self.email_label = Label(self.mainframe, text="Email:", font=("Calibri", 11, "bold"), bg=appLib.default_background)
        self.email_textbox = Entry(self.mainframe, width=44, bg=appLib.color_light_orange, justify="center")
        self.email_label.pack(anchor="center", pady=(50,0))
        self.email_textbox.pack(anchor="center")


        # define password label and textbox
        self.password_label = Label(self.mainframe, text="Password:", font=("Calibri", 11, "bold"), bg=appLib.default_background)
        self.password_textbox = Entry(self.mainframe, width=44, bg=appLib.color_light_orange, show="*", justify="center")
        self.password_label.pack(anchor="center", pady=(45,0))
        self.password_textbox.pack(anchor="center")


        # define login button
        self.login_button = Button(self.mainframe, text="LOGIN", font=("Calibri", 10), width=12, bg=appLib.default_background, command= lambda: self.check_login(master))
        self.login_button.pack(anchor="center", pady=(50,0))

        # define signup_label
        self.signup_label = Label(self.mainframe, text="Registrati", font=("Calibri", 10), width=12, bg=appLib.default_background, fg=appLib.color_orange)
        self.signup_label.pack(anchor="center", pady=(90,0))
        self.signup_label.bind('<Button 1>', lambda event: self.open_new_window(master, Register_Window))

    def check_login(self, master):
            email = str(self.email_textbox.get())
            password = str(self.password_textbox.get())

            if master.using_db == True:
                if email and password:
                    try:
                        config = appLib.load_config()

                        auth_db = mysql.connector.connect(
                            host=config['host'],
                            password=config['password'],
                            username=config['username'],
                            database=config['database'],
                            port=config['port']
                        )

                        cursor = auth_db.cursor()
                        cursor.execute(f"""SELECT * FROM users WHERE email = '{email}' AND pwd = '{password}'""")
                        entry = cursor.fetchone()

                        # se l'utente è presente procedo altrimento errore
                        if entry:

                            valid_postation = False
                            # verifico che il login stia avvenendo da una postazione nota
                            for workstation in json.loads(entry[4]):
                                if workstation == uuid.getnode():
                                    valid_postation = True

                            # se la postazione di accesso è valida entro nel programma
                            if valid_postation:
                                self.open_new_window(master, Home_Window)
                                print(email, password)
                            else:
                                messagebox.showwarning("Access Denied", "Hai già effettuato l'accesso da tre dispositivi diversi. Contatta l'assistenza")

                        else:
                            raise
                    except Exception as e:
                        print(e)
                        messagebox.showerror("Non registrato!", "Sembra che tu non sia registrato!")
                else:
                    messagebox.showerror("Dati mancanti", "Mancano i dati per il login!")
            else:
                self.open_new_window(master, Home_Window)

class Register_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.height = 500
        self.width = 500
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Signup")
        self.margin = 15

        ####################### GRID CONFIGURE
        self.grid_columnconfigure(0, minsize=self.margin, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_columnconfigure(2, minsize=self.margin, weight=1)

        self.grid_rowconfigure(2, minsize=self.margin*2)
        self.grid_rowconfigure(5, minsize=self.margin *2)
        self.grid_rowconfigure(8, minsize=self.margin * 2)
        self.grid_rowconfigure(11, minsize=self.margin*2.5)
        self.grid_rowconfigure(13, minsize=self.margin * 2)



        # define info_labels
        info_text = "BusinessCat funzionerà solo sulla postazione dalla quale ti stai registrando, scegli bene la tua postazione prima di registrati!"
        self.info_label1 = Label(self, text="ATTENZIONE", font=("Calibri", 18), fg=appLib.color_red, bg=appLib.default_background, wraplength=self.width-(self.margin*2))
        self.info_label2 = Label(self, text=info_text, font=("Calibri", 14), fg=appLib.color_red, bg=appLib.default_background, wraplength=self.width-(self.margin*2))
        self.info_label1.grid(row=0, column=1)
        self.info_label2.grid(row=1, column=1)


        # define email_label
        self.email_label = Label(self, text="Inserisci qui la Email", font=("Calibri", 10, "bold"), bg=appLib.default_background)
        self.email_label.grid(row=3, column=1)

        # define email_txtbox
        self.email_txtbox = Entry(self, width=60, bg=appLib.color_light_orange, justify="center")
        self.email_txtbox.grid(row=4, column=1)



        # define password_label
        self.password1_label = Label(self, text="Inserisci qui la Password", font=("Calibri", 10, "bold"), bg=appLib.default_background)
        self.password1_label.grid(row=6, column=1)

        # define password_txtbox
        self.password1_txtbox = Entry(self, width=60, bg=appLib.color_light_orange, show="*", justify="center")
        self.password1_txtbox.grid(row=7, column=1)



        # define repeat_password_label
        self.password2_label = Label(self, text="Ripeti la Password", font=("Calibri", 10, "bold"), bg=appLib.default_background)
        self.password2_label.grid(row=9, column=1)

        # define repeat_password_txtbox
        self.password2_txtbox = Entry(self, width=60, bg=appLib.color_light_orange, show="*", justify="center")
        self.password2_txtbox.grid(row=10, column=1)



        # define sing up button
        self.signup_button = Button(self, width=15, text="REGISTRATI", bg=appLib.default_background, command= lambda: self.signup(master))
        self.signup_button.grid(row=12, column=1)

        # define "back" button frame
        frame_width = 120
        frame_height = 55
        self.back_button_frame = Frame(self, width=frame_width, height=frame_height, bg=appLib.default_background)
        self.back_button_frame.grid(row=14, column=0, columnspan=2, sticky="w")

        # define "back" label
        self.back_paw_label = Label(self.back_button_frame, text="< Torna al Login", font=("Calibri", 10, "bold"), fg=appLib.color_orange, bg=self.back_button_frame['bg'])
        self.back_paw_label.grid(row=0)
        self.back_paw_label.bind('<Button 1>', lambda event: self.back_to_login(master))

        # define "back" paw image
        self.paw_image = Image.open("../config_files/imgs/paws/paw_3.png")
        self.paw_image = self.paw_image.resize((frame_height,frame_width)).rotate(270, expand=True)
        self.paw_image = PIL.ImageTk.PhotoImage(self.paw_image)
        self.back_paw_image = Label(self.back_button_frame, image=self.paw_image, bg=appLib.default_background, borderwidth=0)
        self.back_paw_image.grid(row=1)
        self.bind('<Button 1>', lambda event: self.open_new_window(master, Login_Window))

    def signup(self, master):
        email = self.email_txtbox.get()
        password1 = self.password1_txtbox.get()
        password2 = self.password2_txtbox.get()

        if password1 and password2 and email:

            # check matching password
            if password1 != password2:
                messagebox.showerror("Errore password", "Le password immesse non coincidono!")
                self.clear_fields()
                return
            else:

                try:
                    config = appLib.load_db_config()
                    db = appLib.connect_to_db(config)
                    cursor = db.cursor()

                    already_registered = appLib.check_registered(cursor, email)

                    # se la mail non è già registrata
                    if not already_registered:
                        user_key = uuid.uuid4()
                        mac_address = uuid.getnode()
                        appLib.add_user(db, cursor, email, password1, user_key, [mac_address])
                        messagebox.showinfo("Registrato!", "La registrazione è andata a buon fine, sarai reindirizzato alla finestra per il login!")
                        self.clear_fields(clear_email=True)

                        #ritorna al login
                        self.open_new_window(master, Login_Window)

                    else:
                        messagebox.showwarning("Già registrato!", "La mail risulta già utilizzata!")

                except:
                    raise Exception("Errore registrazione non riuscita")

        else:
            messagebox.showerror("Dati mancanti", "tutti i campi sono obbligatori")

    def clear_fields(self, clear_email=False):
        self.password1_txtbox.delete(0, 'end')
        self.password2_txtbox.delete(0, 'end')

        if clear_email:
            self.email_txtbox.delete(0, 'end')

class Home_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 600
        self.height = 500
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Home")
        self.margin = 15

        self.buttons_width = 20
        self.buttons_height = 2
        self.buttons_y_padding = 20

        #self.menu_background_color = "#f2f2f2"
        #self.menu_buttons_color = appLib.default_background

        self.menu_background_color = appLib.color_light_orange
        self.menu_buttons_color = "#cddbe5"

        #### define left menu
        self.menu_frame = Frame(self, width=200, height=self.height, highlightbackground="black", highlightthickness=0.5, bg=self.menu_background_color)
        self.menu_frame.grid(row=0, column=0, rowspan=2, padx=(1,1), pady=(0,1))
        self.menu_frame.pack_propagate(False)

        # PDF splitter button
        self.button1 = Button(self.menu_frame, text="PDF Splitter", font=("Calibri", 10, "bold"), width=self.buttons_width, height=self.buttons_height, fg=appLib.color_orange, bg=self.menu_buttons_color, command=lambda:self.open_new_window(master, Splitter_Window))
        self.button1.pack(anchor="center", pady=self.buttons_y_padding)

        # Mail sender button
        self.button2 = Button(self.menu_frame, text="Mail Sender", font=("Calibri", 10, "bold"), width=self.buttons_width, height=self.buttons_height, fg=appLib.color_orange, bg=self.menu_buttons_color, command=lambda:self.open_new_window(master, Mail_Sender_Window))
        self.button2.pack(anchor="center", pady=self.buttons_y_padding)

        # Verify Paychecks button
        self.button3 = Button(self.menu_frame, text="Verify Paychecks", font=("Calibri", 10, "bold"), width=self.buttons_width, height=self.buttons_height, fg=appLib.color_orange, bg=self.menu_buttons_color, command=lambda:self.open_new_window(master, Verificator_Window))
        self.button3.pack(anchor="center", pady=self.buttons_y_padding)

        # Billing Manager button
        self.button4 = Button(self.menu_frame, text="Billing Manager", font=("Calibri", 10, "bold"), width=self.buttons_width, height=self.buttons_height, fg=appLib.color_orange, bg=self.menu_buttons_color, command=lambda:self.open_new_window(master, Billing_Landing_Window))
        self.button4.pack(anchor="center", pady=self.buttons_y_padding)

        # "Back to login" Label (packed only if app is using db)
        if master.using_db:
            self.back_label = Label(self.menu_frame, text="<< Torna al Login", font=("Calibri", 11, "bold"), fg=appLib.color_orange, bg=self.menu_background_color)
            self.back_label.pack(expand=True, anchor="sw", pady=self.buttons_y_padding, padx=(20,0))
            self.back_label.bind('<Button 1>', lambda event: self.open_new_window(master, Login_Window))

        #### Splashart frame
        self.splash_frame = Frame(self, width=396, bg=appLib.color_light_orange)
        self.splash_frame.grid(row=0, column=1, sticky="ns", padx=(1,1))
        self.grid_rowconfigure(0, weight=1)
        self.test_label = Label(self.splash_frame)

        try:
            self.splashart_img = PIL.Image.open("../config_files/imgs/home_splash.png")
            self.splashart = PIL.ImageTk.PhotoImage(self.splashart_img)
            self.splashart_label = Label(self.splash_frame, image=self.splashart)
            self.splashart_label.pack()
        except:
            pass

        #### credits frame
        self.credits_frame = Frame(self, width=396, height=20, bg=appLib.default_background)
        self.credits_frame.grid(row=1, column=1, sticky="s", padx=(1,1), pady=(0,1))
        self.credits_frame.pack_propagate(False)
        self.credits_label = Label(self.credits_frame, text="Developed by CSI - Centro Servizi Industriali\tBRONI (PV) IT", font=("Calibri", 8, "bold"), bg=appLib.default_background)
        self.credits_label.pack(fill="y", anchor="w")

class Splitter_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 520
        self.height = 480
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "PDF Splitter")
        self.margin = 15
        self.paycheck_var = IntVar(self, name="paycheck_var")
        self.PAYCHECKS_PATH = "BUSTE PAGA"
        self.BADGES_PATH = "CARTELLINI"

        # verify paycheck and badges status
        self.done_paycheck, self.done_badges = appLib.check_paycheck_badges()

        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, Home_Window))
        self.back_button.pack(anchor="w", pady=(self.margin,0), padx=(self.margin,0))

        # define BusinessCat logo
        self.logo_frame = Frame(self,bg=appLib.default_background)
        self.logo_frame.pack(fill="x", pady=(0,self.margin*2))
        img = PIL.Image.open("../config_files/imgs/BusinessCat.png")
        self.logo = img.resize((110,120))
        self.Businesscat_logo = PIL.ImageTk.PhotoImage(self.logo)
        self.businesscat_img_label = Label(self.logo_frame, image=self.Businesscat_logo, bg=appLib.default_background)
        self.businesscat_img_label.pack()

        # define Paychecks data
        self.PAYCHECK_FRAME = Frame(self, bg=appLib.default_background)
        self.PAYCHECK_FRAME.pack(fill="x", pady=(0,5))
        self.paycheck_button = Button(self.PAYCHECK_FRAME, text="Buste Paga", width=16, height=1, font=("Calibri", 11), command=lambda:self.changeContent(self.paycheck_textbox))
        self.paycheck_button.grid(row=0,column=0, padx=(self.margin*2,self.margin*2), pady=(0,0), sticky="w")
        self.paycheck_textbox = Entry(self.PAYCHECK_FRAME, width=50, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.paycheck_textbox.grid(row=0, column=1, sticky="ew")
        self.paycheck_checkbox = Checkbutton(self.PAYCHECK_FRAME, text="Le Buste Paga contengono il Cartellino", font=("Calibri", 10, "italic"), bg=appLib.default_background, variable=self.paycheck_var, command=lambda:self.toggle_paycheck_checkbox(self.paycheck_var.get()))
        self.paycheck_checkbox.grid(row=1,column=1, padx=(self.margin*2,self.margin*2), pady=(0,0), sticky="w")
        self.paycheck_checkbox.setvar(name="paycheck_var", value=1) # set default

        # define Badges data
        self.BADGES_FRAME = Frame(self, bg=appLib.default_background)
        self.BADGES_FRAME.pack(fill="x", pady=(self.margin, self.margin))
        self.badges_button = Button(self.BADGES_FRAME, text="Cartellini", width=16, height=1, font=("Calibri", 11), command=lambda:self.changeContent(self.badges_textbox))
        self.badges_button.grid(row=0,column=0, padx=(self.margin*2,self.margin*2), pady=(0,5), sticky="w")
        self.badges_textbox = Entry(self.BADGES_FRAME, width=50, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.badges_textbox.grid(row=0, column=1, sticky="ew")

        # define Buttons
        self.BUTTONS_FRAME = Frame(self, bg=appLib.default_background)
        self.BUTTONS_FRAME.pack(fill="x", pady=(self.margin,self.margin))
        self.BUTTONS_FRAME.grid_columnconfigure(0, minsize=12, weight=1)
        #self.BUTTONS_FRAME.grid_columnconfigure(1, minsize=12, weight=1)
        self.BUTTONS_FRAME.grid_columnconfigure(2, minsize=12, weight=1)
        self.button_START = Button(self.BUTTONS_FRAME, text="Dividi", width=12, height=1, command=lambda:appLib.START_in_Thread(self.splitting))
        self.button_START.grid(row=0, column=0)
        #self.button_CLEAR = Button(self.BUTTONS_FRAME, text="Pulisci", width=12, height=1, command=self.clearContent)
        #self.button_CLEAR.grid(row=0, column=1)
        self.button_SEND = Button(self.BUTTONS_FRAME, text="Invia per Email", width=12, height=1, state=DISABLED,command=lambda: self.open_new_window(master, Mail_Sender_Window))
        self.button_SEND.grid(row=0, column=2)
        self.check_send_mail() # check disable for button_SEND

        # define Status_Circle
        self.STATUS_CIRCLE_FRAME = Frame(self, bg=appLib.color_light_orange)
        self.STATUS_CIRCLE_FRAME.pack()
        self.canvas = Canvas(self.STATUS_CIRCLE_FRAME, width=55, height=55, highlightthickness=0, bg=appLib.default_background)
        self.canvas.pack(fill="both", expand=True)
        self.update()
        topleft_coord = (5,5)
        bottomright_coord = (self.canvas.winfo_width()-5, self.canvas.winfo_height()-5)
        self.status_circle = self.canvas.create_oval(topleft_coord, bottomright_coord, outline="", fill=appLib.color_green)
        self.status_label = Label(self, text="Pronto", font=("Calibri",12,"bold"), bg=appLib.default_background, fg=appLib.color_green)
        self.status_label.pack(pady=(0,self.margin))
        self.update()

        # update based on defaults
        self.toggle_paycheck_checkbox(self.paycheck_var.get())
        self.check_send_mail()


    def __BADGES_FROM_PAYCHECKS(self, filename):
        inputpdf = fitz.open(filename)

        if not os.path.exists(self.BADGES_PATH):
            os.mkdir(self.BADGES_PATH)

        full_name = ""
        # per ogni pagina
        for i in range(inputpdf.pageCount):
            page = inputpdf.loadPage(i)
            blocks = page.getTextBlocks()
            blocks.sort(key=lambda block: block[1])

            # per ogni blocco di testo
            for index, x in enumerate(blocks):
                block_data = x[4].split("\n")

                # per ogni parola nel blocco di testo
                for v in block_data:
                    if "cognome" in v.lower():

                        # retriving name
                        try:
                            name = blocks[index + 1][4].split("\n")[1]
                            name = [w[0].upper() + w[1:].lower() for w in name.split(" ")]
                        except IndexError:
                            name = blocks[index + 2][4].split("\n")[1]
                            name = [w[0].upper() + w[1:].lower() for w in name.split(" ")]

                        current_name = ""
                        for i_, word in enumerate(name):
                            current_name += word + " "
                        current_name = current_name[:-1]

                        if "riepilogo" in current_name.lower():
                            break

                        # salva il cartellino
                        if current_name != full_name and current_name != "Datasassunzione" and current_name:
                            badge = fitz.Document()
                            badge.insertPDF(inputpdf, from_page=i, to_page=i)
                            badge.save(f"{self.BADGES_PATH}/" + current_name + ".pdf")
                            full_name = current_name

    def __SPLIT_PAYCHECKS(self, file_to_split):
        inputpdf = fitz.open(file_to_split)
        check_name = ""

        # possible slots where the name is stored in the pdf {slot:array_split_at}
        lookup_range = {
            19: 1,
            21: 1,
            117: 0
        }

        # chech if file_to_split exists
        if not os.path.exists(file_to_split):
            raise Exception(f"File {file_to_split} inesistente!")

        # create dir if it does not exists
        if not os.path.exists(self.PAYCHECKS_PATH):
            os.mkdir(self.PAYCHECKS_PATH)

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

            if not paycheck_owner or paycheck_owner == "Detrazioni":
                print(f"Non ho trovato il proprietario della busta numero {i}, potrebbe essere un cartellino")
                continue
            else:
                paycheck = fitz.Document()

                # se il nome è lo stesso del foglio precedente salvo il pdf come due facciate
                if paycheck_owner == check_name:
                    paycheck.insertPDF(inputpdf, from_page=i - 1, to_page=i)
                # altrimenti salvo solo la pagina corrente
                else:
                    paycheck.insertPDF(inputpdf, from_page=i, to_page=i)

                check_name = paycheck_owner
                paycheck.save(f"{self.PAYCHECKS_PATH}/" + paycheck_owner + ".pdf")

    def __SPLIT_BADGES(self, file_to_split):
        """ this method is used to split old type badges, badges are now mainly extracted throught __BADGES_FROM_PAYCHECKS """
        try:

            inputpdf = fitz.open(file_to_split)

            if not os.path.exists(file_to_split):
                raise Exception(f"File {file_to_split} inesistente!")

            # create dir if it does not exists
            if not os.path.exists(self.BADGES_PATH):
                os.mkdir(self.BADGES_PATH)

            # per ogni pagina
            for i in range(inputpdf.pageCount):
                page = inputpdf.loadPage(i)
                obj = page.getTextBlocks()
                full_name = ""

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

                # salvo il cartellino
                badge = fitz.Document()
                badge.insertPDF(inputpdf, from_page=i, to_page=i)
                badge.save(f"{self.BADGES_PATH}/" + full_name + ".pdf")

        except:
            raise Exception("#######################################\n"
                            "ERRORE nella divisione dei cartellini\n"
                            "#######################################\n")

    def toggle_paycheck_checkbox(self, checkbox_status):

        # if check was unchecked disable all badges elements
        if checkbox_status:
            self.badges_button["state"]=DISABLED
            self.badges_textbox.configure(disabledbackground=appLib.color_grey)
        elif not checkbox_status:
            self.badges_button["state"] = NORMAL
            self.badges_textbox.configure(disabledbackground=appLib.color_light_orange)

    def changeContent(self, txtbox):
        '''
        change the content of a given Entry Object
        '''
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select a File",
                                              filetype=[("PDF files", "*.pdf*")])
        if not filename:
            return

        txtbox.configure(state='normal')
        txtbox.delete(0, "end")
        txtbox.insert(0, filename)
        txtbox.configure(state='disabled')

    def clearContent(self):
        self.paycheck_textbox.configure(state='normal')
        self.badges_textbox.configure(state='normal')
        self.paycheck_textbox.delete(0, "end")
        self.badges_textbox.delete(0, "end")
        self.paycheck_textbox.configure(state='disabled', disabledbackground=self.cget('bg'))
        self.badges_textbox.configure(state='disabled', disabledbackground=self.cget('bg'))

    def gather_data(self):
        to_return = {
            "paychecks_path": self.paycheck_textbox.get(),
            "badges_path": self.badges_textbox.get(),
            "paychecks_checkbox": self.paycheck_var.get()
        }
        return to_return

    def splitting(self):
        splitting_data = self.gather_data()

        # remove directories conditionals
        remove_dir = True
        if os.path.exists(self.PAYCHECKS_PATH) or os.path.exists(self.BADGES_PATH):
            remove_dir = messagebox.askyesno("File già esistenti", f"Una di queste cartelle\n\n{self.PAYCHECKS_PATH}\n{self.BADGES_PATH}\n\nè già presente. Procedere con questa operazione le eliminerà entrambe, vuoi continuare?\n\nSI CONSIGLIA DI SCEGLIERE NO E EFFETTUARNE UNA COPIA SE IN DUBBIO")
        if remove_dir:
            shutil.rmtree(self.PAYCHECKS_PATH, ignore_errors=True)
            shutil.rmtree(self.BADGES_PATH, ignore_errors=True)
            self.done_paycheck, self.done_badges = appLib.check_paycheck_badges()
        else:
            return

        # imposto lo stato dello status_circle
        self.canvas.itemconfig(self.status_circle, fill=appLib.color_yellow)
        self.status_label.config(text="Sto Dividendo", fg=appLib.color_yellow)

        try:

            # CASE 1 (badge in paychecks in the same file)
            if splitting_data["paychecks_checkbox"] == 1:
                if not splitting_data["paychecks_path"]:
                    raise Exception("Scegli il file prima di dividerlo!")
                self.__SPLIT_PAYCHECKS(splitting_data["paychecks_path"])
                self.__BADGES_FROM_PAYCHECKS(splitting_data["paychecks_path"])
                messagebox.showinfo("Done", "File divisi con successo!")

            # CASE 2 (badges and paychecks in two different files)
            if splitting_data["paychecks_checkbox"] == 0:
                if splitting_data["paychecks_path"]:
                    self.__SPLIT_PAYCHECKS(splitting_data["paychecks_path"])
                if splitting_data["badges_path"]:
                    self.__SPLIT_BADGES(splitting_data["badges_path"])

                if splitting_data["paychecks_path"] or splitting_data["badges_path"]:
                    messagebox.showinfo("Done", "File divisi con successo!")
                else:
                    raise Exception("Non è stato specificato nessun file da dividere!")

        except Exception as e:
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_red)
            self.status_label.config(text="ERRORE", fg=appLib.color_red)
            e = str(e).replace("#", "")
            messagebox.showerror("Error", e)
            return

        # reimposto lo stato dello status_circle
        self.canvas.itemconfig(self.status_circle, fill=appLib.color_green)
        self.status_label.config(text="Pronto", fg=appLib.color_green)

        # check mail send button
        self.check_send_mail()

    def check_send_mail(self):
        self.done_paycheck, self.done_badges = appLib.check_paycheck_badges()
        if self.done_badges and self.done_paycheck:
            self.button_SEND.config(state=NORMAL)
            return True
        else:
            return False

class Mail_Sender_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 1200
        self.height = 750
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Mail Sender")
        self.margin = 15
        self.done_paychecks, self.done_badges = self.verify_paycheck_badges()
        #self.email = MIMEMultipart()
        self.email = EmailMessage()


        ############### ROW CONFIGURE
        self.grid_rowconfigure(0, minsize=self.margin) #empty
        self.grid_rowconfigure(2, minsize=self.margin * 3)  # empty
        self.grid_rowconfigure(3, minsize=480)  # table and text frame
        self.grid_rowconfigure(4, minsize=self.margin)  # empty

        ############### COLUMN CONFIGURE
        self.grid_columnconfigure(0, minsize=self.margin) #empty
        self.grid_columnconfigure(1, weight=1) # contact table
        self.grid_columnconfigure(2, minsize=self.margin) #empty
        self.grid_columnconfigure(3, minsize=600, weight=1)  # text frame
        self.grid_columnconfigure(4, minsize=self.margin) #empty


        # define header label
        h = "Digita le informazioni di contatto richieste per ogni contatto oppure importa i contatti da un file Excel"
        self.header = Label(self, text=h, font=("Calibri", 13, "bold"), width=self.width, fg=appLib.color_orange, bg=appLib.default_background)
        self.header.grid(row=1, column=0, columnspan=4, pady=(0,10))

        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, self.prior_window))
        self.back_button.grid(row=1, column=1, sticky="w")

        # define excel import
        self.xls_logo = PIL.Image.open("../config_files/imgs/xlsx_logo.ico")
        self.xls_logo = self.xls_logo.resize((28,28))
        self.xls_logo = PIL.ImageTk.PhotoImage(self.xls_logo)
        self.excel_label = Label(self, image=self.xls_logo, font=("Calibri", 8, "bold"), bg=appLib.default_background)
        self.excel_label.grid(row=2, column=1, sticky=W)
        self.excel_label.bind('<Button 1>', lambda event: self.import_Excel())
        self.apply_balloon_to_widget(self.excel_label, text="Importa da file Excel")

        # define mail title textbox
        self.mail_title_txtbox = Entry(self, bg=appLib.default_background, font=("Calibri", 10))
        self.mail_title_txtbox.insert(0, "Titolo")
        self.mail_title_txtbox.grid(row=2, column=3, sticky="nsew", pady=(10,10))


        # define contact table
        cfg = {
            'align': 'w',
            'cellbackgr': appLib.default_background,
            'cellwidth': 100,
            'colheadercolor': appLib.color_orange,
            'floatprecision': 2,
            'font': 'Calibri',
            'fontsize': 12,
            'fontstyle': '',
            'grid_color': '#ABB1AD',
            'linewidth': 1,
            'rowheight': 22,
            'rowselectedcolor': appLib.color_light_orange,
            'textcolor': 'black'
        }
        self.table_frame = Frame(self, bg=appLib.default_background)
        self.table_frame.grid(row=3, column=1, sticky="nsew")
        self.table = Table(self.table_frame)
        tab_config.apply_options(cfg, self.table)

        self.table.model.df = pd.DataFrame(index=[x for x in range(100)], columns = ['COGNOME', 'NOME', 'EMAIL'])
        self.table.columnwidths["EMAIL"] = 350
        self.table.show()


        # define Entry box
        self.mail_text_frame = Text(self, bg=appLib.default_background, font=("Calibri", 10))
        self.mail_text_frame.insert(1.0, "Inserisci qui il testo della mail . . .")
        self.mail_text_frame.grid(row=3, column=3, sticky="nsew")

        # define file to send
        self.send_label = Label(self, text="Cosa vuoi allegare ai messaggi?", font=("Calibri", 14, "bold"), bg=appLib.default_background)
        self.send_label.grid(row=5, column=1, sticky=W)

        # paycheck checkbox
        self.paycheck_var = IntVar()
        self.paycheck_checkbox = Checkbutton(self, text="Buste paga", font=("Calibri", 12), bg=appLib.default_background, variable=self.paycheck_var)
        self.paycheck_checkbox.grid(row=6, column=1, sticky=W)

        if not self.done_paychecks:
            self.paycheck_checkbox.configure(state=DISABLED)

        # badges checkbox
        self.badges_var = IntVar()
        self.badges_checkbox = Checkbutton(self, text="Cartellini", font=("Calibri", 12), bg=appLib.default_background, variable=self.badges_var)
        self.badges_checkbox.grid(row=7, column=1, sticky=W, pady=(0,10))

        if not self.done_badges:
            self.badges_checkbox.configure(state=DISABLED)

        # other attachment
        self.other_attachment = Button(self, text="Allega un altro file", font=("Calibri", 10, "bold"), width=30, bg=appLib.color_light_orange, command=self.attach_to_mail)
        self.other_attachment.grid(row=8, column=1, sticky=W)

        # send email button
        self.send_email_button = Button(self, text="Invia Email", font=("Calibri", 12, "bold"), bg=appLib.color_light_orange, command=lambda:appLib.START_in_Thread(self.send_mail))
        self.send_email_button.grid(row=5, column=3, sticky="nsew")

        # check attachment
        self.check_attachments_button = Button(self, text=" Visualizza allegati", font=("Calibri", 10, "bold"), bg=appLib.color_light_orange, width=30, command=self.check_attachments)
        self.check_attachments_button.grid(row=8, column=1, sticky="E")

        # status circle
        self.canvas_frame = Frame(self)
        self.canvas_frame.grid(row=6, column=3, rowspan=2, sticky="nsew")
        self.canvas = Canvas(self.canvas_frame, width= self.canvas_frame.winfo_width(), height=self.canvas_frame.winfo_height(), highlightthickness=0, bg=appLib.default_background)
        self.canvas.pack(fill="both" ,expand=True)
        self.canvas_frame.update()
        raggio = 20
        top_left_coord = [self.canvas_frame.winfo_width()/2 - raggio, 15]
        bottom_right_coord = [self.canvas_frame.winfo_width()/2 + raggio, self.canvas_frame.winfo_height() - 15]
        self.status_circle = self.canvas.create_oval(top_left_coord, bottom_right_coord, outline="", fill=appLib.color_green)
        self.circle_label = Label(self, text="Pronto", font=("Calibri", 14, "bold"), fg=appLib.color_green, bg=appLib.default_background)
        self.circle_label.grid(row=8, column=3)


    def verify_paycheck_badges(self):

        done_paychecks = True if os.path.exists("BUSTE PAGA") else False
        done_badges = True if os.path.exists("CARTELLINI") else False

        return (done_paychecks, done_badges)

    def attach_to_mail(self):
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select a File",filetype=[("File", "*.*")])
        with open(filename, "rb") as f:
            file_data = f.read()
            file_type = mimetypes.guess_type(filename)[0].split("/")
            file_name = f.name.split('/')[-1]
            maintype = file_type[0]
            subtype = file_type[1]

            '''
            #other way around. not working though. must fix something here to make it work
            attachment = MIMEApplication(file_data, _subtype=subtype)
            #if maintype == 'text' and subtype == 'plain': encoders.encode_base64(attachment)
            attachment.add_header('Content-Decomposition', 'attachment', filename=file_name)
            self.email.attach(attachment)
            '''

            self.email.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

            messagebox.showinfo("Attachment success", f"File allegato con successo\n\n{f.name}")

    def check_attachments(self):

        def empty_attachments(self, att_window):
            to_empty = []
            # spot widget to remove
            for child in att_window.children:
                to_empty.append(att_window.children[child])

            # removing widgets
            for widget in to_empty:
                widget.destroy()

            for k, v in self.email.get_params()[1:]:
                print(k, v)
                self.email.del_param(k)

            self.email.set_type('text/plain')
            del self.email['Content-Transfer-Encoding']
            del self.email['Content-Disposition']

            att_window.minsize(width=350, height=50)
            Label(att_window, text="Lista allegati vuota", font=("Calibri", 12), bg=appLib.default_background).pack(fill=Y, expand=True)



        # iter throught email attachments
        attachments = [x.get_filename() for x in self.email.iter_attachments()]

        # creating attachment window
        att_window = Toplevel(self, bg=appLib.default_background)
        att_window.title("Attachments")
        att_window.minsize(width=350, height=150)
        att_window.resizable(False, False)
        att_window.focus_set()
        att_window.grab_set()

        # pack attachments names
        if attachments or self.paycheck_var.get() or self.badges_var.get():
            top_text = Label(att_window, text="Questi sono i file che hai allegato:", font=("Calibri", 14), fg=appLib.color_orange, bg= appLib.default_background)
            top_text.pack(expand=1, fill=BOTH)

            if self.paycheck_var.get():
                Label(att_window, text="Busta paga", font=("Calibri", 12), bg=appLib.default_background).pack(fill=Y, expand=True)
            if self.badges_var.get():
                Label(att_window, text="Cartellino", font=("Calibri", 12), bg=appLib.default_background).pack(fill=Y, expand=True)

            for a in attachments:
                text = Label(att_window, text=a, font=("Calibri", 12), bg=appLib.default_background)
                text.pack(pady=(1, 1))

            # pack clear list button
            clear_list_btn = Button(att_window, text="Rimuovi gli allegati (work in progress)", font=("Calibri", 9), bg=appLib.default_background, command= lambda: empty_attachments(self, att_window), state=DISABLED)
            clear_list_btn.pack(expand=True)

        else:
            att_window.minsize(width=350, height=50)
            Label(att_window, text="Lista allegati vuota", font=("Calibri", 12), bg=appLib.default_background).pack(fill=Y, expand=True)



        return attachments

    def import_Excel(self, filename=None):
        """
        excel must have at least those four columns COGNOME NOME EMAIL ASSUNTO
        """
        if filename == None:
            filename = filedialog.askopenfilename(parent=self.master, initialdir=os.getcwd(), filetypes=[("xlsx","*.xlsx"),("xls","*.xls")])
        if not filename:
            return

        df = pd.read_excel(filename).fillna(0)

        # column names we want to export
        keep_column = [
            "COGNOME",
            "NOME",
            "EMAIL",
        ]

        # check if columns are in dataframe
        for column in keep_column:
            if not column in df.columns.values:
                messagebox.showerror("Bad Data", f"Excel file is not providing requested columns >> {keep_column}")
                return

        # set uppercase columns names
        df = df.rename(columns={key: key.upper() for key in df.columns.values})

        # drop not hired workers
        row_to_drop = []
        for row in df.iterrows():
            if row[1].loc["ASSUNTO"] == "NO" or row[1].loc["ASSUNTO"] == 0:
                row_to_drop.append(row[0])
        df = df.drop(labels=row_to_drop, axis=0)

        # drop not wanted columns
        col_to_drop = []
        for row in df.T.iterrows():
            if row[0].upper() not in keep_column:
                col_to_drop.append(row[0])
        df = df.drop(labels=col_to_drop, axis=1)

        # update model
        model = TableModel(dataframe=df)#to_return)
        self.table.updateModel(model)
        self.table.redraw()
        return

    def send_mail(self):
        connection = appLib.connect_to_mail_server()

        try:
            # setting status circle and label
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_yellow)
            self.circle_label.config(text="Sto Inviando...", fg=appLib.color_yellow)

            # parsing attachments path
            ATTACHMENTS_PATH = []
            if self.paycheck_var.get():
                ATTACHMENTS_PATH.append("BUSTE PAGA")
            if self.badges_var.get():
                ATTACHMENTS_PATH.append("CARTELLINI")

            # get mail title
            mail_title = self.mail_title_txtbox.get()
            if mail_title.lower().strip() == "titolo":
                mail_title = None

            # get mail text
            mail_text = self.mail_text_frame.get("1.0",END)
            if mail_text.lower().strip() == "inserisci qui il testo della mail . . .":
                mail_text = ""

            # setting email
            self.email['Subject'] = mail_title
            self.email['From'] = (appLib.load_email_server_config())['email']
            self.email.add_attachment(MIMEText(mail_text, 'plain')) # this is the mail body

            # SEND MAIL TO CONTACTS

            # reset dump file
            if os.path.exists("../LAST_SENT.txt"):
                with open("../LAST_SENT.txt", "w"): pass

            df = self.table.model.df
            for row in df.iterrows():

                #check row validity
                valid = True
                for val in row[1]:
                    if pd.isnull(val):
                        valid = False

                if valid:

                    # create a copy of the email object
                    msg_obj = copy.deepcopy(self.email)

                    msg_obj['To'] = row[1]['EMAIL'].strip()

                    name = row[1]['NOME'].strip()
                    surname = row[1]['COGNOME'].strip()
                    full_name = f"{surname} {name}".lower()
                    for path in ATTACHMENTS_PATH:
                        if os.path.exists(path):
                            rel_path = os.path.relpath(path)
                            for f in os.listdir(rel_path):

                                # check if file's owner is the one in iterating row
                                if (f.split('.')[0]).lower() == full_name:

                                    with open(rel_path + "\\" + f, 'rb') as f_d:
                                        file_data = f_d.read()
                                        file_name = path + ' ' + f_d.name
                                        file_type = mimetypes.guess_type(f)[0].split("/")
                                        maintype = file_type[0]
                                        subtype = file_type[1]
                                        msg_obj.add_attachment(file_data, maintype=maintype, subtype=subtype,filename=file_name)
                        else:
                            print(f"WARNING: cannot find attachment path: {path}")

                    # send email to row contact
                    connection.send_message(msg_obj)
                    print(f"SENT -> {msg_obj['To']}")

                    # dump past sent
                    with open("../LAST_SENT.txt", "a+") as contact_dump:
                        contact_dump.write(f"{msg_obj['To']}\n")

            # end connection with server
            connection.quit()
            print(f"\n--> END CONNECTION {(appLib.load_email_server_config())['server']}")


            # set status circle and label back to ready
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_green)
            self.circle_label.config(text="Pronto", fg=appLib.color_green)
            messagebox.showinfo("Success", "Email inviate ai contatti in lista")

        except Exception as e:
            messagebox.showerror("Errore", f"Errore riscontrato: {e}")
            print(e)
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_red)
            self.circle_label.config(text="Errore", fg=appLib.color_red)

class Verificator_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 600
        self.height = 700
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Verify Paychecks")
        self.margin = 15

        self.radio_buttons_val = IntVar()
        self.choosen_drivedoc_val = StringVar()
        self.choose_worksheet_val = StringVar()

        self.downloaded_df = None
        self.gdrive_filelist = []

        # controller class
        self.Controller = appLib.PaycheckController()

        # Radio Buttons values
        self.radio_values = {
            1: "CARTELLINI",
            2: ""
        }


        ############### ROW CONFIGURE
        self.grid_rowconfigure(0, minsize=self.margin) # empty
        self.grid_rowconfigure(1, minsize=self.margin)  # back button
        self.grid_rowconfigure(2, minsize=self.margin*2) # empty
        self.grid_rowconfigure(3, minsize=self.margin)  # configuration row
        self.grid_rowconfigure(4, minsize=self.margin*2)  # empty
        self.grid_rowconfigure(5, minsize=self.margin)  # check_file row
        self.grid_rowconfigure(6, minsize=self.margin * 2)  # empty
        self.grid_rowconfigure(7, minsize=self.margin * 3)  # badges label
        self.grid_rowconfigure(8, minsize=self.margin * 2)  # radio buttons
        self.grid_rowconfigure(9, minsize=self.margin * 2)  # empty
        self.grid_rowconfigure(10, minsize=self.margin*4)  # Gdrive logo and label
        self.grid_rowconfigure(11, minsize=self.margin * 15)  # Gdrive list and file data Frame
        self.grid_rowconfigure(12, minsize=self.margin)  # Gdrive horizontal scrollbar
        self.grid_rowconfigure(13, minsize=self.margin * 4)  # status circle
        self.grid_rowconfigure(14, minsize=self.margin)  # verify button
        self.grid_rowconfigure(15, minsize=self.margin)  # empty

        ############### COLUMN CONFIGURE
        self.grid_columnconfigure(0, minsize=self.margin) # empty
        self.grid_columnconfigure(1, minsize=self.margin) # back button
        self.grid_columnconfigure(2, minsize=self.margin*11) # window start from here (buttons and listbox)
        self.grid_columnconfigure(3, minsize=self.margin) # empty
        self.grid_columnconfigure(4, minsize=self.margin) # textboxes column
        self.grid_columnconfigure(5, minsize=self.margin)  # Gdrive Scrollbar
        self.grid_columnconfigure(6, minsize=self.margin)  # data column
        self.grid_columnconfigure(7, minsize=self.margin*10) # empty


        # define << back Button
        self.back_button = Button(self, text="<<", width=2, height=1, command=lambda:self.open_new_window(master, self.prior_window))
        self.back_button.grid(row=1, column=1, sticky="w")

        # define import configuration Button
        self.import_config_button = Button(self, width=20, text="Importa Configurazione", command=lambda: [self.changeConfigContent(self.config_txtbox), self.Controller.create_config_from_csv(self.config_txtbox.get())])
        self.import_config_button.grid(row=3, column=2)

        # define config Textbox
        self.config_txtbox = Entry(self, width=60, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.config_txtbox.grid(row=3, column=4, columnspan=4)

        # verify conversion_config path
        if os.path.exists(self.Controller.conversion_table_path):
            self.config_txtbox.configure(state='normal')
            self.config_txtbox.delete(0, "end")
            self.config_txtbox.insert(0, f"CONFIGURAZIONE TROVATA >> {self.Controller.conversion_table_path}")
            self.config_txtbox.configure(state='disabled')

        # define file select Button
        self.check_file_button = Button(self, width=20, text="LUL controllo", command=lambda:self.changeCheckContent(self.check_file_txtbox))
        self.check_file_button.grid(row=5, column=2)

        # define file_to_check Textbox
        self.check_file_txtbox = Entry(self, width=60, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.check_file_txtbox.grid(row=5, column=4, columnspan=4)

        # define where_badges Label
        self.where_badges_label = Label(self, text="Dove sono i Cartellini?", font=("Calibri", 20, "bold"), fg=appLib.color_orange, bg=appLib.default_background)
        self.where_badges_label.grid(row=7, column=2, columnspan=6)

        # define badges Radio Buttons
        self.radio_BusinessCat = Radiobutton(self, text="Cartella di BusinessCat", font=("Calibri", 12), padx=20, variable=self.radio_buttons_val, value=1, bg=appLib.default_background)
        self.radio_BusinessCat.grid(row=8, column=2, columnspan=6, sticky="w")
        self.radio_Other = Radiobutton(self, text="Altro (Seleziona)", font=("Calibri", 12), padx=20, variable=self.radio_buttons_val, value=2, bg=appLib.default_background, command=lambda:self.set_choosen_radio())
        self.radio_Other.grid(row=8, column=2, columnspan=6, sticky="e")
        self.radio_buttons_val.set(1)

        # define google drive img Label and text Label
        self.drive_logo = PIL.Image.open("../config_files/imgs/drive_logo.ico")
        self.drive_logo = self.drive_logo.resize((38,38))
        self.drive_logo = PIL.ImageTk.PhotoImage(self.drive_logo)
        self.gdrive_frame = Frame(self, bg=appLib.default_background)
        self.gdrive_frame.grid(row=10, column=2, columnspan=4, sticky="nsew")
        self.gdrive_img_label = Label(self.gdrive_frame,image=self.drive_logo, bg=appLib.default_background).place(rely=0.58,anchor="w")
        self.gdrive_text_label = Label(self.gdrive_frame, text="Fogli di Calcolo sul tuo Drive", font=("Calibri", 12, "bold"), fg=appLib.color_orange, bg=appLib.default_background).place(relx=0.50, rely=0.55,anchor="center")

        # define Gdrive files Listbox
        self.gdrive_yscrollbar = Scrollbar(self, orient=VERTICAL)
        self.gdrive_yscrollbar.grid(row=11, column=5, rowspan=3, sticky="nsw")
        self.gdrive_xscrollbar = Scrollbar(self, orient=HORIZONTAL)
        self.gdrive_xscrollbar.grid(row=14, column=2, columnspan=3, sticky="ew")
        self.gdrive_listbox = Listbox(self, {"font":("Calibri",12), "yscrollcommand":self.gdrive_yscrollbar.set, "xscrollcommand":self.gdrive_xscrollbar.set, "activestyle":"none"})
        self.gdrive_listbox.grid(row=11, column=2, columnspan=3, rowspan=3, sticky="nsew")
        self.gdrive_yscrollbar.config(command=self.gdrive_listbox.yview)
        self.gdrive_xscrollbar.config(command=self.gdrive_listbox.xview)
        #self.gdrive_listbox.bind("<<ListboxSelect>>", lambda event: self.set_choosen_file_from_list(event))
        self.gdrive_listbox.bind("<Double-1>", lambda event: self.set_choosen_file_from_list(event))
        self.refresh_gdrive_values()

        # define file data Frame
        self.file_data_frame = Frame(self, bg=appLib.default_background)
        self.file_data_frame.grid(row=11, column=6, columnspan=3, sticky="nsew")

        # define choosen file Label
        self.choosen_file_label = Label(self.file_data_frame, text="File da Comparare", font=("Calibri", 14, "bold"), bg=appLib.default_background).pack(pady=(0,5))

        # define selected file Label
        self.selected_file = Label(self.file_data_frame, width=28, text="<< Seleziona", font=("Calibri", 12), bg=appLib.default_background)
        self.selected_file.pack(pady=(0,20))

        # define worksheet Label
        self.choose_worksheet_label = Label(self.file_data_frame, text="Scheda dei dati", font=("Calibri", 14, "bold"), bg=appLib.default_background).pack(pady=(0,5))

        # define worksheet Combobox
        self.choose_worksheet_combobox = ttk.Combobox(self.file_data_frame, width=28, textvariable=self.choose_worksheet_val, state="readonly")
        self.choose_worksheet_combobox.pack()
        self.choose_worksheet_combobox.bind("<<ComboboxSelected>>", lambda event: self.choose_worksheet_val.set(self.choose_worksheet_combobox.get()))

        # status circle
        self.canvas_frame = Frame(self, height=64, width=32, bg=appLib.default_background)
        self.canvas_frame.grid(row=12, column=6, columnspan=2, rowspan=2)
        self.canvas_frame.update()
        self.canvas = Canvas(self.canvas_frame, width=self.canvas_frame.winfo_width(), height=self.canvas_frame.winfo_height()/2, highlightthickness=0, bg=appLib.default_background)
        self.canvas.grid()
        self.canvas_frame.update()
        top_left_coord = [0, 0]
        bottom_right_coord = [self.canvas.winfo_width()-1, self.canvas.winfo_height()-1]
        self.status_circle = self.canvas.create_oval(top_left_coord, bottom_right_coord, outline="", fill=appLib.color_green)
        self.circle_label = Label(self.canvas_frame, text="Pronto", font=("Calibri", 10, "bold"), fg=appLib.color_green, bg=appLib.default_background)
        self.circle_label.grid(pady=(2,4))

        # define VERIFY Button
        self.verify_button = Button(self, text="Verifica", width=16, height=1, command=lambda:self.verify())
        self.verify_button.grid(row=14, column=6, columnspan=2, sticky="ns")


    """    METHODS    """

    def refresh_gdrive_values(self):
        files = appLib.get_sheetlist()
        for f in files:
            self.gdrive_listbox.insert(END, f["name"])
        self.gdrive_filelist = files

    def __populate_sheetnames_combobox(self, array):
        self.choose_worksheet_combobox['values'] = [(*self.choose_worksheet_combobox['values'], val) for val in array]

    def set_choosen_file_from_list(self, event):

        selection = event.widget.curselection()
        if selection:
            index = selection[0]
            data = event.widget.get(index)

            valid = messagebox.askyesno("Conferma", f"{data} è il file che vuoi indicare come riferimento?")

            if valid:

                # set status circle and label in download
                self.canvas.itemconfig(self.status_circle, fill=appLib.color_yellow)
                self.circle_label.config(text="Recupero il file...", fg=appLib.color_yellow)
                self.update()

                # set the choosen filename as the file to be downloaded
                self.choosen_drivedoc_val.set(data)

                # find id of the document based on name
                file_ID = None
                for doc in self.gdrive_filelist:
                    if doc["name"] == self.choosen_drivedoc_val.get():
                        file_ID = doc["id"]
                if not file_ID:
                    raise Exception("Nessun documento sul Drive (id) corrisponde al nome scelto!")

                # set downloaded_df
                self.downloaded_df = appLib.get_df_bytestream(file_ID)

                # populate combobox
                sheetnames = appLib.get_sheetnames_from_bytes(self.downloaded_df)
                self.__populate_sheetnames_combobox(sheetnames)

                if len(self.choosen_drivedoc_val.get()) > 26:
                    self.choosen_drivedoc_val.set(self.choosen_drivedoc_val.get().replace(" ", "\n", 2))
                self.selected_file.config(text=self.choosen_drivedoc_val.get())

                # set status circle and label in download
                self.canvas.itemconfig(self.status_circle, fill=appLib.color_green)
                self.circle_label.config(text="Pronto", fg=appLib.color_green)
                self.update()
                return

            self.choosen_drivedoc_val.set("<< Seleziona")
            self.downloaded_df = None

    def set_choosen_radio(self):
        path = filedialog.askdirectory(initialdir=os.getcwd(), title="Select Badges Directory")
        if not path:
            return
        self.radio_values[2] = path

    def changeConfigContent(self, txtbox):
        '''
        change the content of a given Entry Object
        '''

        overwrite = True
        if os.path.exists(self.Controller.conversion_table_path):
            overwrite = messagebox.askyesno("Configurazione già presente!", "Esiste già una configurazione per i valori da estrarre, vuoi sceglierne un'altra?\n(ATTENZIONE: scegliendo 'SI' i dati di configurazione precedenti andranno persi)")

        if overwrite:
            filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select a File",
                                                  filetype=[("CSV files", "*.csv*")])
            if not filename:
                return

            txtbox.configure(state='normal')
            txtbox.delete(0, "end")
            txtbox.insert(0, filename)
            txtbox.configure(state='disabled')

    def changeCheckContent(self, txtbox):
        '''
        change the content of a given Entry Object
        '''
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select a File",
                                              filetype=[("PDF Files", "*.pdf*")])
        if not filename:
            return

        txtbox.configure(state='normal')
        txtbox.delete(0, "end")
        txtbox.insert(0, filename)
        txtbox.configure(state='disabled')

    def verify(self):

        try:

            if os.path.exists(self.Controller.verify_filename):
                check_input = messagebox.askyesno("File Esistente", f"il file {self.Controller.verify_filename} esiste già, vuoi crearne uno nuovo?")
                if check_input:
                    os.remove(self.Controller.verify_filename)
                else:
                    return


            # get values from window
            self.Controller.set_paychecks_to_check_path(self.check_file_txtbox.get())
            rb_val = self.radio_buttons_val.get()
            self.Controller.set_badges_path(self.radio_values[rb_val])
            choosen_sheet = self.choose_worksheet_combobox.get()

            # verify values
            valid = self.Controller.validate_data()
            if valid:
                # set status circle and label back to ready
                self.canvas.itemconfig(self.status_circle, fill=appLib.color_yellow)
                self.circle_label.config(text="Verifica in corso", fg=appLib.color_yellow)
                self.update()

                self.Controller.paycheck_verification()
                self.Controller.badges_verification()
                self.Controller.compare_badges_to_paychecks()
                problems = self.Controller.compare_paychecks_to_drive(df_bytestream=self.downloaded_df, sheet=choosen_sheet)

                # verify continue
                if problems["uncommon_indexes"]:
                    messagebox.showwarning(f"ATTENZIONE - {problems['error']}", f"I seguenti lavoratori sul Drive non sono presenti tra le Buste Paga\n\n{problems['uncommon_indexes']}\n\n\nNON VERRANNO INSERITI NEL FILE DI RISULTA.")

                # set status circle and label back to ready
                self.canvas.itemconfig(self.status_circle, fill=appLib.color_green)
                self.circle_label.config(text="Pronto", fg=appLib.color_green)
                self.update()

                openfile = messagebox.askyesno("Verifica terminata!", "Vuoi aprire il file di risulta?")
                if openfile:
                    os.system("start"+ " " + self.Controller.verify_filename)

        except Exception as e:
            messagebox.showerror("Errore", e)
            # set status circle and label back to ready
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_red)
            self.circle_label.config(text="Errore", fg=appLib.color_red)
            self.update()
            raise


# BILLING WINDOWS
class Billing_Landing_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 500
        self.height = 240
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Billing Manager")
        self.margin = 15


        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, Home_Window))
        self.back_button.pack(anchor="w", pady=(self.margin,0), padx=(self.margin,0))

        #define welcome message and first instruction
        self.welcome_lbl = Label(self, text="""Benvenuto nel Billing Manager di BusinessCat!""", font=("Calibri", 16, "bold"), bg=appLib.default_background, fg=appLib.color_orange)
        self.welcome_lbl.pack(anchor="center", pady=self.margin)

        # define buttons frame
        self.BUTTONS_FRAME = Frame(self, bg=appLib.default_background)
        self.BUTTONS_FRAME.pack(anchor="center", padx=self.margin, fill="both", expand=True)
        self.jobs_button = Button(self.BUTTONS_FRAME, width=15, text="Jobs", font=("Calibri", 10, "bold"), command=lambda:self.open_new_window(master, Edit_Jobs_Window))
        self.jobs_button.grid(column=0, row=0, pady=5, padx=5)
        self.apply_balloon_to_widget(self.jobs_button, "Visualizza e edita le possibili mansioni nelle fatture")
        self.billing_profiles_button = Button(self.BUTTONS_FRAME, width=15, text="Billing Profiles", font=("Calibri", 10, "bold"), command=lambda:self.open_new_window(master, Edit_BillingProfiles_Window))
        self.billing_profiles_button.grid(column=2, row=0, pady=5, padx=5)
        self.apply_balloon_to_widget(self.billing_profiles_button, "Visualizza e edita i possiibli profili di fatturazione")

        self.start_billing_btn = Button(self.BUTTONS_FRAME, width=20, text="CREA FATTURA", font=("Calibri", 10, "bold"))
        self.start_billing_btn.grid(column=1, row=2, pady=self.margin, padx=5)

        #resize grid
        size = self.BUTTONS_FRAME.grid_size()
        for col in range(size[0]):
            self.BUTTONS_FRAME.grid_columnconfigure(col, weight=1)
        for row in range(size[1]):
            self.BUTTONS_FRAME.grid_rowconfigure(row, minsize=self.margin, weight=1)

class Edit_Jobs_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 500
        self.height = 490
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Edit Jobs")
        self.margin = 15
        self.__jobs_path = "../config_files/BusinessCat billing/jobs.json"
        self.__load_jobs()
        self.displayed_job = IntVar()
        self.displayed_job.set(0)
        self.untouchable_keys = ["id", "billing_profile_id"]
        self.default_new_job = {
            "id":"",
            "name":"",
            "billing_profile_id":""
        }



        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, self.prior_window))
        self.back_button.pack(anchor="w", pady=(self.margin,0), padx=(self.margin,0))

        ####### define main frame structure
        self.MASTER_FRAME = Frame(self, bg=appLib.default_background)
        self.MASTER_FRAME.pack(anchor="center", pady=(self.margin,0))

        self.ADD_RMV_FRAME = Frame(self.MASTER_FRAME, bg=appLib.default_background)
        self.ADD_RMV_FRAME.grid(sticky="nsew", columnspan=4)

        self.DATA_FRAME = Frame(self.MASTER_FRAME,width=420, height=320, bg=appLib.default_background)
        self.DATA_FRAME.grid(row=1, columnspan=4, sticky="nsew")
        self.DATA_FRAME.pack_propagate(0)
        self.CANVAS = Canvas(self.DATA_FRAME)
        self.CANVAS.pack(side="left", fill="both", expand=True)
        self.CANVAS.config(bg=appLib.color_light_orange)

        # define add/rmv buttons
        self.add_lbl = Label(self.ADD_RMV_FRAME, text="+", width=2,height=1, font=("Calibri",20,"bold"), bg=appLib.color_grey, fg=appLib.color_green)
        self.add_lbl.bind("<Button-1>", lambda e: self.add_job())
        self.apply_balloon_to_widget(self.add_lbl, "Aggiungi un nuovo Job")
        self.add_lbl.pack(side="right", padx=5)
        self.rmv_lbl = Label(self.ADD_RMV_FRAME, text="-", width=2,height=1, font=("Calibri",20,"bold"), bg=appLib.color_grey, fg=appLib.color_red)
        self.rmv_lbl.bind("<Button-1>", lambda e: self.remove_job())
        self.apply_balloon_to_widget(self.rmv_lbl, "Rimuovi il Job corrente")
        self.rmv_lbl.pack(side="right", padx=5)

        # define navbar
        navbar_background = appLib.default_background
        navbar_padx = 40
        self.NAV_FRAME = Frame(self, width=420, height=50, bg=navbar_background)
        self.NAV_FRAME.pack(anchor="center", pady=(5,0))
        self.NAV_FRAME.grid_propagate(0)

        # define previous button
        self.previous_button = Button(self.NAV_FRAME, text="<", font=("bold"), width=2, height=1, command=self.previous_job)
        self.previous_button.grid(row=0, column=0, padx=(0,navbar_padx), sticky="w")

        # define info label
        self.info_label = Label(self.NAV_FRAME, font=("Calibri", 12, "bold"), bg=navbar_background)
        self.info_label.grid(row=0, column=1, padx=navbar_padx, sticky="nsew")

        # define next button
        self.next_button = Button(self.NAV_FRAME, text=">", font=("bold"), width=2, height=1, command=self.next_job)
        self.next_button.grid(row=0, column=2, padx=(navbar_padx,0), sticky="e")

        #resize grid
        size = self.NAV_FRAME.grid_size()
        for col in range(size[0]):
            self.NAV_FRAME.grid_columnconfigure(col, minsize=self.NAV_FRAME.winfo_width()/3, weight=1)

        #fetch and display json data
        self.display_data()


    """ PRIVATE METHODS """
    def __load_jobs(self):
        with open(self.__jobs_path, "r") as f:
            self.jobs = json.load(f)

    def __parse_data(self):
        keys = []
        values = []
        for child in self.JSON_CONTAINER.winfo_children():
            if isinstance(child, Label):
                keys.append(child.cget("text"))
            elif isinstance(child, Entry):
                values.append(child.get().strip())
        new_job = dict(zip(keys,values))
        return new_job

    def __save_data(self, new_data=None):
        """ save current displayed job """
        if new_data and new_data != self.jobs[self.displayed_job.get()]:
            save = messagebox.askyesno("Salvare?", "I dati di questa mansione sono stati modificati\nsi desidera salvare i cambiamenti?")
            if save:
                self.jobs[self.displayed_job.get()] = new_data

        with open(self.__jobs_path, "w") as f:
            f.write(json.dumps(self.jobs, indent=4))

    def __get_new_job_id(self):
        id_lenght = 4

        check_high = 0
        for job in self.jobs:
            if int(job["id"]) > check_high:
                check_high = int(job["id"])
        check_high = str(check_high + 1) # increment 1 from the highest id found among all jobs

        new_id = "0"*(id_lenght-len(check_high)) + check_high
        return new_id


    """ PUBLIC METHODS """
    def next_job(self):

        self.__save_data(self.__parse_data())

        max_lenght = len(self.jobs) -1 if len(self.jobs) > 0 else 0
        if self.displayed_job.get() < max_lenght:
            self.displayed_job.set(self.displayed_job.get() + 1)
        else:
            self.displayed_job.set(0)

        self.display_data()

    def previous_job(self):

        self.__save_data(self.__parse_data())

        max_lenght = len(self.jobs) -1 if len(self.jobs) > 0 else 0
        if self.displayed_job.get() > 0:
            self.displayed_job.set(self.displayed_job.get() - 1)
        else:
            self.displayed_job.set(max_lenght)

        self.display_data()

    def display_data(self):

        # clear previous view
        try:
            self.JSON_CONTAINER.destroy()
        except:
            pass

        padx = 5
        label_size = 60
        txt_box_size = 300
        job_name = ""

        # set new container
        self.JSON_CONTAINER = Frame(self.CANVAS, width=400, bg=appLib.color_light_orange)
        self.JSON_CONTAINER.pack(fill="both", expand=True)
        self.JSON_CONTAINER.grid_columnconfigure(0, minsize=label_size, weight=1)
        self.JSON_CONTAINER.grid_columnconfigure(1, minsize=txt_box_size, weight=3)

        # senza lavori da mostrare
        if not self.jobs:
            no_jobs_lbl = Label(self.JSON_CONTAINER, text="NON CI SONO JOBS DA MOSTRARE", font=("Calibri", 12), bg=appLib.color_light_orange)
            no_jobs_lbl.pack(anchor="center", ipady=100)

        # iterate throught jobs and create new grid in the new container
        try:
            i = 0
            for k,v in self.jobs[self.displayed_job.get()].items():

                if k == "name":
                    job_name = v

                key = Label(self.JSON_CONTAINER, text=k, bg=appLib.color_light_orange)
                key.grid(row=i, column=0, pady=(2,1), padx=padx, sticky="nsew")

                value = Entry(self.JSON_CONTAINER, bg=appLib.default_background)
                value.grid(row=i, column=1, pady=(2,1), padx=padx, sticky="nsew")
                value.insert(0, v)

                # conditionally disable fields
                if k in self.untouchable_keys:
                    value.config(state=DISABLED)

                i+=1

            # define displayed data scrollbar
            try:
                self.YSCROLL.destroy()
            except:
                pass
            self.YSCROLL = Scrollbar(self.DATA_FRAME, orient="vertical", command=self.CANVAS.yview)
            self.YSCROLL.pack(side="right", fill="y", expand=False)
            self.CANVAS.configure(yscrollcommand=self.YSCROLL.set)

            self.CANVAS.create_window((0,0), window=self.JSON_CONTAINER, anchor="nw")
            self.CANVAS.bind("<Configure>", lambda e: self.CANVAS.configure(scrollregion=self.CANVAS.bbox("all")))

        except:
            pass

        # change label
        self.info_label.config(text=f"{job_name}")

    def add_job(self):
        create_new = messagebox.askyesno("", "Vuoi creare un nuovo Job?")
        if create_new:
            new_job = copy.deepcopy(self.default_new_job)
            new_job["id"] = self.__get_new_job_id()
            self.jobs.append(new_job)
            go_to_last = messagebox.askyesno("", "Il nuovo Job è stato inizializzato, vuoi visualizzarlo?")
            self.__save_data(self.__parse_data())  # save current job

            if go_to_last:
                self.displayed_job.set(len(self.jobs)-1) # set destination job to the new one
                self.display_data() # display destination job

    def remove_job(self):
        confirm_delete = messagebox.askyesno("", "Vuoi davvero eliminare questo Job?")
        if confirm_delete:
            self.jobs.pop(self.displayed_job.get())
            self.displayed_job.set(self.displayed_job.get() - 1)  # set destination job to the previous one
            self.__save_data()  # save current job
            self.display_data() # display destination job

class Edit_BillingProfiles_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 500
        self.height = 490
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Edit Billing Profiles")
        self.margin = 15
        self.__billing_profiles_path = "../config_files/BusinessCat billing/billing_profiles.json"
        self.__load_billing_profiles()
        self.displayed_profile = IntVar()
        self.displayed_profile.set(0)
        self.untouchable_keys = ["id"]
        self.float_fields = ["threshold_hours", "time_to_add", "price"]
        self.default_profile = {
            "threshold_hours": 0.0,
            "add_over_threshold": True,
            "time_to_add": 0.0,
            "pattern": [],
            "pricelist": [
                {
                    "tag": "OR",
                    "name": "ore_ordinarie",
                    "price": 0.0
                },
                {
                    "tag": "ST",
                    "name": "ore_straordinarie",
                    "price": 0.0
                },
                {
                    "tag": "MN",
                    "name": "ore_notturne",
                    "price": 0.0
                },
                {
                    "tag": "OF",
                    "name": "ore_festive",
                    "price": 0.0
                },
                {
                    "tag": "SF",
                    "name": "ore_straordinarie_festive",
                    "price": 0.0
                },
                {
                    "tag": "SN",
                    "name": "ore_straordinarie_notturne",
                    "price": 0.0
                },
                {
                    "tag": "FN",
                    "name": "ore_festive_notturne",
                    "price": 0.0
                }
            ],
            "id": "",
            "name":""
        }
        self.default_pattern_object = {
            "perform": "/",
            "amount": 2,
            "keep": True
        }



        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, self.prior_window))
        self.back_button.pack(anchor="w", pady=(self.margin,0), padx=(self.margin,0))

        ####### define main frame structure
        self.MASTER_FRAME = Frame(self, bg=appLib.default_background)
        self.MASTER_FRAME.pack(anchor="center", pady=(self.margin,0))

        self.ADD_RMV_FRAME = Frame(self.MASTER_FRAME, bg=appLib.default_background)
        self.ADD_RMV_FRAME.grid(sticky="nsew", columnspan=4)

        self.DATA_FRAME = Frame(self.MASTER_FRAME,width=420, height=320, bg=appLib.default_background)
        self.DATA_FRAME.grid(row=1, columnspan=4, sticky="nsew")
        self.DATA_FRAME.pack_propagate(0)
        self.CANVAS = Canvas(self.DATA_FRAME)
        self.CANVAS.pack(side="left", fill="both", expand=True)
        self.CANVAS.config(bg=appLib.color_light_orange)

        # define add/rmv buttons
        self.add_lbl = Label(self.ADD_RMV_FRAME, text="+", width=2,height=1, font=("Calibri",20,"bold"), bg=appLib.color_grey, fg=appLib.color_green)
        self.add_lbl.bind("<Button-1>", lambda e: self.add_profile())
        self.apply_balloon_to_widget(self.add_lbl, "Aggiungi un nuovo Job")
        self.add_lbl.pack(side="right", padx=5)
        self.rmv_lbl = Label(self.ADD_RMV_FRAME, text="-", width=2,height=1, font=("Calibri",20,"bold"), bg=appLib.color_grey, fg=appLib.color_red)
        self.rmv_lbl.bind("<Button-1>", lambda e: self.remove_profile())
        self.apply_balloon_to_widget(self.rmv_lbl, "Rimuovi il Job corrente")
        self.rmv_lbl.pack(side="right", padx=5)

        # define navbar
        navbar_background = appLib.default_background
        navbar_padx = 40
        self.NAV_FRAME = Frame(self, width=420, height=50, bg=navbar_background)
        self.NAV_FRAME.pack(anchor="center", pady=(5,0))
        self.NAV_FRAME.grid_propagate(0)

        # define previous button
        self.previous_button = Button(self.NAV_FRAME, text="<", font=("bold"), width=2, height=1, command=self.previous_profile)
        self.previous_button.grid(row=0, column=0, padx=(0,navbar_padx), sticky="w")

        # define info label
        self.info_label = Label(self.NAV_FRAME, font=("Calibri", 12, "bold"), bg=navbar_background)
        self.info_label.grid(row=0, column=1, padx=navbar_padx, sticky="nsew")

        # define next button
        self.next_button = Button(self.NAV_FRAME, text=">", font=("bold"), width=2, height=1, command=self.next_profile)
        self.next_button.grid(row=0, column=2, padx=(navbar_padx,0), sticky="e")

        #resize grid
        size = self.NAV_FRAME.grid_size()
        for col in range(size[0]):
            self.NAV_FRAME.grid_columnconfigure(col, minsize=self.NAV_FRAME.winfo_width()/3, weight=1)

        #fetch and display json data
        self.display_data()


    """ PRIVATE METHODS """
    def __load_billing_profiles(self):
        with open(self.__billing_profiles_path, "r") as f:
            self.billing_profiles = json.load(f)

    def __parse_data(self):
        keys = []
        values = []
        for child in self.JSON_CONTAINER.winfo_children():
            if isinstance(child, Label):
                keys.append(child.cget("text"))
            elif isinstance(child, Entry):
                values.append(child.get().strip())
        new_job = dict(zip(keys,values))
        return new_job

    def __save_data(self, new_data=None):
        """ save current displayed job """
        if new_data and new_data != self.billing_profiles[self.displayed_profile.get()]:
            save = messagebox.askyesno("Salvare?", "I dati di questo profilo sono stati modificati\nsi desidera salvare i cambiamenti?")
            if save:
                self.billing_profiles[self.displayed_profile.get()] = new_data

        with open(self.__billing_profiles_path, "w") as f:
            f.write(json.dumps(self.billing_profiles, indent=4))

    def __get_new_job_id(self):
        id_lenght = 4

        check_high = 0
        for job in self.jobs:
            if int(job["id"]) > check_high:
                check_high = int(job["id"])
        check_high = str(check_high + 1) # increment 1 from the highest id found among all jobs

        new_id = "0"*(id_lenght-len(check_high)) + check_high
        return new_id


    """ PUBLIC METHODS """
    def next_profile(self):

        self.__save_data(self.__parse_data())

        max_lenght = len(self.billing_profiles) -1 if len(self.billing_profiles) > 0 else 0
        if self.displayed_profile.get() < max_lenght:
            self.displayed_profile.set(self.displayed_profile.get() + 1)
        else:
            self.displayed_profile.set(0)

        self.display_data()

    def previous_profile(self):

        self.__save_data(self.__parse_data())

        max_lenght = len(self.billing_profiles) -1 if len(self.billing_profiles) > 0 else 0
        if self.displayed_profile.get() > 0:
            self.displayed_profile.set(self.displayed_profile.get() - 1)
        else:
            self.displayed_profile.set(max_lenght)

        self.display_data()

    def display_data(self):

        # clear previous view
        try:
            self.JSON_CONTAINER.destroy()
        except:
            pass

        padx = 5
        default_pady = 5
        label_size = 60
        txt_box_size = 40
        profile_name = ""

        # set new container
        self.JSON_CONTAINER = Frame(self.CANVAS, width=400, bg=appLib.color_light_orange)
        self.JSON_CONTAINER.pack(fill="both", expand=True)
        self.JSON_CONTAINER.grid_columnconfigure(0, minsize=label_size, weight=1)
        self.JSON_CONTAINER.grid_columnconfigure(1, minsize=txt_box_size, weight=3)

        # senza lavori da mostrare
        if not self.billing_profiles:
            no_jobs_lbl = Label(self.JSON_CONTAINER, text="NON CI SONO PROFILI DA MOSTRARE", font=("Calibri", 12), bg=appLib.color_light_orange)
            no_jobs_lbl.pack(anchor="center", ipady=100)

        # iterate throught jobs and create new grid in the new container
        try:
            i = 0
            for k,v in self.billing_profiles[self.displayed_profile.get()].items():

                if k == "name":
                    profile_name = v

                key = Label(self.JSON_CONTAINER, text=k, bg=appLib.color_light_orange)
                key.grid(row=i, column=0, pady=default_pady, padx=padx, sticky="nsew")

                # not array values are packed as label(key) entry(value)
                if not isinstance(v, list):
                    value = Entry(self.JSON_CONTAINER, width=txt_box_size, bg=appLib.default_background)
                    value.grid(row=i, column=1, pady=default_pady, padx=padx, sticky="nsw")
                    value.insert(0, v)

                    # conditionally disable fields
                    if k in self.untouchable_keys:
                        value.config(state=DISABLED)

                    i += 1

                # array values are packed as label(key) scrollable frame(value)
                else:
                    key = Label(self.JSON_CONTAINER, text=k, bg=appLib.color_light_orange)
                    key.grid(row=i, column=0, pady=default_pady, padx=padx, sticky="nsew")
                    i1 = 1

                    if k == "pattern":
                        self.PATTERN_FRAME_MASTER = Frame(self.JSON_CONTAINER, width=250, height=120, bg=appLib.color_red)
                        self.PATTERN_FRAME_MASTER.grid(row=i, column=1, padx=padx, sticky="nsw")
                        if not v:
                            self.PATTERN_FRAME_MASTER.grid(sticky="w")
                        self.PATTERN_FRAME_MASTER.pack_propagate(0)

                        self.PATTERN_CANVAS = Canvas(self.PATTERN_FRAME_MASTER)
                        self.PATTERN_CANVAS.pack(side="left", fill="both", expand=True)
                        self.PATTERN_CANVAS.config(bg=appLib.color_orange)
                        self.PATTERN_CANVAS.bind("<Configure>", lambda e: self.PATTERN_CANVAS.configure(scrollregion=self.PATTERN_CANVAS.bbox("all")))

                        self.PATTERN_FRAME = Frame(self.PATTERN_FRAME_MASTER, bg=appLib.color_orange)

                        self.ADD_PATTERN_BTN = Label(self.PATTERN_FRAME, text="+", width=15, bg=appLib.color_grey, fg=appLib.color_green)
                        self.ADD_PATTERN_BTN.pack(anchor="w")

                        for index, value in enumerate(v):
                            subframe = Frame(self.PATTERN_FRAME, bg=appLib.default_background)
                            subframe.pack(anchor="w", pady=default_pady, padx=padx)
                            rmv_lbl = Label(subframe, text="-", width=2, bg=appLib.color_grey, fg=appLib.color_red)
                            rmv_lbl.grid(row=0, column=2, pady=default_pady, padx=padx, sticky="w")

                            i2 = 0
                            for k1, v1 in value.items():
                                Label(subframe, text=k1, bg=appLib.default_background).grid(row=i2, column=0, pady=default_pady, padx=padx, sticky="w")
                                value = Entry(subframe, bg=appLib.default_background)
                                value.grid(row=i2, column=1, pady=default_pady, padx=padx, sticky="nsew")
                                value.insert(0, v1)
                                # disable field conditionally
                                if k1 in self.untouchable_keys:
                                    value.configure(state="disabled")
                                i2 += 1
                            i1 += 1
                        i += 1

                        self.PATTERN_YSCROLL = Scrollbar(self.PATTERN_CANVAS, orient="vertical", command=self.PATTERN_CANVAS.yview)
                        self.PATTERN_YSCROLL.pack(side="right", fill="y", expand=False)
                        self.PATTERN_CANVAS.configure(yscrollcommand=self.PATTERN_YSCROLL.set)
                        self.PATTERN_CANVAS.create_window((0, 20), window=self.PATTERN_FRAME, anchor="nw")




            # define displayed data scrollbar
            try:
                self.YSCROLL.destroy()
            except:
                pass
            self.YSCROLL = Scrollbar(self.DATA_FRAME, orient="vertical", command=self.CANVAS.yview)
            self.YSCROLL.pack(side="right", fill="y", expand=False)
            self.CANVAS.configure(yscrollcommand=self.YSCROLL.set)

            self.CANVAS.create_window((0,0), window=self.JSON_CONTAINER, anchor="nw")
            self.CANVAS.bind("<Configure>", lambda e: self.CANVAS.configure(scrollregion=self.CANVAS.bbox("all")))

        except Exception as e:
            print(e)
            pass

        # change label
        self.info_label.config(text=f"{profile_name}")

    def add_profile(self):
        create_new = messagebox.askyesno("", "Vuoi creare un nuovo profilo?")

    def remove_profile(self):
        confirm_delete = messagebox.askyesno("", "Vuoi davvero eliminare questo profilo?")











if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    # change this to False only for testing purposes
    root.using_db = False

    # app entry point
    app = Home_Window(root)
    app.mainloop()