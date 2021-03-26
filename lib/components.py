from lib import appLib
import PIL.Image, PIL.ImageTk
import os
import copy
import uuid
import json
import shutil
import datetime
import calendar
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

# TEMPLATE CLASSES

# general purpose
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

        self.lift()

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

# billing windows
class Billing_template(Custom_Toplevel):
    def __init__(self, master=None, width=100, height=100):
        super().__init__(master)

        self.width = width
        self.height = height
        self.__screen_width = master.winfo_screenwidth()
        self.__screen_height = master.winfo_screenheight()
        self.geometry(f"{self.width}x{self.height}+{int((self.__screen_width / 2 - self.width / 2))}+{int((self.__screen_height / 2 - self.height / 2))}")
        self.title("BusinessCat Billing Manager")
        self.iconbitmap(appLib.icon_path)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", lambda: self._close(master))
        self.prior_window = None
        self.margin = 15
        self.info_wraplength = 200

        self.Biller = appLib.BillingManager()

        self.lift()



#############################################################################

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


    """ PRIVATE METHODS """
    def __BADGES_FROM_PAYCHECKS(self, filename):
        inputpdf = fitz.open(filename)

        if not os.path.exists(self.BADGES_PATH):
            os.mkdir(self.BADGES_PATH)

        check_name = ""
        # per ogni pagina
        for i in range(inputpdf.pageCount):
            page = inputpdf.loadPage(i)
            page_info = self.__get_page_owner(page)
            is_badge = page_info[0]
            page_owner = page_info[1]

            if "riepilogo" in page_owner.lower():
                break

            # salva il cartellino
            if is_badge:
                if page_owner and page_owner != check_name and "datasassunzione" not in page_owner:
                    badge = fitz.Document()
                    badge.insertPDF(inputpdf, from_page=i, to_page=i)
                    badge.save(f"{self.BADGES_PATH}/" + page_owner + ".pdf")
                    check_name = page_owner

    def __SPLIT_PAYCHECKS(self, file_to_split):
        inputpdf = fitz.open(file_to_split)
        check_name = ""

        # chech if file_to_split exists
        if not os.path.exists(file_to_split):
            raise Exception(f"File {file_to_split} inesistente!")

        # create dir if it does not exists
        if not os.path.exists(self.PAYCHECKS_PATH):
            os.mkdir(self.PAYCHECKS_PATH)

        # per ogni pagina
        for i in range(inputpdf.pageCount):
            page = inputpdf.loadPage(i)
            page_info = self.__get_page_owner(page=page)
            is_badge = page_info[0]
            paycheck_owner =page_info[1]

            if is_badge:
                #print(f"La pagina numero {i} è un cartellino")
                continue

            # check integrity
            if not paycheck_owner:
                raise Exception(f"La pagina numero {i} non ha un nome al suo interno!")
            if "riepilogo generale" in paycheck_owner.lower() or "detrazioni" in paycheck_owner.lower():
                continue

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
                page_info = self.__get_page_owner(page)
                page_owner = page_info[1]

                # salvo il cartellino
                badge = fitz.Document()
                badge.insertPDF(inputpdf, from_page=i, to_page=i)
                badge.save(f"{self.BADGES_PATH}/" + page_owner + ".pdf")

        except:
            raise Exception("#######################################\n"
                            "ERRORE nella divisione dei cartellini\n"
                            "#######################################\n")

    def __get_page_owner(self, page):
        blocks = page.getText("blocks")
        blocks.sort(key=lambda block: block[1])  # sort vertically ascending

        months = [
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


        name = ""
        is_badge = False

        for index, b in enumerate(blocks):
            if "COGNOME" in b[4].upper() and "NOME" in b[4].upper():
                name_array = blocks[index+1][4].split("\n")

                # check cessato or assunto
                if "cessato" in name_array[0].lower() or "assunto" in name_array[0].lower():
                    name_array = blocks[index+2][4].split("\n")

                if "/" in name_array[0]:
                    is_badge = True
                name = name_array[1]

                # if name is riepilogo generale fix it
                if "riepilogo generale" in name_array[0].lower():
                    name = name_array[0]

                # if name is month fix it
                try:
                    if name.split()[0].lower() in months:
                        name = name_array[0]
                except:
                    raise Exception("ERROR: name unclear in a badge")

                break

        # parse name
        new_name = ""
        for word in name.split():
            new_name += (word[0].upper() + word[1:].lower() + " ")
        page_owner = new_name[:-1]

        return (is_badge, page_owner)


    """ PUBLIC METHODS """
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

                openfile = messagebox.askyesno("Verifica terminata", "Vuoi aprire il file di risulta?")
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
class Billing_Landing_Window(Billing_template):
    def __init__(self, master=None):
        self.width = 500
        self.height = 240
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Billing Manager")

        self.choosen_month = IntVar()
        self.choosen_year = IntVar()
        self.choosen_month.set(0)
        self.choosen_year.set(0)


        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, Home_Window))
        self.back_button.pack(anchor="w", pady=(self.margin,0), padx=(self.margin,0))

        #define welcome message and first instruction
        self.welcome_lbl = Label(self, text="""Benvenuto nel Billing Manager di BusinessCat!""", font=("Calibri", 16, "bold"), bg=appLib.default_background, fg=appLib.color_orange)
        self.welcome_lbl.pack(anchor="center", pady=self.margin)

        # define buttons frame
        self.BUTTONS_FRAME = Frame(self, bg=appLib.default_background)
        self.BUTTONS_FRAME.pack(anchor="center", padx=self.margin, fill="both", expand=True)
        self.jobs_button = Button(self.BUTTONS_FRAME, width=15, text="Mansioni", font=("Calibri", 10, "bold"), command=lambda:self.open_new_window(master, Edit_Jobs_Window))
        self.jobs_button.grid(column=0, row=0, pady=5, padx=5)
        self.clients_button = Button(self.BUTTONS_FRAME, width=15, text="Clienti", font=("Calibri", 10, "bold"), command=lambda:self.open_new_window(master, Edit_Clients_Window))
        self.clients_button.grid(column=1, row=0, pady=5, padx=5)
        self.billing_profiles_button = Button(self.BUTTONS_FRAME, width=15, text="Profili", font=("Calibri", 10, "bold"), command=lambda:self.open_new_window(master, Edit_Profiles_Window))
        self.billing_profiles_button.grid(column=2, row=0, pady=5, padx=5)

        self.create_model_btn = Button(self.BUTTONS_FRAME, width=20, text="CREA MODELLO", font=("Calibri", 10, "bold"), command=self.__insert_model_data)
        self.create_model_btn.grid(column=0, row=2, pady=self.margin, padx=5)
        self.bill_btn = Button(self.BUTTONS_FRAME, width=20, text="CREA FATTURA", font=("Calibri", 10, "bold"), command=self.__insert_billing_data)
        self.bill_btn.grid(column=2, row=2, pady=self.margin, padx=5)

        #resize grid
        size = self.BUTTONS_FRAME.grid_size()
        for col in range(size[0]):
            self.BUTTONS_FRAME.grid_columnconfigure(col, weight=1)
        for row in range(size[1]):
            self.BUTTONS_FRAME.grid_rowconfigure(row, minsize=self.margin, weight=1)


    """     PRIVATE METHODS     """
    def __insert_model_data(self):
        window = Toplevel(self, bg=appLib.color_orange)
        window.resizable(height=False, width=True)
        window.title("Inserimento Dati")
        window.iconbitmap(appLib.icon_path)

        x = self.winfo_x() + 80
        y = self.winfo_y() + 25
        window.geometry(f"+{x}+{y}")

        pady = 10
        padx = 10

        Button(window, text="File Cartellini", bg=appLib.default_background, command=lambda: set_badges_path(self)).grid(row=0, column=0, pady=pady, padx=padx)
        badges_entry = Entry(window, state="disabled", disabledbackground=appLib.color_light_orange)
        badges_entry.grid(row=0, column=1, pady=pady, padx=padx, sticky="nsew")

        month_label = Label(window, text="Mese", bg=appLib.color_orange)
        month_label.grid(row=1, column=0, padx=padx, pady=pady)
        month_combobox = ttk.Combobox(window, justify="center", state="readonly")
        month_combobox['values'] = [x for x in range(1,13)]
        month_combobox.grid(row=1, column=1, padx=padx, pady=pady, sticky="nsew")
        month_combobox.bind("<<ComboboxSelected>>", lambda e: self.choosen_month.set(int(month_combobox.get())))

        year_label = Label(window, text="Anno", bg=appLib.color_orange)
        year_label.grid(row=2, column=0, padx=padx, pady=pady)
        current_year = datetime.datetime.now().year
        year_combobox = ttk.Combobox(window, justify="center", state="readonly")
        year_combobox['values'] = [x for x in range(current_year, current_year+3)]
        year_combobox.grid(row=2, column=1, padx=padx, pady=pady, sticky="nsew")
        year_combobox.bind("<<ComboboxSelected>>", lambda e: self.choosen_year.set(int(year_combobox.get())))

        # setting view
        window.grid_columnconfigure(0, minsize=100)
        window.grid_columnconfigure(1, minsize=250, weight=1)

        Button(window, text="Genera modello", bg=appLib.default_background, command=lambda:generate_model(self)).grid(row=3, column=0, columnspan=2, pady=pady, padx=padx, sticky="nsew")

        # setting minsize
        window.update()
        window.minsize(window.winfo_width(), window.winfo_height())

        def set_badges_path(self):
            filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Seleziona i Cartellini", filetype=[("Excel File", "*.xlsx*"), ("Excel File", "*.xls*")])
            if not filename:
                messagebox.showerror("ERRORE", "Per continuare la generazione del modello è necessario un file excel contenente i cartellini")
                return
            badges_entry.configure(state="normal")
            badges_entry.delete(0, 'end')
            badges_entry.insert(0, filename)
            badges_entry.configure(state="disabled")
            self.Biller._set_badges_path(filename)

        def generate_model(self):
            try:
                if self.Biller.badges_path and self.choosen_month.get() != 0 and self.choosen_year.get() != 0:
                    self.Biller._set_billing_time(self.choosen_month.get(), self.choosen_year.get())
                    self.Biller._create_model()
                    messagebox.showinfo("Successo", "Modello generato con successo!")
                    window.destroy()
                else:
                    messagebox.showerror("ERRORE", "Dati mancanti")
            except Exception as e:
                messagebox.showerror("ERRORE", f"{e}")

        window.lift()
        window.grab_set()

    def __insert_billing_data(self):
        window = Toplevel(self)
        window.resizable(height=False, width=True)
        window.title("Inserimento Dati")
        window.iconbitmap(appLib.icon_path)

        x = self.winfo_x() + 80
        y = self.winfo_y() + 25
        window.geometry(f"+{x}+{y}")

        pady = 10
        padx = 10

        TOP_CONTAINER = Frame(window, bg=appLib.color_orange)
        TOP_CONTAINER.pack(fill="both", expand=True)
        TOP_CONTAINER.grid_columnconfigure(0, minsize=200)
        TOP_CONTAINER.grid_columnconfigure(1, minsize=350, weight=1)

        Button(TOP_CONTAINER, text="Importa Modello Fatturazione", font=("Calibri", 10, "bold"), command=lambda:set_badges_path()).grid(row=0, column=0, padx=padx, pady=pady)
        model_path_txtbox = Entry(TOP_CONTAINER, disabledbackground=appLib.color_light_orange, state="disabled")
        model_path_txtbox.grid(row=0, column=1, padx=padx, pady=pady, sticky="ew")

        Label(TOP_CONTAINER, text="Mese", bg=appLib.color_orange, font=("Calibri", 10, "bold")).grid(row=1, column=0, padx=padx, pady=pady)
        month_combobox = ttk.Combobox(TOP_CONTAINER, values=[x for x in range(1,13)], state="readonly")
        month_combobox.grid(row=1, column=1, padx=padx, pady=pady, sticky="ew")
        month_combobox.bind("<<ComboboxSelected>>", lambda e: self.choosen_month.set(int(month_combobox.get())))

        Label(TOP_CONTAINER, text="Anno", bg=appLib.color_orange, font=("Calibri", 10, "bold")).grid(row=2, column=0, padx=padx, pady=pady)
        current_year = datetime.datetime.now().year
        year_combobox = ttk.Combobox(TOP_CONTAINER, values=[x for x in range(current_year, current_year+3)], state="readonly")
        year_combobox.grid(row=2, column=1, padx=padx, pady=pady, sticky="ew")
        year_combobox.bind("<<ComboboxSelected>>", lambda e: self.choosen_year.set(int(year_combobox.get())))

        sep = ttk.Separator(window, orient='horizontal')
        sep.pack(fill="x")

        BOTTOM_CONTAINER = Frame(window, bg=appLib.default_background)
        BOTTOM_CONTAINER.pack(fill="both", expand=True)
        BOTTOM_CONTAINER.grid_columnconfigure(0, minsize=200)
        BOTTOM_CONTAINER.grid_columnconfigure(1, minsize=350, weight=1)

        Label(BOTTOM_CONTAINER, text="Profilo da fatturare", bg=appLib.default_background, font=("Calibri", 10, "bold")).grid(row=0, column=0,padx=padx,pady=pady, sticky="nsew")
        billing_profile_combobox = ttk.Combobox(BOTTOM_CONTAINER, values=[f"{x['id']} {x['name']}" for x in self.Biller.billing_profiles], state="readonly")
        billing_profile_combobox.grid(row=0, column=1, padx=padx, pady=pady, sticky="ew")

        Button(BOTTOM_CONTAINER, text="GENERA FATTURA", width=50, font=("Calibri", 10, "bold"), command=lambda:make_bill()).grid(row=1, column=0, columnspan=2, padx=padx, pady=pady*2)

        # setting minsize
        window.update()
        window.minsize(window.winfo_width(), window.winfo_height())

        def set_badges_path():
            filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Seleziona i Cartellini",filetype=[("Excel File", "*.xlsx*"), ("Excel File", "*.xls*")])
            if not filename:
                return
            model_path_txtbox.configure(state="normal")
            model_path_txtbox.delete(0, 'end')
            model_path_txtbox.insert(0, filename)
            model_path_txtbox.configure(state="disabled")

        def make_bill():
            try:
                # check for missing data
                if not model_path_txtbox.get()\
                or not year_combobox.get()\
                or not month_combobox.get()\
                or not billing_profile_combobox.get():
                    messagebox.showerror("ERRORE","Dati mancanti")
                    return

                self.Biller._set_billing_time(self.choosen_month.get(), self.choosen_year.get())
                filename = self.Biller._bill(model_path=model_path_txtbox.get(), profile_to_bill=billing_profile_combobox.get())
                show_bill = messagebox.askyesno("Successo", "Fattura redatta, vuoi visualizzarla?")
                if show_bill:
                    os.system(f'"{filename}"')

            except Exception as e:
                messagebox.showerror("ERRORE", f"{e}")

class Edit_Jobs_Window(Billing_template):
    def __init__(self, master=None):
        self.width = 500
        self.height = 490
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Edit Jobs")

        self.displayed_ = IntVar()
        self.displayed_.set(0)


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
        self.add_lbl.bind("<Button-1>", lambda e: self.add_())
        self.apply_balloon_to_widget(self.add_lbl, "Aggiungi un nuovo Job")
        self.add_lbl.pack(side="right", padx=5)
        self.rmv_lbl = Label(self.ADD_RMV_FRAME, text="-", width=2,height=1, font=("Calibri",20,"bold"), bg=appLib.color_grey, fg=appLib.color_red)
        self.rmv_lbl.bind("<Button-1>", lambda e: self.remove_())
        self.apply_balloon_to_widget(self.rmv_lbl, "Rimuovi il Job corrente")
        self.rmv_lbl.pack(side="right", padx=5)

        # define navbar
        navbar_background = appLib.default_background
        navbar_padx = 40
        self.NAV_FRAME = Frame(self, width=420, height=50, bg=navbar_background)
        self.NAV_FRAME.pack(anchor="center", pady=(5,0))
        self.NAV_FRAME.grid_propagate(0)

        # define previous button
        self.previous_button = Button(self.NAV_FRAME, text="<", font=("bold"), width=2, height=1, command=self.previous_)
        self.previous_button.grid(row=0, column=0, padx=(0,navbar_padx), sticky="w")

        # define info label
        self.info_label = Label(self.NAV_FRAME, font=("Calibri", 12, "bold"), bg=navbar_background, wraplength=self.info_wraplength)
        self.info_label.grid(row=0, column=1, padx=navbar_padx, sticky="nsew")

        # define next button
        self.next_button = Button(self.NAV_FRAME, text=">", font=("bold"), width=2, height=1, command=self.next_)
        self.next_button.grid(row=0, column=2, padx=(navbar_padx,0), sticky="e")

        #resize grid
        size = self.NAV_FRAME.grid_size()
        for col in range(size[0]):
            self.NAV_FRAME.grid_columnconfigure(col, minsize=self.NAV_FRAME.winfo_width()/3, weight=1)

        #fetch and display json data
        self.display_data()


    """ PRIVATE METHODS """
    def __parse_data(self):
        keys = []
        values = []
        for child in self.JSON_CONTAINER.winfo_children():
            if isinstance(child, Label):
                keys.append(child.cget("text"))
            elif isinstance(child, ttk.Combobox):
                values.append(self.cmb_var.get().split(" ")[0])
            elif isinstance(child, Entry):
                values.append(child.get().strip())
        new_ = dict(zip(keys,values))
        return new_

    def __save_data(self, new_data=None):
        """ save current displayed job """
        if new_data and new_data != self.Biller.jobs[self.displayed_.get()]:
            save = messagebox.askyesno("Salvare?", "I dati di questa mansione sono stati modificati\nsi desidera salvare i cambiamenti?")
            if save:
                self.Biller.jobs[self.displayed_.get()] = new_data

        with open(self.Biller._jobs_path, "w") as f:
            f.write(json.dumps(self.Biller.jobs, indent=4))


    """ PUBLIC METHODS """
    def next_(self):

        self.__save_data(self.__parse_data())

        max_lenght = len(self.Biller.jobs) -1 if len(self.Biller.jobs) > 0 else 0
        if self.displayed_.get() < max_lenght:
            self.displayed_.set(self.displayed_.get() + 1)
        else:
            self.displayed_.set(0)

        self.display_data()

    def previous_(self):

        self.__save_data(self.__parse_data())

        max_lenght = len(self.Biller.jobs) -1 if len(self.Biller.jobs) > 0 else 0
        if self.displayed_.get() > 0:
            self.displayed_.set(self.displayed_.get() - 1)
        else:
            self.displayed_.set(max_lenght)

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
        if not self.Biller.jobs:
            no_jobs_lbl = Label(self.JSON_CONTAINER, text="NON CI SONO JOBS DA MOSTRARE", font=("Calibri", 12), bg=appLib.color_light_orange)
            no_jobs_lbl.pack(anchor="center", ipady=100)

        # iterate throught jobs and create new grid in the new container
        try:
            i = 0
            for k,v in self.Biller.jobs[self.displayed_.get()].items():

                if k == "name":
                    job_name = v

                key = Label(self.JSON_CONTAINER, text=k, bg=appLib.color_light_orange)
                key.grid(row=i, column=0, pady=(2,1), padx=padx, sticky="nsew")

                if k == "billing_profile_id":
                    profile = self.Biller.get_billing_profile_obj(v)
                    display_name = f"{profile['id']} {profile['name']}" if profile else ""
                    self.cmb_var = StringVar()
                    self.billing_profile_combobox = ttk.Combobox(self.JSON_CONTAINER, width=28, textvariable=self.cmb_var, state="readonly")
                    self.billing_profile_combobox['values'] = [f'{x["id"]} {x["name"]}' for x in self.Biller.billing_profiles]
                    self.billing_profile_combobox.grid(row=i, column=1, pady=(2,1), padx=padx, sticky="nsew")
                    self.billing_profile_combobox.bind('<Map>', lambda event: self.cmb_var.set(display_name))

                else:
                    value = Entry(self.JSON_CONTAINER, bg=appLib.default_background)
                    value.grid(row=i, column=1, pady=(2,1), padx=padx, sticky="nsew")
                    value.insert(0, v)

                    # conditionally disable fields
                    if k in self.Biller.untouchable_keys:
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

    def add_(self):
        create_new = messagebox.askyesno("", "Vuoi creare un nuovo Job?")
        if create_new:
            success = self.Biller._add_job()
            if success:
                go_to_last = messagebox.askyesno("", "Il nuovo Job è stato inizializzato, vuoi visualizzarlo?")
                self.__save_data(self.__parse_data())  # save current job

                if go_to_last:
                    self.displayed_.set(len(self.Biller.jobs)-1) # set destination job to the new one
            self.display_data()

    def remove_(self):
        confirm_delete = messagebox.askyesno("", "Vuoi davvero eliminare questo Job?")
        if confirm_delete:
            self.Biller._rmv_job(self.displayed_.get())
            self.displayed_.set(self.displayed_.get() - 1)  # set destination job to the previous one
            self.__save_data()  # save current job
            self.display_data() # display destination job

    # override
    def open_new_window(self, master, new_window):
        """ use this method to pass from one window to another"""
        data = self.__parse_data()
        self.__save_data(data)

        self.destroy()
        new = new_window(master)
        new.prior_window = type(self)

class Edit_Clients_Window(Billing_template):
    def __init__(self, master=None):
        self.width = 500
        self.height = 490
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Edit Clients")

        self.displayed_ = IntVar()
        self.displayed_.set(0)

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
        self.add_lbl.bind("<Button-1>", lambda e: self.add_())
        self.add_lbl.pack(side="right", padx=5)
        self.rmv_lbl = Label(self.ADD_RMV_FRAME, text="-", width=2,height=1, font=("Calibri",20,"bold"), bg=appLib.color_grey, fg=appLib.color_red)
        self.rmv_lbl.bind("<Button-1>", lambda e: self.remove_())
        self.rmv_lbl.pack(side="right", padx=5)

        # define navbar
        navbar_background = appLib.default_background
        navbar_padx = 40
        self.NAV_FRAME = Frame(self, width=420, height=50, bg=navbar_background)
        self.NAV_FRAME.pack(anchor="center", pady=(5,0))
        self.NAV_FRAME.grid_propagate(0)

        # define previous button
        self.previous_button = Button(self.NAV_FRAME, text="<", font=("bold"), width=2, height=1, command=self.previous_)
        self.previous_button.grid(row=0, column=0, padx=(0,navbar_padx), sticky="w")

        # define info label
        self.info_label = Label(self.NAV_FRAME, font=("Calibri", 12, "bold"), bg=navbar_background, wraplength=self.info_wraplength)
        self.info_label.grid(row=0, column=1, padx=navbar_padx, sticky="nsew")

        # define next button
        self.next_button = Button(self.NAV_FRAME, text=">", font=("bold"), width=2, height=1, command=self.next_)
        self.next_button.grid(row=0, column=2, padx=(navbar_padx,0), sticky="e")

        #resize grid
        size = self.NAV_FRAME.grid_size()
        for col in range(size[0]):
            self.NAV_FRAME.grid_columnconfigure(col, minsize=self.NAV_FRAME.winfo_width()/3, weight=1)

        #fetch and display json data
        self.display_data()


    """ PRIVATE METHODS """
    def __parse_data(self):
        keys = []
        values = []
        for child in self.JSON_CONTAINER.winfo_children():
            if isinstance(child, Label):
                keys.append(child.cget("text"))
            elif isinstance(child, ttk.Combobox):
                values.append(self.cmb_var.get().split(" ")[0])
            elif isinstance(child, Entry):
                values.append(child.get().strip())
        new_ = dict(zip(keys,values))
        return new_

    def __save_data(self, new_data=None):
        """ save current displayed job """
        if new_data and new_data != self.Biller.clients[self.displayed_.get()]:
            save = messagebox.askyesno("Salvare?", "I dati di questo cliente sono stati modificati\nsi desidera salvare i cambiamenti?")
            if save:
                self.Biller.clients[self.displayed_.get()] = new_data

        with open(self.Biller._clients_path, "w") as f:
            f.write(json.dumps(self.Biller.clients, indent=4))

    """ PUBLIC METHODS """
    def next_(self):
        self.__save_data(self.__parse_data())

        max_lenght = len(self.Biller.clients) -1 if len(self.Biller.clients) > 0 else 0
        if self.displayed_.get() < max_lenght:
            self.displayed_.set(self.displayed_.get() + 1)
        else:
            self.displayed_.set(0)

        self.display_data()

    def previous_(self):
        self.__save_data(self.__parse_data())

        max_lenght = len(self.Biller.clients) -1 if len(self.Biller.clients) > 0 else 0
        if self.displayed_.get() > 0:
            self.displayed_.set(self.displayed_.get() - 1)
        else:
            self.displayed_.set(max_lenght)

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
        name = ""

        # set new container
        self.JSON_CONTAINER = Frame(self.CANVAS, width=400, bg=appLib.color_light_orange)
        self.JSON_CONTAINER.pack(fill="both", expand=True)
        self.JSON_CONTAINER.grid_columnconfigure(0, minsize=label_size, weight=1)
        self.JSON_CONTAINER.grid_columnconfigure(1, minsize=txt_box_size, weight=3)

        # senza lavori da mostrare
        if not self.Biller.clients:
            no_jobs_lbl = Label(self.JSON_CONTAINER, text="NON CI SONO CLIENTI DA MOSTRARE", font=("Calibri", 12), bg=appLib.color_light_orange)
            no_jobs_lbl.pack(anchor="center", ipady=100)

        # iterate throught jobs and create new grid in the new container
        try:
            i = 0
            for k,v in self.Biller.clients[self.displayed_.get()].items():

                if k == "name":
                    name = v

                key = Label(self.JSON_CONTAINER, text=k, bg=appLib.color_light_orange)
                key.grid(row=i, column=0, pady=(2,1), padx=padx, sticky="nsew")

                value = Entry(self.JSON_CONTAINER, bg=appLib.default_background)
                value.grid(row=i, column=1, pady=(2,1), padx=padx, sticky="nsew")
                value.insert(0, v)

                # conditionally disable fields
                if k in self.Biller.untouchable_keys:
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
        self.info_label.config(text=f"{name}")

    def add_(self):
        create_new = messagebox.askyesno("", "Vuoi aggiungere un cliente?")
        if create_new:
            success = self.Biller._add_client()
            if success:
                go_to_last = messagebox.askyesno("", "Il nuovo cliente è stato inizializzato, vuoi visualizzarlo?")
                self.__save_data(self.__parse_data())  # save current

                if go_to_last:
                    self.displayed_.set(len(self.Biller.clients)-1) # set destination job to the new one
            self.display_data()

    def remove_(self):
        confirm_delete = messagebox.askyesno("", "Vuoi davvero eliminare questo cliente?")
        if confirm_delete:
            self.Biller._rmv_client(self.displayed_.get())
            self.displayed_.set(self.displayed_.get() - 1)  # set destination client to the previous one
            self.__save_data()
            self.display_data()

    # override
    def open_new_window(self, master, new_window):
        """ use this method to pass from one window to another"""
        data = self.__parse_data()
        self.__save_data(data)

        self.destroy()
        new = new_window(master)
        new.prior_window = type(self)

class Edit_Profiles_Window(Billing_template):
    def __init__(self, master=None):
        self.width = 500
        self.height = 490
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Edit Billing Profiles")

        self.displayed_ = IntVar()
        self.displayed_.set(0)


        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda: self.open_new_window(master, self.prior_window))
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
        self.add_lbl.bind("<Button-1>", lambda e: self.add_())
        self.add_lbl.pack(side="right", padx=5)
        self.rmv_lbl = Label(self.ADD_RMV_FRAME, text="-", width=2,height=1, font=("Calibri",20,"bold"), bg=appLib.color_grey, fg=appLib.color_red)
        self.rmv_lbl.bind("<Button-1>", lambda e: self.remove_())
        self.rmv_lbl.pack(side="right", padx=5)

        # define navbar
        navbar_background = appLib.default_background
        navbar_padx = 40
        self.NAV_FRAME = Frame(self, width=420, height=50, bg=navbar_background)
        self.NAV_FRAME.pack(anchor="center", pady=(5,0))
        self.NAV_FRAME.grid_propagate(0)

        # define previous button
        self.previous_button = Button(self.NAV_FRAME, text="<", font=("bold"), width=2, height=1, command=self.previous_)
        self.previous_button.grid(row=0, column=0, padx=(0,navbar_padx), sticky="w")

        # define info label
        self.info_label = Label(self.NAV_FRAME, font=("Calibri", 12, "bold"), bg=navbar_background)
        self.info_label.grid(row=0, column=1, padx=navbar_padx, sticky="nsew")

        # define next button
        self.next_button = Button(self.NAV_FRAME, text=">", font=("bold"), width=2, height=1, command=self.next_)
        self.next_button.grid(row=0, column=2, padx=(navbar_padx,0), sticky="e")

        #resize grid
        size = self.NAV_FRAME.grid_size()
        for col in range(size[0]):
            self.NAV_FRAME.grid_columnconfigure(col, minsize=self.NAV_FRAME.winfo_width()/3, weight=1)

        #fetch and display json data
        self.display_data()


    """ PRIVATE METHODS """
    def __parse_data(self):
        keys = []
        values = []

        for child in self.JSON_CONTAINER.winfo_children():
            if isinstance(child, Label):
                if child.cget("text") != "pricelist":
                    keys.append(child.cget("text"))
            elif isinstance(child, ttk.Combobox):
                values.append(self.cmb_var.get().split(" ")[0])
            elif isinstance(child, Entry):
                values.append(child.get().strip())
        new_profile = dict(zip(keys, values))

        # parse pattern and pricelist
        new_profile["pricelist"] = []
        for child in self.PRICELIST_FRAME.winfo_children():
            if isinstance(child, Frame):
                pricelist_keys = []
                pricelist_values = []
                for subchild in child.winfo_children():
                    if isinstance(subchild, Label):
                        pricelist_keys.append(subchild.cget("text"))
                    elif isinstance(subchild, Entry):
                        pricelist_values.append(subchild.get().strip())
                new_profile["pricelist"].append(dict(zip(pricelist_keys, pricelist_values)))

        # parse values
        try:
            for key in new_profile:
                if key in ["threshold_hours","time_to_add", "amount", "price"]:
                    val = str(new_profile[key]).replace(",", ".")
                    new_profile[key] = float(val)
                if key in ["keep", "add_over_threshold"]:
                    new_profile[key] = bool(int(new_profile[key]))

                if isinstance(new_profile[key], list):
                    for index, obj in enumerate(new_profile[key]):
                        for subkey in obj:
                            if subkey in ["threshold_hours", "time_to_add", "amount", "price"]:
                                val = str(new_profile[key][index][subkey]).replace(",",".")
                                try:
                                    new_profile[key][index][subkey] = int(val)
                                except:
                                    new_profile[key][index][subkey] = float(val)
                            if subkey in ["keep", "add_over_threshold"]:
                                new_profile[key][index][subkey] = bool(int(new_profile[key][index][subkey]))
        except Exception as e:
            messagebox.showerror("ERRORE", f"Il campo {key} presenta un errore\n\nERRORE: {e}\n\nNecessaria correzione dell'errore per proseguire.")
            raise

        return new_profile

    def __save_data(self, new_data=None):
        """ save current displayed job """
        if new_data and new_data != self.Biller.billing_profiles[self.displayed_.get()]:
            save = messagebox.askyesno("Salvare?", "I dati di questo profilo sono stati modificati\nsi desidera salvare i cambiamenti?")
            if save:
                self.Biller.billing_profiles[self.displayed_.get()] = new_data

        with open(self.Biller._billing_profiles_path, "w") as f:
            f.write(json.dumps(self.Biller.billing_profiles, indent=4))


    """ PUBLIC METHODS """
    def next_(self):
        data = self.__parse_data()
        self.__save_data(data)

        max_lenght = len(self.Biller.billing_profiles) -1 if len(self.Biller.billing_profiles) > 0 else 0
        if self.displayed_.get() < max_lenght:
            self.displayed_.set(self.displayed_.get() + 1)
        else:
            self.displayed_.set(0)

        self.display_data()

    def previous_(self):
        self.__save_data(self.__parse_data())

        max_lenght = len(self.Biller.billing_profiles) -1 if len(self.Biller.billing_profiles) > 0 else 0
        if self.displayed_.get() > 0:
            self.displayed_.set(self.displayed_.get() - 1)
        else:
            self.displayed_.set(max_lenght)

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
        name_ = ""

        # set new container
        self.JSON_CONTAINER = Frame(self.CANVAS, width=400, bg=appLib.color_light_orange)
        self.JSON_CONTAINER.pack(fill="both", expand=True)
        self.JSON_CONTAINER.grid_columnconfigure(0, minsize=label_size, weight=1)
        self.JSON_CONTAINER.grid_columnconfigure(1, minsize=txt_box_size, weight=3)

        # senza profili da mostrare
        if not self.Biller.billing_profiles:
            no_jobs_lbl = Label(self.JSON_CONTAINER, text="NON CI SONO PROFILI DA MOSTRARE", font=("Calibri", 12), bg=appLib.color_light_orange)
            no_jobs_lbl.pack(anchor="center", ipady=100)

        try:
            i = 0
            for k,v in self.Biller.billing_profiles[self.displayed_.get()].items():

                if k == "name":
                    name_ = v

                key = Label(self.JSON_CONTAINER, text=k, bg=appLib.color_light_orange)
                key.grid(row=i, column=0, pady=default_pady, padx=padx, sticky="nsew")


                if k == "client":
                    profile = self.Biller.get_client_object(v)
                    display_name = f"{profile['id']} {profile['name']}" if profile else ""
                    self.cmb_var = StringVar()
                    self.client_combobox = ttk.Combobox(self.JSON_CONTAINER,textvariable=self.cmb_var, state="readonly")
                    self.client_combobox['values'] = [f'{x["id"]} {x["name"]}' for x in self.Biller.clients]
                    self.client_combobox.grid(row=i, column=1, pady=(2, 1), padx=padx, sticky="ew")
                    self.client_combobox.bind('<Map>', lambda event: self.cmb_var.set(display_name))

                elif k == "pricelist":
                    i1 = 1
                    self.PRICELIST_FRAME_MASTER = Frame(self.JSON_CONTAINER, width=300, height=240, bg=appLib.color_orange)
                    self.PRICELIST_FRAME_MASTER.grid(row=i, column=1, padx=padx, pady=default_pady, sticky="nsw")
                    if not v:
                        self.PRICELIST_FRAME_MASTER.grid(sticky="w")
                    self.PRICELIST_FRAME_MASTER.pack_propagate(0)

                    self.PRICELIST_CANVAS = Canvas(self.PRICELIST_FRAME_MASTER, bg=appLib.color_green)
                    self.PRICELIST_CANVAS.pack(side="left", fill="both", expand=True)
                    self.PRICELIST_CANVAS.bind("<Configure>", lambda e: self.PRICELIST_CANVAS.configure(scrollregion=self.PRICELIST_CANVAS.bbox("all")))

                    self.PRICELIST_FRAME = Frame(self.PRICELIST_FRAME_MASTER, bg=appLib.color_orange)

                    for index, value in enumerate(v):
                        subframe = Frame(self.PRICELIST_FRAME, bg=appLib.default_background)
                        subframe.pack(anchor="w", fill="x", pady=default_pady, padx=padx)

                        i2 = 0
                        for k1, v1 in value.items():
                            Label(subframe, text=k1, bg=appLib.default_background).grid(row=i2, column=0, pady=default_pady, padx=padx, sticky="w")
                            value = Entry(subframe, bg=appLib.default_background)
                            value.grid(row=i2, column=1, pady=default_pady, padx=padx, sticky="nsew")
                            value.insert(0, v1)
                            # disable field conditionally
                            if k1 in self.Biller.untouchable_keys:
                                value.configure(state="disabled")
                            i2 += 1

                        subframe.grid_columnconfigure(1, weight=1)

                        i1 += 1
                    i += 1

                    self.PRICELIST_YSCROLL = Scrollbar(self.PRICELIST_CANVAS, orient="vertical", width=18, command=self.PRICELIST_CANVAS.yview)
                    self.PRICELIST_YSCROLL.pack(side="right", fill="y", expand=False)
                    self.PRICELIST_CANVAS.configure(yscrollcommand=self.PRICELIST_YSCROLL.set)
                    self.PRICELIST_CANVAS.create_window((0, 0), width=280, window=self.PRICELIST_FRAME, anchor="nw")

                else:
                    value = Entry(self.JSON_CONTAINER, width=txt_box_size, bg=appLib.default_background)
                    value.grid(row=i, column=1, pady=default_pady, padx=padx, sticky="nsw")
                    value.insert(0, v)

                    # conditionally disable fields
                    if k in self.Biller.untouchable_keys:
                        value.config(state=DISABLED)

                i += 1


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
        self.info_label.config(text=f"{name_}")

    def add_(self):
        create_new = messagebox.askyesno("", "Vuoi creare un nuovo profilo?")
        if create_new:
            success = self.Biller._add_billing_profile()
            if success:
                go_to_last = messagebox.askyesno("", "Il nuovo profilo è stato inizializzato, vuoi visualizzarlo?")
                self.__save_data(self.__parse_data())  # save current

                if go_to_last:
                    self.displayed_.set(len(self.Biller.billing_profiles)-1) # set destination to the new one
            self.display_data() # display destination

    def remove_(self):
        confirm_delete = messagebox.askyesno("", "Vuoi davvero eliminare questo profilo?")
        if confirm_delete:
            self.Biller._rmv_billing_profile(self.displayed_.get())
            self.displayed_.set(self.displayed_.get() - 1)
            self.__save_data()
            self.display_data()

    # override
    def open_new_window(self, master, new_window):
        """ use this method to pass from one window to another"""
        data = self.__parse_data()
        self.__save_data(data)

        self.destroy()
        new = new_window(master)
        new.prior_window = type(self)




'''
class Billing_Window(Billing_template):
    def __init__(self, master=None):
        self.width = 500
        self.height = 450
        super().__init__(master, self.width, self.height)

        self.master_ = master # keep refer to master
        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Billing")
        self.margin = 15

        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, self.prior_window))
        self.back_button.pack(anchor="w", pady=(self.margin,0), padx=self.margin)

        #display first view
        self.__first_view()


    """ PRIVATE METHODS """

    def __clear_view(self, widget=None):

        if not widget:
            for child in self.winfo_children():
                child.pack_forget()
                child.grid_forget()
        else:
            for child in widget.winfo_children():
                child.pack_forget()
                child.grid_forget()

    def __unlock_widget(self, widget):
        if str(widget.cget("state")) == "disabled":
            try:
                widget.configure(state="readonly")
            except:
                widget.configure(state="normal")

    # first_view
    def __set_badges(self):
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select a File",filetype=[("Excel File", "*.xlsx*"), ("Excel File", "*.xls*")])
        if not filename:
            return
        if filename != self.choosen_badge_var.get():
            self.choosen_badge_var.set(filename)
            self.dynamic_text.set(self.instruction_steps[1])
            self.__unlock_widget(self.month_combobox)

    def __set_month(self):
        month = self.month_combobox_var.get()
        choosen_month = self.months_checkup[month]
        if choosen_month != self.choosen_month.get():
            self.choosen_month.set(choosen_month)
            self.dynamic_text.set(self.instruction_steps[2])
            self.__unlock_widget(self.year_combobox)

    def __set_year(self):
        choosen_year = self.year_combobox_var.get()
        if choosen_year != self.choosen_year.get():
            self.choosen_year.set(choosen_year)
            self.dynamic_text.set(self.instruction_steps[3])
            self.__unlock_widget(self.start_bill_btn)

    # second_view
    def __init_BillingManager(self):

        bill_name = self.bill_name_var.get()
        if not bill_name:
            bill_name = "Fattura"

        self.Biller = appLib.BillingManager(bill_name, month=self.choosen_month.get(), year=self.choosen_year.get())
        self.Biller.set_badges_path(self.choosen_badge_var.get())

    def __select_all(self):
        if self.SELECT_ALL_CHKBOX_var.get() == True:
            self.SELECT_ALL_CHKBOX_var.set(1)
            for widget in self.NAMES_FRAME.winfo_children():
                if isinstance(widget, Checkbutton):
                    widget.setvar(widget.cget("variable"), True)
        elif self.SELECT_ALL_CHKBOX_var.get() == False:
            self.SELECT_ALL_CHKBOX_var.set(0)
            for widget in self.NAMES_FRAME.winfo_children():
                if isinstance(widget, Checkbutton):
                    widget.setvar(widget.cget("variable"), False)

    def __get_selected_workers(self):
        """ sets a dict of filter workers (only selected workers are used as dict keys)"""
        self.SELECTED_NAMES = []
        all_names = self.NAMES_FRAME.winfo_children()
        for index, widget in enumerate(all_names):
            if isinstance(widget, Checkbutton):
                if bool(int(widget.getvar(widget.cget("variable")))):
                    self.SELECTED_NAMES.append(all_names[index].cget("text"))

    # third_view
    def __create_billing_data(self):
        """
        upon selecting workers from previous view create three dict containing
        1. worker_name (key) : hours (value)
        2. worker_name (key) : jobs (value)
        3. worker_name (key) : billing_profiles (value)

        every value is a dict using days of the choosen month as keys and a variable value in this form
        1. a dict containing hour types during that day
        2. a string containing the id of the job done that day
        3. a string containig the id of the billing_profile to use for that day
        """

        # get total_content
        self.TOTAL_CONTENT = self.Biller.parse_badges(names=self.SELECTED_NAMES)

        self.WORKER_HOURS = self.Biller.parse_days(self.TOTAL_CONTENT)
        #days = {k: "" for k in range(calendar.monthrange(self.choosen_year.get(), self.choosen_month.get())[1])}
        days = {k:"" for k in list(self.WORKER_HOURS.items())[0][1]} # get empty days string in month from the first worker
        self.WORKER_JOBS = {worker: days for worker in self.WORKER_HOURS}
        self.WORKER_BILLING_PROFILES = {worker: days for worker in self.WORKER_HOURS}

    def __display_worker(self):
        try:
            self.DAYS_CANVAS.destroy()
        except:
            pass
        self.DAYS_CANVAS = Canvas(self.DAYS_FRAME_MASTER, bg=appLib.color_light_orange)
        self.DAYS_CANVAS.pack(side="left", fill="both", expand=True)
        self.DAYS_FRAME = Frame(self.DAYS_CANVAS, bg=appLib.color_light_orange)
        self.DAYS_FRAME.pack(anchor="center", padx=10, pady=10, fill="both", expand=True)
        self.DAYS_FRAME.bind("<Configure>",lambda e: self.DAYS_CANVAS.configure(scrollregion=self.DAYS_CANVAS.bbox("all")))


        # putting days in frame
        for day in self.WORKER_JOBS[self.SELECTED_NAMES[self.DISPLAYED_WORKER.get()]].keys():
            day_frame = Frame(self.DAYS_FRAME, bg=appLib.color_light_orange)
            day_frame.pack(anchor="center")
            Label(day_frame, text=day, bg=appLib.color_light_orange).grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

            job_name = self.Biller.get_jobname(self.WORKER_JOBS[self.SELECTED_NAMES[self.DISPLAYED_WORKER.get()]][day])
            cmb = ttk.Combobox(day_frame, state="readonly", width=20, values=self.Biller.jobs_namelist)
            cmb.set(job_name)
            cmb.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
            for col in range(day_frame.grid_size()[0]):
                if col == 0:
                    day_frame.grid_columnconfigure(col, minsize=100)
                else:
                    day_frame.grid_columnconfigure(col, minsize=350)

        try:
            self.DAYS_YSCROLL.destroy()
        except:
            pass
        self.DAYS_YSCROLL = Scrollbar(self.DAYS_FRAME_MASTER, orient="vertical", command=self.DAYS_CANVAS.yview)
        self.DAYS_YSCROLL.pack(side="right", fill="y", expand=False)
        self.DAYS_CANVAS.configure(yscrollcommand=self.DAYS_YSCROLL.set)

        self.DAYS_CANVAS.create_window((0, 0), window=self.DAYS_FRAME, anchor="nw")


        # reset widgets
        self.INFO_LBL.configure(text=f"""LAVORATORE:\t{self.SELECTED_NAMES[self.DISPLAYED_WORKER.get()]}""")
        self.SET_ALL_JOBS_CMBOX_var.set("")
        self.WEEKENDS_CHECK_var.set(False)

    def __parse_data(self):
        cached_keys = {}
        w_data = {}
        for elem in self.DAYS_FRAME.winfo_children():
            key = ""
            value= ""
            if isinstance(elem, Frame):
                for widget in elem.winfo_children():
                    if isinstance(widget, Label):
                        key = widget.cget("text")
                        if not key:
                            raise Exception("ERROR: Cannot find key in worker days!")
                    if isinstance(widget, ttk.Combobox):
                        value_ = widget.get()
                        if value_:
                            if value_ not in cached_keys:
                                for job in self.Biller.jobs:
                                    if str(job["name"]) == value_:
                                        cached_keys[value_] = job["id"]
                            try:
                                value = cached_keys[value_]
                            except:
                                raise Exception(f"ERRORE: {value_} non presente tra i Jobs")
            w_data[key] = value
        return w_data

    def __save_data(self):
        # gather data and save it to this worker
        new_data = self.__parse_data()
        w = list(self.WORKER_JOBS.keys())[self.DISPLAYED_WORKER.get()]
        self.WORKER_JOBS[w] = new_data

    def __next_worker(self):
        self.__save_data()

        current = int(self.DISPLAYED_WORKER.get())
        all_workers_len = len(self.TOTAL_CONTENT)-1
        if current < all_workers_len:
            next_w = current+1
        else:
            next_w = 0

        # reset view
        self.__clear_view(widget=self.DAYS_FRAME_MASTER)

        self.DISPLAYED_WORKER.set(next_w)
        self.__display_worker()

    def __previous_worker(self):
        self.__save_data()

        current = int(self.DISPLAYED_WORKER.get())
        all_workers_len = len(self.TOTAL_CONTENT)-1
        if current > 0:
            prev_w = current - 1
        else:
            prev_w = all_workers_len

        self.__clear_view(widget=self.DAYS_FRAME_MASTER)
        self.DISPLAYED_WORKER.set(prev_w)
        self.__display_worker()

    def __assign_all_jobs(self):

        job_to_set = self.SET_ALL_JOBS_CMBOX_var.get()
        also_weekend = self.WEEKENDS_CHECK_var.get()

        for elem in self.DAYS_FRAME.winfo_children():
            if isinstance(elem, Frame):
                ignore = False
                for widget in elem.winfo_children():

                    if isinstance(widget, Label):
                        key = widget.cget("text")
                        if not key:
                            raise Exception("ERROR: Cannot find key in worker days!")
                        current_day = int(key[:-1])
                        wd = datetime.datetime.strptime(f"{current_day}-{self.choosen_month.get()}-{self.choosen_year.get()}", "%d-%m-%Y").weekday()
                        if wd in [5,6] and not also_weekend:
                            ignore = True

                    if isinstance(widget, ttk.Combobox) and not ignore:
                        widget.set(job_to_set)

    def __confirm_and_bill(self):
        self.__save_data() # save current data
        self.WORKER_BILLING_PROFILES = self.Biller.parse_jobs_to_profiles(self.WORKER_JOBS) #create billing profiles from jobs
        self.Biller.bill(self.WORKER_HOURS, self.WORKER_JOBS, self.WORKER_BILLING_PROFILES, dump_detailed=False, dump_values=False, bill_by_job=False)
        self.Biller.bill(self.WORKER_HOURS, self.WORKER_JOBS, self.WORKER_BILLING_PROFILES, dump_detailed=False, dump_values=False, bill_by_job=True)
        messagebox.showinfo("Fatturazione conclusa", f"Documento {self.Biller.bill_name} redatto con successo")
        self.open_new_window(self.master_, Billing_Landing_Window) # returning to Billing_Landing_Window


    # views
    def __third_view(self):

        continue_process = messagebox.askokcancel("", "Una volta confermati i lavoratori non è più possibile tornare indietro, si desidera continuare?")
        if not continue_process:
            return

        default_padx = 5
        default_pady = 5

        # filter selected workers and clear view
        self.__get_selected_workers()
        self.__create_billing_data()
        self.__clear_view(self.TOP_MASTER_FRAME)

        self.DISPLAYED_WORKER = IntVar()
        self.DISPLAYED_WORKER.set(0)

        # INFO FRAME
        self.INFO_FRAME = Frame(self.TOP_MASTER_FRAME, bg=appLib.color_orange)
        self.INFO_FRAME.pack(anchor="center", padx=default_padx, pady=default_pady, fill="both")
        self.INFO_LBL = Label(self.INFO_FRAME, text=f"""LAVORATORE:\t{self.SELECTED_NAMES[self.DISPLAYED_WORKER.get()]}""", font=("Calibri", 14, "bold"), bg=appLib.color_light_orange)
        self.INFO_LBL.pack(anchor="center", fill="x", padx=default_padx, pady=(default_pady, 50))
        instruction_lbl = Label(self.INFO_FRAME,text="Indica la mansione che ho svolto nei giorni in cui ho lavorato",font=("Calibri", 14, "bold"), bg=appLib.color_orange, fg=appLib.color_light_orange)
        instruction_lbl.pack(fill="x", padx=default_padx, pady=default_pady * 2)

        # OPTIONS FRAME
        self.OPTIONS_FRAME = Frame(self.TOP_MASTER_FRAME, bg=appLib.color_light_orange)
        self.OPTIONS_FRAME.pack(anchor="center", padx=default_padx, pady=default_pady, fill="both")
        set_all_lbl = Label(self.OPTIONS_FRAME, text="Mansione", font=("Calibri", 10), bg=appLib.color_light_orange).grid(row=0, column=0)
        self.SET_ALL_JOBS_CMBOX_var = StringVar()
        self.WEEKENDS_CHECK_var = BooleanVar()
        set_all_cmbox = ttk.Combobox(self.OPTIONS_FRAME, state="readonly", textvariable=self.SET_ALL_JOBS_CMBOX_var, values=self.Biller.jobs_namelist)
        set_all_cmbox.grid(row=0, column=1, padx=default_padx, pady=default_pady, sticky="ew")
        weekend_check = Checkbutton(self.OPTIONS_FRAME, text="Sab/Dom", font=("Calibri", 10), variable=self.WEEKENDS_CHECK_var, bg=appLib.color_light_orange)
        weekend_check.grid(row=0, column=2, padx=default_padx, pady=default_pady)
        assign_btn = Button(self.OPTIONS_FRAME, text="Assegna a tutti i giorni", command=self.__assign_all_jobs).grid(row=0, column=3, padx=default_padx, pady=default_pady)
        self.OPTIONS_FRAME.grid_columnconfigure(1, weight=1)

        # DAYS FRAME
        self.DAYS_FRAME_MASTER = Frame(self.TOP_MASTER_FRAME, bg=appLib.color_light_orange)
        self.DAYS_FRAME_MASTER.pack(anchor="center", padx=10, pady=10, fill="both", expand=True)
        self.DAYS_FRAME_MASTER.pack_propagate(0)

        # display current worker
        self.__display_worker()

        # NAVBAR
        self.NAVBAR_FRAME = Frame(self.TOP_MASTER_FRAME, bg=appLib.color_orange)
        self.NAVBAR_FRAME.pack(side="top", pady=(default_pady*2, default_pady*4), padx=default_padx, fill="both")
        self.CONTINUE_BTN = Button(self.NAVBAR_FRAME, width=20, height=2, text="CONFERMA E PROCEDI", command=self.__confirm_and_bill)
        self.CONTINUE_BTN.grid(row=0, column=0, padx=(20,0), sticky="w")
        self.PREVIOUS_WORKER_BTN = Button(self.NAVBAR_FRAME, text="<<", height=2, width=4, bg=appLib.default_background, command=self.__previous_worker)
        self.PREVIOUS_WORKER_BTN.grid(row=0, column=2, padx=(0,100), sticky="e")
        self.NEXT_WORKER_BTN = Button(self.NAVBAR_FRAME, text=">>", height=2, width=4, bg=appLib.default_background, command=self.__next_worker)
        self.NEXT_WORKER_BTN.grid(row=0, column=2, padx=(0,20), sticky="e")
        for col in range(self.NAVBAR_FRAME.grid_size()[0]):
            self.NAVBAR_FRAME.grid_columnconfigure(col, weight=1)

    def __second_view(self):

        view_width = 650
        view_height = 700
        default_padx = 5
        default_pady = 5

        # check if bill name has been specified
        if not self.bill_name_var.get():
            contine_billing = messagebox.askokcancel("Campo Vuoto","Il Nome Fattura è vuoto. Si desidera continuare senza un nome fattura?")
            if not contine_billing:
                return

        # clear previous view and resize
        self.__clear_view()
        self.geometry(f"{view_width}x{view_height}+{500}+{50}")

        self.__init_BillingManager()

        # TOP_MASTER_FRAME
        self.TOP_MASTER_FRAME = Frame(self, bg=appLib.color_orange)
        self.TOP_MASTER_FRAME.pack(anchor="center", padx=self.margin, pady=self.margin, fill="both", expand=True)

        # INFO FRAME
        self.INFO_FRAME = Frame(self.TOP_MASTER_FRAME, bg=appLib.color_orange)
        self.INFO_FRAME.pack(anchor="center", padx=default_padx, pady=default_pady, fill="both")

        lbl_text = f"""'{self.bill_name_var.get()}'\tPeriodo: {self.choosen_month.get()}/{self.choosen_year.get()}"""
        info_label = Label(self.INFO_FRAME, text=lbl_text, font=("Calibri", 14, "bold"), bg=appLib.color_light_orange)
        info_label.pack(anchor="center", fill="x", padx=default_padx, pady=(default_pady, 60))

        instruction_lbl = Label(self.INFO_FRAME, text="Ho trovato questi lavoratori nei cartellini, quali devo inserire in fattura?", font=("Calibri", 14, "bold"), bg=appLib.color_orange, fg=appLib.color_light_orange)
        instruction_lbl.pack(fill="x", padx=default_padx, pady=default_pady*2)

        # SELECT ALL
        self.SELECT_ALL_CHKBOX_var = BooleanVar()
        self.SELECT_ALL_CHKBOX = Checkbutton(self.INFO_FRAME, text="Seleziona tutti", var=self.SELECT_ALL_CHKBOX_var, font=("Calibri", 12, "bold"), bg=appLib.color_orange, command=self.__select_all)
        self.SELECT_ALL_CHKBOX.pack(anchor="center", padx=default_padx)

        # NAMES FRAME
        self.NAMES_FRAME_MASTER = Frame(self.TOP_MASTER_FRAME, bg=appLib.color_light_orange)
        self.NAMES_FRAME_MASTER.pack(anchor="center", padx=default_padx*2, pady=default_pady*2, fill="both", expand=True)
        self.NAMES_FRAME_MASTER.pack_propagate(0)
        self.NAMES_CANVAS = Canvas(self.NAMES_FRAME_MASTER, bg=appLib.color_light_orange)
        self.NAMES_CANVAS.pack(side="left", fill="both", expand=True)
        self.NAMES_FRAME = Frame(self.NAMES_CANVAS, bg=appLib.color_light_orange)
        self.NAMES_FRAME.pack(anchor="center", padx=default_padx*2, pady=default_pady*2, fill="both", expand=True)
        self.NAMES_FRAME.bind("<Configure>",lambda e: self.NAMES_CANVAS.configure(scrollregion=self.NAMES_CANVAS.bbox("all")))

        self.NAMES_YSCROLL = Scrollbar(self.NAMES_FRAME_MASTER, orient="vertical", command=self.NAMES_CANVAS.yview)
        self.NAMES_YSCROLL.pack(side="right", fill="y", expand=False)
        self.NAMES_CANVAS.configure(yscrollcommand=self.NAMES_YSCROLL.set)

        self.NAMES_CANVAS.create_window((0, 0), window=self.NAMES_FRAME, anchor="nw")


        # putting names in frame
        names = self.Biller.get_all_badges_names()
        for worker in names:
            name_checkbtn = Checkbutton(self.NAMES_FRAME, text=worker, font=("Calibri", 12), bg=appLib.color_light_orange)
            name_checkbtn.pack(anchor="w", pady=default_pady, padx=default_padx)

        # CONTINUE BUTTON
        self.CONTINUE_BTN = Button(self.TOP_MASTER_FRAME, width=30, text="CONFERMA E PROCEDI", command=self.__third_view)
        self.CONTINUE_BTN.pack(anchor="center", pady=(0,default_pady*2))

    def __first_view(self):
        self.months_checkup = {
            "Gennaio":1,
            "Febbraio":2,
            "Marzo": 3,
            "Aprile": 4,
            "Maggio": 5,
            "Giugno": 6,
            "Luglio": 7,
            "Agosto": 8,
            "Settembre": 9,
            "Ottobre": 10,
            "Novembre": 11,
            "Dicembre": 12
        }
        self.instruction_steps = [
            "1. Imposta un Nome Fattura e scegli il file excel contenente i cartellini",
            "2. Seleziona il mese",
            "3. Seleziona l'anno",
            "4. Pronti a fatturare!"
        ]

        default_padx = 5
        default_pady = 5

        # define master frame
        self.MASTER_FRAME = Frame(self, bg=appLib.color_orange)
        self.MASTER_FRAME.pack(anchor="center", padx=self.margin, pady=self.margin, fill="both")

        # define master Label
        self.MASTER_INSTRUCTION = Label(self.MASTER_FRAME, text="Segui l'istruzione qui sotto per procedere con la fatturazione", font=("Calibri", 16, "bold"), wraplength=250, bg=appLib.color_orange, fg=appLib.color_light_orange)
        self.MASTER_INSTRUCTION.pack(anchor="center")

        # define dynamic instruction
        self.dynamic_text = StringVar()
        self.dynamic_text.set(self.instruction_steps[0])
        self.dynamic_instruction = Label(self.MASTER_FRAME, textvariable=self.dynamic_text, font=("Calibri", 11), bg=appLib.color_orange)
        self.dynamic_instruction.pack(anchor="center", pady=10)

        # FLOW FRAME
        self.FLOW_FRAME = Frame(self.MASTER_FRAME, bg=appLib.default_background)
        self.FLOW_FRAME.pack(anchor="center", padx=self.margin, pady=self.margin, fill="both")

        # BILL NAME
        self.BILL_NAME_FRAME = Frame(self.FLOW_FRAME, bg=appLib.default_background)
        self.BILL_NAME_FRAME.pack(side="top", padx=self.margin, pady=default_pady, fill="x")

        bill_name_lbl = Label(self.BILL_NAME_FRAME, text="Nome Fattura", bg=appLib.default_background)
        bill_name_lbl.grid(row=0, column=0, padx=default_padx, pady=default_pady, sticky="nsew")

        self.bill_name_var = StringVar()
        self.bill_name_entry = Entry(self.BILL_NAME_FRAME, textvariable=self.bill_name_var)
        self.bill_name_entry.grid(row=0, column=1, padx=default_padx, pady=default_pady, sticky="nsew")

        # BADGES FRAME
        self.BADGES_FRAME = Frame(self.FLOW_FRAME, bg=appLib.default_background)
        self.BADGES_FRAME.pack(side="top", padx=self.margin, pady=default_pady, fill="x")

        choose_badge_btn = Button(self.BADGES_FRAME, text="Cartellini", command= self.__set_badges)
        choose_badge_btn.grid(row=0, column=0, padx=default_padx, pady=default_pady, sticky="nsew")

        self.choosen_badge_var = StringVar()
        self.choosen_badge_entry = Entry(self.BADGES_FRAME, textvariable=self.choosen_badge_var, state="disabled")
        self.choosen_badge_entry.grid(row=0, column=1, padx=default_padx, pady=default_pady, sticky="nsew")

        # MONTH FRAME
        self.MONTH_FRAME = Frame(self.FLOW_FRAME, bg=appLib.default_background)
        self.MONTH_FRAME.pack(side="top", padx=self.margin, pady=default_pady, fill="x")

        month_label = Label(self.MONTH_FRAME, text="Mese", bg=appLib.default_background)
        month_label.grid(row=0, column=0, padx=default_padx, pady=default_pady, sticky="nsew")

        self.month_combobox_var = StringVar()
        self.choosen_month = IntVar()
        self.month_combobox = ttk.Combobox(self.MONTH_FRAME, textvariable=self.month_combobox_var, state="disabled", width=30) #change state to readonly when active
        self.month_combobox['values'] = [x for x in self.months_checkup.keys()]
        self.month_combobox.grid(row=0, column=1, padx=default_padx, pady=default_pady, sticky="nsew")
        self.month_combobox.bind("<<ComboboxSelected>>", lambda e: self.__set_month())

        # YEAR FRAME
        self.YEAR_FRAME = Frame(self.FLOW_FRAME, bg=appLib.default_background)
        self.YEAR_FRAME.pack(side="top", padx=self.margin, pady=default_pady, fill="x")

        year_label = Label(self.YEAR_FRAME, text="Anno", bg=appLib.default_background)
        year_label.grid(row=0, column=0, padx=default_padx, pady=default_pady, sticky="nsew")

        self.year_combobox_var = StringVar()
        self.choosen_year = IntVar()
        self.year_combobox = ttk.Combobox(self.YEAR_FRAME, state="disabled", textvariable=self.year_combobox_var, width=30)
        self.year_combobox['values'] = [(datetime.datetime.now().year)+x for x in range(5)]
        self.year_combobox.grid(row=0, column=1, padx=default_padx, pady=default_pady, sticky="nsew")
        self.year_combobox.bind("<<ComboboxSelected>>", lambda e: self.__set_year())

        # resize all grids in self.FLOW_FRAME
        for child in self.FLOW_FRAME.winfo_children():
            if isinstance(child, Frame):
                size = child.grid_size()
                for col in range(size[0]):
                    if col == 0:
                        child.grid_columnconfigure(col, minsize=100)
                    else:
                        child.grid_columnconfigure(col, weight=1)

        # START BILL BUTTON
        self.start_bill_btn = Button(self.FLOW_FRAME, text="Fattura", state="disabled", command= self.__second_view)
        self.start_bill_btn.pack(side="bottom", fill="x", pady=default_pady*2, padx=default_padx*2)
'''



if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    # change this to False only for testing purposes
    root.using_db = False

    # app entry point
    app = Home_Window(root)

    app.mainloop()