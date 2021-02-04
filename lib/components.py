from lib import appLib
import PIL.Image, PIL.ImageTk
import os
import copy
import uuid
import json
import pandas as pd
from pandastable import Table, TableModel, config as tab_config
from email.message import EmailMessage
import mimetypes
import smtplib
from tkinter import *
from ttkwidgets.frames import Balloon
from tkinter import filedialog
from tkinter import messagebox

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
        self.widget_balloon = Balloon(master=widget, text=text, timeout=1, height=100)


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
        self.width = 500
        self.height = 420
        super().__init__(master, self.width, self.height)

        self.title(self.title().split("-")[:1][0] + " - " + "PDF Splitter")

        # verify paycheck and badges status
        self.done_paycheck, self.done_badges = appLib.check_paycheck_badges()

        # define Canvas
        self.canvas = Canvas(self, width = self.width, height = self.height, bg="white")
        self.canvas.grid()

        # define back button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, Home_Window))
        self.canvas.create_window(30,30, window=self.back_button)

        # define Logo
        self.logo_image = PhotoImage(file=appLib.logo_path)
        self.canvas.create_image((self.width/2), 80, image=self.logo_image)

        # status_circle
        raggio = 10
        x_offset = 0    # tweak x offset to move circle horizontally
        y_offset = 15   #tweak y offset to move circle vertically
        top_left_coord = [(self.width/2) - raggio*2, (self.height/4)*3]
        bottom_right_coord = [(self.width/2) + raggio*2, (self.height/4)*3 + raggio*4]
        top_left_coord[0] += x_offset
        bottom_right_coord[0] += x_offset
        top_left_coord[1] += y_offset
        bottom_right_coord[1] += y_offset
        self.status_circle = self.canvas.create_oval(top_left_coord, bottom_right_coord, outline="", fill=appLib.color_green)

        # define status label
        self.status_label = self.canvas.create_text((self.width/2), (self.height - 30), text="Pronto", fill=appLib.color_green, font=("Calibri",12,"bold"))


        ############################################### BROWSE FILES (TXTBOXES & BUTTONS)
        paycheck_heights = 180
        badges_heights = 240

        # define Paycheck Textbox
        self.paycheck_textbox = Entry(self.canvas, width = 40, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.canvas.create_window((self.width/2) + 95, paycheck_heights, window=self.paycheck_textbox)

        # define Badges Textbox
        self.badges_textbox = Entry(self.canvas, width = 40, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.canvas.create_window((self.width/2) + 95, badges_heights, window=self.badges_textbox)

        # define Browse Paycheck
        self.button_explore = Button(self.canvas, text = "Scegli il file delle Buste", width = 20, height = 1, command = lambda:self.changeContent(self.paycheck_textbox))
        self.canvas.create_window((self.width/4), paycheck_heights, window=self.button_explore)

        # define Browse Badges
        self.button_explore = Button(self.canvas, text = "Scegli il file dei Cartellini", width = 20, height = 1, command = lambda:self.changeContent(self.badges_textbox))
        self.canvas.create_window((self.width/4), badges_heights, window=self.button_explore)



        ################################################# START /CLEAR BUTTONS
        command_buttons_height = 290

        # define START button
        self.button_START = Button(self.canvas, text = "Dividi", width = 12, height = 1, command=lambda:appLib.START_in_Thread(self.splitting))
        self.canvas.create_window((self.width/5), command_buttons_height, window=self.button_START)

        # define CLEAR button
        self.button_CLEAR = Button(self.canvas, text = "Pulisci", width = 12, height = 1, command=self.clearContent)
        self.canvas.create_window((self.width/5 *2.5), command_buttons_height, window=self.button_CLEAR)

        # define SEND button
        self.button_SEND = Button(self.canvas, text="Invia per Email", width=12, height=1, state=DISABLED,command=lambda: self.open_new_window(master, Mail_Sender_Window))
        self.canvas.create_window((self.width/5 *4), command_buttons_height, window=self.button_SEND)
        self.check_send_mail()

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

    def splitting(self):
        paycheck_path = self.paycheck_textbox.get()
        badges_path = self.badges_textbox.get()

        # verify paycheck and badges status
        self.done_paycheck, self.done_badges = appLib.check_paycheck_badges()

        # se ho già generato i file printo errore
        if (self.done_paycheck and paycheck_path) and (self.done_badges and badges_path):
            messagebox.showerror("Error", "File già divisi!")
        elif (self.done_paycheck and paycheck_path) and not (self.done_badges and badges_path):
            messagebox.showerror("Error", "Buste già divise!")
        elif (self.done_badges and badges_path) and not (self.done_paycheck and paycheck_path):
            messagebox.showerror("Error", "Cartellini già divisi!")

        else:
            # imposto lo stato di lavoro dello status_circle e della status_label
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_yellow)
            self.canvas.itemconfig(self.status_label, fill=appLib.color_yellow, text="Sto Dividendo...")

            if paycheck_path and not self.done_paycheck:
                try:
                    appLib.CREATE_BUSTE(paycheck_path, "BUSTE PAGA")
                    self.paycheck_textbox.configure(disabledbackground=appLib.color_green)
                    self.done_paycheck = True
                except Exception as e:
                    self.paycheck_textbox.configure(disabledbackground=appLib.color_red)
                    self.canvas.itemconfig(self.status_circle, fill=appLib.color_red)
                    self.canvas.itemconfig(self.status_label, fill=appLib.color_red, text="ERRORE")
                    e = str(e).replace("#", "")
                    messagebox.showerror("Error", e)

            if badges_path and not self.done_badges:
                try:
                    appLib.CREATE_CARTELLINI(badges_path, "CARTELLINI")
                    self.badges_textbox.configure(disabledbackground=appLib.color_green)
                    self.done_badges = True
                except Exception as e:
                    self.badges_textbox.configure(disabledbackground=appLib.color_red)
                    self.canvas.itemconfig(self.status_circle, fill=appLib.color_red)
                    self.canvas.itemconfig(self.status_label, fill=appLib.color_red, text="ERRORE")
                    e = str(e).replace("#", "")
                    messagebox.showerror("Error", e)

            # reimposto lo stato di lavoro dello status_circle
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_green)
            self.canvas.itemconfig(self.status_label, fill=appLib.color_green, text="Pronto")

            # end messages
            if paycheck_path and self.done_paycheck and not self.done_badges:
                messagebox.showinfo("Done", "Buste divise con successo!")
            elif badges_path and self.done_badges and not self.done_paycheck:
                messagebox.showinfo("Done", "Cartellini divisi con successo!")
            elif (self.done_paycheck and self.done_badges) and paycheck_path or badges_path:
                messagebox.showinfo("Done", "File divisi con successo!")

        # check mail send button
        self.check_send_mail()

    def check_send_mail(self):
        if self.done_badges and self.done_paycheck:
            self.button_SEND.config(state=ACTIVE)
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
        self.width = 550
        self.height = 700
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Verify Paychecks")
        self.margin = 15

        self.radio_buttons_val = IntVar()

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
        self.grid_rowconfigure(9, minsize=self.margin * 3)  # empty
        self.grid_rowconfigure(10, minsize=self.margin)  # Gdrive logo and file label

        ############### COLUMN CONFIGURE
        self.grid_columnconfigure(0, minsize=self.margin) # empty
        self.grid_columnconfigure(1, minsize=self.margin) # back button
        self.grid_columnconfigure(2, minsize=self.margin) # window start from here
        self.grid_columnconfigure(3, minsize=self.margin) # empty
        self.grid_columnconfigure(4, minsize=self.margin) # textboxes column
        self.grid_columnconfigure(5, minsize=self.margin*2) # data column
        self.grid_columnconfigure(6, minsize=self.margin)  # empty

        # define << back Button
        self.back_button = Button(self, text="<<", width=2, height=1, command= lambda:self.open_new_window(master, self.prior_window))
        self.back_button.grid(row=1, column=1, sticky="w")

        # define import configuration Button
        self.import_config_button = Button(self, width=20, text="Importa Configurazione")
        self.import_config_button.grid(row=3, column=2)

        # define config Textbox
        self.config_txtbox = Entry(self, width=50, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.config_txtbox.grid(row=3, column=4, columnspan=2)

        # define file select Button
        self.check_file_button = Button(self, width=20, text="File da Controllare")
        self.check_file_button.grid(row=5, column=2)

        # define file_to_check Textbox
        self.check_file_txtbox = Entry(self, width=50, state=DISABLED, disabledbackground=appLib.color_light_orange)
        self.check_file_txtbox.grid(row=5, column=4, columnspan=2)

        # define where_badges Label
        self.where_badges_label = Label(self, text="Dove sono i Cartellini?", font=("Calibri", 20, "bold"), fg=appLib.color_orange, bg=appLib.default_background)
        self.where_badges_label.grid(row=7, column=2, columnspan=4)

        # define badges Radio Buttons
        self.radio_BusinessCat = Radiobutton(self, text="Cartella di BusinessCat", font=("Calibri", 12), padx=20, variable=self.radio_buttons_val, value=1, bg=appLib.default_background)
        self.radio_BusinessCat.grid(row=8, column=2, columnspan=4, sticky="w")
        self.radio_BusinessCat.select() # selected by default

        self.radio_Other = Radiobutton(self, text="Altro (Seleziona)", font=("Calibri", 12), padx=20, variable=self.radio_buttons_val, value=2, bg=appLib.default_background)
        self.radio_Other.grid(row=8, column=2, columnspan=4, sticky="e")

        # define google drive img Label
        self.gdrive_img_label = Label(self, width=4, height=2, bg=appLib.color_green)
        self.gdrive_img_label.grid(row=10, column=1, columnspan=3, sticky="e")

        # define choosen file Label
        self.choosen_file_label = Label(self, text="File da Comparare", font=("Calibri", 12, "bold"), bg=appLib.default_background)
        self.choosen_file_label.grid(row=10, column=5, columnspan=2, sticky="nsew")




if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    # change this to False only for testing purposes
    root.using_db = False

    # app entry point
    #app = Home_Window(root)
    app = Verificator_Window(root)
    app.mainloop()