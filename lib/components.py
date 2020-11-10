from lib import appLib
import os
import uuid
import hashlib
import json
from threading import Thread
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import mysql.connector

class Custom_Toplevel(Toplevel):
    def __init__(self, master=None, width=100, height=100):
        super().__init__(master)

        self.width = width
        self.height = height
        self.__screen_width = master.winfo_screenwidth()
        self.__screen_height = master.winfo_screenheight()
        self.geometry(f"{self.width}x{self.height}+{int((self.__screen_width / 2 - self.width / 2))}+{int((self.__screen_height / 2 - self.height / 2))}")
        self.title("BusinessCat - PDF Splitter")
        self.iconbitmap(appLib.icon_path)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", lambda: self.close(master))

    def close(self, master):
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
        self.image = Image.open(appLib.logo_path)
        self.image = ImageTk.PhotoImage(self.image)
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
        self.signup_label.bind('<Button 1>', lambda event: self.signup(master))

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
                                self.destroy()
                                Splitter_Window(master)
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
                self.destroy()
                Splitter_Window(master)

    def signup(self, master):
        self.destroy()
        Register_Window(master)

class Register_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.height = 500
        self.width = 500
        super().__init__(master, self.width, self.height)

        self.config(bg=appLib.default_background)
        self.title(self.title().split("-")[:1][0] + " - " + "Registrazione")
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
        self.paw_image = ImageTk.PhotoImage(self.paw_image)
        self.back_paw_image = Label(self.back_button_frame, image=self.paw_image, bg=appLib.default_background, borderwidth=0)
        self.back_paw_image.grid(row=1)
        self.bind('<Button 1>', lambda event: self.back_to_login(master))

    def back_to_login(self, master):
        self.destroy()
        Login_Window(master)

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
                    config = appLib.load_config()
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
                        self.destroy()
                        Login_Window(master)

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

class Splitter_Window(Custom_Toplevel):
    def __init__(self, master=None):
        self.width = 500
        self.height = 420
        super().__init__(master, self.width, self.height)


        self.done_paycheck = False
        self.done_badges = False

        # define Canvas
        self.canvas = Canvas(self, width = self.width, height = self.height, bg="white")
        self.canvas.grid()

        # define Logo
        self.logo_image = PhotoImage(file=appLib.logo_path)
        self.canvas.create_image((self.width/2), 80, image=self.logo_image)

        # status_circle
        raggio = 10
        x_offset = -5    # tweak x offset to move circle horizontally
        y_offset = 15   #tweak y offset to move circle vertically
        top_left_coord = [(self.width/2) - raggio*2, (self.height/4)*3]
        bottom_right_coord = [(self.width/2) + raggio*2, (self.height/4)*3 + raggio*4]
        top_left_coord[0] += x_offset
        bottom_right_coord[0] += x_offset
        top_left_coord[1] += y_offset
        bottom_right_coord[1] += y_offset
        self.status_circle = self.canvas.create_oval(top_left_coord, bottom_right_coord, outline="", fill=appLib.color_green)

        # define credits
        credits = "BusinessCat - Developed by CSI Centro Servizi Industriali - Broni (PV) Italy"
        self.canvas.create_text((self.width/2), (self.height - 15), font="Calibri 10 bold", text=credits)



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
        button_explore = Button(self.canvas, text = "Scegli il file dei Cartellini", width = 20, height = 1, command = lambda:self.changeContent(self.badges_textbox))
        self.canvas.create_window((self.width/4), badges_heights, window=button_explore)



        ################################################# START /CLEAR BUTTONS
        command_buttons_height = 295

        # define START button
        button_START = Button(self.canvas, text = "Dividi", width = 10, height = 1, command=self.START)
        self.canvas.create_window((self.width/4 + 50), command_buttons_height, window=button_START)

        # define CLEAR button
        button_CLEAR = Button(self.canvas, text = "Pulisci", width = 10, height = 1, command=self.clearContent)
        self.canvas.create_window(((self.width/4 + 200)), command_buttons_height, window=button_CLEAR)

    def changeContent(self, txtbox):
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select a File",
                                              filetype=[("PDF files", "*.pdf*")])
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

        # check if directories exist
        if os.path.exists("BUSTE PAGA"):
            if len(os.listdir("BUSTE PAGA")) > 0:
                self.done_paycheck = True
        if os.path.exists("CARTELLINI"):
            if len(os.listdir("CARTELLINI")) > 0:
                self.done_badges = True

        # se ho già generato i file printo errore
        if (self.done_paycheck and paycheck_path) and (self.done_badges and badges_path):
            messagebox.showerror("Error", "File già divisi!")
        elif (self.done_paycheck and paycheck_path) and not (self.done_badges and badges_path):
            messagebox.showerror("Error", "Buste già divise!")
        elif (self.done_badges and badges_path) and not (self.done_paycheck and paycheck_path):
            messagebox.showerror("Error", "Cartellini già divisi!")

        else:
            # imposto lo stato di lavoro dello status_circle
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_yellow)

            if paycheck_path and not self.done_paycheck:
                try:
                    appLib.CREATE_BUSTE(paycheck_path, "BUSTE PAGA")
                    self.paycheck_textbox.configure(disabledbackground=appLib.color_green)
                    done_paycheck = True
                except Exception as e:
                    self.paycheck_textbox.configure(disabledbackground=appLib.color_red)
                    self.canvas.itemconfig(self.status_circle, fill=appLib.color_red)
                    e = str(e).replace("#", "")
                    messagebox.showerror("Error", e)

            if badges_path and not self.done_badges:
                try:
                    if os.path.exists("CARTELLINI"):
                        if len(os.listdir("CARTELLINI")) == 0:
                            appLib.CREATE_CARTELLINI(badges_path, "CARTELLINI")
                            self.badges_textbox.configure(disabledbackground=appLib.color_green)
                            done_badges = True
                    else:
                        appLib.CREATE_CARTELLINI(badges_path, "CARTELLINI")
                        self.badges_textbox.configure(disabledbackground=appLib.color_green)
                        done_badges = True
                except Exception as e:
                    self.badges_textbox.configure(disabledbackground=appLib.color_red)
                    self.canvas.itemconfig(self.status_circle, fill=appLib.color_red)
                    e = str(e).replace("#", "")
                    messagebox.showerror("Error", e)

            # reimposto lo stato di lavoro dello status_circle
            self.canvas.itemconfig(self.status_circle, fill=appLib.color_green)

            # end messages
            if paycheck_path and self.done_paycheck and not self.done_badges:
                messagebox.showinfo("Done", "Buste divise con successo!")
            elif badges_path and self.done_badges and not self.done_paycheck:
                messagebox.showinfo("Done", "Cartellini divisi con successo!")
            elif (self.done_paycheck and self.done_badges) and paycheck_path or badges_path:
                messagebox.showinfo("Done", "File divisi con successo!")

            done_paycheck = False
            done_badges = False

    def START(self):
        # creating a thread preventing the program to freeze during the loop
        Thread(target=self.splitting).start()



if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    # change this to False only for testing purposes
    root.using_db = False

    # app entry point
    app = Register_Window(root)
    app.mainloop()