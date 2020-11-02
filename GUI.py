import main
import os
from threading import Thread
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

done_paycheck = False
done_badges = False
logo_path = "config_files/imgs/BusinessCat.png"
color_green = "#80ff80"
color_yellow = "#ffeb64"
color_red = "#ff4d4d"

def changeContent(txtbox):
    filename = filedialog.askopenfilename(initialdir = os.getcwd(), title = "Select a File", filetype = [("PDF files","*.pdf*")])
    txtbox.configure(state='normal')
    txtbox.delete(0,"end")
    txtbox.insert(0, filename)
    txtbox.configure(state='disabled')

def clearContent():
    paycheck_textbox.configure(state='normal')
    badges_textbox.configure(state='normal')
    paycheck_textbox.delete(0,"end")
    badges_textbox.delete(0,"end")
    paycheck_textbox.configure(state='disabled', disabledbackground=root.cget('bg'))
    badges_textbox.configure(state='disabled', disabledbackground=root.cget('bg'))

def splitting():
    paycheck_path = paycheck_textbox.get()
    badges_path = badges_textbox.get()
    global done_paycheck
    global done_badges

    #check if directories exist
    if os.path.exists("BUSTE PAGA"):
        if len(os.listdir("BUSTE PAGA")) > 0:
            done_paycheck = True
    if os.path.exists("CARTELLINI"):
        if len(os.listdir("CARTELLINI")) > 0:
            done_badges = True

    # se ho già generato i file printo errore
    if (done_paycheck and paycheck_path) and (done_badges and badges_path):
        messagebox.showerror("Error", "File già divisi!")
    elif (done_paycheck and paycheck_path) and not (done_badges and badges_path):
        messagebox.showerror("Error", "Buste già divise!")
    elif (done_badges and badges_path) and not (done_paycheck and paycheck_path):
        messagebox.showerror("Error", "Cartellini già divisi!")

    else:
        # imposto lo stato di lavoro dello status_circle
        canvas.itemconfig(status_circle, fill=color_yellow)

        if paycheck_path and not done_paycheck:
            try:
                main.CREATE_BUSTE(paycheck_path, "BUSTE PAGA")
                paycheck_textbox.configure(disabledbackground=color_green)
                done_paycheck = True
            except Exception as e:
                paycheck_textbox.configure(disabledbackground=color_red)
                canvas.itemconfig(status_circle, fill=color_red)
                e = str(e).replace("#", "")
                messagebox.showerror("Error", e)

        if badges_path and not done_badges:
            try:
                if os.path.exists("CARTELLINI"):
                    if len(os.listdir("CARTELLINI")) == 0:
                        main.CREATE_CARTELLINI(badges_path, "CARTELLINI")
                        badges_textbox.configure(disabledbackground=color_green)
                        done_badges = True
                else:
                    main.CREATE_CARTELLINI(badges_path, "CARTELLINI")
                    badges_textbox.configure(disabledbackground=color_green)
                    done_badges = True
            except Exception as e:
                badges_textbox.configure(disabledbackground=color_red)
                canvas.itemconfig(status_circle, fill=color_red)
                e = str(e).replace("#", "")
                messagebox.showerror("Error", e)

        # reimposto lo stato di lavoro dello status_circle
        canvas.itemconfig(status_circle, fill=color_green)

        # end messages
        if paycheck_path and done_paycheck and not done_badges:
            messagebox.showinfo("Done", "Buste divise con successo!")
        elif badges_path and done_badges and not done_paycheck:
            messagebox.showinfo("Done", "Cartellini divisi con successo!")
        elif (done_paycheck and done_badges) and paycheck_path or badges_path:
            messagebox.showinfo("Done", "File divisi con successo!")

        done_paycheck = False
        done_badges = False

def START():
    # creating a thread preventing the program to freeze during the loop
    Thread(target=splitting).start()

# root
root = Tk()
root_width = 500
root_height = 420
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
root.title("BusinessCat - PDF Splitter")
root.iconbitmap("config_files/imgs/Cat.ico")
root.geometry(f"{root_width}x{root_height}+{int((screen_width/2 - root_width/2))}+{int((screen_height/2 - root_height/2))}")
root.resizable(False, False)
root.margin = 15


############################################## GRAPHICS

# define Canvas
canvas = Canvas(root, width = root_width, height = root_height, bg="white")
canvas.grid()

# define Logo
logo_image = PhotoImage(file=logo_path)
canvas.create_image((root_width/2), 80, image=logo_image)

# status_circle
raggio = 10
x_offset = -5    # tweak x offset to move circle horizontally
y_offset = 15   #tweak y offset to move circle vertically
top_left_coord = [(root_width/2) - raggio*2, (root_height/4)*3]
bottom_right_coord = [(root_width/2) + raggio*2, (root_height/4)*3 + raggio*4]
top_left_coord[0] += x_offset
bottom_right_coord[0] += x_offset
top_left_coord[1] += y_offset
bottom_right_coord[1] += y_offset
status_circle = canvas.create_oval(top_left_coord, bottom_right_coord, outline="", fill=color_green)

# define credits
credits = "BusinessCat - Developed by CSI Centro Servizi Industriali - Broni (PV) Italy"
canvas.create_text((root_width/2), (root_height - root.margin), font="Calibri 10 bold", text=credits)



############################################### BROWSE FILES (TXTBOXES & BUTTONS)
paycheck_heights = 180
badges_heights = 240

# define Paycheck Textbox
paycheck_textbox = Entry(canvas, width = 40, state=DISABLED)
canvas.create_window((root_width/2) + 95, paycheck_heights, window=paycheck_textbox)

# define Badges Textbox
badges_textbox = Entry(canvas, width = 40, state=DISABLED)
canvas.create_window((root_width/2) + 95, badges_heights, window=badges_textbox)

# define Browse Paycheck
button_explore = Button(canvas, text = "Scegli il file delle Buste", width = 20, height = 1, command = lambda:changeContent(paycheck_textbox))
canvas.create_window((root_width/4), paycheck_heights, window=button_explore)

# define Browse Badges
button_explore = Button(canvas, text = "Scegli il file dei Cartellini", width = 20, height = 1, command = lambda:changeContent(badges_textbox))
canvas.create_window((root_width/4), badges_heights, window=button_explore)



################################################# START /CLEAR BUTTONS
command_buttons_height = 295

# define START button
button_START = Button(canvas, text = "Dividi", width = 10, height = 1, command=START)
canvas.create_window((root_width/4 + 50), command_buttons_height, window=button_START)

# define CLEAR button
button_CLEAR = Button(canvas, text = "Pulisci", width = 10, height = 1, command=clearContent)
canvas.create_window(((root_width/4 + 200)), command_buttons_height, window=button_CLEAR)


root.mainloop()