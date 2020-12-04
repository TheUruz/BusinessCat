from lib.components import *


if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    # change this to False only for testing purposes
    root.using_db = False

    # defining entry point
    if root.using_db:
        app = Login_Window(root)
    else:
        app = Home_Window(root)

    app.mainloop()


