from lib.components import *


if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    # change this to False only for testing purposes
    root.using_db = True

    # app entry point
    app = Login_Window(root)
    app.mainloop()


