from openpyxl import load_workbook
from tkinter import Tk, RIGHT, BOTH, RAISED, X, N, LEFT, Text, IntVar, StringVar, BOTTOM, W, Listbox, END, Y, TOP, Menu, \
    DISABLED, NORMAL, Toplevel, E, S, SW
from tkinter.ttk import Frame, Button, Style, Label, Entry, Checkbutton, LabelFrame, Scrollbar, Radiobutton, Progressbar


class Example(Frame):

    def __init__(self, root):
        super().__init__()
        self.initUI()
        self.root = root

    def initUI(self):
        self.grid()
        self.master.title("Grid Manager")

        for r in range(10):
            self.master.rowconfigure(r, weight=1)
        for c in range(8):
            self.master.columnconfigure(c, weight=1)
            Button(self, text="Button {0}".format(c)).grid(row=6, column=c, sticky=E + W)

        Frame1 = Frame(self)
        Frame1.grid(row=0, column=0, rowspan=5, columnspan=3, sticky=W + E + N + S)
        scrollbar = Scrollbar(Frame1)
        scrollbar.pack(side=RIGHT, anchor=N, fill=Y, padx=2, pady=5)
        # Example
        self.mylist = Listbox(Frame1, yscrollcommand=scrollbar.set)
        for line in range(200):
            self.mylist.insert(END, "I am EMPTY, PLease don't surprise")

        self.mylist.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.mylist.yview)

        Frame3 = Frame(self)
        Frame3.grid(row=0, column=3, rowspan=2, columnspan=2, sticky=W + E + N + S)
        AddButton = Button(Frame3, text="--> Add -->")
        AddButton.pack(side=TOP, padx=5, pady=5)

        DeleteButton = Button(Frame3, text="<-- Delete <--")
        DeleteButton.pack(side=BOTTOM, padx=5, pady=5)

        Frame4 = Frame(self)
        Frame4.grid(row=0, column=5, rowspan=6, columnspan=13, sticky=W + E + N + S)
        scrollbar = Scrollbar(Frame4)
        scrollbar.pack(side=RIGHT, anchor=N, fill=Y, padx=5, pady=5)
        # Example
        self.mylist1 = Listbox(Frame4, yscrollcommand=scrollbar.set)
        for line in range(200):
            self.mylist1.insert(END, "I am EMPTY, PLease don't surprise")
        self.mylist1.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.mylist1.yview)

        Frame5 = Frame(self)
        Frame5.grid(row=15, column=0, rowspan=3, columnspan=2, sticky=W + E + N + S)
        lbl2 = Label(Frame5, text="Algorithm :", width=10)
        lbl2.pack(side=LEFT, padx=5, pady=5)

        algorithm = StringVar()
        Frame6 = Frame(self)
        Frame6.grid(row=18, column=0, rowspan=8, columnspan=13, sticky=W + E + N + S)
        entry2 = Text(Frame6, height=8)
        entry2.pack(fill=X, padx=5, expand=False)

        frame7 = Frame(self, relief=RAISED, borderwidth=1)
        frame7.grid(row=26, column=0, rowspan=1, columnspan=13, sticky=W + E + N + S)
        #
        self.pack(fill=BOTH, expand=True)
        #
        Frame8 = Frame(self, relief=RAISED, borderwidth=1)
        Frame8.grid(row=27, column=0, rowspan=1, columnspan=13, sticky=W + E + N + S)
        RunButton = Button(Frame8, text="<--> Run <-->")
        RunButton.pack(side=RIGHT, padx=5, pady=5)


# ----------------Generate function must be a top-level module function---------------------------------------


def main():
    root = Tk()
    root.geometry("650x400")
    root.resizable(0, 0)
    app = Example(root)
    root.mainloop()


if __name__ == '__main__':
    main()
