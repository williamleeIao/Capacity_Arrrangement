from openpyxl import load_workbook
from tkinter import Tk, RIGHT, BOTH, RAISED, X, N, LEFT, Text, IntVar, StringVar, BOTTOM, W, Listbox, END, Y, TOP, Menu, \
    DISABLED, NORMAL, Toplevel, E, S, SW
from tkinter.ttk import Frame, Button, Style, Label, Entry, Checkbutton, LabelFrame, Scrollbar, Radiobutton, Progressbar
from File_handle import Excel_Operation
import re


class Example(Frame):

    def file_run(self, file_path):
        path = file_path.get()
        self.value = self.excel_op.file_load(path, 'CONT Forecast', 'A', 'Z')
        for i in self.value:
            self.mylist.insert(END, i)

    def add_listbox(self):
        # get from selection and put into another one
        value = self.mylist.get(self.mylist.curselection())
        # put value into another listbox
        self.mylist1.insert(END, value)
        print(value)
        sel = self.mylist.curselection()
        for index in sel[::-1]:
            self.mylist.delete(index)

    def delete_listbox(self):
        value = self.mylist1.get(self.mylist1.curselection())
        # put value into another listbox
        self.mylist.insert(END, value)
        print(value)
        sel = self.mylist1.curselection()
        for index in sel[::-1]:
            self.mylist1.delete(index)

    def run_all(self, algorithm, value_to_skip):
        # separator ;
        # Get Value to skip and get Algorithm
        print(algorithm)
        print(value_to_skip)  # it is a list
        df = self.excel_op.run_all_rule(algorithm, list(value_to_skip), self.value)
        self.excel_op.save_new_file(df, r"C:\Users\willlee\Desktop\testing.xlsx")

    def __init__(self, root):
        super().__init__()
        self.initUI()
        self.root = root
        self.excel_op = Excel_Operation()

    def initUI(self):
        self.grid()
        self.master.title("Grid Manager")

        for r in range(10):
            self.master.rowconfigure(r, weight=1)
        for c in range(8):
            self.master.columnconfigure(c, weight=1)
            Button(self, text="Button {0}".format(c)).grid(row=6, column=c, sticky=E + W)

        Frame0 = Frame(self)
        Frame0.grid(row=0, column=0, rowspan=1, columnspan=3, sticky=W + E + N + S)
        lbl1 = Label(Frame0, text="File Path :", width=10)
        lbl1.pack(side=LEFT, padx=5, pady=5)

        file_path = StringVar()
        frame = Frame(self)
        frame.grid(row=1, column=0, rowspan=3, columnspan=7, sticky=W + E + N + S)
        entry2 = Entry(frame, textvariable=file_path)
        entry2.pack(fill=X, padx=5, expand=False)

        frameA = Frame(self)
        frameA.grid(row=1, column=7, rowspan=3, columnspan=1, sticky=W + E + N + S)
        Button_load_file = Button(frameA, text='load file', command=lambda: self.file_run(file_path))
        Button_load_file.pack(fill=X, padx=5, expand=False)

        frame_1 = Frame(self, relief=RAISED, borderwidth=1)
        frame_1.grid(row=4, column=0, rowspan=1, columnspan=13, sticky=W + E + N + S)

        Frame1 = Frame(self)
        Frame1.grid(row=5, column=0, rowspan=5, columnspan=3, sticky=W + E + N + S)
        scrollbar = Scrollbar(Frame1)
        scrollbar.pack(side=RIGHT, anchor=N, fill=Y, padx=2, pady=5)
        # Example
        self.mylist = Listbox(Frame1, yscrollcommand=scrollbar.set)
        # for line in range(200):
        #     self.mylist.insert(END, "I am EMPTY, PLease don't surprise")

        self.mylist.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.mylist.yview)

        Frame3 = Frame(self)
        Frame3.grid(row=5, column=3, rowspan=2, columnspan=2, sticky=W + E + N + S)
        AddButton = Button(Frame3, text="--> Add -->", command=lambda: self.add_listbox())
        AddButton.pack(side=TOP, padx=5, pady=5)

        DeleteButton = Button(Frame3, text="<-- Delete <--", command=lambda: self.delete_listbox())
        DeleteButton.pack(side=BOTTOM, padx=5, pady=5)

        Frame4 = Frame(self)
        Frame4.grid(row=5, column=5, rowspan=6, columnspan=13, sticky=W + E + N + S)
        scrollbar = Scrollbar(Frame4)
        scrollbar.pack(side=RIGHT, anchor=N, fill=Y, padx=5, pady=5)
        # Example
        self.mylist1 = Listbox(Frame4, yscrollcommand=scrollbar.set)
        # for line in range(200):
        #     self.mylist1.insert(END, "I am EMPTY, PLease don't surprise")
        self.mylist1.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.config(command=self.mylist1.yview)

        Frame5 = Frame(self)
        Frame5.grid(row=17, column=0, rowspan=3, columnspan=2, sticky=W + E + N + S)
        lbl2 = Label(Frame5, text="Algorithm :", width=10)
        lbl2.pack(side=LEFT, padx=5, pady=5)

        algorithm = StringVar()
        Frame6 = Frame(self)
        Frame6.grid(row=20, column=0, rowspan=8, columnspan=13, sticky=W + E + N + S)
        entry2 = Entry(Frame6, textvariable=algorithm)
        entry2.pack(fill=X, padx=5, expand=False)

        Frame10 = Frame(self)
        Frame10.grid(row=40, column=0, rowspan=3, columnspan=13, sticky=W + E + N + S)
        lbl2 = Label(Frame10, text="File_Destination_Saving :", width=25)
        lbl2.pack(side=LEFT, padx=5, pady=5)

        save_file_path = StringVar()
        Frame9 = Frame(self)
        Frame9.grid(row=46, column=0, rowspan=8, columnspan=13, sticky=W + E + N + S)
        entry2 = Entry(Frame9, textvariable=save_file_path)
        entry2.pack(fill=X, padx=5, expand=False)

        frame7 = Frame(self, relief=RAISED, borderwidth=1)
        frame7.grid(row=60, column=0, rowspan=1, columnspan=13, sticky=W + E + N + S)
        #
        self.pack(fill=BOTH, expand=True)
        #
        Frame8 = Frame(self, relief=RAISED, borderwidth=1)
        Frame8.grid(row=97, column=0, rowspan=1, columnspan=13, sticky=W + E + N + S)
        RunButton = Button(Frame8, text="<--> Run <-->", command=lambda: self.run_all(algorithm.get(),
                                                                                      self.mylist1.get(0,
                                                                                                       self.mylist1.size() - 1)))
        RunButton.pack(side=RIGHT, padx=5, pady=5)


# ----------------Generate function must be a top-level module function---------------------------------------


def main():
    root = Tk()
    root.geometry("650x380")
    root.resizable(0, 0)
    app = Example(root)
    root.mainloop()


if __name__ == '__main__':
    main()
