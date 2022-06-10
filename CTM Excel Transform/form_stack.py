from tkinter import ttk,Tk,StringVar,Label,Button,E,W,N,S,Frame
from tkinter.filedialog import askopenfilename
from win32com.shell import shell, shellcon
from transform import transform

class FST:

    def __init__(self, master):

        master.minsize(width=400, height=100)

        # create title
        master.title("Formstack Transform")

        # create import raw_data button and label
        self.raw_label_text = StringVar()
        self.raw_label_text.set("")
        self.raw_label = Label(master, textvariable=self.raw_label_text)
        self.raw = Button(master, text="Select Import Spreadsheet", command= lambda: self.OpenFile(self.raw_label_text))
        self.raw.grid(sticky=E,pady=5)
        self.raw_label.grid(row=0,column=1,pady=5)

        # add padding between rows
        self.separator = Frame(height=10, bd=1)
        self.separator.grid(row=1,columnspan=2,sticky=W+E)

        # create load master data spreadsheet
        self.data_label_text = StringVar()
        self.data_label_text.set("")
        self.data_label = Label(master, textvariable=self.data_label_text)
        self.data = Button(master, text="Select Update Spreadsheet", command= lambda: self.OpenFile(self.data_label_text))
        self.data.grid(sticky=E,pady=5)
        self.data_label.grid(row=2,column=1,pady=5)

        # create Transform button
        self.transform = Button(master, text="Update Spreadsheet", command= lambda: transform(self.raw_label_text.get(),self.data_label_text.get()))
        self.transform.grid(row=3,column=2,columnspan=2,sticky=W+E+N+S,ipadx=10,padx=5,pady=5)
        

    def OpenFile(self, textvar):
        name = askopenfilename(initialdir=shell.SHGetFolderPath(0, shellcon.CSIDL_PERSONAL, None, 0),
                                filetypes =(("Excel File", "*.xlsx"), ("All", "*.*")),
                                title = "Choose a file."
                                )
        textvar.set(name)
        #Using try in case user types in unknown file or closes without choosing a file.
        try:
            with open(name,'r') as UseFile:
                print(UseFile.read())
        except:
            pass

root = Tk()
app = FST(root)

root.mainloop()
