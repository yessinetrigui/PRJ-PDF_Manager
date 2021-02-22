import PyPDF2
import os
import tkinter as tk
from tkinter import ttk, Tk, messagebox
import ctypes
from PIL import ImageTk, Image
from tkinter import filedialog
import docx2pdf
class APP:
    height = ''
    width = ''
    Wscreen = ''
    Hscreen = ''

    def __init__(self, width, height, Wscreen, Hscreen):
        FinalW = (Wscreen - width) / 2
        FinalH = (Hscreen - height) / 2
        self.frame = frame
        self.frame.geometry('%dx%d+%d+%d' % (width, height, FinalW, FinalH))
        self.frame.title('PDF MANAGER')
        self.frame.resizable(0, 0)
        self.frame.state('normal')

        # ICO FOR THE APP
        self.frame.iconbitmap(default='DATA/IMG/Pdf.ico')
        myappid = 'mycompany.myproduct.subproduct.version'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        # BACKGROUND FOR APP
        bgpath = 'DATA/IMG/bg.png'
        bg = ImageTk.PhotoImage(Image.open(bgpath))
        self.bg = ttk.Label(self.frame, image=bg)
        self.bg.image = bg
        self.bg.place(x=-2, y=0)
        # BUTTONS OF APP
        ph1 = tk.PhotoImage(file='Data/IMG/bumerg.png')
        bu1 = ttk.Button(self.frame, image=ph1)
        bu1.image = ph1
        bu1.config(command=self.PDF_MRG)
        bu1.place(x=30, y=38)
        ph2 = tk.PhotoImage(file='Data/IMG/buconv.png')
        bu2 = ttk.Button(self.frame, image=ph2)
        bu2.image = ph2
        bu2.config(command=self.PDF_CONV)
        bu2.place(x=200, y=38)


    def PDF_MRG(self):
        self.wind = tk.Toplevel()
        self.frame.overrideredirect(True)
        self.wind.protocol("WM_DELETE_WINDOW", self.update)
        self.wind.grab_set()
        Wscreen = self.wind.winfo_screenwidth()
        Hscreen = self.wind.winfo_screenheight()
        width = 350
        height = 280
        FinalW = (Wscreen - width) / 2
        FinalH = (Hscreen - height) / 2
        FinalH -=161
        FinalW -= 8
        self.wind.geometry('%dx%d+%d+%d' % (width, height, FinalW, FinalH))
        self.wind.resizable(0, 0)
        self.wind.title('PDF Merger')
        # lbl 1 (img)
        path = 'DATA/IMG/pdfmerg.png'
        img1 = ImageTk.PhotoImage(Image.open(path))
        lbl1 = ttk.Label(self.wind, image=img1)
        lbl1.image = img1
        lbl1.pack()
        path2 = 'DATA/IMG/AddF.png'
        img2 = ImageTk.PhotoImage(Image.open(path2))
        lbl2 = ttk.Label(self.wind, image=img2)
        lbl2.image = img2
        lbl2.pack()
        # the treeview of all pdf files added
        self.treev = ttk.Treeview(self.wind, selectmode='browse', height=6)
        # Calling pack
        self.treev.pack()
        # Constructing vertical scrollbar
        # with treeview
        verscrlbar = ttk.Scrollbar(self.wind, orient="vertical", command=self.treev.yview)
        # Calling pack method w.r.to verical
        # scrollbar
        verscrlbar.place(x=30, y=110)
        # Configuring treeview
        self.treev.configure(xscrollcommand=verscrlbar.set)
        # Defining number of columns
        self.treev["columns"] = ("1", "2")
        # Defining heading
        self.treev['show'] = 'headings'
        # Assigning the width and anchor to  the
        # respective columns
        self.treev.column("1", width=40, minwidth=50, anchor='c')
        self.treev.column("2", width=210, anchor='c')
        # Assigning the heading names
        # respective columns
        self.treev.heading("1", text="N°")
        self.treev.heading("2", text="File Name")
        # Add the adding files button
        path3 = 'DATA/IMG/add.png'
        img3 = ImageTk.PhotoImage(Image.open(path3))
        bu1 = ttk.Button(self.wind, image=img3)
        bu1.image = img3
        bu1.config(command=self.ADD_Fpdf)
        bu1.place(x=196, y=40)

        # Add the remove files button
        path4 = 'DATA/IMG/remove.png'
        img4 = ImageTk.PhotoImage(Image.open(path4))
        bu2 = ttk.Button(self.wind, image=img4)
        bu2.image = img4
        bu2.config(command=self.REM_F)
        bu2.place(x=220, y=40)
        # Add the DONE
        path5 = 'DATA/IMG/DONE.png'
        img5 = ImageTk.PhotoImage(Image.open(path5))
        bu3 = ttk.Button(self.wind, image=img5)
        bu3.image = img5
        bu3.config(command=self.Merg_F)
        bu3.pack(pady=5)
        # List of FILES pdf
        self.PDFL =[]
    def ADD_Fdoc(self):
        global filename
        if len(self.PDFL) > 0:
            messagebox.showerror('PDF Manager', 'You Cannot Add More Than 1 File')
        else:
            path = filedialog.askopenfilename(filetypes=[('Files DOCX', '*.docx')])
            if path in self.PDFL:
                messagebox.showerror('PDF Manager', 'This File Is already Added')
            else:
                self.PDFL.append(path)
                # find the index of file name
                for i in range(len(path)):
                    if path[i] == "/":
                        filename = path[i + 1:]
                self.treev.insert("", index=len(self.PDFL), text=filename, values=("%d" % (len(self.PDFL)), "%s" % (filename)))

    def ADD_Fpdf(self):
        global filename
        path = filedialog.askopenfilename(filetypes=[('Files PDF', '*.pdf')])
        if path in self.PDFL:
            messagebox.showerror('PDF Manager', 'This File Is already Added')
        else:
            self.PDFL.append(path)
            # find the index of file name
            for i in range(len(path)):
                if path[i] == "/":
                    filename = path[i + 1:]
            self.treev.insert("", index=len(self.PDFL), text=filename, values=("%d" % (len(self.PDFL)), "%s" % (filename)))

    def REM_F(self):
        try:
            global item_text
            # get the values of item selected and put it in item_text var
            for item in self.treev.selection():
                item_text = self.treev.item(item, "text")
            # then we find the value in the list to remove it
            for filepath in self.PDFL:
                if str(filepath).endswith(item_text):
                    self.PDFL.remove(self.PDFL[self.PDFL.index(filepath)])
            deletedfile = self.treev.selection()[0]
            self.treev.delete(deletedfile)
        except:
            messagebox.showerror('PDF MANAGER', 'No File Selected!')

    def Merg_F(self):
        pdfmerg = PyPDF2.PdfFileMerger()
        for ListPaths in self.PDFL:
            print(ListPaths)
            pdfmerg.append(ListPaths)
        filesave = filedialog.asksaveasfilename(filetypes=[('Files PDF', '*.pdf')])
        print(filesave)
        pdfmerg.write(filesave+'.pdf')
        pdfmerg.close()

    def PDF_CONV(self):
        self.wind = tk.Toplevel()
        self.frame.overrideredirect(True)
        self.wind.protocol("WM_DELETE_WINDOW", self.update)
        self.wind.grab_set()
        Wscreen = self.wind.winfo_screenwidth()
        Hscreen = self.wind.winfo_screenheight()
        width = 350
        height = 280
        FinalW = (Wscreen - width) / 2
        FinalH = (Hscreen - height) / 2
        FinalH -= 161
        FinalW -= 8
        self.wind.geometry('%dx%d+%d+%d' % (width, height, FinalW, FinalH))
        self.wind.resizable(0, 0)
        self.wind.title('PDF Merger')
        # lbl 1 (img)
        path = 'DATA/IMG/pdfconv.png'
        img1 = ImageTk.PhotoImage(Image.open(path))
        lbl1 = ttk.Label(self.wind, image=img1)
        lbl1.image = img1
        lbl1.pack()
        path2 = 'DATA/IMG/AddF.png'
        img2 = ImageTk.PhotoImage(Image.open(path2))
        lbl2 = ttk.Label(self.wind, image=img2)
        lbl2.image = img2
        lbl2.pack()
        # the treeview of all pdf files added
        self.treev = ttk.Treeview(self.wind, selectmode='browse', height=6)
        # Calling pack
        self.treev.pack()
        # Constructing vertical scrollbar
        # with treeview
        verscrlbar = ttk.Scrollbar(self.wind, orient="vertical", command=self.treev.yview)
        # Calling pack method w.r.to verical
        # scrollbar
        verscrlbar.place(x=30, y=110)
        # Configuring treeview
        self.treev.configure(xscrollcommand=verscrlbar.set)
        # Defining number of columns
        self.treev["columns"] = ("1", "2")
        # Defining heading
        self.treev['show'] = 'headings'
        # Assigning the width and anchor to  the
        # respective columns
        self.treev.column("1", width=40, minwidth=50, anchor='c')
        self.treev.column("2", width=210, anchor='c')
        # Assigning the heading names
        # respective columns
        self.treev.heading("1", text="N°")
        self.treev.heading("2", text="File Name")
        # Add the adding files button
        path3 = 'DATA/IMG/add.png'
        img3 = ImageTk.PhotoImage(Image.open(path3))
        bu1 = ttk.Button(self.wind, image=img3)
        bu1.image = img3
        bu1.config(command=self.ADD_Fdoc)
        bu1.place(x=196, y=40)

        # Add the remove files button
        path4 = 'DATA/IMG/remove.png'
        img4 = ImageTk.PhotoImage(Image.open(path4))
        bu2 = ttk.Button(self.wind, image=img4)
        bu2.image = img4
        bu2.config(command=self.REM_F)
        bu2.place(x=220, y=40)
        # Add the DONE
        path5 = 'DATA/IMG/doctoodf.png'
        img5 = ImageTk.PhotoImage(Image.open(path5))
        bu3 = ttk.Button(self.wind, image=img5)
        bu3.image = img5
        bu3.config(command=self.CONdoctopdf)
        bu3.pack(pady=5)
        # List of FILES pdf
        self.PDFL = []
    def CONdoctopdf(self):
        global indexfinal
        for files in self.PDFL:
            filesave = filedialog.asksaveasfilename(filetypes=[('Files PDF', '*.pdf')])
            filesave += ".pdf"
            i = len(filesave) -1
            while i > 0:
                if filesave[i] == '/':
                    i += 1
                    indexfinal = i
                    i = -10
                else:
                    i -= 1
            docx2pdf.convert("%s" % (files))
            docx2pdf.convert("%s" % (files), "%s" % (filesave))
            docx2pdf.convert("%s" % (filesave[:indexfinal]))

    def update(self):
        self.frame.overrideredirect(False)
        self.wind.destroy()

frame = Tk()
w = frame.winfo_screenwidth()
h = frame.winfo_screenheight()
APP(350, 272, w, h)
frame.mainloop()
