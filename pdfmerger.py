from tkinter import *
from tkinter import ttk, messagebox, filedialog
from functools import partial
from PyPDF2 import PdfFileMerger, PdfFileReader
import os 

class PDFMerger():
    def __init__(self, master):

        # Initialising Variables
        self.image_path = os.getcwd()
        self.entries = []
        self.btn_id = 0
        self.filecopy = ''
        self.pos_x = 150
        self.pos_y = 20

        # Styling Root, Labels and Frames
        master.title('PDF Merger by Kaif Kutchwala')
        master.geometry('500x600')
        master.resizable(False, False)
        master.config(background = '#25274d')

        self.style = ttk.Style()
        self.style.configure('TLabel', background = '#25274d')
        self.style.configure('TFrame', background = '#25274d')
        
        # Headers Section: includes Title, Author and Logo
        self.headers = ttk.Frame(master)
        self.headers.pack()

        self.logo = PhotoImage(file = self.image_path + '\icon1.png').subsample(2,2)
        self.logo_label = ttk.Label(self.headers, image = self.logo).grid(row = 0, column = 0, rowspan = 2, padx = 10)
        self.title = ttk.Label(self.headers, text = "PDF Merger", font = ('Code Bold', 25, 'bold'), foreground = '#ffffff').grid(row = 0, column = 1, pady = (10,0))
        self.author = ttk.Label(self.headers, text = "by Kaif Kutchwala", font = ('Quicksand', 15, 'bold'), foreground = '#69d3ff').grid(row = 1, column = 1, pady = (0,15))

        # Files to be Merged Section:
        self.merge_files = ttk.Frame(master)
        self.merge_files.pack()

        ttk.Label(self.merge_files, text = 'Files to Merge', foreground = '#fef9ff').grid(row = 0, column = 0, sticky = 'w')

        self.canvas = Canvas(self.merge_files, background = '#29648a')
        self.sbar = ttk.Scrollbar(self.merge_files, orient = VERTICAL, command = self.canvas.yview)
        self.canvas.configure(yscrollcommand = self.sbar.set, scrollregion = (0,0, 400, 2000))
        self.sbar.grid(row = 1, column = 1, sticky = 'ns')
        self.canvas.grid(row = 1, column = 0)

        # Destination File Section:
        self.dst_info = ttk.Frame(master)
        self.dst_info.pack(pady = 10)

        ttk.Label(self.dst_info, text = 'Destination File', foreground = '#fef9ff').grid(row = 0, column = 0, sticky = 'w')
        self.dst_entry = ttk.Entry(self.dst_info, width = 50)
        self.dst_entry.insert(0, 'Select the destination folder')
        self.dst_entry.grid(row = 1, column = 0)

        self.dst_browse = Button(self.dst_info, text = 'Browse', bg = '#29648a', foreground = '#69d3ff', command = self.dstAskFile)
        self.dst_browse.grid(row = 1, column = 1, padx = 10)

        self.dst_file = ttk.Entry(self.dst_info, width = 25)
        self.dst_file.insert(0, 'Enter file name here.')
        self.dst_file.grid(row = 2, column = 0, columnspan = 2)

        # Buttons Section: Includes 'Add PDF' and 'Merge' Button
        self.btns = ttk.Frame(master)
        self.btns.pack(pady = 10)

        self.btn_image = PhotoImage(file = self.image_path + '\PDFMerger_button.png').subsample(9,9)
        self.merge_btn = Button(self.btns, text = 'Merge', borderwidth = 0, bg = '#25274d', compound = CENTER, image = self.btn_image, font = ('Trebuchet MS', 15), foreground = '#f2f2f2', command = self.merge)
        self.merge_btn.grid(row = 0, column = 0)

        # Adds Entry and Button to Files to be merged section each time it is pressed.
        self.add_btn = Button(self.btns, text = 'Add PDF', borderwidth = 0, bg = '#25274d', compound = CENTER, image = self.btn_image, font = ('Trebuchet MS', 15), foreground = '#f2f2f2', command = self.addPDF)
        self.add_btn.grid(row = 0, column = 1)

    def addPDF(self):
        # Create Entry and Button
        # Button Id ensures that when file path is selected it is sent to the correct entry
        btn = Button(self.canvas, text = 'Browse', bg = '#29648a', foreground = '#69d3ff', command = partial(self.askFile, self.btn_id))
        en = ttk.Entry(self.canvas, width = 45)
        # Add Entry to 'entries' for future reference
        self.entries.append(en)
        # Add  Entry and Button to the canvas.
        self.canvas.create_window(self.pos_x, self.pos_y, window = en)
        self.canvas.create_window(self.pos_x + 190, self.pos_y, window = btn)
        # Increment Button ID (same as index number of entry in 'entries' list)
        # Increment pos_y so next set of entry and button are add below previous
        self.btn_id += 1
        self.pos_y += 50

    def askFile(self, id):
        # Open File Dialog for only PDF files
        self.filecopy = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("pdf files","*.pdf"),("all files","*.*")))
        #print(self.filecopy)
        #print(id)
        # Insert filepath in entry with the same index as Button ID.
        self.entries[id].delete(0, 'end')
        self.entries[id].insert(0, self.filecopy)
    
    def dstAskFile(self):
        # Same as askFile method but for destination.
        self.filecopy = filedialog.askdirectory()
        print(self.filecopy)
        self.dst_entry.delete(0, 'end')
        self.dst_entry.insert(0, self.filecopy)

    def merge(self):
        # Create PdfFileMerger obejct
        self.mergedObject = PdfFileMerger()
        # Create empty list that stores all filepaths
        self.files = []
        # Add all file paths to list
        for entry in self.entries:
            # Avoid empty entries
            if entry.get() != '':
                self.files.append(str(entry.get()))
                # Clear Entries
                entry.delete(0, 'end')
        # Loop through paths and merge all files
        for f in self.files:
            self.mergedObject.append(PdfFileReader(str(f), 'rb'))
        # Write Merged File
        try:
            self.mergedObject.write(str(self.dst_entry.get()) + "/" +str(self.dst_file.get()) + '.pdf')
            # Clear Entry and Reset Values
            self.dst_entry.delete(0, 'end')
            self.dst_file.delete(0, 'end')
            self.dst_entry.insert(0, 'Select the destination folder')
            self.dst_file.insert(0, 'Enter file name here.')
        except:
            messagebox.showerror('Destination File Error', 'Problem with Destination Info. Please check again.')

        # Show Completion Message
        messagebox.showinfo("Process Complete", str(self.dst_file.get()) + ".pdf File Can Now be Found in Destination Folder")

def main():
    # Creating Root Window
    root = Tk()
    # Creating object of PDFMerger class and passing in root window.
    pdfmerger = PDFMerger(root)
    root.mainloop()
    
if __name__ == "__main__":
    main()
