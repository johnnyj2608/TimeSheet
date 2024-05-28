import customtkinter as ctk
import os
from customtkinter import filedialog 
from datetime import datetime
from modSheet import modifySheets, printSheets, getExcelCount, months

class ProcessStop:
    def __init__(self):
        self.value = False

class TimesheetApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Automate Timesheets")
        self.processStop = ProcessStop()
        self.processRunning = False
        self.folderPath = ''
        self.prevDir = None

        self.root.geometry(self.centerWindow(self.root, 500, 450, self.root._get_window_scaling()))
        self.frame = ctk.CTkFrame(master=self.root)
        self.frame.pack(pady=20, padx=70, fill="both", expand=True)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        
        self.titleLabel = ctk.CTkLabel(master=self.frame, text="Automate Timesheets", font=(None, 25, "bold"))
        self.titleLabel.grid(row=0, column=0, columnspan=3, pady=12, padx=10, sticky="ew")

        self.folderLabel = ctk.CTkLabel(master=self.frame, text="No folder selected")
        self.folderLabel.grid(row=1, column=0, columnspan=3, pady=0, padx=10)

        self.browseButton = ctk.CTkButton(master=self.frame, text="Select Folder", command=self.browseFolder)
        self.browseButton.grid(row=2, column=0, columnspan=3, pady=(0, 6), padx=10)

        self.rangeFrameInit()
        self.dateFrameInit()

        self.modifyButton = ctk.CTkButton(master=self.frame, text="Apply Changes", command=self.toggleModifyButton, state="disabled")
        self.modifyButton.grid(row=6, column=0, columnspan=3, pady=6, padx=10)

        self.printButton = ctk.CTkButton(master=self.frame, text="Print Files", command=self.printPressed, state="disabled")
        self.printButton.grid(row=7, column=0, columnspan=3, pady=6, padx=10)

        self.statusLabel = ctk.CTkLabel(master=self.frame, text="")
        self.statusLabel.grid(row=8, column=0, columnspan=3, pady=0, padx=10)

        self.frame.grid_columnconfigure((0, 2), weight=1)

    def rangeFrameInit(self):
        self.rangeFrame = ctk.CTkFrame(master=self.frame, fg_color=self.frame.cget("fg_color"))
        self.rangeFrame.grid(row=3, column=0, columnspan=3, pady=6, padx=10)

        self.rangelabel = ctk.CTkLabel(master=self.rangeFrame, text="File Range")
        self.rangelabel.grid(row=0, column=0, columnspan=3, pady=0, padx=10, sticky="ew")

        self.startRangeEntry = ctk.CTkEntry(master=self.rangeFrame, width=40)
        self.startRangeEntry.grid(row=1, column=0, columnspan=1, pady=(0, 12), padx=10, sticky="e")
        self.startRangeEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateMember), "%P"), state="disabled")

        self.memberHyphenlabel = ctk.CTkLabel(master=self.rangeFrame, text="to")
        self.memberHyphenlabel.grid(row=1, column=1, columnspan=1, pady=(0, 12), padx=10, sticky="ew")

        self.endRangeEntry = ctk.CTkEntry(master=self.rangeFrame, width=40)
        self.endRangeEntry.grid(row=1, column=2, columnspan=1, pady=(0, 12), padx=10, sticky="w")
        self.endRangeEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateMember), "%P"), state="disabled")

        self.rangeFrame.grid_columnconfigure((0, 1, 2), weight=1)

    def dateFrameInit(self):
        self.dateFrame = ctk.CTkFrame(master=self.frame, fg_color=self.frame.cget("fg_color"))
        self.dateFrame.grid(row=4, column=0, columnspan=4, pady=6, padx=10, sticky="ew")

        self.monthLabel = ctk.CTkLabel(master=self.dateFrame, text="Month:")
        self.monthLabel.grid(row=0, column=0, columnspan=2, pady=6, padx=10, sticky="e")

        self.monthCombo = ctk.CTkComboBox(master=self.dateFrame, values=list(months.keys()), width=110, state="disabled")
        self.monthCombo.grid(row=0, column=2, columnspan=2, pady=6, padx=10, sticky="w")

        self.yearLabel = ctk.CTkLabel(master=self.dateFrame, text="Year:")
        self.yearLabel.grid(row=1, column=0, columnspan=2, pady=6, padx=10, sticky="e")
        
        self.yearEntry = ctk.CTkEntry(master=self.dateFrame, width=110)
        self.yearEntry.grid(row=1, column=2, columnspan=2, pady=6, padx=10, sticky="w")
        self.yearEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateYear), "%P"), state="disabled")

        self.dateFrame.grid_columnconfigure((0, 2), weight=1)

    def browseFolder(self):
        initialDir = self.prevDir
        folderPath = filedialog.askdirectory(initialdir=initialDir)
        if not folderPath:
            return

        self.folderPath = folderPath
        self.enableUserActions()
        parentDir = os.path.dirname(self.folderPath)
        folderName = os.path.basename(self.folderPath)
        self.folderLabel.configure(text=folderName, text_color="gray84")
        self.prevDir = parentDir

        self.startRangeEntry.delete(0, "end")
        self.startRangeEntry.insert(0, 1)

        self.endRangeEntry.delete(0, "end")
        self.endRangeEntry.insert(0, getExcelCount(self.folderPath))

        currentMonth = datetime.now().strftime("%B")
        self.monthCombo.set(currentMonth)

        currentYear = datetime.now().year
        self.yearEntry.delete(0, "end")
        self.yearEntry.insert(0, currentYear)

    def toggleModifyButton(self):
        valid, response = self.validateInputs()
        if not valid:
            self.statusLabel.configure(text=response, text_color="red")
            return
        self.statusLabel.configure(text="", text_color="gray84")
        self.disableUserActions()
        if not self.processRunning:
            self.modifyButton.configure(text="Stop Changes", fg_color='#800000', hover_color='#98423d')
            self.processRunning = True
            selectedMonth = self.monthCombo.get()
            selectedYear = self.yearEntry.get()
            self.statusLabel.configure(text="")
            self.statusLabel.update()
            startMember = int(self.startRangeEntry.get())-1
            endMember = int(self.endRangeEntry.get())

            self.modifyButton.configure(state="normal")
            modifySheets(self.folderPath, selectedMonth, selectedYear, startMember, endMember, self.statusLabel, self.processStop)
        else:
            self.processStop.value = True
        self.modifyButton.configure(text="Apply Changes", fg_color='#1f538d', hover_color='#14375e')
        self.processRunning = False
        self.enableUserActions()

    def printPressed(self):
        valid, response = self.validateInputs()
        if not valid:
            self.statusLabel.configure(text=response, text_color="red")
            return
        self.statusLabel.configure(text="", text_color="gray84")
        self.disableUserActions()
        if not self.processRunning:
            self.printButton.configure(text="Stop Printing", fg_color='#800000', hover_color='#98423d')
            self.processRunning = True
            self.statusLabel.configure(text="")
            self.statusLabel.update()
            startMember = int(self.startRangeEntry.get())-1
            endMember = int(self.endRangeEntry.get())

            self.printButton.configure(state="normal")
            printSheets(self.folderPath, self.statusLabel, startMember, endMember, self.processStop)
        else:
            self.processStop.value = True
        self.printButton.configure(text="Print Files", fg_color='#1f538d', hover_color='#14375e')
        self.processRunning = False
        self.enableUserActions()

    def validateYear(self, val):
        return val == "" or (val.isdigit() and len(val) <= 4)
    
    def validateMember(self, val):
        return val == "" or val.isdigit()
    
    def validateInputs(self):
        members = getExcelCount(self.folderPath)

        try:
            startMember = int(self.startRangeEntry.get())
            endMember = int(self.endRangeEntry.get()) 
        except:
            return False, "Member range can not be empty"

        if startMember == 0 or endMember == 0:
            return False, "Member range can not be 0"
        if startMember > endMember:
            return False, "Member start range larger than end range"
        if startMember > members or endMember > members:
            return False, "Member range larger than member count"
        
        selectedYear = self.yearEntry.get()
        if not selectedYear:
            return False, "Year can not be empty"

        return True, ""

    def disableUserActions(self):
        self.browseButton.configure(state="disabled")
        self.startRangeEntry.configure(state="disabled")
        self.endRangeEntry.configure(state="disabled")
        self.monthCombo.configure(state="disabled")
        self.yearEntry.configure(state="disabled")
        self.modifyButton.configure(state="disabled")
        self.printButton.configure(state="disabled")

    def enableUserActions(self):
        self.browseButton.configure(state="normal")
        self.startRangeEntry.configure(state="normal")
        self.endRangeEntry.configure(state="normal")
        self.monthCombo.configure(state="normal")
        self.yearEntry.configure(state="normal")
        self.modifyButton.configure(state="normal")
        self.printButton.configure(state="normal")

    def centerWindow(self, Screen: ctk, width: int, height: int, scale_factor: float = 1.0):
        screen_width = Screen.winfo_screenwidth()
        screen_height = Screen.winfo_screenheight()
        x = int(((screen_width/2) - (width/2)) * scale_factor)
        y = int(((screen_height/2) - (height/1.5)) * scale_factor)
        return f"{width}x{height}+{x}+{y}"

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = TimesheetApp()
    app.run()
