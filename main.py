import customtkinter as ctk
import os
from customtkinter import filedialog 
from datetime import datetime
from modSheet import modifySheets, printSheets, months

class ProcessStop:
    def __init__(self):
        self.value = False

class TimesheetApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Timesheet Manager")
        self.processStop = ProcessStop()
        self.processRunning = False
        self.folderPath = ''
        self.prevDir = None

        self.root.geometry(self.centerWindow(self.root, 500, 400, self.root._get_window_scaling()))
        self.frame = ctk.CTkFrame(master=self.root)
        self.frame.pack(pady=20, padx=70, fill="both", expand=True)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        
        self.titleLabel = ctk.CTkLabel(master=self.frame, text="Sign-In Sheets")
        self.titleLabel.grid(row=0, column=0, columnspan=3, pady=12, padx=10, sticky="ew")

        self.folderLabel = ctk.CTkLabel(master=self.frame, text="No folder selected")
        self.folderLabel.grid(row=1, column=0, columnspan=3, pady=0, padx=10)

        self.browseButton = ctk.CTkButton(master=self.frame, text="Select Folder", command=self.browseFolder)
        self.browseButton.grid(row=2, column=0, columnspan=3, pady=(0, 12), padx=10)

        self.monthLabel = ctk.CTkLabel(master=self.frame, text="Month:")
        self.monthLabel.grid(row=3, column=0, columnspan=1, pady=6, padx=10, sticky="e")

        currentMonth = datetime.now().strftime("%B")
        self.monthCombo = ctk.CTkComboBox(master=self.frame, values=list(months.keys()), width=110)
        self.monthCombo.grid(row=3, column=2, columnspan=1, pady=12, padx=10, sticky="w")
        self.monthCombo.set(currentMonth)

        self.yearLabel = ctk.CTkLabel(master=self.frame, text="Year:")
        self.yearLabel.grid(row=4, column=0, columnspan=1, pady=6, padx=10, sticky="e")

        currentYear = datetime.now().year
        self.yearEntry = ctk.CTkEntry(master=self.frame, width=110)
        self.yearEntry.grid(row=4, column=2, columnspan=1, pady=12, padx=10, sticky="w")
        self.yearEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateYear), "%P"))
        self.yearEntry.insert(0, currentYear)
        self.yearEntry.bind("<KeyRelease>", self.enableModify)

        self.modifyButton = ctk.CTkButton(master=self.frame, text="Apply Changes", command=self.toggleModifyButton, state="disabled")
        self.modifyButton.grid(row=5, column=0, columnspan=3, pady=6, padx=10)

        self.printButton = ctk.CTkButton(master=self.frame, text="Print Files", command=self.printPressed, state="disabled")
        self.printButton.grid(row=6, column=0, columnspan=3, pady=6, padx=10)

        self.statusLabel = ctk.CTkLabel(master=self.frame, text="")
        self.statusLabel.grid(row=7, column=0, columnspan=3, pady=6, padx=10)

        self.frame.grid_columnconfigure((0, 2), weight=1)

    def browseFolder(self):
        initialDir = self.prevDir
        self.folderPath = filedialog.askdirectory(initialdir = initialDir)
        if self.folderPath:
            parentDir = os.path.dirname(self.folderPath)
            folderName = os.path.basename(self.folderPath)
            self.folderLabel.configure(text=folderName)
            self.prevDir = parentDir
        self.enableModify()

    def toggleModifyButton(self):
        self.disableUserActions()
        if not self.processRunning:
            self.modifyButton.configure(text="Stop Changes", fg_color='#800000', hover_color='#98423d')
            self.processRunning = True
            selectedMonth = self.monthCombo.get()
            selectedYear = self.yearEntry.get()
            self.statusLabel.configure(text="")
            self.statusLabel.update()

            self.modifyButton.configure(state="normal")
            modifySheets(self.folderPath, selectedMonth, selectedYear, self.statusLabel, self.processStop)
        else:
            self.processStop.value = True
        self.modifyButton.configure(text="Apply Changes", fg_color='#1f538d', hover_color='#14375e')
        self.processRunning = False
        self.enableUserActions()

    def printPressed(self):
        self.disableUserActions()
        if not self.processRunning:
            self.printButton.configure(text="Stop Printing", fg_color='#800000', hover_color='#98423d')
            self.processRunning = True
            self.statusLabel.configure(text="")
            self.statusLabel.update()

            self.printButton.configure(state="normal")
            printSheets(self.folderPath, self.statusLabel, self.processStop)
        else:
            self.processStop.value = True
        self.printButton.configure(text="Print Files", fg_color='#1f538d', hover_color='#14375e')
        self.processRunning = False
        self.enableUserActions()

    def validateYear(self, val):
        return val == "" or (val.isdigit() and len(val) <= 4)

    def enableModify(self, *args):
        if self.yearEntry.get() and int(self.yearEntry.get()) >= 2000 and self.folderLabel.cget("text") != "No folder selected":
            self.modifyButton.configure(state="normal")
            self.printButton.configure(state="normal")
        else:
            self.modifyButton.configure(state="disabled")
            self.printButton.configure(state="disabled")

    def disableUserActions(self):
        self.monthCombo.configure(state="disabled")
        self.yearEntry.configure(state="disabled")
        self.browseButton.configure(state="disabled")
        self.modifyButton.configure(state="disabled")
        self.printButton.configure(state="disabled")

    def enableUserActions(self):
        self.monthCombo.configure(state="normal")
        self.yearEntry.configure(state="normal")
        self.browseButton.configure(state="normal")
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
