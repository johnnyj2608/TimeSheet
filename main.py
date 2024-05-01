import customtkinter as ctk
from customtkinter import filedialog 
from datetime import datetime
from modSheet import modifySheets, printSheets, months

class TimesheetApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Timesheet Manager")

        self.root.geometry(self.centerWindow(self.root, 500, 400, self.root._get_window_scaling()))
        self.frame = ctk.CTkFrame(master=self.root)
        self.frame.pack(pady=20, padx=70, fill="both", expand=True)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.titleLabel = ctk.CTkLabel(master=self.frame, text="Sign-In Sheets")
        self.titleLabel.grid(row=0, column=0, columnspan=2, pady=12, padx=10)

        self.folderLabel = ctk.CTkLabel(master=self.frame, text="No folder selected")
        self.folderLabel.grid(row=1, column=0, columnspan=2, pady=0, padx=10)

        self.browseButton = ctk.CTkButton(master=self.frame, text="Select Folder", command=self.browseFolder)
        self.browseButton.grid(row=2, column=0, columnspan=2, pady=(0, 12), padx=10)

        self.monthLabel = ctk.CTkLabel(master=self.frame, text="Month:")
        self.monthLabel.grid(row=3, column=0, pady=6, padx=10, sticky="e")

        currentMonth = datetime.now().strftime("%B")
        self.monthCombo = ctk.CTkComboBox(master=self.frame, values=list(months.keys()))
        self.monthCombo.grid(row=3, column=1, pady=12, padx=10, sticky="w")
        self.monthCombo.set(currentMonth)

        self.yearLabel = ctk.CTkLabel(master=self.frame, text="Year:")
        self.yearLabel.grid(row=4, column=0, pady=6, padx=10, sticky="e")

        currentYear = datetime.now().year
        self.yearEntry = ctk.CTkEntry(master=self.frame)
        self.yearEntry.grid(row=4, column=1, pady=12, padx=10, sticky="w")
        self.yearEntry.configure(validate="key", validatecommand=(self.frame.register(self.validateYear), "%P"))
        self.yearEntry.insert(0, currentYear)
        self.yearEntry.bind("<KeyRelease>", self.enableModify)

        self.modifyButton = ctk.CTkButton(master=self.frame, text="Apply Changes", command=self.modifyPressed, state="disabled")
        self.modifyButton.grid(row=5, column=0, columnspan=2, pady=6, padx=10)

        self.printButton = ctk.CTkButton(master=self.frame, text="Print Files", command=self.printPressed, state="disabled")
        self.printButton.grid(row=6, column=0, columnspan=2, pady=6, padx=10)

        self.statusLabel = ctk.CTkLabel(master=self.frame, text="")
        self.statusLabel.grid(row=7, column=0, columnspan=2, pady=6, padx=10)

        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_columnconfigure(1, weight=1)

    def browseFolder(self):
        filename = filedialog.askdirectory()
        if filename:
            self.folderLabel.configure(text=filename)
        self.enableModify()

    def modifyPressed(self):
        selectedFolder = self.folderLabel.cget("text")
        selectedMonth = self.monthCombo.get()
        selectedYear = self.yearEntry.get()
        self.statusLabel.configure(text="")
        self.statusLabel.update()
        self.processRunning()
        modifySheets(selectedFolder, selectedMonth, selectedYear, self.statusLabel)
        self.processComplete()

    def printPressed(self):
        selectedFolder = self.folderLabel.cget("text")
        self.printButton.configure(state="disabled")
        self.statusLabel.configure(text="")
        self.statusLabel.update()
        self.processRunning()
        printSheets(selectedFolder, self.statusLabel)
        self.processComplete()

    def validateYear(self, val):
        return val == "" or (val.isdigit() and len(val) <= 4)

    def enableModify(self, *args):
        if self.yearEntry.get() and int(self.yearEntry.get()) >= 2000 and self.folderLabel.cget("text") != "No folder selected":
            self.modifyButton.configure(state="normal")
            self.printButton.configure(state="normal")
        else:
            self.modifyButton.configure(state="disabled")
            self.printButton.configure(state="disabled")

    def processRunning(self):
        self.monthCombo.configure(state="disabled")
        self.yearEntry.configure(state="disabled")
        self.browseButton.configure(state="disabled")
        self.modifyButton.configure(state="disabled")
        self.printButton.configure(state="disabled")

    def processComplete(self):
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
