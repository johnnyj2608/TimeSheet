import customtkinter as ctk
from customtkinter import filedialog 
from datetime import datetime
from modSheet import modifySheets
from modSheet import months

class TimesheetApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("An Kang Timesheet Window")
        self.root.geometry("500x350")

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.frame = ctk.CTkFrame(master=self.root)
        self.frame.pack(pady=20, padx=60, fill="both", expand=True)

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

        self.modifyButton = ctk.CTkButton(master=self.frame, text="Confirm Changes", command=self.modifyPressed, state="disabled")
        self.modifyButton.grid(row=5, column=0, columnspan=2, pady=12, padx=10)

        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_columnconfigure(1, weight=1)

    def browseFolder(self):
        filename = filedialog.askdirectory()
        self.folderLabel.configure(text=filename)
        self.enableModify()

    def modifyPressed(self):
        selectedFolder = self.folderLabel.cget("text")
        selectedMonth = self.monthCombo.get()
        selectedYear = self.yearEntry.get()
        modifySheets(selectedFolder, selectedMonth, selectedYear)

    def validateYear(self, val):
        return val == "" or (val.isdigit() and len(val) <= 4)

    def enableModify(self, *args):
        if self.yearEntry.get() and int(self.yearEntry.get()) >= 2000 and self.folderLabel.cget("text") != "No folder selected":
            self.modifyButton.configure(state="normal")
        else:
            self.modifyButton.configure(state="disabled")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = TimesheetApp()
    app.run()
