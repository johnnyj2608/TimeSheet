import customtkinter as ctk
from datetime import datetime

def browseFolder():
    print("test")

def modifySheets():
    print("test")

def validateYear(val):
    return val == "" or (val.isdigit() and len(val) <= 4)

def validYear(*args):
    if yearEntry.get() and int(yearEntry.get()) >= 2000:
        modifyButton.configure(state="normal")
    else:
        modifyButton.configure(state="disabled")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

root = ctk.CTk()
root.geometry("500x350")

frame = ctk.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

titleLabel = ctk.CTkLabel(master=frame, text="Sign-In Sheets")
titleLabel.grid(row=0, column=0, columnspan=2, pady=12, padx=10)

browseButton = ctk.CTkButton(master=frame, text="Select Folder", command=browseFolder)
browseButton.grid(row=1, column=0, columnspan=2, pady=12, padx=10)

monthLabel = ctk.CTkLabel(master=frame, text="Month:")
monthLabel.grid(row=2, column=0, pady=6, padx=10, sticky="e")

months = ["January", "February", "March", "April", "May", 
           "June", "July", "August", "September", "October", 
           "November", "December"]
currentMonth = datetime.now().strftime("%B")
monthCombo = ctk.CTkComboBox(master=frame, values=months)
monthCombo.grid(row=2, column=1, pady=12, padx=10, sticky="w")
monthCombo.set(currentMonth)

yearLabel = ctk.CTkLabel(master=frame, text="Year:")
yearLabel.grid(row=3, column=0, pady=6, padx=10, sticky="e")

currentYear = datetime.now().year
yearEntry = ctk.CTkEntry(master=frame)
yearEntry.grid(row=3, column=1, pady=12, padx=10, sticky="w")
yearEntry.configure(validate="key", validatecommand=(frame.register(validateYear), "%P"))
yearEntry.insert(0, currentYear)
yearEntry.bind("<KeyRelease>", validYear)

modifyButton = ctk.CTkButton(master=frame, text="Confirm Changes", command=modifySheets)
modifyButton.grid(row=4, column=0, columnspan=2, pady=12, padx=10)

frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)

root.mainloop()