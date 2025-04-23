import os, shutil
from openpyxl import Workbook, load_workbook
import tkinter as tk
import tkinter.font as font
from tkinter import messagebox, StringVar, OptionMenu, ttk

def validate_input():
    user_input = combo.get()
    try:
        if int(user_input) in TEFCSR:
            createQuote(int(user_input))
        else:
            messagebox.showerror("Site Number Not Found!", "Please check that you have entered the correct site number.")
    except:
        messagebox.showerror("Site Number Not Found!", "Please check that you have selected a site number.")

def createQuote(site_num, msgbox=True):
    i = TEFCSR.index(site_num)
    access_type = {
        "" : 0,
        "Ladder" : 0,
        "Podium Steps" : 9,
        "Mobile Scaffolding" : 10,
        "Cherry Picker" : 11
    }
    filename = "NetCS - " + str(site_num) + " - " + site_name[i] + " - CS.xlsx"
    wb = load_workbook("bins/VMO2 CS TEMPLATE.xlsx")
    ws = wb["Summary"]
    ws.cell(4, 2).value = site_num
    ws.cell(5, 2).value = site_name[i]
    ws.cell(6, 2).value = site_address[i]
    ws.cell(7, 2).value = site_postcode[i]
    ws = wb["Ratecard"]
    ws.cell(8, 3).value = cost_recieved[i]
    if site_access[i] != None:
        if access_type[site_access[i]] > 0:
            ws.cell(access_type[site_access[i]], 5).value = 1
    if split_decom[i] == "Y":
        ws.cell(6, 5).value = 1
    if ooh_decom[i] == "Y":
        ws.cell(7, 5).value = 1
    wb.save(filename)
    if msgbox:
        messagebox.showinfo("Quote Created!", "Quote for site: " + str(site_num) + " - " + site_name[i] + " has been created!")

def createAllQuotes():
    for id in TEFCSR:
        try:
            createQuote(id, False)
        except Exception as e:
            messagebox.showerror("Error", f"Site Skipped: {id}, {e}")

    messagebox.showinfo("Quotes Created", "Quotes have been created for every site.")

url = "https://newedge-my.sharepoint.com/:x:/g/personal/tima_netcs_co_uk/EQD9EciysPNPpW5OzvV71pYBS0sjqyyCOM4QpxkF29Vl3w?download=1"

os.system(".\\bins\\wget.exe -O download/tracker.xlsx " + url)

wb = load_workbook("download/tracker.xlsx")
ws = wb["Decom Project 2025 - Surveys"]
TEFCSR = [int(ws.cell(cell_no, 1).value) for cell_no in range(3, ws.max_row)]
site_name = [ws.cell(cell_no, 3).value.replace("/", "-").replace("\t", "") for cell_no in range(3, ws.max_row)]
site_address = [ws.cell(cell_no, 4).value for cell_no in range(3, ws.max_row)]
site_postcode = [ws.cell(cell_no, 5).value for cell_no in range(3, ws.max_row)]
cost_recieved = [ws.cell(cell_no, 61).value for cell_no in range(3, ws.max_row)]
site_access = [ws.cell(cell_no, 63).value for cell_no in range(3, ws.max_row)]
split_decom = [ws.cell(cell_no, 64).value for cell_no in range(3, ws.max_row)]
ooh_decom = [ws.cell(cell_no, 65).value for cell_no in range(3, ws.max_row)]

root = tk.Tk()
root.geometry("300x130")
root.resizable(False, False)
root.title("VMO2 Quote Generator")

label = tk.Label(root, text="Select Site Number", font=font.Font(size=12))
label.pack(pady=10)

combo = ttk.Combobox(
    state="readonly",
    values=sorted(TEFCSR)
)
combo.pack()
button = tk.Button(root, text="Submit", command=validate_input, font=font.Font(size=12))
button.pack(pady=10)

#button2 = tk.Button(root, text="Create Quote for all sites", command=createAllQuotes, font=font.Font(size=12))
#button2.pack(pady=10)

root.mainloop()