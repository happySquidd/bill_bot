from scrape_dates import get_dates
from scrape_price import get_price
from edit_xlsx import create_xlsx
import os
import subprocess
from datetime import datetime
import time
import shutil
from colorama import Fore, Style
import sys

#change for public version: excel_name, roommate1_name, roommate2_name, ALL ROOMMATE PATHS THAT DISPLAY THE ADDRESS
all_pdf = []
folder_name = "Tenant_Billing"
excel_name = "123 Main St., CityName_Utilities_" # add cost and date later
pdf_count = 0
duke_path = ""
spectrum_path = ""
roommate1_name = "name1"
roommate2_name = "name2"
date_now = datetime.now()
todays_year = date_now.year
todays_day = date_now.day
todays_month = date_now.month
if todays_month == 1:
    bill_month = 12
else:
    bill_month = todays_month


def get_year(month):
    global bill_year
    if month == 12:
        bill_year = todays_year-1
    else:
        bill_year = todays_year
    return bill_year
bill_year = get_year(bill_month)


desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
documents_path = os.path.join(os.path.expanduser("~"), "Documents")
if not os.path.exists(documents_path):
    one_drive_path = os.path.join(os.path.expanduser("~"), "OneDrive")
    if os.path.exists(one_drive_path):
        if os.path.exists(os.path.join(one_drive_path, "Documents")):
            documents_path = os.path.join(one_drive_path, "Documents")
        else:
            os.mkdir(documents_path)
roommate1_path = os.path.join(documents_path, str("123 Roommate " + str(roommate1_name)), "Utilities", str(todays_year), str("For "+str(bill_month)+"-"+str(bill_year)+"_Billed on "+str(todays_month)+"-"+str(todays_day)+"-"+str(todays_year)))
roommate2_path = os.path.join(documents_path, str("123 Roommate " + str(roommate2_name)), "Utilities", str(todays_year), str("For "+str(bill_month)+"-"+str(bill_year)+"_Billed on "+str(todays_month)+"-"+str(todays_day)+"-"+str(todays_year)))
folder_path = os.path.join(desktop_path, folder_name)


if not os.path.exists(folder_path):
    print("Creating 'Tenant_Billing' folder on Desktop")
    os.mkdir(folder_path)



# super janky code to create a string of nested folders but idc to work on this
if not os.path.exists(os.path.join(documents_path, str("123 Roommate " + str(roommate1_name)))):
    os.mkdir(os.path.join(documents_path, str("123 Roommate " + str(roommate1_name))))
if not os.path.exists(os.path.join(documents_path, str("123 Roommate " + str(roommate1_name)), "Utilities")):
    os.mkdir(os.path.join(documents_path, str("123 Roommate " + str(roommate1_name)), "Utilities"))
if not os.path.exists(os.path.join(documents_path, str("123 Roommate " + str(roommate1_name)), "Utilities", str(todays_year))):
    os.mkdir(os.path.join(documents_path, str("123 Roommate " + str(roommate1_name)), "Utilities", str(todays_year)))
if not os.path.exists(roommate1_path):
    print(f"Creating a billing folder for {roommate1_name}")
    os.mkdir(roommate1_path)
# else:
#     print("Folder for this billing period already exists, exiting")
#     time.sleep(2)
#     exit()
if not os.path.exists(os.path.join(documents_path, str("123 Roommate " + str(roommate2_name)))):
    os.mkdir(os.path.join(documents_path, str("123 Roommate " + str(roommate2_name))))
if not os.path.exists(os.path.join(documents_path, str("123 Roommate " + str(roommate2_name)), "Utilities")):
    os.mkdir(os.path.join(documents_path, str("123 Roommate " + str(roommate2_name)), "Utilities"))
if not os.path.exists(os.path.join(documents_path, str("123 Roommate " + str(roommate2_name)), "Utilities", str(todays_year))):
    os.mkdir(os.path.join(documents_path, str("123 Roommate " + str(roommate2_name)), "Utilities", str(todays_year)))
if not os.path.exists(roommate2_path):
    print(f"Creating a billing folder for {roommate2_name}")
    os.mkdir(roommate2_path)

print("Opening folder. Drop your bill pdf files into this folder")
subprocess.Popen(["explorer", folder_path])
input("Press enter once you're done")

def check_folder_contents():
    global pdf_count
    global all_pdf
    global duke_path
    global spectrum_path

    pdf_count = 0
    all_pdf = []
    duke_path = ""
    spectrum_path = ""

    folder_contents = os.listdir(folder_path)
    for i in folder_contents:
        if i[-3:] == "pdf":
            pdf_count += 1
            cur_pdf_path = os.path.abspath(os.path.join(folder_path, i))
            all_pdf.append(cur_pdf_path)

            words = i.split()
            if "bill" in words or "bill" in i:
                duke_path = cur_pdf_path
            else:
                spectrum_path = cur_pdf_path

check_folder_contents()
while pdf_count != 2:
    if pdf_count < 2:
        input("Missing bill file(s) (make sure bill files are .pdf) \nPress enter to retry")
        check_folder_contents()
    else:
        input("Too many files in the folder. Try leaving only the neccesarry files \nPress enter to retry")
        check_folder_contents()

while len(duke_path) < 2:
    input("Please ensure files aren't renamed from original bill files. (We're assuming that duke pdf starts with 'bill')\nPress enter to retry")
    check_folder_contents()
    print()

print("Calculating...")
steps = 20
for i in range(steps + 1):
    sys.stdout.write("\r")
    bar = "=" * i + " " * (steps - i) 
    percentage = int((i / steps) * 100) 
    sys.stdout.write(f"[{bar}] {percentage}%")
    sys.stdout.flush()
    time.sleep(0.05)

print()
duke_bill_date, duke_from, duke_to = get_dates(duke_path)
duke_amount = get_price(duke_path)
spectrum_bill_date, spectrum_from, spectrum_to = get_dates(spectrum_path)
spectrum_amount = get_price(spectrum_path)

total_amount = round(spectrum_amount+duke_amount, 2)

# converts date formatting from Nov 29 to 11/29/2024
months = {"Jan": 1,"Feb": 2,"Mar": 3,"Apr": 4,"May": 5,"Jun": 6,"Jul": 7,"Aug": 8,"Sep": 9,"Oct": 10,"Nov": 11,"Dec": 12}
# all_dates = [duke_bill_date, duke_from, duke_to, spectrum_bill_date, spectrum_from, spectrum_to] #can't do this to loop through values

def assign_date(date, compare_year_from=""):
    words = date.split()
    month = words[0]
    day = words[1]

    for key, value in months.items():
        if key == month:
            month=value
    year = bill_year
    
    if not len(compare_year_from) < 2:
        old_date = compare_year_from.split()
        old_month = old_date[0]
        if old_month == "Dec":
            year = str(bill_year+1)
    
    return str(month) + "/" + str(day) + "/" + str(year)


duke_bill_date=assign_date(duke_bill_date)
duke_to=assign_date(duke_to, duke_from)
duke_from=assign_date(duke_from)
spectrum_bill_date=assign_date(spectrum_bill_date)
spectrum_to=assign_date(spectrum_to, spectrum_from)
spectrum_from=assign_date(spectrum_from)

# create excel name and copy over files
roommates_to_copy_dir = [roommate1_path, roommate2_path]
files_to_copy = [duke_path, spectrum_path]

def convert_date(date): #converts from mm/dd/yyyy to mmddyy
    split_date = date.split("/")
    new_date = ""
    for i in split_date:
        if i == split_date[2]: #shorten the year
            new_date += i[2:]
        else:
            if len(i) < 2:
                new_date += ("0"+i)
            else:
                new_date += i

    return new_date

for roommate_copy in roommates_to_copy_dir:
    for file in files_to_copy:
        if file == files_to_copy[0]: # if duke
            from_date = convert_date(duke_from)
            to_date = convert_date(duke_to)
            if to_date[:2] == "12": #if ending is dec then change year on starting
                from_date = from_date[:4]
                last_year = to_date[4:]
                from_date += last_year
            destination = os.path.join(roommate_copy, str("Duke_"+str(from_date)+"-"+str(to_date)+"_"+"$"+str(duke_amount)+".pdf"))
            duke_new_name = str("Duke_"+str(from_date)+"-"+str(to_date)+"_"+"$"+str(duke_amount)+".pdf")
        else:
            from_date = convert_date(spectrum_from)
            to_date = convert_date(spectrum_to)
            if to_date[:2] == "12": #if ending is dec then change year on starting
                from_date = from_date[:4]
                last_year = to_date[4:]
                from_date += last_year
            destination = os.path.join(roommate_copy, str("Spectrum_"+str(from_date)+"-"+str(to_date)+"_"+"$"+str(spectrum_amount)+".pdf"))
            spectrum_new_name = str("Spectrum_"+str(from_date)+"-"+str(to_date)+"_"+"$"+str(spectrum_amount)+".pdf")
        shutil.copy(file, destination)

excel_name += "$"+str(total_amount)+"_"+str(todays_month)+"."+str(todays_day)+"."+str(todays_year)+".xlsx"
excel_path = os.path.join(folder_path, excel_name)
create_xlsx(excel_path, spectrum_bill_date, spectrum_from, spectrum_to, spectrum_amount, duke_bill_date, duke_from, duke_to, duke_amount)

duke_new_path = os.path.abspath(os.path.join(folder_path, duke_new_name))
os.rename(duke_path, duke_new_path)
spectrum_new_path = os.path.abspath(os.path.join(folder_path, spectrum_new_name))
os.rename(spectrum_path, spectrum_new_path)
for roommate_dir in roommates_to_copy_dir:
    shutil.copy(excel_path, roommate_dir)

print(Fore.GREEN+"Successfully created all files. Xlsx filled out")
print(Style.RESET_ALL)
print(f"Double check total amounts just in case. Total for this period: \nSpectrum: ${spectrum_amount}     Duke: ${duke_amount}     50%: ${round(total_amount/2, 2)}")

input("Press Enter to exit")

