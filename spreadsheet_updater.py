import tkinter.ttk

import selenium.common.exceptions as exceptions  # To use SessionNotCreatedException from selenium
from selenium import webdriver  # To use the ChromeDriver
from selenium.webdriver.chrome.options import Options  # To set selenium options

import gspread  # Handles spreadsheet
from oauth2client.service_account import ServiceAccountCredentials  # Handles bot account

from tkinter import messagebox  # Pop up messages
from tkinter import *  # GUI
from tkinter.ttk import *


# Creates a chromedriver, checking for 2 errors, and returns a driver to be used in other functions
def create_driver():
    options = Options()  # Selenium options
    options.add_argument('--headless')  # No Chrome popup
    try:
        driver = webdriver.Chrome(executable_path="/usr/local/bin/chromedriver",
                                  options=options)  # Trying to create driver
        return driver  # returns driver if successful

    except exceptions.SessionNotCreatedException as error:  # This is an exception from Selenium
        messagebox.showinfo("Please Update ChromeDriver",  # If wrong version, throw error and give url for download
                            f"{str(error)[30:132]}\n\n"  # relevant section of Selenium exception
                            f"Please RE-DOWNLOAD the correct ChromeDriver version from:\n"
                            f"https://chromedriver.chromium.org/downloads\n\n"  # Chrome driver website
                            f"REPLACE 'ChromeDriver' executable in the following path with the new one:\n"
                            f"\n/usr/local/bin\n"
                            f"\n⌘+⇧+G on Finder (use above path)")  # Correct path to place ChromeDriver for program
        return  # ends function

    except exceptions.WebDriverException:  # This is another exception from Selenium
        messagebox.showinfo("ChromeDriver Not Found",  # If ChromeDriver is in the wrong path, throw error
                            "ChromeDriver needs to be moved to:\n"
                            "\n/usr/local/bin\n"  # Required (Target) Path
                            "\n⌘+⇧+G on Finder (use above path)")  # Instruction for user
        return  # ends function


# Gets the text elements of every element in every given class and returns a 1D list of all these text
def scrape_texts(driver_given_target, classes):
    ret = []
    for cls in classes:  # loop through all the relevant classes, ex. name, price, date
        elements = driver_given_target.find_elements_by_class_name(cls)  # get all elements of cls in target
        for e in elements:  # loop through each element that was found
            s = e.text  # get text attribute of each element
            if s:  # skipping blank texts
                ret.append(s)
    return ret


# Gets and processes Thai stock data from Google Finance for given quotes
def th_stocks():  # uses GoogleFinance
    driver = create_driver()
    th_quotes = set(get_latest_config("s"))
    relevant_classes = ["kHAtIb", "YMlKec.fxKbKc", "ygUjEc"]  # "P2Luy.ZYVHBb" for change, "M2CUtd"[5] for P/E ratio
    final = []

    for quote in th_quotes:
        target = f"https://www.google.com/finance/quote/{quote}"  # New target
        driver.get(target)  # driver has now been given the target website
        driver.implicitly_wait(3)  # Ensure page load
        final.append(scrape_texts(driver, relevant_classes))  # passed driver given target and the relevant classes

    driver.quit()

    for lst in final:
        lst[1] = float(lst[1][1:])  # removing Baht sign
        lst[2] = " ".join(lst[2].split()[:4])  # processing date and time

    return sorted(final, key=lambda ar: ar[0])  # sorting by first index in each sub-array, i.e., name of stock


# Gets and processes mutual fund data from thaifundstoday.com for given funds
def th_mutual_funds():  # uses thaifundstoday.com
    driver = create_driver()
    funds = set(get_latest_config("f"))
    relevant_classes = ["span7", "unchanged", "date"]  # "text-error.price-change" for change
    final = []

    for fund in funds:
        target = f"http://www.thaifundstoday.com/en/funds/{fund}"  # new target
        driver.get(target)
        driver.implicitly_wait(3)  # ensure page load
        final.append(scrape_texts(driver, relevant_classes))

    driver.quit()

    for lst in final:
        lst[1] = float(lst[1][1:])  # removing Baht sign
        lst[2] = lst[2][6:]  # processing date and time

    return sorted(final, key=lambda ar: ar[0])  # sorting by first index in each sub-array, i.e., name of fund


# Collects 5 tables of data on the last 5 years for given quotes from Google Finance
def five_year_summary():
    driver = create_driver()
    quotes = get_latest_config("5")
    working_quotes = []
    non_working = []
    all_data = []  # holds all quote summaries

    for quote in quotes:
        driver.get(f"https://www.google.com/finance/quote/{quote}")

        # Clicking "Annual" button
        try:  # try because user may have entered an invalid quote in the config
            found_buttons = driver.find_elements_by_class_name("VfPpkd-vQzf8d")
            found_buttons[3].click()  # click annual button
            working_quotes.append(quote)
        except:
            non_working.append(quote)  # for troubleshooting later
            continue  # skip to next iteration of quote for loop

        # Clicking year buttons and processing table data
        # need to find buttons again, they changed. next line takes last 5 buttons because those are the "year" buttons
        found_buttons = driver.find_elements_by_class_name("VfPpkd-vQzf8d")[-5:]  # last 5 buttons are year buttons
        found_years = [item.text for item in found_buttons]  # to account for changing years
        quote_summary = [["Info"] + found_years] + [["0" for x in range(6)] for x in range(8)]  # placeholders

        # enumerate buttons to keep track of the columns in quote_summary
        for col_number, button in enumerate(found_buttons, start=1):
            button.click()  # updates target website's table
            rows = driver.find_elements_by_class_name("roXhBd")[1:]  # rows of the table, excluding first (header) row

            # use enumerate to get appropriate row number that needs to be edited in quote_summary
            for row_num, item in enumerate(rows, start=1):  # start at 1 because row 0 is already set
                my_row = item.text.split()[:-1]
                if col_number == 1:
                    quote_summary[row_num][0] = " ".join(my_row[:-1])  # first column set to the words (index=0)
                    quote_summary[row_num][1] = reformat_number(my_row[-1])  # second column set to first year data
                else:
                    quote_summary[row_num][col_number] = reformat_number(my_row[-1])  # third column onward

        all_data.append(quote_summary)  # contains 5 tables of data on the last 5 years for each company

    driver.quit()
    return list(zip(working_quotes, all_data))


# Re-formats numbers ending with T (10^12), B (10^9), or M (10^6) into numbers
def reformat_number(string_num):
    if string_num[-1] == "T":
        return eval(string_num[:-1] + "*10**12")
    elif string_num[-1] == "B":
        return eval(string_num[:-1] + "*10**9")
    elif string_num[-1] == "M":
        return eval(string_num[:-1] + "*10**6")
    else:
        return string_num


# Updates spreadsheet with new data
# Needed to use Google Cloud Platform and Activate Google Drive and Google Sheets APIs.
# Also had to make service account and download "Key" as a JSON file and rename to "client_secret.json"
def update_spreadsheet():
    messagebox.showinfo("Updating", "Updating\nPlease allow up to 1 minute\nClick OK to continue")

    # creating credentials and using them to create a client to interact with spreadsheet
    credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json')
    client = gspread.authorize(credentials)

    # opening sheet by name
    big_sheet = client.open("Investments")

    # editing th_data sheet in the spreadsheet
    current_sheet = big_sheet.worksheet("th_data")
    clear = [["" for x in range(3)] for i in range(999)]

    stocks = th_stocks()
    current_sheet.update('A2:C1000', clear)  # clearing sheet first, in case there are changes in config
    current_sheet.update('A2:C1000', stocks)

    mfs = th_mutual_funds()
    current_sheet.update('E2:G1000', clear)  # clearing sheet first, in case there are changes in config
    current_sheet.update('E2:G1000', mfs)


# Gets five year summary info and puts it into the spreadsheet
def update_five_year_summary():
    messagebox.showinfo("Updating", "Updating\nPlease allow up to 1 minute\nClick OK to continue")

    # creating credentials and using them to create a client to interact with spreadsheet
    credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json')
    client = gspread.authorize(credentials)

    # opening sheet by name
    big_sheet = client.open("Investments")

    # editing five_year_summaries sheet in the spreadsheet
    current_sheet = big_sheet.worksheet("five_year_summaries")
    data = five_year_summary()

    clear = [["" for x in range(7)] for i in range(999)]  # cannot actually clear, so we create a table of empty str
    current_sheet.update("A2:G1000", clear)  # clearing sheet first, in case there are changes in config
    current_sheet.update("B1:G1", [data[0][1][0]])  # sets first row

    start_row, end_row = 2, 9
    for name, table in data:
        current_sheet.update(f"A{start_row}:A{end_row}", [[name] for i in range(8)])  # fill column with quote
        current_sheet.update(f"B{start_row}:G{end_row}", table[1:])  # ignore first row, put table into spreadsheet
        start_row += 8
        end_row += 8


# Gets latest user settings from "config" sheet in their spreadsheet
def get_latest_config(choice):
    # creating credentials and using them to create a client to interact with spreadsheet
    credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json')
    client = gspread.authorize(credentials)

    # opening spreadsheet by name
    big_sheet = client.open("Investments")

    # opening config sheet
    current_sheet = big_sheet.worksheet("config")

    # getting all values as a list of lists (sub-lists are for each row, commas separate columns)
    configs = current_sheet.get_all_values()

    stocks = [item[0] for item in configs if item[0]]  # first column
    funds = [item[1] for item in configs if item[1]]  # second column
    five_year_summary_quotes = [item[2] for item in configs if item[2]]  # third column

    # returning requested set only, otherwise, returning all
    if choice == "s":
        return stocks
    elif choice == "f":
        return funds
    elif choice == "5":
        return five_year_summary_quotes
    else:
        return stocks, funds, five_year_summary_quotes


# Creates Tkinter window for user to click button and shows a messagebox
def start_window():
    root = tkinter.Tk()
    root.title("Spreadsheet Updater")
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.geometry(f"350x350+{screen_width // 2 - 175}+{screen_height // 2 - 175}")

    update = Button(root, text="Click to Update Spreadsheet!", command=update_spreadsheet)
    update.pack(fill="both", expand=True, anchor=CENTER, padx=50, pady=10)

    five_year = Button(root, text="Click to Update 5 Year Summary!", command=update_five_year_summary)
    five_year.pack(fill="both", expand=True, anchor=CENTER, padx=50, pady=10)

    root.mainloop()


start_window()
