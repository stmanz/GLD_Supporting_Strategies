import calendar
import datetime
import time
import glob
import math
import ntpath
import os
import sys
import shutil
from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename

import pandas as pd


# import sys
# sys.modules[__name__].__dict__.clear()
#######################################################################################################################
# Functions
def bank_account_name(account_name_loc):
    # Change Deposit Account name
    acc_name_prev = account_name_loc['Account'].iloc[0]
    name_string = ''.join([p for p in acc_name_prev if not p.isdigit()]).strip()
    gl_num = acc_name_prev[:4]
    if acc_name_prev.find('#') == -1:
        acct_num = gl_num
        bank_name = name_string + ' #' + acct_num
    else:
        pound_loc = acc_name_prev.find('#')
        acct_num = acc_name_prev[pound_loc + 1:pound_loc + 5]
        bank_name = name_string + acct_num

    return gl_num, acct_num, bank_name


# Create letter for column number
def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


# Define formatting for XLSWRITER
def formatting(sheet, col_val):
    # get the xlsx writer workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets[sheet]

    # Add accountancy format
    accounting = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)'})

    # Set format without assigning column width
    worksheet.set_column(col_val + ':' + col_val, None, accounting)


# Define a function to automatically fit columns
def get_col_widths(df, sn):
    xl_sht_name = str(sn)

    for col in df:
        column_wid = max(df[col].astype(str).map(len).max(), len(col))
        col_index = df.columns.get_loc(col)
        writer.sheets[xl_sht_name].set_column(col_index, col_index, math.ceil(column_wid * 1.25))


#######################################################################################################################

# Ask user to define current working directory
print('Choose directory\n')
prev_dir = os.getcwd()
Tk().withdraw()
directory = filedialog.askdirectory()
os.chdir(directory)
if not prev_dir == os.getcwd():
    print('Current Working Directory changed to :')
    print(os.getcwd())
    print()

# Find the correct excel file in path.
types = ('*.xlsx', '*.csv')  # the tuple of file types
file_list_glob = []
for files in types:
    file_list_glob.extend(glob.glob(files))

file_list = [file_list for file_list in file_list_glob if '~' not in file_list]
input_file_list = [input_file for input_file in file_list if 'Daily' in input_file]
if input_file_list:
    print('Possible files to choose from are:\n')
    for i in input_file_list:
        print(i)
        print()
    print()
else:
    print('Choose the file that holds all GLD information.Possibly named in the following format:')
    print('Daily_WL_GL_20220202.csv')
    print()

# Ask user to define file
Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file

# Create input
input_file = ntpath.basename(filename)
print()
print('Chosen file: ' + input_file)
print()
if not os.getcwd() == os.path.dirname(os.path.abspath(filename)):
    shutil.copyfile(filename, os.getcwd() + '\\' + input_file)
    print(input_file + ' moved to current directory.')
    print()

file_numbers = ''.join([p for p in input_file if p.isdigit()]).strip()
file_yr = int(file_numbers[:4])
file_mo = int(file_numbers[4:6])
file_day = int(file_numbers[6:])
output_file = 'GLD_Output_' + file_numbers

# Create a string for due date
# If December, restart month at January
if file_mo == 12:
    due = datetime.datetime(file_yr + 1, 1, 5, 0, 0)
else:
    due = datetime.datetime(file_yr, file_mo + 1, 5, 0, 0)
due = due.strftime('%m/%d/%Y')
datetime_object = datetime.datetime.strptime(str(file_mo), '%m')
full_month_name = datetime_object.strftime('%B')

# Read in file
[file_name, file_ext] = os.path.splitext(input_file)
raw_data = 0
if file_ext == '.xlsx':
    raw_data = pd.read_excel(input_file)

    # Change format of date in Date column
    raw_data['Date'] = raw_data['Date'].dt.strftime('%m/%d/%Y')

elif file_ext == '.csv':
    raw_data = pd.read_csv(input_file)

# Separate by account and type.
Type = ['Check', 'Deposit', 'Transfer', 'Journal', 'Trust check detail']
# Create a dataframe only including these types for later use.
raw_data = raw_data[raw_data['Type'].str.contains('|'.join(Type))]
# Change accounting fields to floats.
d_loc = raw_data.columns.get_loc('Debit')
c_loc = raw_data.columns.get_loc('Credit')
raw_data[raw_data.columns[d_loc:c_loc + 1]] = raw_data[raw_data.columns[d_loc:c_loc + 1]].replace('[\$,]', '',
                                                                                                  regex=True).astype(
    float)
# Determine all account names
account = raw_data['Account'].unique()

# Initialization
Checks_op = pd.DataFrame()
Checks_tr = pd.DataFrame()
Deposits = pd.DataFrame()
Transfers = pd.DataFrame()
Journals = pd.DataFrame()
Payments = pd.DataFrame()
Invoices = pd.DataFrame()
ref_dep = 1  # This will change in the future
ref_journal = 100  # This will change in the future
ref_invoice = 500
ref_pay = 1000

for typ in Type:
    # Separate by account
    List = raw_data[raw_data['Type'].str.contains(typ, regex=False)]

    # If list ends up being empty, just continue
    if List.empty:
        continue

    # Date Extraction
    date_split = List['Date'].iloc[0].split('/')
    month = int(date_split[0])
    year = int(date_split[2])
    # Find number of days in the invoice month
    end_month = calendar.monthrange(year, month)

    # Create a string for invoice date
    invoice = datetime.datetime(year, month, end_month[1], 0, 0)
    invoice = invoice.strftime('%m/%d/%Y')

    # Checks
    if typ == 'Check':

        # Separate based on Operating and Trust
        check_op = List[List['Account'].str.contains('Operating', regex=False)]
        check_trust = List[List['Account'].str.contains('Trust Account', regex=False)]

        if not check_op.empty:
            # Operating
            for i in range(len(check_op['Type'])):
                # Reference Number
                ref_check = check_op['No'].iloc[i]
                # Change Check Account name
                acc_check_op = bank_account_name(check_op.iloc[[i]])

                Checks_op_new = pd.Series(
                    [ref_check, acc_check_op[0], check_op['Date'].iloc[i], check_op['To/From'].iloc[i],
                     check_op['Credit'].iloc[i], check_op['Description'].iloc[i], check_op['Description'].iloc[i],
                     acc_check_op[1], 'Pawley\'s Island'])
                Checks_op = Checks_op.append(Checks_op_new, ignore_index=True)

        if not check_trust.empty:
            # Trust
            for i in range(len(check_trust['Type'])):
                # Reference Number
                ref_check = check_trust['No'].iloc[i]
                # Change Check Account name
                acc_check_trust = bank_account_name(check_trust.iloc[i])

                # Account
                account_name = 'Trust Liability Account:' + check_trust['Matter'].iloc[i]

                Checks_tr_new = pd.Series(
                    [ref_check, acc_check_trust[2], check_trust['Date'].iloc[i], check_trust['To/From'].iloc[i],
                     check_trust['Credit'].iloc[i], check_trust['Matter'].iloc[i], check_trust['Description'].iloc[i],
                     account_name, 'Pawley\'s Island'])
                Checks_tr = Checks_tr.append(Checks_tr_new, ignore_index=True)

    # Deposits and Payments
    if typ == 'Deposit':

        # Separate based on Deposits and Payments
        deposits = List[List['Account'].str.contains('Trust', regex=False)]
        payments = List[List['Account'].str.contains('Operating', regex=False)]

        if not deposits.empty:
            # Deposits
            for i in range(len(deposits['Type'])):
                # Account
                account_name = 'Trust Liability Account:' + deposits['Matter'].iloc[i]
                # Change Deposit Account name
                acc_deposits = bank_account_name(deposits.iloc[[i]])

                Deposit_new = pd.Series(
                    [ref_dep, deposits['Date'].iloc[i], deposits['No'].iloc[i], acc_deposits[2],
                     'Pawley\'s Island', deposits['Description'].iloc[i], deposits['To/From'].iloc[i],
                     deposits['No'].iloc[i], deposits['Debit'].iloc[i], account_name, '', ''])
                Deposits = Deposits.append(Deposit_new, ignore_index=True)

                # Increment Reference Number
                ref_dep += 1

        if not payments.empty:
            # Payments
            for i in range(len(payments['Type'])):
                # Change Deposit Account name
                acc_payments = bank_account_name(payments.iloc[[i]])

                Payments_new = pd.Series(
                    [ref_pay, payments['Date'].iloc[i], payments['No'].iloc[i], payments['To/From'].iloc[i],
                     acc_payments[0], payments['Debit'].iloc[i], payments['Description'].iloc[i], '', '', '', ''])
                Payments = Payments.append(Payments_new, ignore_index=True)

                # Increment Reference Number
                ref_pay += 1

    # Transfers
    if typ == 'Transfer':

        # Find matching amounts for transfer
        amount_match = List['Debit'].unique()

        # Separate by operating and trust
        account_type_operating = List[List['Account'].str.contains('Operating', regex=False)]
        account_type_trust = List[List['Account'].str.contains('Trust', regex=False)]

        try:
            # if not account_type_operating.empty and not account_type_trust.empty:
            for i in range(len(account_type_operating['Type'])):
                # Change Deposit Account name
                acc_operating = bank_account_name(account_type_operating.iloc[[i]])
                acc_trust = bank_account_name(account_type_trust.iloc[[i]])

                # Determine operating matter and find trust counterpart.
                op_matter = account_type_operating['Matter'].iloc[0]

                trust_counter = account_type_trust[account_type_trust['Matter'] == op_matter]

                # ToAccount
                to_account = 'Trust Liability Account:' + account_type_trust['Matter'].iloc[i]

                # Private Note
                private = 'Transfer from ' + acc_trust[2] + ' to ' + acc_operating[2]

                Transfers_new = pd.Series([account_type_operating['Date'].iloc[i], private, to_account, acc_trust[2],
                                           account_type_operating['Debit'].iloc[i]])
                Transfers = Transfers.append(Transfers_new, ignore_index=True)
                break
        except KeyError:
            sys.exit(
                'Oops! It looks like you may have an Operating entry without a corresponding Trust entry or vice versa!')

    # Journal Entries
    if typ == 'Journal':
        if not List.empty:
            for i in range(len(List['Type'])):

                # If Debit, amount is positive. If Credit, amount is negative.
                if pd.notna(List['Credit'].iloc[i]):
                    # Credit
                    amount = -List['Credit'].iloc[i]

                    # Change Deposit Account name
                    acc_journal = bank_account_name(List.iloc[[i]])

                    # Find corresponding Journal entry
                else:
                    # Debit
                    amount = List['Debit'].iloc[i]

                    # Change Deposit Account name
                    acc_journal = bank_account_name(List.iloc[[i]])

                Journals_new = pd.Series([ref_journal, List['Date'].iloc[i], List['No'].iloc[i], acc_journal[0], amount,
                                          List['Description'].iloc[i], 'Pawley\'s Island', '', '', '', '', ''])
                Journals = Journals.append(Journals_new, ignore_index=True)

    # # Invoice
    # if typ == 'Invoice':
    #
    #     # Separate by Soft Cost and Professional Fees
    #     soft = List[List['Account'].str.contains('Soft', regex=False)]
    #     professional = List[List['Account'].str.contains('Professional', regex=False)]
    #
    #     # Change Deposit Account name
    #     [acc_num_professional, acc_name_professional, acc_name_new_professional] = bank_account_name(professional)
    #     [acc_num_soft, acc_name_soft, acc_name_new_soft] = bank_account_name(soft)
    #
    #     # Create an instance to increase the reference number when a new customer comes along
    #     prev_customer = List['Matter'].iloc[0]
    #
    #     # Professional
    #     for i in range(len(professional['Type'])):
    #         # Does previous customer match current customer? If not, increment reference number.
    #         if prev_customer != professional['Matter'].iloc[i]:
    #             ref_invoice += 1
    #
    #         Invoice_new = pd.Series(
    #             [ref_invoice, '', invoice, due, professional['Matter'].iloc[i], professional['Date'].iloc[i],
    #              acc_name_new_professional, professional['Description'].iloc[i], '', '', ''])
    #         Invoices = Invoices.append(Invoice_new, ignore_index=True)
    #
    #         # Save customer name
    #         prev_customer = professional['Matter'].iloc[i]
    #
    #     # Soft
    #     for i in range(len(soft['Type'])):
    #         # Does previous customer match current customer? If not, increment reference number.
    #         if prev_customer != soft['Matter'].iloc[i]:
    #             ref_invoice += 1
    #
    #         Invoice_new = pd.Series(
    #             [ref_invoice, '', invoice, due, soft['Matter'].iloc[i], soft['Date'].iloc[i],
    #              acc_name_new_soft, soft['Description'].iloc[i], '', '', ''])
    #         Invoices = Invoices.append(Invoice_new, ignore_index=True)
    #
    #         # Save customer name
    #         prev_customer = soft['Matter'].iloc[i]

# Create column names for every dataframe and make sure all columns that need to be numeric are.
if not Checks_op.empty:
    Checks_op.columns = ['RefNumber', 'BankAccount', 'TxnDate', 'Vendor', 'ExpenseAmount', 'PrivateNote', 'ExpenseDesc',
                         'ExpenseAccount', 'Location']
    Checks_op['RefNumber'] = pd.to_numeric(Checks_op['RefNumber'], errors='coerce')
    Checks_op['BankAccount'] = pd.to_numeric(Checks_op['BankAccount'], errors='coerce')
    Checks_op['ExpenseAccount'] = pd.to_numeric(Checks_op['ExpenseAccount'], errors='coerce')
if not Checks_tr.empty:
    Checks_tr.columns = ['RefNumber', 'BankAccount', 'TxnDate', 'Vendor', 'ExpenseAmount', 'PrivateNote', 'ExpenseDesc',
                         'ExpenseAccount', 'Location']
    Checks_tr['RefNumber'] = pd.to_numeric(Checks_tr['RefNumber'], errors='coerce')
if not Deposits.empty:
    Deposits.columns = ['RefNumber', 'TxnDate', 'PaymentRefNumber', 'DepositToAccount', 'Location', 'PrivateNote',
                        'Entity', 'LineDesc', 'LineAmount', 'Account', 'PaymentMethod', 'Class']
    Deposits['RefNumber'] = pd.to_numeric(Deposits['RefNumber'], errors='coerce')
if not Transfers.empty:
    Transfers.columns = ['TxnDate', 'PrivateNote', 'ToAccount', 'FromAccount', 'Amount']
if not Journals.empty:
    Journals.columns = ['RefNumber', 'TxnDate', 'PrivateNote', 'Account', 'LineAmount', 'LineDesc', 'Location',
                        'Entity', 'Class', 'IsAdjustment', 'Currency', 'ExchangeRate']
    Journals['RefNumber'] = pd.to_numeric(Journals['RefNumber'], errors='coerce')
    Journals['Account'] = pd.to_numeric(Journals['Account'], errors='coerce')
if not Payments.empty:
    Payments.columns = ['RefNumber', 'TxnDate', 'PaymentRefNumber', 'Customer', 'DepositToAccount', 'LineAmount',
                        'PrivateNote', 'InvoiceApplyTo', 'PaymentMethod', 'Currency', 'ExchangeRate']
    Payments['RefNumber'] = pd.to_numeric(Payments['RefNumber'], errors='coerce')
    Payments['DepositToAccount'] = pd.to_numeric(Payments['DepositToAccount'], errors='coerce')
# if not Invoices.empty:
# Invoices.columns = ['RefNumber', 'Customer', 'TxnDate', 'DueDate', 'PrivateNote', 'LineServiceDate', 'LineItem',
#                     'LineDesc', 'LineQty', 'LineUnitPrice', 'LineAmount']
# Invoices[['RefNumber', 'LineQty']] = pd.to_numeric(Invoices[['RefNumber', 'LineQty']])

# Create necessary directories if not already done so.
directory_month = full_month_name + ' Report'
cwd = os.getcwd()
path = cwd + '\\' + directory_month

# Check whether the specified path exists or not
isExist = os.path.exists(path)

if not isExist:
    # Create a new directory because it does not exist
    os.makedirs(path)
    print('The directory ' + directory_month + ' has been created!')
    print()

# Create Sheet Names
sheet_names = ['Checks (O)', 'Checks (T)', 'Deposits(T)', 'Receive-Payments', 'Transfers', 'Journal-Entries',
               'Invoices']

# When writing to excel, format the data columns with accounting format for easier viewing.
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(path + '\\' + output_file + '.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
for i in sheet_names:

    if i == 'Checks (O)':
        Checks_op.to_excel(writer, index=False, sheet_name=i)

        if not Checks_op.empty:
            # Create a list to change formatting of certain columns
            acc_change = column_string(Checks_op.columns.get_loc('ExpenseAmount') + 1)
            formatting(i, acc_change)

            # Note: It isn't possible to format any cells that already have a format such
            # as the index or headers or any cells that contain dates or datetimes.

            # Auto-adjust columns' width
            get_col_widths(Checks_op, i)

            # Delete variables for next iteration
            to_delete = ['acc_change', 'acc_change_names']
            for _var in to_delete:
                if _var in locals() or _var in globals():
                    exec(f'del {_var}')

    elif i == 'Checks (T)':
        Checks_tr.to_excel(writer, index=False, sheet_name=i)

        if not Checks_tr.empty:
            # Create a list to change formatting of certain columns
            acc_change = column_string(Checks_tr.columns.get_loc('ExpenseAmount') + 1)
            formatting(i, acc_change)

            # Note: It isn't possible to format any cells that already have a format such
            # as the index or headers or any cells that contain dates or datetimes.

            # Auto-adjust columns' width
            get_col_widths(Checks_tr, i)

            # Delete variables for next iteration
            to_delete = ['acc_change', 'acc_change_names']
            for _var in to_delete:
                if _var in locals() or _var in globals():
                    exec(f'del {_var}')

    elif i == 'Deposits(T)':
        Deposits.to_excel(writer, index=False, sheet_name=i)

        if not Deposits.empty:
            # Create a list to change formatting of certain columns
            acc_change = column_string(Deposits.columns.get_loc('LineAmount') + 1)
            formatting(i, acc_change)

            # Note: It isn't possible to format any cells that already have a format such
            # as the index or headers or any cells that contain dates or datetimes.

            # Auto-adjust columns' width
            get_col_widths(Deposits, i)

            # Delete variables for next iteration
            to_delete = ['acc_change', 'acc_change_names']
            for _var in to_delete:
                if _var in locals() or _var in globals():
                    exec(f'del {_var}')

    elif i == 'Receive-Payments':
        Payments.to_excel(writer, index=False, sheet_name=i)

        if not Payments.empty:
            # Create a list to change formatting of certain columns
            acc_change = column_string(Payments.columns.get_loc('LineAmount') + 1)
            formatting(i, acc_change)

            # Note: It isn't possible to format any cells that already have a format such
            # as the index or headers or any cells that contain dates or datetimes.

            # Auto-adjust columns' width
            get_col_widths(Payments, i)

            # Delete variables for next iteration
            to_delete = ['acc_change', 'acc_change_names']
            for _var in to_delete:
                if _var in locals() or _var in globals():
                    exec(f'del {_var}')

    elif i == 'Transfers':
        Transfers.to_excel(writer, index=False, sheet_name=i)

        if not Transfers.empty:
            # Create a list to change formatting of certain columns
            acc_change = column_string(Transfers.columns.get_loc('Amount') + 1)
            formatting(i, acc_change)

            # Note: It isn't possible to format any cells that already have a format such
            # as the index or headers or any cells that contain dates or datetimes.

            # Auto-adjust columns' width
            get_col_widths(Transfers, i)

            # Delete variables for next iteration
            to_delete = ['acc_change', 'acc_change_names']
            for _var in to_delete:
                if _var in locals() or _var in globals():
                    exec(f'del {_var}')

    elif i == 'Journal-Entries':
        Journals.to_excel(writer, index=False, sheet_name=i)

        if not Journals.empty:
            # Create a list to change formatting of certain columns
            acc_change = column_string(Journals.columns.get_loc('LineAmount') + 1)
            formatting(i, acc_change)

            # Note: It isn't possible to format any cells that already have a format such
            # as the index or headers or any cells that contain dates or datetimes.

            # Auto-adjust columns' width
            get_col_widths(Journals, i)

            # Delete variables for next iteration
            to_delete = ['acc_change', 'acc_change_names']
            for _var in to_delete:
                if _var in locals() or _var in globals():
                    exec(f'del {_var}')

    elif i == 'Invoices':
        Invoices.to_excel(writer, index=False, sheet_name=i)

        # if not Invoices.empty:
        #     # Create a list to change formatting of certain columns
        #     acc_change_names = ['LineUnitPrice', 'LineAmount']
        #     for change in acc_change_names:
        #         acc_change = column_string(Invoices.columns.get_loc(change) + 1)
        #         formatting(i, acc_change)
        #
        #     # Note: It isn't possible to format any cells that already have a format such
        #     # as the index or headers or any cells that contain dates or datetimes.
        #
        #     # Auto-adjust columns' width
        #     get_col_widths(Invoices, i)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Tell user all is well
print('Code executed successfully! You can go home now.')
time.sleep(2)
print('Self destruct in 30 seconds:')
print()
time.sleep(2)
print('COUNTDOWN:')
print('====================================')
j = 30
for i in range(31):
    time.sleep(1)
    print(j)
    j -= 1
print("     _.-^^---....,,--       ")
print(" _--                  --_  ")
print("<                        >)")
print("|                         | ")
print(" \._                   _./  ")
print("    ```--. . , ; .--'''       ")
print("          | |   |             ")
print("       .-=||  | |=-.   ")
print("       `-=#$%&%$#=-'   ")
print("          | ;  :|     ")
print(" _____.,-#%&$@%#&#~,._____")

time.sleep(4)
print()
print()
print('Still there?!')
print('Good. Go home now. You deserve it.')
time.sleep(4)
