import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from datetime import date
import sys, os
from nameparser import HumanName

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)


def check_before_using_dict(column, dictionary, df):
    unique_column_values = df[column].dropna().unique()

    keys = dictionary.keys()
    missing = [i for i in unique_column_values if i not in keys]

    if bool(missing):
        print("*--------------Alert Missing Mapping-------------------------------------*")
        print(f"column: {column}")
        print(f"dict: {dictionary}")
        print(f"missing values: {missing}")
        sys.exit("")


def read_file(file_name, path):
    return pd.read_csv(path / file_name)


def determine_file_name(number):
    str_date = date.today().strftime("%m_%d_%y")
    fn = f"deposit_slip_{str_date}_batch_{number}.xlsx"
    return fn


if __name__ == '__main__':

    # paths
    path_data = Path("../data")
    path_google_drive_data = Path("G:/My Drive/programming/data")
    #path_out = Path("C:/Users/liebism/Desktop/audit_slips")
    path_out = Path("G:/My Drive/check_deposit_slips/")

    path_desktop = Path("C:/Users/liebism/Desktop/batches")

    # get batch number and set up file in and out names
    batch_number = input("enter batch number:")
    fn_in = "batch_"+batch_number+".csv"
    fn_out = determine_file_name(batch_number)

    # read file
    df = read_file(fn_in, path_desktop)

    # conditionally remove rows with NA as amount
    mask = df['Amount'].notna()
    df = df[mask]

    # filter for checks
    df['Pay method'] =  df['Pay method'].str.lower()
    mask = df['Pay method'].str.contains('check|other')
    df = df[mask]

    # set up dict for fund name to code
    temp = read_file("fund_to_code.csv",  path_google_drive_data)
    d = dict(zip(temp.Fund, temp.Code))

    df['Amount'] = df['Amount'].str.replace('$', '').str.replace(',', '').astype(float)
    sr = df.groupby('Fund')['Amount'].sum()
    total_amount = df['Amount'].sum()

    df_2 = sr.to_frame().reset_index()
    check_before_using_dict('Fund', d, df_2)

    df_2['Account #'] = df_2['Fund'].map(d)
    number_of_entries = df.shape[0]
    check_number_adjustment = int(input("check adjustment:"))
    number_of_checks = number_of_entries - check_number_adjustment

    workbook = load_workbook( path_data / "template_deposit_slip.xlsx")
    sheet = workbook.active

    # set up date
    today = date.today()
    sheet['C6'] = today.strftime("%m/%d/%y")
    sheet['C7'] = "Advancement"
    sheet['C8'] = "Michael Liebis"
    sheet['C9'] = total_amount
    sheet['C10'] = number_of_checks

    d = {0: "B",
         1: "C",
         2: "D"}

    df_2 = df_2[['Account #', 'Fund', 'Amount']]

    # populate the spreadsheet
    for i, row in df_2.iterrows():
        for j, value in enumerate(row):
            r, c = i + 14, d[j]
            cell_coordinates = f"{c}{r}"
            sheet[cell_coordinates] = value

    # print out the number of checks per amount
    print(df.groupby('Amount').size())

    #save and open
    workbook.save(path_out / fn_out)
    os.startfile(path_out / fn_out)

    # rename and move check scan
    prefix = input("check file prefix, n for skip: ")
    if prefix != 'n':

        path_google_drive = Path("G:/My Drive/")
        path_deposit_checks = Path("G:/My Drive/deposit checks")
        files = os.listdir(path_google_drive)



        temp = [filename for filename in files if filename.startswith(prefix) and filename.endswith('.pdf')]
        fn = temp[0]

        new_fn = batch_number + "_Batch_" + fn

        os.rename(path_google_drive / fn, path_deposit_checks / new_fn)







    # print files in path





