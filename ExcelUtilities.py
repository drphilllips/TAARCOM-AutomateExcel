import os

import numpy as np
import pandas as pd
from xlrd import XLRDError

import EnumTypes


default_sheet_name = "Data"


def saveError(*excel_files):
    """Checks for obstacles with saving the output file

    :param: excel_files: path to output directory
    :return: whether there is an error with saving
    """
    for file in excel_files:
        try:
            open(file, 'r+')
        except FileNotFoundError:
            pass
        except PermissionError:
            return True
    return False


def createExcelFile(filename, sheet_data):
    """Creates an Excel file from a dataframe

    :param filename: name for our created file
    :param sheet_data: dataframe which will be copied to this file
    :return: xlsxwriter object for formatting this file
    """

    # Verify output path
    out_dir = "W:/Output/"
    out_path = out_dir + filename
    if saveError(out_path):
        print("..One or more files are currently open in Excel!\n"
              "..Please close the files and try again.\n"
              "*Program Terminated*")
        return

    # Write the output file
    writer = pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd")
    sheet_data.to_excel(writer, sheet_name=default_sheet_name, index=False)

    print("> New file saved at: " + out_path)

    return writer


def loadLookupFile(filename, sheet_name):
    """Loads the specified sheet from the lookup file to a dataframe

    :param: filename: name of the lookup file
    :param: sheet_name: name of the main sheet we pull data from
    :return: dataframe with sheet data
    """

    # Assume file is in the lookup directory
    look_dir = "W:/Lookup/"
    filepath = look_dir + filename

    if os.path.exists(filepath):
        try:
            sheet_data = pd.read_excel(filepath, sheet_name).fillna("")
        except XLRDError:
            print("..Error reading sheet name for " + filename + "!\n"
                  "..Please make sure the main tab is named \"" + sheet_name + "\".\n"
                  "*Program Terminated*")
            return
    else:
        print("..No " + filename + " file found!\n"
              "..Please make sure " + filename + " is in the directory.\n"
              "*Program Terminated*")
        return

    return sheet_data


def formatSheet(sheet_data, writer):
    """Formats our output file to make it look nice :)

    :param: sheet_data: working data frame for output
    :param: writer: working xlsxwriter for Excel tools
    :return: formatted Excel file
    """

    # If there is nothing to format, return
    if sheet_data.shape[0] == 0:
        return
    # Store the working sheet from the output Excel file
    sheet = writer.sheets[default_sheet_name]
    # Freeze header so it remains stationary when scrolling up or down
    sheet.freeze_panes(1, 0)
    # Set auto filter
    sheet.autofilter(0, 0, sheet_data.shape[0], sheet_data.shape[1]-1)
    # Ignore number stored as text error
    sheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})

    # Define all the different format options we will need
    fmt_default = writer.book.add_format({'font': 'Calibri',
                                          'font_size': 11})
    fmt_center_aligned = writer.book.add_format({'font': 'Calibri',
                                                 'font_size': 11,
                                                 'align': 'center'})
    fmt_right_aligned = writer.book.add_format({'font': 'Calibri',
                                                'font_size': 11,
                                                'align': 'right'})
    fmt_accounting = writer.book.add_format({'font': 'Calibri',
                                             'font_size': 11,
                                             'num_format': 44})
    fmt_number_with_commas = writer.book.add_format({'font': 'Calibri',
                                                     'font_size': 11,
                                                     'num_format': 3})
    fmt_individual = writer.book.add_format({'font': 'Calibri',
                                             'font_size': 11,
                                             'bg_color': '#ccc0da'})
    fmt_new_account = writer.book.add_format({'font': 'Calibri',
                                              'font_size': 11,
                                              'bg_color': 'yellow'})
    fmt_out_of_territory = writer.book.add_format({'font': 'Calibri',
                                                   'font_size': 11,
                                                   'bg_color': '#ff5050'})
    fmt_proper_name_not_associated = writer.book.add_format({'font': 'Calibri',
                                                             'font_size': 11,
                                                             'bg_color': '#99ff66'})
    fmt_proper_name_not_found = writer.book.add_format({'font': 'Calibri',
                                                        'font_size': 11,
                                                        'bg_color': '#66ffff'})

    # Determine which columns need which format
    accounting_cols = ['Unit Price', 'Invoiced Dollars']
    number_with_commas_cols = ['Quantity']
    definitely_text_cols = ['Customer Class']  # Make sure it's not interpreted as a number
    center_aligned_cols = ['OSR', 'Reported Distributor']
    right_aligned_cols = ['Zip Code', 'Phone']

    # -------------------------
    #  Format and size columns
    # -------------------------

    for col in sheet_data.columns:
        # Setting each column's style
        fmt = fmt_default
        if col in accounting_cols:
            fmt = fmt_accounting
        elif col in number_with_commas_cols:
            fmt = fmt_number_with_commas
        elif col in definitely_text_cols:
            fmt = fmt_default
        elif col in center_aligned_cols:
            fmt = fmt_center_aligned
        elif col in right_aligned_cols:
            fmt = fmt_right_aligned

        # Find column width by largest item in that column
        col_width = max(sheet_data[col].astype(str).map(len).max(), len(col)) + 5
        col_width = min(col_width, 35)  # Max width
        # Set column width and formatting
        col_idx = sheet_data.columns.get_loc(col)
        sheet.set_column(col_idx, col_idx, col_width, fmt)

    # -----------------------
    #  Flag individual cells
    # -----------------------

    # +++ Flags +++
    # 1) Purple: customer classed as "individual"
    # 2) Yellow: new account, not found in rootCustomerMappings
    # 3) Red: out of territory, not found in CAZipCode
    # 4) Green: account found in rootCustomerMappings, but not assigned a proper name

    try:
        customer_col_index = list(sheet_data).index('Reported Customer')
        row_index = 1
        for i in sheet_data.index:
            # Filling customer cell with one of these values
            customer_name = sheet_data.loc[i, 'Name']
            customer_company = sheet_data.loc[i, 'Reported Customer']
            # Read flags to determine format
            flag = sheet_data.loc[i, 'Flag']

            # Format based on flags
            # Order of if statements determines precedence of flags
            try:
                if flag == EnumTypes.Flag.OOT.value and not customer_company:  # Out of territory + individual
                    sheet.write(row_index, customer_col_index, customer_name, fmt_out_of_territory)
                elif flag == EnumTypes.Flag.OOT.value:  # Out of territory
                    sheet.write(row_index, customer_col_index, customer_company, fmt_out_of_territory)
                elif flag == EnumTypes.Flag.CNP.value:  # Individual
                    sheet.write(row_index, customer_col_index, customer_name, fmt_individual)
                elif flag == EnumTypes.Flag.CNF.value:  # New account: customer not found in map
                    sheet.write(row_index, customer_col_index, customer_company, fmt_new_account)
                elif flag == EnumTypes.Flag.PNA.value:  # Proper name not associated: customer found, no proper name
                    sheet.write(row_index, customer_col_index, customer_company, fmt_proper_name_not_associated)
                elif flag == EnumTypes.Flag.PNF.value:  # Proper name in customer-map but not in master acct list
                    sheet.write(row_index, customer_col_index, customer_company, fmt_proper_name_not_found)
            except:  # Handle NaN values
                pass

            row_index += 1
    except:
        print("..Unable to format with flags")

    # ---------------
    #  Abracon Flags
    # ---------------

    if 'Abracon Flag' in sheet_data.columns:
        fmt_abr_yellow = writer.book.add_format({'font': 'Calibri',
                                                 'font_size': 11,
                                                 'bg_color': EnumTypes.AbraconFlags.YELLOW_HEX.value})
        fmt_abr_green = writer.book.add_format({'font': 'Calibri',
                                                'font_size': 11,
                                                'bg_color': EnumTypes.AbraconFlags.GREEN_HEX.value})
        fmt_abr_orange = writer.book.add_format({'font': 'Calibri',
                                                 'font_size': 11,
                                                 'bg_color': EnumTypes.AbraconFlags.ORANGE_HEX.value})

        row_index = 1
        abr_flag_col_index = list(sheet_data).index('Abracon Flag')

        for i in sheet_data.index:
            abr_flag = sheet_data.loc[i, 'Abracon Flag']

            if abr_flag == EnumTypes.AbraconFlags.YELLOW_DESC.value:
                sheet.write(row_index, abr_flag_col_index, abr_flag, fmt_abr_yellow)
            elif abr_flag == EnumTypes.AbraconFlags.GREEN_DESC.value:
                sheet.write(row_index, abr_flag_col_index, abr_flag, fmt_abr_green)
            elif abr_flag == EnumTypes.AbraconFlags.ORANGE_DESC.value:
                sheet.write(row_index, abr_flag_col_index, abr_flag, fmt_abr_orange)

            row_index += 1

