import os

import numpy as np
import pandas as pd
from xlrd import XLRDError
import openpyxl
import xlsxwriter

import EnumTypes
import ExcelUtilities


def main(ins_df, company):
    """Matches insight file columns to standard columns to
    organize the data to TAARCOM, Inc. standards

    :param: ins_df: data frame for the working insight file
    :param: company: company that provided the insight file
    :return: insight file with standardized columns
    """

    # -----------------------------
    #  Load the root column libray
    # -----------------------------

    root_column_library = ExcelUtilities.loadLookupFile("RootColumnLibrary.xlsx", "Standardize Columns")

    # --------------------------------
    #  Create standard file dataframe
    # --------------------------------

    # Pull standard columns from root column library
    header = list(root_column_library)
    # Remove first/last name columns
    header = header[:-2]
    # For Digi-Key: add feedback columns
    if company == EnumTypes.Company.DGK:
        header.insert(3, 'Information for Digi-Key')
        header.insert(3, 'How Contacted')
        header.insert(3, '(Suggested) End Product')
        header.insert(3, 'Must Contact')
    # Add flag column for future use
    header.append('Flag')
    # Create dataframe with standard columns
    std_df = pd.DataFrame(columns=header).fillna("")

    # ------------------------------------------------
    #  Match insight file columns to standard columns
    # ------------------------------------------------

    # Make values in root column library all lowercase so the column look-up is not case-sensitive
    for i in root_column_library.index:
        for root_col in root_column_library.columns:
            root_column_library.loc[i, root_col] = str(root_column_library.loc[i, root_col]).lower()

    # Also lower column headers
    ins_df.columns = [str(col_header).lower() for col_header in ins_df.columns]

    # Store names of insight file columns for later formatting
    first_name_col = ""
    last_name_col = ""
    quantity_col = ""
    unit_price_col = ""
    invoiced_dollars_col = ""
    zip_code_col = ""
    phone_number_col = ""

    for ins_col in ins_df.columns:  # For each insight file column
        for root_col in root_column_library.columns:  # Check if it matches any root column
            if ins_col in [col.lower() for col in root_column_library[root_col].values] or ins_col == root_col.lower():
                # Save insight column name for future use (filling in blanks/formatting)
                if root_col == 'First Name':
                    first_name_col = ins_col
                elif root_col == 'Last Name':
                    last_name_col = ins_col
                elif root_col == 'Quantity':
                    quantity_col = ins_col
                elif root_col == 'Unit Price':
                    unit_price_col = ins_col
                elif root_col == 'Invoiced Dollars':
                    invoiced_dollars_col = ins_col
                elif root_col == 'Zip Code':
                    zip_code_col = ins_col
                elif root_col == 'Phone':
                    phone_number_col = ins_col

                # As long as we are not in the first or last name columns
                if not root_col == 'First Name' and not root_col == 'Last Name':
                    # Copy contents of insight column to matching root column
                    for i in ins_df.index:
                        std_df.loc[i, root_col] = ins_df.loc[i, ins_col]

    # ---------------------------------
    #  Fill in blanks / Format columns
    # ---------------------------------

    for i in std_df.index:
        # Fill distributor/principal based on the company that provided this file (dropdown menu)
        if company == EnumTypes.Company.DGK or company == EnumTypes.Company.MOU:
            std_df.loc[i, 'Reported Distributor'] = company.value
        elif company == EnumTypes.Company.ABR:
            std_df.loc[i, 'Principal'] = company.value

        # Combine first and last name into 'Name'
        if first_name_col and last_name_col:
            std_df.loc[i, 'Name'] = str(ins_df.loc[i, first_name_col]) + " " + str(ins_df.loc[i, last_name_col])

        # Invoiced dollars = quantity * unit price
        if quantity_col and unit_price_col and not invoiced_dollars_col:
            std_df.loc[i, 'Invoiced Dollars'] = float(ins_df.loc[i, quantity_col]) * float(ins_df.loc[i, unit_price_col])
        # Unit price = invoiced dollars / quantity
        if quantity_col and invoiced_dollars_col and not unit_price_col:
            quantity = ins_df.loc[i, quantity_col]
            if quantity:  # Make sure we don't divide by zero
                std_df.loc[i, 'Unit Price'] = float(ins_df.loc[i, invoiced_dollars_col]) / float(quantity)
        # Quantity = invoiced dollars / unit price
        if invoiced_dollars_col and unit_price_col and not quantity_col:
            unit_price = ins_df.loc[i, unit_price_col]
            if unit_price:  # Make sure we don't divide by zero
                std_df.loc[i, 'Quantity'] = float(ins_df.loc[i, invoiced_dollars_col]) / float(unit_price)

        # Format all zip codes to #####-#### (no trailing four zeroes)
        if zip_code_col:
            # Convert to string
            zip_code = str(std_df.loc[i, 'Zip Code'])

            # Filter out all non-numeric characters
            numeric_filter = filter(str.isdigit, zip_code)
            zip_code = "".join(numeric_filter)

            # Trim trailing four zeroes
            if zip_code.endswith("0000"):
                zip_code = zip_code[:-4]
            # If length 9, send to XXXXX-XXXX
            elif len(zip_code) == 9:
                zip_code = zip_code[:5] + "-" + zip_code[-4:]

            std_df.loc[i, 'Zip Code'] = zip_code

        # Format all phone numbers to #(XXX) XXX-XXXX
        if phone_number_col:
            # Convert to string
            phone_number = str(std_df.loc[i, 'Phone'])

            # Filter out all non-numeric characters
            numeric_filter = filter(str.isdigit, phone_number)
            phone_number = "".join(numeric_filter)

            if len(phone_number) == 10:
                phone_number = "(" + phone_number[:3] + ") " + phone_number[3:6] + "-" + phone_number[6:]
            elif len(phone_number) == 11:
                phone_number = phone_number[:1] + "(" + phone_number[1:4] + ") " +\
                               phone_number[4:7] + "-" + phone_number[7:]

            std_df.loc[i, 'Phone'] = phone_number

    # Return standardized data frame
    return std_df

