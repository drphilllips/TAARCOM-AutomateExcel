import os
import pandas as pd
from xlrd import XLRDError

import EnumTypes
import ExcelUtilities


def main(std_df):
    """Automatically fills in sales reps based on:
    1) Master Account List (Customer -> Sales Rep)
    2) Territories (Zip Code -> Sales Rep)

    :param std_df: standardized data frame for the insight file
    :return: new file with OSR column filled in
    """

    # ----------------------------------
    #  Load the necessary lookup files
    # ----------------------------------

    customer_to_proper_name_map = ExcelUtilities.loadLookupFile("rootCustomerMappings.xlsx", "Sales Lookup")
    mstr_account_list = ExcelUtilities.loadLookupFile("Master Account List.xlsx", "Allacct")
    mstr_territory_list = ExcelUtilities.loadLookupFile("CAZipCode.xlsx", "CA_BASIC_ROSTER")

    # -----------------
    #  Find sales reps
    # -----------------

    # Pull out columns we need now to expedite indexing later
    customer_col = customer_to_proper_name_map['Root Customer'].tolist()
    customer_to_proper_name_col = customer_to_proper_name_map['ProperName'].tolist()
    proper_name_col = mstr_account_list['ProperName'].tolist()
    proper_name_to_sales_rep_col = mstr_account_list['SLS'].tolist()
    zip_code_col = mstr_territory_list['ZipCode'].tolist()
    zip_sales_rep_col = mstr_territory_list['Sls'].tolist()

    # Change text columns to lower-case so search isn't case-sensitive
    customer_col = [str(item).lower() for item in customer_col]
    customer_to_proper_name_col = [str(item).lower() for item in customer_to_proper_name_col]
    proper_name_col = [str(item).lower() for item in proper_name_col]

    # Perform sales rep lookup on each row of our standard dataframe
    for i in std_df.index:
        # Set pass/fail flags
        account_list_fail = False
        territory_list_fail = False

        # Set desired values
        proper_name = ""
        sales_rep = ""

        # +++ Look at account first +++
        # Search by customer
        customer = str(std_df.loc[i, 'Reported Customer']).lower()  # Not case-sensitive

        # If no customer provided, flag as individual, move on to territory
        if not customer:
            std_df.loc[i, 'Flag'] = EnumTypes.Flag.CNP.value
            account_list_fail = True

        # If customer provided, map customer to its proper name
        if not account_list_fail:
            # Look first in the proper name column of customer-proper name map
            if customer in customer_to_proper_name_col or customer in proper_name_col:
                proper_name = customer
            elif customer in customer_col:  # Then look in customer column
                # Look for all occurrences of this customer within the customer column
                for index, elem in enumerate(customer_col):
                    if elem == customer:
                        customer_index = index
                        proper_name = customer_to_proper_name_col[customer_index]
                        if proper_name:  # Use the first occurrence that has an associated proper name
                            break
                # If the proper name is blank
                if not proper_name:
                    # No proper name associated
                    std_df.loc[i, 'Flag'] = EnumTypes.Flag.PNA.value
                    account_list_fail = True
            else:
                # Customer not found in Customer to Proper Name Map (New Account)
                std_df.loc[i, 'Flag'] = EnumTypes.Flag.CNF.value
                account_list_fail = True

        if not account_list_fail:
            # Find sales rep by proper name
            if proper_name in proper_name_col:
                proper_name_index = proper_name_col.index(str(proper_name).lower())  # Ignore case
                sales_rep = proper_name_to_sales_rep_col[proper_name_index]
            else:
                # Proper name not found in Master Account List
                std_df.loc[i, 'Flag'] = EnumTypes.Flag.PNF.value
                account_list_fail = True

        # +++ If that doesn't work, find by zip code +++
        if account_list_fail:
            # Get zip code
            zip_code = str(std_df.loc[i, 'Zip Code'])

            # Only care about precision to the first five digits
            if len(zip_code) == 10:
                zip_code = zip_code[:5]
            elif len(zip_code) != 5:  # Zip code too small or too large, out of territory
                std_df.loc[i, 'Flag'] = EnumTypes.Flag.OOT.value
                territory_list_fail = True

            if not territory_list_fail:
                try:
                    # Convert to int because that is how it is stored in master territory list
                    zip_code = int(zip_code)
                    zip_code_index = zip_code_col.index(zip_code)
                    sales_rep = zip_sales_rep_col[zip_code_index]
                except (ValueError, IndexError):
                    # Zip code not found in Master Territory List
                    std_df.loc[i, 'Flag'] = EnumTypes.Flag.OOT.value
                    territory_list_fail = True

        # +++ Save sales rep to our data frame +++
        if not account_list_fail or not territory_list_fail:
            std_df.loc[i, 'OSR'] = sales_rep

    return std_df



