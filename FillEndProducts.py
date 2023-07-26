import os

import pandas as pd
from xlrd import XLRDError

import ExcelUtilities


def main(std_df):
    """Automatically fills in end product based on Reported Customer;
    **ASSUME** we are working with Digi-Key

    :param std_df: standardized data frame for the insight file
    :return: new data frame with end product column filled in
    """

    # --------------------------
    #  Load the End Product Map
    # --------------------------

    end_product_map = ExcelUtilities.loadLookupFile("EndProductMap.xlsx", "EndProductLookup")

    # ----------------------------
    #  Fill in End Product column
    # ----------------------------

    # Pull out columns we need now to expedite indexing later
    proper_name_col = end_product_map['Proper Name'].tolist()
    end_product_col = end_product_map['End Product'].tolist()

    for i in std_df.index:
        # Set variable
        end_product = ""

        # Get the proper name from our report
        proper_name = std_df.loc[i, 'Reported Customer']
        # Map it to the end product
        try:
            proper_name_index = proper_name_col.index(proper_name)
            end_product = end_product_col[proper_name_index]
        except (ValueError, IndexError):
            # Proper name not in the map
            pass

        # Set end product
        std_df.loc[i, '(Suggested) End Product'] = end_product

    return std_df

