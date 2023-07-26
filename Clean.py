import os

import pandas as pd

import AssignSalesReps
import EnumTypes
import ExcelUtilities
import FillEndProducts
import GetAbraconFlags
import StandardizeColumns


def main(filepath, company):
    """Standardizes columns, gets proper customers, and
    most importantly, assigns sales reps for each order
    in the insight file; to be sent out to sales reps

    :param filepath: path to insight file
    :param company: company that provided the insight file
    :return: export new, cleaned-up insight file
    """

    # -----------------------
    #  Load insight file
    # -----------------------

    # Load the insight file to a data frame
    ins_df = pd.read_excel(filepath, sheet_name=0).fillna("")

    # ----------------------
    #  Create standard file
    # ----------------------

    # Standardize columns
    std_df = StandardizeColumns.main(ins_df, company)

    # For Digi-Key, fill in end product
    if company == EnumTypes.Company.DGK:
        std_df = FillEndProducts.main(std_df)

    # For Abracon, set flags based on color-coding
    if company == EnumTypes.Company.ABR:
        std_df = GetAbraconFlags.main(std_df, filepath)

    # Assign sales reps
    std_df = AssignSalesReps.main(std_df)

    # ----------------------
    #  Export standard file
    # ----------------------

    # Strip root off filepath and leave just the filename for output
    filename = os.path.basename(filepath)[:-5] + " (Standardized).xlsx"

    # Create file
    writer = ExcelUtilities.createExcelFile(filename, std_df)
    # Format columns in Excel
    ExcelUtilities.formatSheet(std_df, writer)
    # Save the file
    writer.save()

    # Success message
    print("> File successfully standardized!\n"
          "*Program Complete*")
