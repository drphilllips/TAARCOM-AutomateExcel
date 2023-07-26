import os

import pandas as pd
from xlrd import XLRDError

import ExcelUtilities


def main(filepaths):
    """Stitches together Excel files, ensuring consistent columns

    :param filepaths: list of paths to each sales rep's feedback file
    :return: single, compiled feedback report
    """

    # ------------------------
    # Load the feedback files
    # ------------------------

    try:
        fdbk_file_dfs = [pd.read_excel(filepath, sheet_name=0) for filepath in filepaths]
    except XLRDError:
        print('..Error reading in files!\n'
              '*Program Terminated*')
        return

    # ------------------
    #  Create dataframe
    # ------------------

    header = fdbk_file_dfs[0].columns  # First file determines columns
    cmp_df = pd.DataFrame(columns=header).fillna("")

    # -----------------------
    #  Stitch files together
    # -----------------------

    file_number = 1
    for fdbk_file_df in fdbk_file_dfs:
        # Make sure they have the same number of columns as the original file
        if len(fdbk_file_df.columns) != len(header):
            print("..Column mismatch between files 1 and " + str(file_number) + ".\n" +
                  "*Program Terminated*")
            return

        # Make sure each column matches
        for i in range(len(fdbk_file_df.columns)):
            if fdbk_file_df.columns[i] != cmp_df.columns[i]:
                print("..Column mismatch between files 1 and " + str(file_number) + ".\n" +
                      "*Program Terminated*")
                return

        # Append this dataframe to the compiled dataframe
        cmp_df = cmp_df.append(fdbk_file_df, ignore_index=True)

        file_number += 1

    # ----------------------
    #  Export Compiled file
    # ----------------------

    # Use the first filename as the compiled report name
    first_filename = os.path.basename(filepaths[0])
    report_name = first_filename
    # Remove [OSR]
    if "]" in report_name:
        report_name = report_name[report_name.index("]"):]
    # Remove (Standardized) and .xlsx
    if "(" in report_name:
        report_name = report_name[:report_name.index("(")]
    elif "." in report_name:
        report_name = report_name[:report_name.index(".")]
    # Trim leading and trailing whitespace
    report_name = report_name.strip()
    # Add finishing touches
    filename = report_name.strip() + " (Compiled).xlsx"

    # Create the file
    writer = ExcelUtilities.createExcelFile(filename, cmp_df)
    # Format the file
    ExcelUtilities.formatSheet(cmp_df, writer)
    # Save the file
    writer.save()

    print("> Files successfully compiled!\n"
          "*Program Complete*")
