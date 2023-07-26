import os

import pandas as pd

import ExcelUtilities


def main(filepath):
    """Splits up the cleaned insight file into several files,
    one for each sales rep

    :param filepath: path to cleaned file
    :return: export one file for each salesperson
    """

    # ------------------------
    #  Load standardized file
    # ------------------------

    std_df = pd.read_excel(filepath, sheet_name=0).fillna("")
    header = std_df.columns

    # ------------------------------------------
    #  Separate rep data into unique dataframes
    # ------------------------------------------

    reps = []
    rep_dfs = []

    for i in std_df.index:
        sales_rep = std_df.loc[i, 'OSR']

        if sales_rep in reps:
            # Find the rep's index in the dataframe collection
            rep_index = reps.index(sales_rep)
            # Get rep's dataframe
            rep_df = rep_dfs[rep_index]
            # Append rep's row to their dataframe
            rep_df.loc[rep_df.size] = std_df.iloc[i]
        else:
            # Add rep to indexing map
            reps.append(sales_rep)
            # Construct a new dataframe for this rep
            rep_df = pd.DataFrame(columns=header)
            # Add the first entry to the dataframe
            rep_df.loc[0] = std_df.iloc[i]
            # Add dataframe to the collection
            rep_dfs.append(rep_df)

    # ----------------------------------------
    #  Export each dataframe as an Excel file
    # ----------------------------------------

    # Strip root off filepath to get filename
    in_filename = os.path.basename(filepath)

    rep_index = 0  # Track each rep's dataframe
    for rep in reps:

        rep_df = rep_dfs[rep_index]

        out_filename = "[" + rep + "] " + in_filename

        # Create file
        writer = ExcelUtilities.createExcelFile(out_filename, rep_df)
        # Format columns in Excel
        ExcelUtilities.formatSheet(rep_df, writer)
        # Save the file
        writer.save()

        rep_index += 1

    # Success message
    print("> File successfully split!\n"
          "*Program Complete*")
