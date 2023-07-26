
import openpyxl

import EnumTypes


def main(std_df, filepath):

    workbook = openpyxl.load_workbook(filepath, data_only=True)

    sheet = workbook.worksheets[0]

    std_df.loc[0, 'Abracon Flag'] = ""  # initialize Abracon Flag column

    for i in std_df.index:
        cell = "A" + str(i+1)
        color_index = sheet[cell].fill.start_color.index
        color_in_hex = color_index[2:]

        if color_in_hex == EnumTypes.AbraconFlags.YELLOW_HEX.value:
            std_df.loc[i-1, 'Abracon Flag'] = EnumTypes.AbraconFlags.YELLOW_DESC.value
        elif color_in_hex == EnumTypes.AbraconFlags.GREEN_HEX.value:
            std_df.loc[i-1, 'Abracon Flag'] = EnumTypes.AbraconFlags.GREEN_DESC.value
        elif color_in_hex == EnumTypes.AbraconFlags.ORANGE_HEX.value:
            std_df.loc[i-1, 'Abracon Flag'] = EnumTypes.AbraconFlags.ORANGE_DESC.value

    return std_df

