import pandas


''' Function replaced by SkipValuesLikeList'''
# def FilterDataframe(dataframe, column_to_filter, values_to_skip):
#
#     """ This function is designed to filter dataframe. If "values_to_skip" exists in dataframe
#     "column_to_filter" column, then it will be skipped. Function returns filtered dataframe.
#     :param dataframe: [dataframe] Pandas input object
#     :param column_to_filter: [str] Name of column with values to filter
#     :param values_to_skip: list[str] Values to be skipped
#     :return: [dataframe] Filtered pandas output object
#     """
#
#     blank_cells = ['', 'nan', 'NaN']
#     for value_to_skip in values_to_skip:
#         if value_to_skip not in blank_cells:
#             # Retrieve all parameters indexes from dataframe which are different than value_to_skip
#             filtered_data_indexes = dataframe.index[dataframe[column_to_filter] != value_to_skip]
#             # Return filtered data in dataframe
#             dataframe = dataframe.loc[filtered_data_indexes]
#         else:
#             # Remove rows with blank cell in column_to_filter columns
#             dataframe.dropna(subset=[column_to_filter], inplace=True)
#
#     return dataframe

def SkipValuesLikeList(dataframe, column_name, values_to_skip):

    """ This function is designed to filter dataframe. If "values_to_skip" exists in dataframe
    "column_to_filter" column, then it will be skipped. Function returns filtered dataframe.
    :param dataframe: [dataframe] Pandas input object
    :param column_name: [str] Name of column with values to filter
    :param values_to_skip: list[str] Values to be skipped
    :return: [dataframe] Filtered pandas output object
    """

    blank_cells = ['', 'nan', 'NaN']

    # Retrieve all parameters indexes from dataframe with value_to_skip in column_name
    filtered_data_indexes = dataframe.index[dataframe[column_name].isin(values_to_skip)]
    # Return filtered data in dataframe
    dataframe = dataframe.drop(filtered_data_indexes)

    if any(item in values_to_skip for item in blank_cells):
        # Remove rows with blank cell in column_to_filter columns
        dataframe.dropna(subset=[column_name], inplace=True)

    return dataframe


def SkipValuesLikeRegEx(dataframe, column_name, like):

    """ This function is designed to remove unwanted rows from dataframe regarding value in the specified column.
    :param dataframe: [dataframe] Pandas input object
    :param column_name: [str] Name of column with values to remove
    :param like: [str] Value on which column will be removed
    :return: [dataframe] Filtered pandas output object
    """

    # Retrieve all parameters indexes from dataframe with value_to_skip in column_name
    filtered_data_indexes = dataframe.index[dataframe[column_name].str.contains(like, regex=True, na=False)]
    # Return filtered data in dataframe
    dataframe = dataframe.drop(filtered_data_indexes)

    return dataframe


def DataframeModifications(dataframe, sheet_name):

    # Example of regex and removing unwanted rows
    if sheet_name == "Port 12 PCCB Parameters":
        dataframe = SkipValuesLikeRegEx(dataframe, 'Name', 'U1|W1|V1|U2|W2|V2|U3|W3|V3|U4|W4|V4|U5|W5|V5|U6|W6|V6|U7|'
                                                           'W7|V7|U8|W8|V8|U9|W9|V9')
    if sheet_name == "Port 14 PIOB Parameters":
        dataframe = SkipValuesLikeRegEx(dataframe, 'New Parameter Number', 'TBD')

    # Skip parameters with CRs which are not applicable to the drive
    dataframe = SkipValuesLikeList(dataframe, 'Commercial Release', crs_to_skip)

    # Skip parameters with Applicable to column values like No and empty cell
    dataframe = SkipValuesLikeList(dataframe, applicable_to_column_name, applicable_values_to_skip)

    # Skip parameters which are named as "Reserved"
    dataframe = SkipValuesLikeList(dataframe, 'Name', names_to_skip)

    # Delete duplicated parameters based on the priority of CR and overwrite dataframe
    dataframe = DeleteDuplicatedParams(dataframe, crs_priorities_dict)

    # Change displayed "New Parameter Number" value datatype from real to int
    dataframe['New Parameter Number'] = dataframe['New Parameter Number'].astype(int)

    return dataframe


def DeleteDuplicatedParams(dataframe, priorities_dictionary):

    """This function is designed to delete duplications from dataframe. If dataframe contains the same parameter
    name for different CR, then based on priorities_dictionary it will leave only this one with the highest
    priority (the lowest value - the highest priority). Function returns dataframe without duplications.
    :param dataframe: [dataframe] Pandas input object
    :param priorities_dictionary: [dict] Dictionary with key = CR name and value = priority
    :return: [dataframe] Pandas output object without duplications
    """

    # Filling "crs_priorities_list" with values based on CR name from Commercial Release column and
    # value from "priorities_dictionary" assigned to certain CR name
    crs_priorities_list = []
    for i in range(len(dataframe)):
        crs_priorities_list.append(priorities_dictionary[dataframe['Commercial Release'].values[i]])

    # Adding CR priority column with assigned values
    dataframe['CR priority'] = crs_priorities_list
    # Sorting dataframe by the values from CR priority column (from lowest priority [highest value]
    # to the highest [lowest value]
    dataframe = dataframe.sort_values("CR priority", ascending=False)
    # Checking for duplicates in Name column. If exist then leave only last (highest priority).
    dataframe = dataframe.drop_duplicates(subset="Name", keep="last")
    # Sorting dataframe by New Parameter Number column. Return to the previous sort (from lowest to highest).
    dataframe = dataframe.sort_values("New Parameter Number")
    # Delete CR priority column added for duplications deletion purpose
    dataframe.drop('CR priority', inplace=True, axis=1)

    return dataframe


def FilterByFirmwareRev(major_rev, minor_rev, family_text):
    crs_to_skip_list = []
    if major_rev == 1:

        if family_text == "PF6000T":
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10"]
        else:
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR1 R2", "CR1 R3",
                                "CR2 R4", "Dev CR2 R4", "CR2 R5", "CR2 R6", "CR2 R6.003", "CR2 R6.004", "CR3", "CR3 LC",
                                "Temp CR3 R10", "CR3 R10", "CR3 R11", "MV_CR1", "MV_CR2"]
    elif major_rev == 2:
        crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR1 R3", "CR2 R4",
                            "Dev CR2 R4", "CR2 R5", "CR2 R6", "CR2 R6.003", "CR2 R6.004", "CR3", "CR3 LC",
                            "Temp CR3 R10", "CR3 R10", "CR3 R11", "MV_CR1", "MV_CR2"]
    elif major_rev == 3:
        crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR2 R4", "Dev CR2 R4",
                            "CR2 R5", "CR2 R6", "CR2 R6.003", "CR2 R6.004", "CR3", "CR3 LC", "Temp CR3 R10", "CR3 R10",
                            "CR3 R11", "MV_CR1", "MV_CR2"]
    elif major_rev == 4:
        crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR2 R5", "CR2 R6",
                            "CR2 R6.003", "CR2 R6.004", "CR3", "CR3 LC", "Temp CR3 R10", "CR3 R10", "CR3 R11", "MV_CR1",
                            "MV_CR2"]
    elif major_rev == 5:
        crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR2 R5", "CR2 R6",
                            "CR2 R6.003", "CR2 R6.004", "CR3", "CR3 LC", "Temp CR3 R10", "CR3 R10", "CR3 R11", "MV_CR1",
                            "MV_CR2"]
    elif major_rev == 6:
        if minor_rev == 3:
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR2 R6.004",
                                "CR3", "CR3 LC", "Temp CR3 R10", "CR3 R10", "CR3 R11", "MV_CR1", "MV_CR2"]
        elif minor_rev == 4:
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR3", "CR3 LC",
                                "Temp CR3 R10", "CR3 R10", "CR3 R11", "MV_CR1", "MV_CR2"]
    elif major_rev == 7:
        if family_text == "PF6000T":
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10"]
        else:
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR3 R10",
                                "CR3 R11", "MV_CR1", "MV_CR2"]
    elif major_rev == 8:
        if family_text == "PF6000T":
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10"]
        else:
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR3 R10",
                                "CR3 R11", "MV_CR1", "MV_CR2"]
    elif major_rev == 10:
        if family_text == "PF6000T":
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10"]
        else:
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "CR3 R11",
                                "MV_CR1", "MV_CR2"]
    elif major_rev == 11:
        if family_text == "PF6000T":
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10"]
        else:
            crs_to_skip_list = ["nan", "CRx", "Dev CR1", "Dev CR1 R2", "Dev CR2 R4", "Temp CR3 R10", "MV_CR1", "MV_CR2"]
    return crs_to_skip_list


'''**************************************************'''
'''EDIT SECTION - this variables can be edited ******'''

sheets_to_parse = ["Port 0 ICB Parameters",
                   "Port 9 Application Parameters",
                   "Port 13 Converter Control Param",
                   "Port 12 PCCB Parameters",
                   "Port 14 PIOB Parameters"]

columns_to_import = ["Name",
                     "New Parameter Number",
                     "Commercial Release",
                     "Offline Minimum",
                     "Offline Maximum",
                     "Offline Default",
                     "Online Minimum",
                     "Online Maximum",
                     "Online Default",
                     "Applicable to PF6000T?"]

names_to_skip = ["Reserved"]

applicable_values_to_skip = ['No', '']

applicable_to_column_name = "Applicable to PF6000T?"
'''**************************************************'''

'''**************************************************'''
'''MAINTENANCE SECTION - this variables need to be   '''
'''maintained when changes in HPC Database Draft     '''
'''will appear.                                      '''

file_name = "HPC Database - Draft.xlsm"

crs_priorities_dict = {'CRx': 19,
                       'CR1': 18,
                       'Dev CR1': 17,
                       'CR1 R2': 16,
                       'Dev CR1 R2': 15,
                       'CR1 R3': 14,
                       'CR2 R4': 13,
                       'Dev CR2 R4': 12,
                       'CR2 R5': 11,
                       'CR2 R6': 10,
                       'CR2 R6.003': 9,
                       'CR2 R6.004': 8,
                       'CR3': 7,
                       'CR3 LC': 6,
                       'Temp CR3 R10': 5,
                       'CR3 R10': 4,
                       'CR3 R11': 3,
                       'MV_CR1': 2,
                       'MV_CR2': 1}

tested_drive = 'PF6000T'
major_fw_rev = 1
minor_fw_rev = 1

# Based on firmware revision and family text define CRs which should be skipped
crs_to_skip = FilterByFirmwareRev(major_fw_rev, minor_fw_rev, tested_drive)
'''**************************************************'''

if __name__ == "__main__":
    # Create Pandas dataframe object
    hpc_dbase_dataframe = pandas.ExcelFile(file_name)

    # Parse only required sheets from excel file
    ports_dataframes = [hpc_dbase_dataframe.parse(sheet_name=sheet_to_parse, usecols=columns_to_import) for
                        sheet_to_parse in sheets_to_parse]

    for sheet_number, sheet_name in enumerate(sheets_to_parse):
        # Data modifications section
        ports_dataframes[sheet_number] = DataframeModifications(ports_dataframes[sheet_number], sheet_name)

        # Print modified dataframes to excel
        ports_dataframes[sheet_number].to_excel(sheets_to_parse[sheet_number] + '.xlsx', index=False)
        print "File {}.xlsm saved successfully".format(sheets_to_parse[sheet_number])
        print "*" * 100