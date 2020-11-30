import pandas


def FilterDataframe(dataframe, column_to_filter, values_to_skip):
    """ This function is designed to filter dataframe. If "values_to_skip" exists in dataframe
    "column_to_filter" column, then it will be skipped. Function returns filtered dataframe.
    :param dataframe: [dataframe] Pandas input object
    :param column_to_filter: [str] Name of column with values to filter
    :param values_to_skip: [str] or list[str] Values to be skipped
    :return: [dataframe] Filtered pandas output object
    """
    blank_cells = ['', 'nan', 'NaN']
    for value_to_skip in values_to_skip:
        if value_to_skip not in blank_cells:
            # Retrieve all parameters indexes from dataframe which are different than value_to_skip
            filtered_data_indexes = dataframe.index[dataframe[column_to_filter] != value_to_skip]
            # Return filtered data in dataframe
            dataframe = dataframe.loc[filtered_data_indexes]
        else:
            dataframe.dropna(subset=[column_to_filter], inplace=True)
    return dataframe


if __name__ == "__main__":

    file_name = "HPC Database - Draft.xlsm"

    '''*******************************************'''
    '''Edit section - this variables can be edited'''

    tested_drive = 'PF755TR'

    sheets_to_parse = ["Port 0 ICB Parameters",
                       "Port 9 Application Parameters",
                       "Port 13 Converter Control Param",
                       "Port 12 Parameters",
                       "Port 14 Parameters"]

    columns_to_import = ["Name",
                         "New Parameter Number",
                         "Commercial Release",
                         "Offline Minimum",
                         "Offline Maximum",
                         "Offline Default",
                         "Online Minimum",
                         "Online Maximum",
                         "Online Default",
                         "Applicable to Frame 5-6 Panel Mount Drives (PF755TR and PF755TL)?"]

    crs_to_skip = ["CRx"]

    names_to_skip = ["Reserved"]

    applicable_values_to_skip = ['No', '']
    '''*******************************************'''

    # Create Pandas dataframe object
    hpc_dbase_dataframe = pandas.ExcelFile(file_name)

    # Display sheet names from excel file
    for idx, sheet_name in enumerate(hpc_dbase_dataframe.sheet_names):
        print "{} - {}".format(idx, sheet_name)

    # Parse only demanded sheets from excel file
    ports_dataframes = [hpc_dbase_dataframe.parse(sheet_name=sheet_to_parse, usecols=columns_to_import) for
                        sheet_to_parse in sheets_to_parse]

    '''
    Parameters filtration for all used workbook sheets.
    In the table, there shouldn't be parameters assigned to "CRx", "MV_CR1", "MV_CR2" commercial releases.
    There shouldn't also be other parameters than this which are applicable to specified drive.
    Reserved parameters also shouldn't be visible.
    Parameter number column needs to be displayed as int instead of real.
    '''
    for sheet_number in range(0, len(ports_dataframes)):


        # # TODO: Duplications should be deleted by priority from CR
        # ports_dataframes[sheet_number].drop_duplicates(subset="Name", keep="last", inplace=True)

        # Skip parameters with CRs which are not applicable to the drive
        ports_dataframes[sheet_number] = FilterDataframe(ports_dataframes[sheet_number], 'Commercial Release',
                                                         crs_to_skip)

        # Skip parameters which are not applicable to the tested drive
        if tested_drive == "PF755TR":
            # Based on the drive type "applicable_to_column" variable should be rewritten from HPC Database document
            applicable_to_column_name = "Applicable to Frame 5-6 Panel Mount Drives (PF755TR and PF755TL)?"
            ports_dataframes[sheet_number] = FilterDataframe(ports_dataframes[sheet_number], applicable_to_column_name,
                                                             applicable_values_to_skip)

        # Skip parameters which are named as "Reserved"
        ports_dataframes[sheet_number] = FilterDataframe(ports_dataframes[sheet_number], 'Name', names_to_skip)

        # Change displayed "New Parameter Number" value datatype from real to int
        ports_dataframes[sheet_number]['New Parameter Number'] = ports_dataframes[sheet_number][
            'New Parameter Number'].astype(int)

        ports_dataframes[sheet_number].to_excel(sheets_to_parse[sheet_number] + '.xlsx', index=False)

        '''RETRIEVING DUPLICATED VALUES'''
        duplicated = ports_dataframes[sheet_number][ports_dataframes[sheet_number]['Name'].duplicated(keep=False) == True]
        print duplicated[['Commercial Release', 'Name']]
        print "------------"
    # Show 3 dataframe columns: 'Name', 'Commercial Release', 'New Parameter Number'
    print ports_dataframes[0][['Name', 'Commercial Release', 'New Parameter Number']]

    print "*" * 80
    print ports_dataframes[0]['Commercial Release'].values
    print "*" * 80
    print ports_dataframes[0]['Applicable to Frame 5-6 Panel Mount Drives (PF755TR and PF755TL)?'].values[0]

