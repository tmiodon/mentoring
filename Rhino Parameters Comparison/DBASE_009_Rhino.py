import pandas
import numpy
import os
import shutil


def ClearFolder(folder_path):
    """ This function removes every file located in the folder which path is specified by path

    Arguments:
        folder_path: [str] Path to the folder that will be cleared
    """

    if not os.path.exists('Output'):
        os.makedirs('Output')

    folder = folder_path
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


def CompareDataframes(source_df, target_df, column_names):
    """ This function compares two dataframes within provided column names and returns dataframe with true
    (when difference was captured) and false (when values are the same) values.

    Arguments:
        source_df: Pandas dataframe object (source dataframe)
        target_df: Pandas dataframe object (target dataframe)
        column_names: List of columns to compare
    Returns:
        result_df: Pandas dataframe object (results from comparison)
    """
    results_file = open('Output\\Results.txt', 'w')

    if len(source_df) == len(target_df):
        result_df = pandas.DataFrame(columns=column_names)

        for column in result_df.columns.tolist():
            result_df[column] = numpy.where(source_df[column].astype(str) != target_df[column].astype(str),
                                            'True', 'False')
        return result_df
    else:
        print "Tables have different length! Comparison aborted."
        results_file.write("Tables have different length! Comparison aborted.")
        results_file.close()
        print "File Results.txt saved successfully!"
        raise ValueError


def ShowComparison(source_df, target_df, result_df):
    """ This function analyses all true values in result_df and prints ordered differences by column name in console or
    in txt file.

    Arguments:
        source_df: Pandas dataframe object (source dataframe)
        target_df: Pandas dataframe object (target dataframe)
        result_df: Pandas dataframe object (results from comparison)
    """
    df_to_show_columns = ['Port Number',
                          'Parameter Number',
                          'Parameter Name',
                          'Expected Value',
                          'Actual Value']

    differences_found = False

    results_file = open('Output\\Results.txt', 'w')

    for column in result_df.columns.tolist():
        df_to_show = pandas.DataFrame(columns=df_to_show_columns)
        for row in range(len(result_df)):
            if result_df[column].values[row] == 'True':
                df_to_show = df_to_show.append({'Port Number': source_df['Port Number'].values[row],
                                                'Parameter Number': source_df['Parameter Number'].values[row],
                                                'Parameter Name': source_df['Parameter Name'].values[row],
                                                'Expected Value': source_df[column].values[row],
                                                'Actual Value': target_df[column].values[row]}, ignore_index=True)

        if df_to_show.empty is False:
            differences_found = True

            results_file.write("Differences in {}".format(column))
            results_file.write('\n')
            results_file.write("*" * 100)
            results_file.write('\n')
            results_file.write(df_to_show.to_string(index=False))
            results_file.write('\n\n')
        else:
            continue

    if differences_found:
        print "File Results.txt saved successfully!"
    else:
        print "File Results.txt saved successfully!"
        results_file.write('Compared files are the same based on columns: {}.'.format(columns_to_import))

    results_file.close()


def SaveComparisonInExcelFiles(source_df, target_df, result_df):
    """ This function analyses all true values in result_df and prints ordered differences by column name
    in excel files.

    Arguments:
        source_df: Pandas dataframe object (source dataframe)
        target_df: Pandas dataframe object (target dataframe)
        result_df: Pandas dataframe object (comparison results)
    """
    df_to_save_columns = ['Port Number',
                          'Parameter Number',
                          'Parameter Name',
                          'Expected Value',
                          'Actual Value']

    differences_found = False

    for column in result_df.columns.tolist():
        df_to_save = pandas.DataFrame(columns=df_to_save_columns)
        for row in range(len(result_df)):
            if result_df[column].values[row] == 'True':
                df_to_save = df_to_save.append({'Port Number': source_df['Port Number'].values[row],
                                                'Parameter Number': source_df['Parameter Number'].values[row],
                                                'Parameter Name': source_df['Parameter Name'].values[row],
                                                'Expected Value': source_df[column].values[row],
                                                'Actual Value': target_df[column].values[row]}, ignore_index=True)

        if df_to_save.empty is False:
            differences_found = True
            df_to_save.to_excel("Output\\Differences in {}".format(column) + '.xlsx', index=False)
        else:
            continue

    if differences_found:
        print "Excel files saved successfully!"
    else:
        print "Compared files are the same based on columns: {}.".format(columns_to_import)


if __name__ == '__main__':
    # Remove all files from Output folder
    ClearFolder('Output')

    # Path of the original parameter database
    rhino_file_name = 'Input\\Parameter_database_English.csv'
    # Path of the target parameter database. This must be edited based on the current file name.
    emulated_rhino_file_name = 'Input\\Parameter_database_emulation_1_0_181.csv'

    # Columns that we want to import from csv files. If needed, cols can be added to the list and they will be taken
    # into account in the comparison process
    columns_to_import = ['Port Number',
                         'Parameter Number',
                         'Parameter Name',
                         'Parameter Max Value',
                         'Parameter Min Value',
                         'Parameter Default Value',
                         'Parameter Unit',
                         'Parameter Writable',
                         'Value Does Not Default']

    # Create dataframes by reading original and compared parameter databases
    rhino_params = pandas.read_csv(rhino_file_name, usecols=columns_to_import)
    emulated_rhino_params = pandas.read_csv(emulated_rhino_file_name, usecols=columns_to_import)

    # Perform comparison and save result in dataframe
    compared_df = CompareDataframes(rhino_params, emulated_rhino_params, columns_to_import)
    # Print comparison results to the Results.txt file
    ShowComparison(rhino_params, emulated_rhino_params, compared_df)
    # This function can be used to export results separately to different excel files
    # SaveComparisonInExcelFiles(rhino_params, emulated_rhino_params, compared_df)
