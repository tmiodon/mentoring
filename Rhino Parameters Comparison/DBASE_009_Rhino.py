import pandas
import numpy


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
    if len(source_df) == len(target_df):
        result_df = pandas.DataFrame(columns=column_names)

        for column in result_df.columns.tolist():
            result_df[column] = numpy.where(source_df[column] != target_df[column], 'True', 'False')

        return result_df
    else:
        print "Tables have different length! Comparison aborted."
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

    results_file = open('Results.txt', 'w')

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
        results_file.write('Compared files are the same.')

    results_file.close()


def SaveComparisonInExcelFiles(source_df, target_df, result_df):
    """ This function analyses all true values in result_df and prints ordered differences by column name
    in excel files.

    Arguments:
        source_df: Pandas dataframe object (source dataframe)
        target_df: Pandas dataframe object (target dataframe)
        result_df: Pandas dataframe object (results from comparison)
    """
    df_to_save_columns = ['Port Number',
                          'Parameter Number',
                          'Parameter Name',
                          'Expected Value',
                          'Actual Value']

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
            df_to_save.to_excel("Differences in {}".format(column) + '.xlsx', index=False)
        else:
            continue


if __name__ == '__main__':
    rhino_file_name = 'Parameter_database_English.csv'
    emulated_rhino_file_name = 'Parameter_database_emulation_1_0_181.csv'

    columns_to_import = ['Port Number',
                         'Parameter Number',
                         'Parameter Name',
                         'Parameter Max Value',
                         'Parameter Min Value',
                         'Parameter Default Value',
                         'Parameter Unit',
                         'Parameter Writable',
                         'Value Does Not Default']

    rhino_params = pandas.read_csv(rhino_file_name, usecols=columns_to_import)
    emulated_rhino_params = pandas.read_csv(emulated_rhino_file_name, usecols=columns_to_import)
    compared_df = CompareDataframes(rhino_params, emulated_rhino_params, columns_to_import)

    ShowComparison(rhino_params, emulated_rhino_params, compared_df)

    # This function can be used to export results separately to different excel files
    # SaveComparisonInExcelFiles(rhino_params, emulated_rhino_params, compared_df)
