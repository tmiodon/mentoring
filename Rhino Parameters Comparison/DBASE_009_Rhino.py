import pandas
import numpy


def CompareDataframes(df1, df2, column_names):
    result_df = pandas.DataFrame(columns=column_names)

    for column in result_df.columns.tolist():
        result_df[column] = numpy.where(df1[column] != df2[column], 'True', 'False')

    return result_df


def ShowComparsion(df1, df2, result_df):
    df_to_show_columns = ['Port Number',
                          'Parameter Number',
                          'Parameter Name',
                          'Expected Value',
                          'Actual Value']

    for column in result_df.columns.tolist():
        df_to_show = pandas.DataFrame(columns=df_to_show_columns)
        for row in range(len(result_df)):
            if result_df[column].values[row] == 'True':
                df_to_show = df_to_show.append({'Port Number': df1['Port Number'].values[row],
                                                'Parameter Number': df1['Parameter Number'].values[row],
                                                'Parameter Name': df1['Parameter Name'].values[row],
                                                'Expected Value': df1[column].values[row],
                                                'Actual Value': df2[column].values[row]}, ignore_index=True)

        if df_to_show.empty is False:
            print "\nDifferences in {}".format(column)
            print "*" * 100
            print df_to_show.to_string(index=False)
        else:
            continue


def SaveComparisonToFiles(df1, df2, result_df):
    df_to_save_columns = ['Port Number',
                          'Parameter Number',
                          'Parameter Name',
                          'Expected Value',
                          'Actual Value']

    for column in result_df.columns.tolist():
        df_to_save = pandas.DataFrame(columns=df_to_save_columns)
        for row in range(len(result_df)):
            if result_df[column].values[row] == 'True':
                df_to_save = df_to_save.append({'Port Number': df1['Port Number'].values[row],
                                                'Parameter Number': df1['Parameter Number'].values[row],
                                                'Parameter Name': df1['Parameter Name'].values[row],
                                                'Expected Value': df1[column].values[row],
                                                'Actual Value': df2[column].values[row]}, ignore_index=True)

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

    # ShowComparsion(rhino_params, emulated_rhino_params, compared_df)

    SaveComparisonToFiles(rhino_params, emulated_rhino_params, compared_df)