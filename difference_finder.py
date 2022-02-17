import time
import pandas as pd
import glob


def compare(file1, file2, threshold, count=0, no_match=pd.DataFrame()):
    # file1, file2 are prepared dataframes. first two columns are primary and all other columns are secondary.
    # now we need to find matches between these two dataframes.
    # first we need to find the primary column in both dataframes.
    # then we need to find the secondary column & extra columns in both dataframes.

    pd.set_option('display.width', 500)
    pd.set_option('display.max_columns', 7)  # for better visibility
    pd.set_option("display.max_rows", 999)

    for index, i in file1.iterrows():
        date = str(i['Date']).split(' ')[0]
        for index2, j in file2.iterrows():
            date1 = str(j['Date']).split(' ')[0]

            if date == date1 and float(i[1]) == float(j[1]) and float(i[2]) == float(j[2]) and float(i[3]) != float(j[3]):
                if float(i[3]) - float(j[3]) > threshold and float(j[3]) - float(i[3]) < -threshold:

                    count += 1
                    print(count)
                    print(date, int(i[1]),int(i[2]),float(i[3]),'\t',float(j[3]), '\t',  '\n')

                    no_match.loc[count, f'{file1.columns[0]}'] = date
                    no_match.loc[count, f'{file1.columns[1]}'] = i[1]
                    no_match.loc[count, f'{file1.columns[2]}'] = i[2]
                    no_match.loc[count, f'{filename1} {file1.columns[3]}'] = i[3]
                    no_match.loc[count, f'{filename2} {file2.columns[3]}'] = j[3]
                    no_match.loc[count, 'Difference'] = i[3] - j[3]

                    if len(file1.columns) > 4:
                        for k in range(4, len(file1.columns)):
                            no_match.loc[count, file1.columns[k]] = i[k]
                    if len(file2.columns) > 4:
                        for k in range(4, len(file2.columns)):
                            no_match.loc[count, file2.columns[k]] = j[k]

    if no_match.empty:

        print('\nNothing found.\n')
        time.sleep(1)

    return no_match


def launch_ui():
    all_files = glob.glob('*.xlsx')

    # first file
    for index, file in enumerate(all_files):
        print('\t\n', index, file)

    try:
        user_file1 = int(input('\nSelect first file: (number only)\t'))
        sh_f1 = pd.ExcelFile(all_files[user_file1])
        print('\n\n\n')
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    for index, sheet_name in enumerate(sh_f1.sheet_names):
        print('\t\n', index, sheet_name)
    try:
        user_sheet1 = int(input('\nSelect sheet of first file: (number only)\t'))
        file1 = pd.read_excel(all_files[user_file1], sheet_name=sh_f1.sheet_names[user_sheet1])
        print('\n\n\n')
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    file1.columns = file1.columns.str.strip()
    for index, column in enumerate(file1.columns):
        print('\t\n', index, column)
    try:
        pri_column1 = input('\nChoose columns for identification: (numbers separated by comma only)\t')
        pri_column1 = pri_column1.split(',')
        print('\n\n\n')
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    for index, column in enumerate(file1.columns):
        print('\t\n', index, column)
    try:
        sec_column1 = int(input('\nChoose a balance column: (numbers only)\t'))
        print('\n\n\n')
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    for index, column in enumerate(file1.columns):
        print('\t\n', index, column)
    try:
        extra_col1 = input('\nSelect all other columns you wish to be displayed:\n'
                           'Separate numbers by comma ( , )\n'
                           'If none are needed, leave blank and press enter \t')
        print('\n\n\n')
        extra_col1 = extra_col1.split(',')
    except:
        print('\nPlease enter one number or more by seperating whem with comma.\n')
        time.sleep(2)
        exit()
    if extra_col1 != ['']:
        # 3 identifiers
        new_file1 = file1[
            [file1.columns[int(pri_column1[0])], file1.columns[int(pri_column1[1])], file1.columns[int(pri_column1[2])],
             file1.columns[sec_column1]]]
        for i in extra_col1:
            print(file1.columns[int(i)])
            new_file1[file1.columns[int(i)]] = file1.iloc[:, int(i)]
    else:
        new_file1 = file1[
            [file1.columns[int(pri_column1[0])], file1.columns[int(pri_column1[1])], file1.columns[int(pri_column1[2])],
             file1.columns[sec_column1]]]

    # second file
    for index, file in enumerate(all_files):
        print('\t\n', index, file)
    try:
        user_file2 = int(input('\nSelect second file: (number only)\t'))
        print('\n\n\n')
        sh_f2 = pd.ExcelFile(all_files[user_file2])
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    for index, sheet_name in enumerate(sh_f2.sheet_names):
        print('\t\n', index, sheet_name)
    try:
        user_sheet2 = int(input('\nSelect sheet of second file: (number only)\t'))
        print('\n\n\n')
        file2 = pd.read_excel(all_files[user_file2], sheet_name=sh_f2.sheet_names[user_sheet2])
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    file2.columns = file2.columns.str.strip()
    for index, column in enumerate(file2.columns):
        print('\t\n', index, column)
    try:
        pri_column2 = input('\nChoose Identification columns: (numbers separated by comma only)\t')
        pri_column2 = pri_column2.split(',')
        print('\n\n\n')
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    for index, column in enumerate(file2.columns):
        print('\t\n', index, column)
    try:
        sec_column2 = int(input('\nChoose a column for balance: (numbers only)\t'))
        print('\n\n\n')
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    for index, column in enumerate(file2.columns):
        print('\t\n', index, column)
    try:
        extra_col2 = input('\nSelect all other columns you wish to be displayed:\n'
                           'Separate numbers by comma ( , )\n'
                           'If none are needed, leave blank and press enter\t')
        print('\n\n\n')
        extra_col2 = extra_col2.split(',')
    except:
        print('\nPlease enter one number or more by separating whem with comma.\n')
        time.sleep(2)
        exit()

    if extra_col2 != ['']:
        # made for Viggo. 3 identifiers
        new_file2 = file2[
            [file2.columns[int(pri_column2[0])], file2.columns[int(pri_column2[1])], file2.columns[int(pri_column2[2])],
             file2.columns[sec_column2]]]
        for i in extra_col2:
            print(file2.columns[int(i)])
            new_file2[file2.columns[int(i)]] = file2.iloc[:, int(i)]
    else:
        new_file2 = file2[
            [file2.columns[int(pri_column2[0])], file2.columns[int(pri_column2[1])], file2.columns[int(pri_column2[2])],
             file2.columns[sec_column2]]]

    filename1 = all_files[user_file1].replace('.xlsx', '')
    filename2 = all_files[user_file2].replace('.xlsx', '')

    return new_file1, new_file2, filename1, filename2


if __name__ == '__main__':
    file1, file2, filename1, filename2 = launch_ui()

    try:
        threshold = int(input('\nEnter a threshold for the difference between two balances: \t'))
    except:
        print('\nPlease enter a valid number.\n')
        time.sleep(2)
        exit()

    no_match1 = compare(file1, file2, threshold)
    length = len(no_match1.index)
    no_match2 = compare(file2, file1, threshold, length)

    no_match_final = pd.concat([no_match1, no_match2])
    no_match_final.drop_duplicates(subset=no_match_final.columns[3], keep='first', inplace=True)
    no_match_final.to_excel(f'Compared output.xlsx', index=False)