import os
import pandas as pd

QC_TESTS = [['FILE_NAME', 'PART NUM', 'CUSTOMER', 'FLASH TIME', 'DRY TIME&TEMP','METHOD JOIN']] 
#List of all QC tests, used to decide on the columns needed in the dataframe XLSC_FOR_DEACOM
FOLDER_PATH  = ("C:/BLANK")

def extract_test_list(FOLDER_PATH):
    """
    Iterates through Excel files in a folder and extracts the row which contains all QC tests
    
    """
    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            file_path = os.path.join(FOLDER_PATH, filename)
            try:
                df = pd.read_excel(file_path, usecols="C:P",skiprows=8, nrows=1, header=None)
            except:
                print("error in : ", str(file_path))
            add_tests(df, 0)
    return df

def add_tests(df, row):
    #Converts rows to columns, then filters all NAN values, then converts them to a list for QC_TESTS
    df = df.transpose()
    qc_test_list = (df[row].dropna().tolist())
    qc_test_list = [x.strip(' ').upper() for x in qc_test_list]
    QC_TESTS.append(qc_test_list)

def qc_test_specs(FOLDER_PATH):
    '''
    The goal of this function is to take a dataframe slice of the original dataframe in the format where the spec test names are on the bottom, the methods are on top, and Min, Max, and Target are in the middle.
    The prototype of this function is as follows:

    METHOD_JOIN is used in this for cases where there is a blank qc test value, but there is a value in the method row, which we concatenate for evaluation by SM later
    '''
    df_for_xlsx = pd.DataFrame(columns=QC_TESTS)
    for filename in os.listdir(FOLDER_PATH):
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            file_path = os.path.join(FOLDER_PATH, filename)
            try:
                df = pd.read_excel(file_path, usecols="B:Q", nrows=9, header=None)
            except:
                print("error in : ", str(file_path))
            
            #copy the section which will be processed for QC Test specifications
            df_qc_slice = df.iloc[4:, 1:].copy()

            #Make a dict to use for the dataframe creation as a row into the main dataframe
            qc_dict ={}
            
            qc_dict['FILE_NAME'] = filename

            #Get Part num
            qc_dict['PART NUM'] = df.iloc[0,0]

            #grab the entire customer section, and reverse it -- the blue_rhine / stock value is always a hidden column before the customer name
            customer_section = df.iloc[0,12:]
            customer_section = customer_section.iloc[::-1]
            #print(customer_section)
            for value in customer_section[customer_section.notnull()]:
                #print(value)
                if str(value) != 'CUSTOMER:':
                    qc_dict['CUSTOMER'] = str(value)
                    break
                else:
                    pass

            #capture flash time
            qc_dict['FLASH TIME'] = df.iloc[3,0]
            
            #explicitly set all values to str
            df.iloc[3,5:9] = df.iloc[3,5:9].apply(lambda x: str(x) if pd.notna(x) else x)
            qc_dict['DRY TIME&TEMP'] = df.iloc[3,5:9].str.cat(sep=' ', na_rep='')
            
            #Catch all for tests which are only listed as methods in the spreadsheets
            qc_dict['METHOD JOIN'] = ''

            
            qc_tests = df_qc_slice.iloc[4]
            qc_tests = qc_tests.str.upper()
            qc_tests = qc_tests.str.strip(" ")

            df_qc_slice.columns = qc_tests
            df_qc_slice.reset_index(drop=True, inplace = True)

            for series_name, series in df_qc_slice.items():
                #print(series_name)
                #print(series)
                if str(series_name) == '':
                    qc_dict['METHOD JOIN'] += str(series[0]) + ' - '
                elif str(series_name) == 'nan' and series.isnull().all() == True:
                    pass
                else:
                    if series[0:3].isnull().all() == False:
                        series = series.apply(lambda x: str(x) if pd.notna(x) else x)
                        result = series[0:4].str.cat(sep=' - ', na_rep='')
                        #print(result)
                        qc_dict[series_name] = result
            qc_dict_df = pd.DataFrame(qc_dict, index=[0])
            df_for_xlsx = pd.concat([df_for_xlsx, qc_dict_df], ignore_index=True)
    return df_for_xlsx

def send_to_spreadsheet(df):
    '''
    This function will take the qc_test_specs dataframe and send it into an excel formated document
    '''
    df.to_excel("C:/BLANK")

def flatten(xss):
    #From stack overflow answers, condenses a list of lists into one big list
    return [x for xs in xss for x in xs if x]

#iterate through all files and build a list of QC tests to make the column of spreadsheet
extract_test_list(FOLDER_PATH) 

#convert the qc test global list from set data type to list and use to build the excel spreadsheet header
QC_TESTS = list(set(flatten(QC_TESTS)))

#build the final dataframe which will be exported to excel
final_frame = qc_test_specs(FOLDER_PATH)

#export the final dataframe to excel
send_to_spreadsheet(final_frame)