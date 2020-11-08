import pandas as pd
import numpy as np
import datetime
from datetime import date
from datetime import datetime

def main():
    path1 = '/Users/USER/Dropbox/台北最愛排房表/最愛109年訂房房況表.xls'
    path2 = '/Users/USER/Dropbox/台北最愛排房表/最愛110年訂房房況表.xls'
    xls1 = pd.ExcelFile(path1)
    xls2 = pd.ExcelFile(path2)
    # read in the slide of excel
    df = pd.read_excel(xls1,'109年7-12月')
    df2 = pd.read_excel(xls2,'110年1-6月')
    df3 = pd.read_excel(xls2,'110年7-12月')
    def clean(df):
        df = df.iloc[:18,1:].drop(1).T
        column_name = ['days_in_week',201,301,501,601,202,302,502,602,203,303,503,
        603,205,305,505,605]
        #get relevant column
        df.columns = column_name
        # get relevant rows
        df = df.iloc[1:]
        return df
    df = clean(df)
    df2 = clean(df2)
    df3 = clean(df3)
    df =df.append(df2).append(df3)
    # #data cleaning
    # df = df.iloc[:18,1:].drop(1).T
    # column_name = ['days_in_week',201,301,501,601,202,302,502,602,203,303,503,
    # 603,205,305,505,605]
    # #get relevant column
    # df.columns = column_name
    # # get relevant rows
    # df = df.iloc[1:]

    #data mnipulation
    #get the boolean of current relevant day
    current_date_boolean = pd.Series(df.index) >= date.today()
    # resettig index to allow date to be accessed
    df = df.reset_index()
    df.rename(columns={'index':'date'}, inplace=True)
    #filtering the date
    df = df[current_date_boolean]
    #get rooms
    rooms = [201,301,501,601,202,302,502,602,203,303,503,603,205,305,505,605]

    #stratify the dates that allow reservation
    final_room_ava = pd.DataFrame()
    #pre-allocation of number of days possible
    ava_series_econ = pd.Series([0]*df.shape[0])
    ava_series_stand = pd.Series([0]*df.shape[0])
    ava_series_deluxe = pd.Series([0]*df.shape[0])
    #stratify each room for 15 days continuous vacancies
    for j in range(len(rooms)):
        temp_name = df.iloc[:,j+2].name
        temp_room_vacancy = df[temp_name].isna()
        for i in range(df.shape[0]-14):
            if str(temp_name)[1:] =='01':
                if temp_room_vacancy.iloc[i:i+15].sum() == 15:
                    ava_series_econ[i] += 1
            elif str(temp_name)[1:] =='05':
                if temp_room_vacancy.iloc[i:i+15].sum() == 15:
                    ava_series_deluxe[i] += 1
            else:
                if temp_room_vacancy.iloc[i:i+15].sum() == 15:
                    ava_series_stand[i] += 1
    final_room_ava['date'] = list(df['date'])
    final_room_ava['經濟'] = ava_series_econ
    final_room_ava['標準'] = ava_series_stand
    final_room_ava['豪華'] = ava_series_deluxe

    export = (final_room_ava.set_index('date').T)
    #changing it to format friendly to excel
    export.columns = pd.Series(export.columns).apply(lambda x: str(x.date()
    .strftime('%m-%d')))
    export.T.to_excel('今天空房表.xls')
    export.T.to_excel(r'/Users/USER/Dropbox/台北最愛排房表/今天空房表.xls')

if __name__ == "__main__":
    main();
