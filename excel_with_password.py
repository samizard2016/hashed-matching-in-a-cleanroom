"""
Dec 3 2019 Paul, Samir
reading an excel workbook protected with password
get_df writes the firstsheet back to pandas DataFrame
"""

import sys
import win32com.client
import pandas as pd
import logging

class ExcelWithPassword:
    def __init__(self,file,read_pass,write_pass):
        self.__xlApp__ = win32com.client.Dispatch("Excel.Application")
        self.__xlApp__.Visible = False
        self.__file__ = file
        self.__read_pass__ = read_pass
        self.__write_pass__ = write_pass
        self.logger = logging.getLogger("PCRPT")
    def open_xlsx(self):
        try:
            self.__xlwb__ = self.__xlApp__.Workbooks.Open(self.__file__, False, True, None,
            self.__read_pass__,self.__write_pass__)
            self.logger.info(f"{self.__file__} has successfully been opened")
            return True
        except Exception as err:
            self.logger.error(f"Problem with opening {self.__file__} with password - {err}")
            return False
    def print(self):
        xlws = self.__xlwb__.Sheets(1) # counts from 1, not from 0
        print(xlws.Name)
        print(xlws.Cells(1, 1))
    def get_pdf(self):
        xlws = self.__xlwb__.Sheets(1)
        lastCol = xlws.UsedRange.Columns.Count
        lastRow = xlws.UsedRange.Rows.Count
        content = xlws.Range(xlws.Cells(1, 1), xlws.Cells(lastRow, lastCol)).Value
        df = pd.DataFrame(list(content))
        df = df[1:]
        # df = pd.read_excel(self.__xlwb__,sheet_name=xlws.name)
        self.__xlApp__.Quit()
        # print(df.columns)
        return df
    def get_df(self):
        xlws = self.__xlwb__.Sheets(1)
        lastCol = xlws.UsedRange.Columns.Count
        lastRow = xlws.UsedRange.Rows.Count
        content = xlws.Range(xlws.Cells(1, 1), xlws.Cells(lastRow, lastCol)).Value
        df = pd.DataFrame(list(content))
        df = df[1:]
        df.columns = ['phone_no','NCCS']
        # df = pd.read_excel(self.__xlwb__,sheet_name=xlws.name)
        self.__xlApp__.Quit()
        # print(df.columns)
        return df
        # file_to_save_to = f"{self.__file__}_1.xlsx"
        # df.to_excel(file_to_save_to,index=False)
        # print(f"Successfully saved a copy to {file_to_save_to}")

    #Shashank 12032020 added for kwp
    def get_df_kwp(self):
        xlws = self.__xlwb__.Sheets(1)
        lastCol = xlws.UsedRange.Columns.Count
        lastRow = xlws.UsedRange.Rows.Count
        content = xlws.Range(xlws.Cells(1, 1), xlws.Cells(lastRow, lastCol)).Value
        df = pd.DataFrame(list(content))
        df = df[1:]
        df.columns = ['ad_id','household_id']
        # df = pd.read_excel(self.__xlwb__,sheet_name=xlws.name)
        self.__xlApp__.Quit()
        # print(df.columns)
        return df
    # Shashank 12032020 added for kwp

if __name__=="__main__":
    ewp = ExcelWithPassword("C:\\Users\\PaulSa\\Documents\\self-talking.xlsx",u"self-talk",u"self-talk")
    print("Successfully read the excel  file with password")
    ewp.print()
    ewp.get_df()
