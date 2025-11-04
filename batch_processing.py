import pandas as pd
import logging
from excel_with_password import ExcelWithPassword
from matching import NewMatching
from datetime import datetime
import win32timezone

class BatchProcessing:
    """
    batch processing on multiple input files from HS whereas input of kantar list would remain fixed
    """
    def __init__(self,schema_xlsx,source_pass,dataset):
        print("Inside the batch processing")
        self.logger = logging.getLogger("PCRPT")
        self.schema_xlsx = schema_xlsx
        self.source_pass = source_pass
        self.data_set = dataset
        self.logger.info(f"batch processing schema file traced as {schema_xlsx}")
        print("Schema definition is complete")
        try:
            ewp = ExcelWithPassword(self.schema_xlsx,self.ensureUtf(self.source_pass),self.ensureUtf(self.source_pass))
        except Exception as err:
            self.logger.error(f"Problem with opening the batch schema {schema_xlsx} - {err}")
        if ewp.open_xlsx():
            self.schema_df = ewp.get_pdf()[3:]
            if self.data_set == 'digital':
                self.headers = ["source_xlsx", "password_source", "prof_csv", "viewership_csv", "device_csv", "salt"]
            else:
                self.headers = ["source_xlsx", "password_source", "prof_csv", "viewership_csv", "salt"]
            self.schema_df.columns = self.headers
            self.schema_df = self.schema_df.reset_index()
            self.logger.info(f"A total of {self.schema_df.shape[0]} processes have been traced")
            print(f"A total of {self.schema_df.shape[0]} processes have been traced")
        else:
            self.logger.info(f"problem with opening {self.schema_xlsx}")

    def run_batches(self):
        csv_files_to_remove = []
        for indx in self.schema_df.index:
            xlsx_file = self.schema_df[self.headers[0]][indx]
            source_pass = self.schema_df[self.headers[1]][indx]
            ewp = ExcelWithPassword(xlsx_file,source_pass,source_pass)
            if ewp.open_xlsx():
                source = ewp.get_df()
                self.logger.info(f"{xlsx_file} has been opened")
            else:
                self.logger.error(f"failed to open {xlsx_file}")
                continue
            try:
                prof = self.schema_df[self.headers[2]][indx]
                df_prof = pd.read_csv(prof)
                self.logger.info(f"{prof} has been opened")
            except Exception as err:
                self.logger.error(f"failed to open/ read {prof}")
                continue
            try:
                viewership = self.schema_df[self.headers[3]][indx]
                df_viewership = pd.read_csv(viewership)
                self.logger.info(f"{viewership} has been opened")
            except Exception as err:
                self.logger.error(f"failed to open/ read {viewership}")
                continue
            if self.data_set=='digital':
                try:
                    device = self.schema_df[self.headers[4]][indx]
                    df_device = pd.read_csv(device)
                    self.logger.info(f"{device} has been opened")
                except Exception as err:
                    self.logger.error(f"failed to open/ read {device}")
                    continue
            targets = [df_prof,df_viewership,df_device] if self.data_set=='digital' else [df_prof,df_viewership]
            matching_for = ["Profile","Viewership","Device"]
            self.logger.info(f"input files have been processed successfully - approaching matching for batch-{indx+1}")
            if self.data_set=='digital':
                salt = self.schema_df[self.headers[5]][indx]
            results = []
            for i,target in enumerate(targets):
                self.logger.info(f"Matching for {matching_for[i]} has been approached")

                tkeys = {"source": source,"target":target,'salt':salt} if self.data_set=='digital' else {"source": source,"target":target}
                match = NewMatching(**tkeys)
                match_key = 'phone_no' if self.data_set=='digital' else 'ad_id'
                res = match.get_matches(match_key)
                results.append(res)
            m_dfs = self.save_matched_data(targets,results)
            csv_files = self.write_back_to_csvs(m_dfs)
            csv_files_to_remove.append(csv_files)
            self.logger.info(f"Batch {indx+1} completed")
        self.logger.info("batch processing over!")
        # flatten the list of the csv files
        files = [elem for elems in csv_files_to_remove for elem in elems]
        return files

    def save_matched_data(self,dfs,flags):
        # filter on matching
        tdfs = []
        for ind,df in enumerate(dfs):
            tdf = df[flags[ind]]
            tdfs.append(tdf)
        return tdfs
    def write_back_to_csvs(self,dfs):
        now = datetime.now()
        dt_string = f"{now.year}{now.month}{now.day}_{now.hour}{now.minute}"
        if self.data_set == 'digital':
            self.csv_files = [f"hs_prof_matched_{dt_string}.csv",f"hs_viewership_matched_{dt_string}.csv",f"hs_device_matched_{dt_string}.csv"]
        else:
            self.csv_files = [f"hs_prof_matched_{dt_string}.csv", f"hs_viewership_matched_{dt_string}.csv"]
        for ind,df in enumerate(dfs):
            df.to_csv(self.csv_files[ind],index=False)
            self.logger.info(f"{self.csv_files[ind]} has been saved to the current folder")
        # update the csv_file names for right_widget
        return self.csv_files

    def replace_ad_id(self,dfs,r_map):
        tdfs = []
        for ind, df in enumerate(dfs):
            #df = df['ad_id'].map(r_map)
            #w_df = w_df.drop('ad_id',axis=1)
            t_df = self.replace_with_household_id(df,r_map)
            t_df = t_df.rename(columns={'ad_id':'household_id'})
            tdfs.append(t_df)
        return tdfs

    def replace_with_household_id(self,df,d_map):
        for row in range(df.shape[0]):
            ad_id = df.iloc[row]['ad_id']
            df.iloc[row]['ad_id'] = d_map[ad_id]
        return df


    def ensureUtf(self,s, encoding='utf8'):
        """Converts input to unicode if necessary.
        If `s` is bytes, it will be decoded using the `encoding` parameters.
        This function is used for preprocessing /source/ and /filename/ arguments
        to the builtin function `compile`.
        """
        if type(s) == bytes:
            return s.decode(encoding, 'ignore')
        else:
            return s

def main():
    print("Please wait while batches are in progress ...")
    schema_xlsx = "cleanroom_batch_processing_kantar_digital.xlsx"
    batch = BatchProcessing(schema_xlsx,'1235@kantar')
    batch.run_batches()
    print("Done!")

if __name__=="__main__":
    main()
