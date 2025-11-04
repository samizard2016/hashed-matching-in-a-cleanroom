'''
PCRPTool is the toolset application for cleanroom process for kantar panels(KWP/WAM) in India
21 January 2020
modified to csv inputs on 27 January
Paul, Samir
Pune, Maharashtra
'''
from PyQt5.QtCore import *
from PyQt5.QtGui  import *
from PyQt5.QtWidgets import *
import logging
import pandas as pd
import numpy as np
import openpyxl
import math
import sys
import os
from matching import NewMatching
from disp_df import PandasModel
import shutil
from excel_with_password import ExcelWithPassword
from archiving import ArchivingApp
from datetime import datetime
import glob
from batch_processing import BatchProcessing
from reg import Registry 

class PCRPTool(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.color_scheme = '#000000'
        self.main_widget = MainWidget(self)
        self.setCentralWidget(self.main_widget)
        self.main_widget.right_widget.tabWidget.tabBar().setVisible(False)
        self.main_widget.left_widget.setStyleSheet(f"background-color: {self.color_scheme};")
        self.showMaximized()
        # self.handle_interactions()
        # self.setStyleSheet("background-color: #646665;")
        self.setStyleSheet(f"background-color: {self.color_scheme}")
        self.setWindowTitle("Kantar Cleanroom Process Toolkit")
        self.setWindowIcon(QIcon("toolset.jpg"))

        self.statusBar = QStatusBar(self)
        self.statusBar.setObjectName("statusBar")
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage('Status:')
        self.statusBar.setStyleSheet("color: white;")
        self.source = False
        # restoring from the archived log
        try:
            self.arch_app = ArchivingApp(os.getcwd(),"kcrpt.kdpx",'kcrpt.kdpx',os.getcwd(),["*.log"])
            self.arch_app.uncompress()
        except Exception as err:
            msg = f"Problem in restoring the log - {err}"
            self.logger.error(msg)
            self.message_box(msg)
        self.wam_pass = None
        self.kwp_pass = None #Shashank 05032020 added for KWP
        self.csv_files = []
        self._status = ""
        #start logging
        logging.basicConfig(
                        format='%(asctime)s %(levelname)-8s [%(lineno)d %(filename)s] %(message)s',
                        datefmt = '%d-%m-%Y %H:%M:%S',
                        level=logging.DEBUG,
                        filename='pcrpt.log'
                        )
        self.logger = logging.getLogger("PCRPT")
        """DEBUG INFO WARNING ERROR CRITICAL"""
        self.logger.info("A new session started")
        if self.check_reg():
            self.update_status("ready to go")
            self.handle_interactions()
        else:
            self.update_status('registration failed. If you are running first time, send the reg file to update')
    def check_reg(self):
        if os.path.isfile('crreg.bin'):
            self.reg = Registry.restore('crreg.bin')
            if self.reg.check():
                return True
        else:
            self.reg = Registry(**{'expiry_date':"2022-11-15"})     
            return False
    def update_status(self,message):
        self.statusBar.showMessage(f"Status: {message}")
        self.statusBar.repaint()
    @property
    def Status(self):
        return self._status
    @Status.setter
    def Status(self,current_status):
        self._status = current_status
        self.update_status(self._status)

    def closeEvent(self,event):
        # pass
        log_file = open("pcrpt.log",'a')
        log_file.flush()
        log_file.close()
        try:
            self.arch_app.compress(self.wam_pass)
            self.arch_app.compress(self.kwp_pass) #Shashank 05032020 added for KWP
        except Exception as err:
            msg = f"Problem in archiving the log - {err}"
            self.logger.error(msg)
            self.message_box(msg)
        # unlinking matched files - clean the room
        if len(self.csv_files) > 0:
            for file in self.csv_files:
                try:
                    os.remove(file)
                except Exception as err:
                    self.logger.error(f"{file} wasn't found to remove - please check and remove if found")
            # remove if there are old matching files too
            fileList = glob.glob('hs_*_matched_*.csv')
            for filePath in fileList:
                try:
                    os.remove(filePath)
                    self.logger.info(f"{filePath} has been successfully removed and unlinked")
                except:
                    self.logger.warning(f"couldn't remove {filePath}")

    def handle_interactions(self):
        self.main_widget.left_widget.cleanProcess.clicked.connect(self.show_pcrpt)
        self.main_widget.left_widget.export.clicked.connect(self.export)
        self.main_widget.left_widget.log.clicked.connect(self.show_logs)
        self.main_widget.left_widget.matchedData.clicked.connect(self.show_matchedData)
        # self.main_widget.right_widget.source_kwp_data.clicked.connect(self.get_kwp_id_csv)
        # self.main_widget.right_widget.target_hs_prof_kwp_data.clicked.connect(self.get_hs_prof_csv)
        # self.main_widget.right_widget.target_hs_viewership_kwp_data.clicked.connect(self.get_hs_viewership_csv)
        # self.main_widget.right_widget.target_hs_device_kwp_data.clicked.connect(self.get_hs_device_csv)
        # self.main_widget.right_widget.source_wam_data.clicked.connect(self.get_wam_rid_csv)
        # self.main_widget.right_widget.wam_pass.editingFinished.connect(self.on_wam_pass)
        # self.main_widget.right_widget.kwp_pass.editingFinished.connect(self.on_kwp_pass) #Shashank 05032020 added for KWP
        # self.main_widget.right_widget.target_hs_prof_wam_data.clicked.connect(self.get_hs_prof_csv)
        # self.main_widget.right_widget.target_hs_viewership_wam_data.clicked.connect(self.get_hs_viewership_csv)
        # self.main_widget.right_widget.target_hs_device_wam_data.clicked.connect(self.get_hs_device_csv)
        # self.main_widget.right_widget.batch_wam_schema.clicked.connect(self.get_batch_schema)
        # self.main_widget.right_widget.batch_kwp_schema.clicked.connect(self.get_batch_schema) #Shashank 05032020 added for KWP

        # run matching for digital and hs data
        # self.main_widget.right_widget.run_matching_wam.clicked.connect(self.run_matching)
        # self.main_widget.right_widget.run_matching_kwp.clicked.connect(self.run_matching)
        # self.main_widget.right_widget.run_batch_processing_wam.clicked.connect(self.run_batch_wam)
        # self.main_widget.right_widget.run_batch_processing_kwp.clicked.connect(self.run_batch_kwp) #Shashank 05032020 added for KWP

    def run_batch_wam(self):
        self.update_status("Batches are being run - please wait ...")
        batch_schema = self.main_widget.right_widget.batch_wam_schema.text()
        schema_pass = self.main_widget.right_widget.schema_pass.text()
        data_set = 'kwp' if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex() == 0 else 'digital'
        batch = BatchProcessing(batch_schema,schema_pass,data_set)
        self.csv_files = batch.run_batches()
        self.update_status("Batches processing over")

    def run_batch_kwp(self):

        self.update_status("Batches are being run - please wait ...")
        batch_schema = self.main_widget.right_widget.batch_kwp_schema.text()
        schema_kwp_pass = self.main_widget.right_widget.schema_kwp_pass.text()
        data_set = 'kwp' if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==0 else 'digital'
        # self.message_box(f"calling from {data_set}\n{batch_schema}\n{schema_kwp_pass}")
        try:
            batch = BatchProcessing(batch_schema,schema_kwp_pass,data_set)
            self.csv_files = batch.run_batches() #Shashank 12032020 added for kwp
        except Exception as err:
            self.logger.error(f"Problem in running batch processes - {err}")
        self.update_status("Batches processing over")

    def export(self):
        path = QFileDialog.getExistingDirectory(self, "Select Directory to copy matched files to")
        if path != "":
            try:
                for file in self.csv_files:
                    shutil.copy2(file,path)
                msg = "matched files are successfully copied to the selected folder"
                self.update_status(msg)
                self.logger.info(msg)
            except Exception as err:
                self.update_status(f"problem in copying - {err}")

    def get_dfs(self,files):
        dfs = []
        for file in files:
            df = pd.read_csv(file)
            dfs.append(df)
        return dfs
    def logging(self):
        layout = QVBoxLayout()
        self.main_widget.right_widget.tabLog.setLayout(layout)
        layout.setContentsMargins(0,0,0,0)
        logTextBox = QTextEdit()
        logTextBox.clear()
        logTextBox.setStyleSheet('''
                QTextEdit {
                    font: 10pt "Consolas";
                }
            ''')
        self.log_file = open('pcrpt.log','r')
        with self.log_file:
            text = self.log_file.read()
            logTextBox.setText(text)
        logTextBox.repaint()
        layout.addWidget(logTextBox)
        self.setLayout(layout)

    def run_matching(self):
        self.update_status("Please wait while matching processes are in progress")
        salt = ''
        if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==0:
            df_source = self.df_source_kwp
            df_prof = self.df_prof_kwp
            df_viewership = self.df_viewership_kwp
            #df_device = self.df_device_kwp
            salt = self.main_widget.right_widget.salt_kwp.text()
            dfs = [df_source, df_prof, df_viewership]
        elif self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==1:
            if self.source:
                df_source = self.df_source_wam
            else:
                self.logger.error("Kantar Digital source list not found")
                self.update_status('error - Kantar Digital source list not found')
                return
            df_prof = self.df_prof_wam
            df_viewership = self.df_viewership_wam
            df_device = self.df_device_wam
            salt = self.main_widget.right_widget.salt.text()
            dfs = [df_source, df_prof, df_viewership, df_device]
        if salt == "":
            self.update_status("salt has not been entered")
            self.logger.error("Salt not found")
            return

        proc_name = ['prof','viewership','device']
        results = []
        self.update_status("preprocessing complete")
        # hashed_source = Matching.get_hashed(source_list,salt,iters)
        # self.update_status("hash keys were successfully generated for the source list")
        for ind,df in enumerate(dfs[1:]): #looping over from prof
            if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex() == 1:
                if 'phone_no' not in df.columns:
                    self.message_box(f"The key for matching 'phone_no' \nis not found in the target for {proc_name[ind]}")
                    self.update_status("Please make sure the target contains 'phone_no' in hashed format")
                    return
                settings = {
                        "source":dfs[0],
                        "target": df,
                        "salt": salt
                        }
                self.update_status(f"matching for {proc_name[ind]} is being approached")
                match = NewMatching(**settings)
                res, d_nccs = match.get_matches('phone_no')
                results.append(res)
                self.logger.info(f"matching for {proc_name[ind]} completed")
                if ind < 2:
                    self.update_status(f"matching for {proc_name[ind]} completed")
                else:
                    self.update_status(f"matching for {proc_name[ind]} completed. Please wait for few moments more...")

            elif self.main_widget.right_widget.tabWidgetPCRPT.currentIndex() == 0:
                if 'ad_id' not in df.columns:
                    self.message_box(f"The key for matching 'ad_id' \nis not found in the target for {proc_name[ind]}")
                    self.update_status("Please make sure the target contains 'ad_id' in hashed format")
                    return
                settings = {
                    "source": dfs[0],
                    "target": df,
                    "salt": salt
                }
                self.update_status(f"matching for {proc_name[ind]} is being approached")
                match = NewMatching(**settings)
                res = match.get_matches('ad_id')  # Shashank 12032020 added for kwp
                results.append(res)
                self.logger.info(f"matching for {proc_name[ind]} completed")
                if ind < 2:
                    self.update_status(f"matching for {proc_name[ind]} completed")
                else:
                    self.update_status(f"matching for {proc_name[ind]} completed. Please wait for few moments more...")

        # m_dfs = self.unlink_target(dfs,results,salt,iters)
        # NOT SAVING NOW FOR EVALUATION ONLY - sharechat Need to uncomment once the evaluation is over
        m_dfs = self.save_matched_data(dfs, results,d_nccs)
        self.write_back_to_csvs(m_dfs)

        # Shashank 12032020 added for kwp
        if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex() == 0:
            #self.df_source_kwp["ad_id"] = NewMatching.get_hashed(self.df_source_kwp["ad_id"].apply(int).values,salt=salt) # temp commented
            final_dict = dict(zip(self.df_source_kwp["ad_id"], self.df_source_kwp["household_id"]))
            m_dfs = self.replace_ad_id(m_dfs,final_dict)
        self.update_status("")
        log_file = open("pcrpt.log", 'a')
        log_file.flush()
        log_file.close()
        PCRPTool.message_box("Matching process completed")


    def run_matching_ids(self):
        self.update_status("Please wait while matching processes are in progress")
        if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==0:
            df_source = self.df_source_kwp
            df_prof = self.df_prof_kwp
            df_viewership = self.df_viewership_kwp
            df_device = self.df_device_kwp
        elif self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==1:
            if self.source:
                df_source = self.df_source_wam
            else:
                self.logger.error("Kantar Digital source list not found")
                self.update_status('error - Kantar Digital source list not found')
                return
            df_prof = self.df_prof_wam
            df_viewership = self.df_viewership_wam
            df_device = self.df_device_wam
        salt = self.main_widget.right_widget.salt.text()
        if salt == "":
            self.update_status("salt has not been entered")
            self.logger.error("Salt not found")
            return
        iters = 1
        source = df_source['phone_no'].apply(str).values
        # source_list = [str(item).strip() for item in source['phone_no'].values]
        dfs = [df_source,df_prof,df_viewership,df_device]
        dfs = [elem for elem in dfs if elem != None]
        proc_name = ['prof','viewership','device']
        results = []
        self.update_status("preprocessing complete")
        # hashed_source = Matching.get_hashed(source_list,salt,iters)
        # self.update_status("hash keys were successfully generated for the source list")
        df_source_with_NCCS = None
        for ind,df in enumerate(dfs[1:]): #looping over from prof
            if 'phone_no' not in df.columns:
                self.message_box(f"The key for matching 'phone_no' \nis not found in the target for {proc_name[ind]}")
                self.update_status("Please make sure the target contains 'phone_no' in hashed format")
                return
            settings = {
                    "source_list":source,
                    "target_list": df['phone_no'].apply(str).values,
                    "source_hashed": False,
                    "target_hashed": True,
                    "salt": salt,
                    "iterations": iters,
                    }
            self.update_status(f"matching for {proc_name[ind]} is being approached")
            match = Matching(**settings)
            res = match.get_matches()
            results.append(res)
            self.logger.info(f"matching for {proc_name[ind]} completed")
            if ind < 2:
                self.update_status(f"matching for {proc_name[ind]} completed")
            else:
                self.update_status(f"matching for {proc_name[ind]} completed. Please wait for few moments more...")

        # m_dfs = self.unlink_target(dfs,results,salt,iters)
        m_dfs = self.save_matched_data(dfs,results,d_nccs)
        self.write_back_to_csvs(m_dfs)
        log_file = open("pcrpt.log",'a')
        log_file.flush()
        log_file.close()
        PCRPTool.message_box("Matching process completed")

    def save_matched_data(self,dfs,flags,d_nccs):
        # filter on matching
        tdfs = []
        for ind,df in enumerate(dfs[1:]):
            tdf = df[flags[ind]]
            tdf['NCCS'] = tdf['phone_no'].map(d_nccs)
            tdfs.append(tdf)
        return tdfs
 
                                                            
    def replace_ad_id(self,dfs,r_map):
        tdfs = []
        for ind, df in enumerate(dfs):
            #w_df['ad_id'] = w_df['ad_id'].map(r_map)
            #w_df = w_df.rename(columns={"ad_id":"household_id"})
            #w_df['household_id'] = w_df['ad_id'].map(r_map)
            #w_df = w_df.drop('ad_id', axis=1)
            t_df = self.replace_with_household_id(df, r_map)
            t_df = t_df.rename(columns={'ad_id': 'household_id'})
            tdfs.append(t_df)
        return tdfs

    def replace_with_household_id(self,df,d_map):
        for row in range(df.shape[0]):
            ad_id = df.iloc[row]['ad_id']
            df.iloc[row]['ad_id'] = d_map[ad_id]
        return df


    def unlink_target(self,dfs,flags,salt,iters):
        hashed_phone_wam1 = Matching.get_hashed(dfs[0]['phone_no'].apply(str).values,salt,iters)
        hashed_phone_wam2 = Matching.get_hashed(dfs[0]['phone_no'].apply(str).values,'zen::',iters)
        self.update_status('Close to completing ...few moments more')
        # filter on matching
        tdfs = []
        for ind,df in enumerate(dfs[1:]):
            tdf = df[flags[ind]]
            tdfs.append(tdf)
        map_dict = dict(zip(hashed_phone_wam1,hashed_phone_wam2))
        _dfs = []
        for df in tdfs:
            df['phone_no'] = df['phone_no'].map(map_dict)
            _dfs.append(df)
        self.update_status('')
        return _dfs

    def write_back_to_csvs(self,dfs):
        now = datetime.now()
        dt_string = f"{now.year}{now.month}{now.day}_{now.hour}{now.minute}"
        if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex() == 1:
            self.csv_files = [f"_prof_matched_{dt_string}.csv",f"_viewership_matched_{dt_string}.csv",f"_device_matched_{dt_string}.csv"]
        else:
            self.csv_files = [f"_prof_matched_{dt_string}.csv", f"_viewership_matched_{dt_string}.csv"]

        for ind,df in enumerate(dfs):
            df.to_csv(self.csv_files[ind],index=False)
            self.logger.info(f"{self.csv_files[ind]} has been saved to the current folder")
        # update the csv_file names for right_widget
        self.main_widget.right_widget.csv_files = self.csv_files

    def get_vars_from_file(self,file,file_type="csv"):
        if file_type=="csv":
            try:
                df = pd.read_csv(file)
            except Exception as err:
                msg = f"Couldn't open {file} - {err}"
                self.logger.error(msg)
                self.update_status(msg)
                return (None,None)
        elif file_type=="json":
            df = pd.read_json(file)
        heads = df.columns
        return (heads,df)

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

    def get_wam_rid_csv(self):
        file = self.get_file("Select Kantar Digital RID File",'xlsx')
        if file != None:
            self.main_widget.right_widget.source_wam_data.setText(file)

    def get_batch_schema(self):
        file = self.get_file("Select the schema file for batch processing",'xlsx')
        if file != None:
            self.main_widget.right_widget.batch_wam_schema.setText(file)
            self.main_widget.right_widget.batch_kwp_schema.setText(file) #Shashank 05032020 added for KWP

    def on_wam_pass(self):
        xlsx_file = self.main_widget.right_widget.source_wam_data.text()
        wam_pass = self.main_widget.right_widget.wam_pass.text()
        try:
            self.source = False
            self.wam_pass = self.ensureUtf(wam_pass)
            ewp = ExcelWithPassword(xlsx_file,self.wam_pass,self.wam_pass)
            if ewp.open_xlsx():
                self.df_source_wam = ewp.get_df()
                msg = f"A total of {self.df_source_wam.shape[0]} records have been traced for the source list in {xlsx_file}"
                self.logger.info(msg)
                self.update_status(msg)
                self.source = True
            else:
                self.update_status(f"Failed to read the source list - {xlsx_file}")
        except Exception as err:
            msg = f"Problem with opening Kantar Digital source list - {err}"
            self.logger.error(msg)
            self.update_status(msg)

    def get_kwp_id_csv(self):
        file = self.get_file("Select KWP Id File",'xlsx')
        if file != None:
            self.main_widget.right_widget.source_kwp_data.setText(file)

    #Shashank 05032020 added for KWP
    def on_kwp_pass(self):
        xlsx_file = self.main_widget.right_widget.source_kwp_data.text()
        kwp_pass = self.main_widget.right_widget.kwp_pass.text()
        try:
            self.source = False
            self.kwp_pass = self.ensureUtf(kwp_pass)
            ewp = ExcelWithPassword(xlsx_file,self.kwp_pass,self.kwp_pass)
            if ewp.open_xlsx():
                #self.df_source_kwp = ewp.get_df()
                self.df_source_kwp = ewp.get_df_kwp() #Shashank 12032020 added for kwp
                msg = f"A total of {self.df_source_kwp.shape[0]} records have been traced for the source list in {xlsx_file}"
                self.logger.info(msg)
                self.update_status(msg)
                self.source = True
            else:
                self.update_status(f"Failed to read the source list - {xlsx_file}")
        except Exception as err:
            msg = f"Problem with opening Kantar Worldpanel source list - {err}"
            self.logger.error(msg)
            self.update_status(msg)
    # Shashank 05032020 added for KWP

    def get_hs_prof_csv(self):
        file = self.get_file("Select Profile",'csv')
        list_widget = self.main_widget.right_widget.list_prof_kwp
        if file != None:
            vars,df = self.get_vars_from_file(file,"csv")
            if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==0:
                self.main_widget.right_widget.target_hs_prof_kwp_data.setText(file)
                self.df_prof_kwp = df
            elif self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==1:
                self.main_widget.right_widget.target_hs_prof_wam_data.setText(file)
                list_widget = self.main_widget.right_widget.list_prof_wam
                self.df_prof_wam = df
            list_vars = ["".join(var) for var in vars]
            self.logger.info(f"vars uploaded for prof - {list_vars}")
            msg = f"A total of {df.shape[0]} records have been traced from {file}"
            self.logger.info(msg)
            self.update_status(msg)
            self.load_listWidget(vars,list_widget)


    def get_hs_viewership_csv(self):
        file = self.get_file("Select Viewership Data",'csv')
        if file != None:
            list_widget = self.main_widget.right_widget.list_viewership_kwp
            if file != None:
                vars,df = self.get_vars_from_file(file,"csv")
                if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==0:
                    self.main_widget.right_widget.target_hs_viewership_kwp_data.setText(file)
                    self.df_viewership_kwp = df
                elif self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==1:
                    self.main_widget.right_widget.target_hs_viewership_wam_data.setText(file)
                    list_widget = self.main_widget.right_widget.list_viewership_wam
                    self.df_viewership_wam = df

            list_vars = ["".join(var) for var in vars]
            self.logger.info(f"vars uploaded for Viewership - {list_vars}")
            msg = f"A total of {df.shape[0]} records have been traced from {file}"
            self.logger.info(msg)
            self.update_status(msg)
            self.load_listWidget(vars,list_widget)

    def get_hs_device_csv(self):
        file = self.get_file("Select Device Data",'csv')
        if file != None:
            list_widget = self.main_widget.right_widget.list_device_kwp
            vars, df = self.get_vars_from_file(file,"csv")
            if file != None:
                if self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==0:
                    self.main_widget.right_widget.target_hs_device_kwp_data.setText(file)
                    self.df_device_kwp = df
                elif self.main_widget.right_widget.tabWidgetPCRPT.currentIndex()==1:
                    self.main_widget.right_widget.target_hs_device_wam_data.setText(file)
                    list_widget = self.main_widget.right_widget.list_device_wam
                    self.df_device_wam = df
            list_vars = ["".join(var) for var in vars]
            self.logger.info(f"vars uploaded for device - {list_vars}")
            msg = f"A total of {df.shape[0]} records have been traced from {file}"
            self.logger.info(msg)
            self.update_status(msg)
            self.load_listWidget(vars,list_widget)

    def get_file(self,title,type):
        cwd = os.getcwd()
        openFileName = QFileDialog.getOpenFileName(self, title, cwd,("CSV (*.csv)" if type=='csv' else ("MS Excel (*.xlsx)")))
        if openFileName != ('', ''):
            file = openFileName[0]
            self.logger.info(f"{file} has been selected")
            return file
        else:
            return None

    def get_xlsx(self):
        cwd = os.getcwd()
        openFileName = QFileDialog.getOpenFileName(self, 'Open File', cwd,"Excel Files (*.xlsx)")
        if openFileName != ('', ''):
            file = openFileName[0]
            try:
                wb = openpyxl.load_workbook(file)
                ws = wb.worksheets
                sheets = []
                for sheet in ws:
                    sheets.append(sheet.title)
                return (file,sheets)
            except Exception as err:
                self.message_box(str(err))
                self.logger.error(err)
        else:
            return (None,None)
    def get_vars_from_listbox(self,listBox):
        vars_all = []
        vars_checked = []
        indices = []
        count = listBox.count()
        indx = 0
        for item in range(count):
            itm = listBox.item(item)
            vars_all.append(itm.text())
            if itm.checkState() == Qt.Checked:
                vars_checked.append(itm.text())
                indices.append(indx)
            indx += 1
        return {"all":vars_all,"checked":vars_checked,"indices":indices}

    def load_listWidget(self,var_list,list_widget,checkable=True):
        list_widget.clear()
        for var in var_list:
            item = QListWidgetItem(var, list_widget)
            if checkable==True:
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
        list_widget.repaint()

    def show_pcrpt(self):
        self.main_widget.right_widget.tabWidget.setCurrentIndex(0)
    def show_export(self):
        self.main_widget.right_widget.tabWidget.setCurrentIndex(1)
    def show_logs(self):
        self.main_widget.right_widget.tabWidget.setCurrentIndex(2)
        self.update_status("")
        self.logging()
    def show_matchedData(self):
        self.main_widget.right_widget.tabWidget.setCurrentIndex(3)
        self.main_widget.right_widget.show_match_datadfs()
        self.main_widget.right_widget.tabWidgetMatchedData.currentChanged.connect(self.onChange_matched_data) #changed!

    def onChange_matched_data(self,i):
        shapes = self.main_widget.right_widget.shapes
        dataset = ["Profile","Viewership","Device"]
        self.update_status(f"{dataset[i]} dataset has {shapes[i][0]} records with {shapes[i][1]} variables")

    @staticmethod
    def message_box(text_message):
        mb = QMessageBox()
        mb.setStyleSheet("background-color: #646665;color: white")
        mb.setIcon(QMessageBox.Information)
        mb.setWindowTitle('Panel Clean Room Process Tool')
        mb.setText(text_message)
        mb.setStandardButtons(QMessageBox.Ok)
        mb.exec()




class MainWidget(QWidget):
    def __init__(self,parent=None):
        super().__init__(parent)
        layout = QHBoxLayout()
        self.left_widget = LeftWidget(self)
        self.right_widget = RightWidget(self)
        layout.addWidget(self.left_widget)
        layout.addWidget(self.right_widget)
        self.setLayout(layout)
        self.showMaximized()

class LeftWidget(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.layout = QVBoxLayout(self)
        self.color_scheme = '#6E7073'
        # pcrpt
        self.cleanProcess = QPushButton("") # cleanProcess
        self.cleanProcess.setToolTip("Matching Tools")
        self.cleanProcess.setStyleSheet("QPushButton{ background-color: #ffffff; color: k ; font: 10pt;border-radius:25;margin: 2px}")
        self.cleanProcess.setIcon(QIcon("mapping.jpg"))
        self.cleanProcess.setIconSize(QSize(128,128))
        self.setStyleSheet("QToolTip{border: 1px solid white; font-size: 8pt;}")
        self.layout.addWidget(self.cleanProcess)

        # show matched data
        self.matchedData = QPushButton("") # matched data
        self.matchedData.setToolTip("Show Matched Data for Profile/ Viewership/ Device")
        self.matchedData.setIcon(QIcon("matched.png"))
        self.matchedData.setIconSize(QSize(128,128))
        self.matchedData.setStyleSheet("QPushButton{ background-color: #ffffff; color: k ; font: 10pt;border-radius:25;margin: 2px}")
        self.setStyleSheet("QToolTip{border: 1px solid white; font-size: 8pt;}")
        self.layout.addWidget(self.matchedData)

        # export
        self.export = QPushButton("") # exports
        self.export.setToolTip("Export Matched Files to External Device")
        self.export.setIcon(QIcon("file_copy.jpg"))
        self.export.setIconSize(QSize(128,128))
        self.export.setStyleSheet("QPushButton{ background-color: #ffffff; color: k ; font: 10pt;border-radius:25;margin: 2px}")
        self.setStyleSheet("QToolTip{border: 1px solid white; font-size: 8pt;}")
        self.layout.addWidget(self.export)

        # logging
        self.log = QPushButton("") # exports
        self.log.setToolTip("Show Logs")
        self.log.setIcon(QIcon("log.jpg"))
        self.log.setIconSize(QSize(128,128))
        self.log.setStyleSheet("QPushButton{ background-color: #ffffff; color: k ; font: 10pt;border-radius:25;margin: 2px}")
        self.setStyleSheet("QToolTip{border: 1px solid white; font-size: 8pt;}")
        self.layout.addWidget(self.log)

        self.layout.addStretch(1)
        self.setStyleSheet("padding: 0px; margin: 0px; background-color: black;")
        self.setLayout(self.layout)

class ClickLabel(QLabel):
    clicked = pyqtSignal()

    def mousePressEvent(self, event):
        self.clicked.emit()
        QLabel.mousePressEvent(self, event)

class QCLineEdit(QLineEdit):
    clicked = pyqtSignal()
    def mousePressEvent(self, event):
        self.clicked.emit()
        QLineEdit.mousePressEvent(self, event)

class RightWidget(QWidget):
    def __init__(self, parent):
        super().__init__()
        layout = QHBoxLayout(self)
        self.color_scheme = '#ffffff' #'#737373' #'#6E7073'
        self.tabWidget = QTabWidget()
        self.tabPCRPTool = QWidget()
        self.tabWidget.addTab(self.tabPCRPTool ,"Panel Cleanroom Process")
        self.tabExport = QWidget()
        self.tabWidget.addTab(self.tabExport ,"Export")
        self.tabLog = QWidget()
        self.tabWidget.addTab(self.tabLog ,"Logs")
        self.tabMatchedData = QWidget()
        self.tabWidget.addTab(self.tabMatchedData,"Matched Data")
        layout.addWidget(self.tabWidget)
        self.csv_files = None

        # self.layout.addStretch(1)
        layout.setStretch(0, 1)
        self.setStyleSheet(f"padding: 0px; margin: 0px; background-color: {self.color_scheme};")
        self.setLayout(layout)
        self.logger = logging.getLogger("PCRPT")
        self.PCRPT()
        # self.Export()

        self.logger = logging.getLogger("PCRPT")

    def PCRPT(self):
        layout = QHBoxLayout(self)
        self.tabPCRPTool.setLayout(layout)
        layout.setContentsMargins(0,0,0,0)

        self.color_scheme = '#6E7073'
        self.tabWidgetPCRPT = QTabWidget()
        self.tabCR = QWidget()
        self.tabWidgetPCRPT.addTab(self.tabCR ,"Kantar Clean-room 2.0")
        layout.addWidget(self.tabWidgetPCRPT)
        layout.setStretch(0, 1)
        self.clean_room_ui()


    def clean_room_ui(self):
        grid_layout = QGridLayout()
        self.tabCR.setLayout(grid_layout)
        grid_layout.setContentsMargins(0,0,0,0)
        grid_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)    # align grid to top and left

        kantar_logo = QLabel(self)
        pixmap = QPixmap('kantar.jpg')
        kantar_logo.setPixmap(pixmap)
        kantar_logo.setStyleSheet("margin: 10px")
        # Optional, resize window to image size
        self.resize(pixmap.width(),pixmap.height())
        grid_layout.addWidget(kantar_logo,0,4,Qt.AlignTop | Qt.AlignRight)
        title = QLabel("Kantar Clean-room for Matching Data on PII")
        title.setStyleSheet("font-size: 18pt;color: black;text-align: center; margin: 5px 0px")
        grid_layout.addWidget(title,1,0,1,4,Qt.AlignTop | Qt.AlignLeft)

        self.source_lab = QLabel("Source dataset(*.xlsx), password and PII")
        self.source_lab.setStyleSheet("font-size: 10pt;color: black;margin: 5px 0px;")
        self.source_data = QCLineEdit() #QFileDialog.getExistingDirectory()
        self.source_data.setStyleSheet("font-size: 10pt;background-color: #bac8e0;color: black;width: 500px;")
        self.source_pass = QLineEdit()
        # self.source_pass.setText("kt346B") # fixing the password for unattending matching session
        self.source_pass.setStyleSheet("font-size: 10pt;background-color: #bac8e0;lineedit-password-character: 42;width: 150px")
        self.source_pass.setEchoMode(QLineEdit.Password)
        self.cb_pii_source = QComboBox()
        self.cb_pii_source.setStyleSheet("font-size: 10pt;color: black;margin: 5px 0px; width: 150px") 
        self.target_lab = QLabel("Target Dataset(*.xlsx), password and PII")
        self.target_lab.setStyleSheet("font-size: 10pt;color: black;margin: 5px 0px")
        self.target_data = QCLineEdit() #QFileDialog.getExistingDirectory()
        self.target_data.setStyleSheet("font-size: 10pt;background-color: #bac8e0;color: black;width: 500px;")
        self.target_pass = QLineEdit()
        # self.target_pass.setText("kt346B") # fixing the password for unattending matching session
        self.target_pass.setStyleSheet("font-size: 10pt;background-color: #bac8e0;lineedit-password-character: 42;width: 150px")
        self.target_pass.setEchoMode(QLineEdit.Password)
        self.cb_pii_target = QComboBox()
        self.cb_pii_target.setStyleSheet("font-size: 10pt;color: black;margin: 5px 0px;width: 150px") 
        empty_line = QLabel("")
        grid_layout.addWidget(empty_line,2,0)
        grid_layout.addWidget(self.source_lab,3,0, Qt.AlignTop | Qt.AlignRight)
        grid_layout.addWidget(self.source_data,3,1,Qt.AlignLeft)
        grid_layout.addWidget(self.source_pass,3,2,Qt.AlignLeft | Qt.AlignTop)
        grid_layout.addWidget(self.cb_pii_source,3,3,Qt.AlignLeft | Qt.AlignTop)
        grid_layout.addWidget(self.target_lab,4,0,Qt.AlignRight)
        grid_layout.addWidget(self.target_data,4,1,Qt.AlignLeft)
        grid_layout.addWidget(self.target_pass,4,2,Qt.AlignLeft | Qt.AlignTop)
        grid_layout.addWidget(self.cb_pii_target,4,3,Qt.AlignLeft | Qt.AlignTop)
        # select_data = QLabel("Variables from source and target")
        # select_data.setStyleSheet("font-size: 10pt;color: black;margin: 5px 0px;")
        # self.list_source = ListWidget(self.tabCR)
        # self.list_source.setToolTip("variables to export from source dataset")
        # self.list_source.setStyleSheet("""QScrollBar:vertical {width: 30px;margin: 3px 0 3px 0;}
        #                                     QListWidget {width: 250px;height: 50px;}
        #                                     QListWidget {background-color: #bac8e0;color: black;}
        #                                     """)
        # self.list_target = ListWidget(self.tabCR)
        # self.list_target.setStyleSheet("""QScrollBar:vertical {width: 30px;margin: 3px 0 3px 0;}
        #                                         QListWidget {width: 250px;height: 50px;}
        #                                         QListWidget {background-color: #bac8e0;color: black;}
        #                                         """)
        # self.list_target.setToolTip("variables to export from target dataset")
        self.chkbtn_clean_room = QCheckBox("Clean Room on Exit")
        self.chkbtn_clean_room.setChecked(True)
        self.chkbtn_clean_room.setToolTip("Removes the matched files on exit from the app")
        self.chkbtn_clean_room.setStyleSheet("font-size: 10pt;background-color: white;color: black;")
        self.chkbtn_clean_room.setEnabled(False)
        # grid_layout.addWidget(select_data,5,0,Qt.AlignTop | Qt.AlignRight)
        # grid_layout.addWidget(self.list_source,6,1,Qt.AlignLeft | Qt.AlignTop)
        # grid_layout.addWidget(self.list_target,6,2,Qt.AlignLeft | Qt.AlignTop)
        grid_layout.addWidget(self.chkbtn_clean_room,5,1,Qt.AlignLeft)
       
        grid_layout.addWidget(empty_line,7,0)
        self.run_matching = QPushButton("Run Matching Process")
        self.run_matching.setStyleSheet("font-size: 10pt; height: 75px; width: 500px;background-color: black; color: white;margin: 0px; border-radius:20px")
        grid_layout.addWidget(self.run_matching,8,1,1,2,Qt.AlignTop | Qt.AlignLeft)
       
    def show_match_datadfs(self):
        layout = QHBoxLayout(self)
        self.tabMatchedData.setLayout(layout)
        layout.setContentsMargins(0,0,0,0)
        self.color_scheme = '#6E7073'
        self.tabWidgetMatchedData = QTabWidget()
        self.tabProf, self.tabViewership, self.tabDevice = QWidget(),QWidget(),QWidget()
        self.tabWidgetMatchedData.addTab(self.tabProf ,"Profile")
        self.tabWidgetMatchedData.addTab(self.tabViewership ,"Viewership")
        self.tabWidgetMatchedData.addTab(self.tabDevice,"Device")
        layout.addWidget(self.tabWidgetMatchedData)
        self.shapes = []
        tabs = [self.tabProf,self.tabViewership,self.tabDevice]
        for i,file in enumerate(self.csv_files):
            self.shapes.append(self.upload_df(file,tabs[i]))
        # self.shapes.append(self.upload_df("hs_viewership_matched.csv",self.tabViewership))
        # self.shapes.append(self.upload_df("hs_device_matched.csv",self.tabDevice))

    def upload_df(self,csv_file,widget):
        layout = QVBoxLayout(self)
        widget.setLayout(layout)
        tableView = QTableView(self)
        tableView.setStyleSheet("font-size: 10 pt;")
        layout.addWidget(tableView)
        tableView.setSortingEnabled(True)
        if os.path.isfile(csv_file):
            df = pd.read_csv(csv_file)
            model = PandasModel(df)
            tableView.setModel(model)
            return df.shape
        else:
            self.logger.info(f"{csv_file} was not found for displaying")
        return (0,0)

    def invoke_explorer(self):
        path = self.img_path_folder.text()
        path=os.path.realpath(path)
        os.startfile(path)

    def get_xlsx(self):
        cwd = os.getcwd()
        openFileName = QFileDialog.getOpenFileName(self, 'Open File', cwd,"Excel Files (*.xlsx)")
        if openFileName != ('', ''):
            file = openFileName[0]
            try:
                wb = openpyxl.load_workbook(file)
                ws = wb.worksheets
                sheets = []
                for sheet in ws:
                    sheets.append(sheet.title)
                self.logger.info(f"{file} has been uploaded")
                return (file,sheets)
            except Exception as err:
                self.message_box(str(err))
                self.logger.error(err)
        else:
            return (None,None)


class ListWidget(QListWidget):
  def __init__(self, parent):
    super().__init__(parent)
    # self.setAcceptDrops(True)
  def sizeHint(self):
    s = QSize()
    s.setHeight(super(ListWidget,self).sizeHint().height())
    s.setHeight(200)
    s.setWidth(500)
    return s

    def mimeTypes(self):
        mimetypes = super().mimeTypes()
        mimetypes.append('text/plain')
        return mimetypes

    def dropMimeData(self, index, data, action):
        if data.hasText():
            self.addItem(data.text())
            return True
        else:
            return super().dropMimeData(index, data, action)

if __name__=="__main__":
    if QApplication.instance():
        app = QApplication.instance()
    else:
        app = QApplication(sys.argv)
    pcrpt = PCRPTool()
    pcrpt.show()
    sys.exit(app.exec_())
