'''
https://docs.python.org/3/library/hashlib.html
paul, samir
29 January 2020
Pune, Maharashtra
'''
import hashlib
import logging
import pandas as pd

class Matching:
    def __init__(self,**kwargs):
        self.options = dict()
        self.options.update(**kwargs)
        self.logger = logging.getLogger("PCRPT")

    def get_matches(self):
        self.__source__ = self.options["source_list"]
        self.__target__ = self.options['target_list']
        self.__source_hashed__ = self.options['source_hashed']
        self.__target_hashed__ = self.options['target_hashed']
        self.logger.info(f"lists of {len(self.__source__)} and {len(self.__target__)} were traced for source and target respectively")
        if 'salt' in self.options:
            self.__salt__ = self.options['salt']
        if 'iterations' in self.options:
            self.__iterations__ = self.options['iterations']

        if self.__source_hashed__ and self.__target_hashed__:
            matches = [True if item in self.__source__ else False for item in self.__target__]
            self.logger.info(f"A total of {matches.count(True)} records were matched")
        else:
            hashed_source = self.get_hashed(self.__source__,self.__salt__,self.__iterations__)
            matches = [True if item in hashed_source else False for item in self.__target__]
            self.logger.info(f"A total of {matches.count(True)} records were matched")
        return matches
    @classmethod
    def get_hashed(cls,source,salt,iterations):
        hashed_items = [hashlib.pbkdf2_hmac('sha256', str.encode(item), str.encode(salt), iterations).hex()
            for item in source]
        return hashed_items

class NewMatching:
    def __init__(self,**kwargs):
        self.__source__ = kwargs['source']
        self.__target__ = kwargs['target']
        if 'salt' in kwargs:
            self.__salt__ = kwargs['salt']
        self.logger = logging.getLogger("PCRPT")
    def get_matches(self,match_key):
        # get hashed values
        # self.__source__[match_key].to_csv('source values.csv')
        hashed_phone_no = NewMatching.get_hashed(self.__source__[match_key].apply(self.convert_to_int).values,salt=self.__salt__)#temp commented
        tdf = pd.DataFrame({"hashed": hashed_phone_no,match_key: self.__source__[match_key].values,'NCCS':self.__source__['NCCS']})
        d_nccs = dict(zip(tdf['hashed'],tdf['NCCS'])) 
        try:
            # tdf = pd.DataFrame({"hashed": self.__source__[match_key].values, match_key: self.__source__[match_key].values})
            self.logger.info(f"lists of {self.__source__.shape[0]} and {self.__target__.shape[0]} were traced for source and target respectively")
            # xdf = pd.DataFrame({'source':tdf['hashed'].values[:300],'target':self.__target__[match_key].values})
            # xdf.to_csv('xdf.csv')
            matches = self.__target__[match_key].isin(tdf['hashed']).astype(bool).to_list()
            self.logger.info(f"A total of {matches.count(True)} records were matched from a total of {len(matches)}")
            return matches, d_nccs
        except Exception as e:
            self.logger.error(f"problem in matching: {e}")
    def convert_to_int(self,val):
        if val != None:
            try:
                return int(val)
            except Exception as e:
                self.logger.error(f"matching: convert_to_int - invalid phone no ({val}): {e}")
        else:
            self.logger.error("Empty match key")
            return

    #Shashank 12032020 added for kwp
    # def get_matches_kwp(self):
    #     # get hashed values
    #     hashed_ad_id = NewMatching.get_hashed(self.__source__['ad_id'].apply(self.convert_to_int_kwp).values,salt=self.__salt__)
    #     tdf = pd.DataFrame({"hashed_ad_id": hashed_ad_id,"ad_id": self.__source__['ad_id'].values})
    #     # tdf.to_csv("dummy.csv",index=False)
    #     self.logger.info(f"lists of {self.__source__.shape[0]} and {self.__target__.shape[0]} were traced for source and target respectively")
    #     matches = self.__target__['ad_id'].isin(tdf['hashed_ad_id']).astype(bool).to_list()
    #     self.logger.info(f"A total of {matches.count(True)} records were matched from a total of {len(matches)}")
    #     return matches
    # def convert_to_int_kwp(self,val):
    #     if val != None:
    #         return int(val)
    #     else:
    #         self.logger.error("Empty ad_id")
    #         return
    # Shashank 12032020 added for kwp

    @classmethod
    def get_hashed(cls,source,salt):
        hashed_items = [NewMatching.generate_hash(item,salt) for item in source]
        return hashed_items
    @classmethod
    def generate_hash(cls,item,salt):
        try:
            xitem = str(item).strip()+str(salt).strip()
            hp = hashlib.sha256()
            hp.update(str.encode(xitem))
            return hp.hexdigest()
        except Exception as err:
            cls.logger.error("Missing phone number or empty row")
            return None
def main():
    # from datetime import datetime
    # t1 = datetime.now()
    # df = pd.read_csv("kantar_digital_phone_nos.csv")
    # hashed_items = NewMatching.get_hashed(df.phone_no.values,salt="abc12xyz")
    # tdf = pd.DataFrame({"phone_no":df.phone_no.values,"hashed_phone_no":hashed_items})
    # tdf.to_csv("test_hashed_phone_no.csv",index=False)
    # print("Saved to test_hashed_phone_no.csv")
    # t2 = datetime.now()
    # print(f"Time taken : {t2-t1} secs")
    # p = '8018250746'
    p = '6000030759'
    _hp = NewMatching.generate_hash(p,'qwq23hsd')
    print(_hp)
if __name__ == "__main__":
    main()
    # p = '8018250746'+'abc12xyz'
    # hp = hashlib.sha256()
    # hp.update(str.encode(p))
    # ph = hp.hexdigest()
    # print(ph)
