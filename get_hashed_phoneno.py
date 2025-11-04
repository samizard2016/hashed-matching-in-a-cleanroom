from matching import Matching
import pandas as pd

df = pd.read_csv("test_phone_no.csv")
act_phone_no = df.phone_no.values

def add_salt(item):
    return str(item)+"abc12xyz"
phone_no = df.phone_no.apply(add_salt).values
hashed_phone_nos = Matching.get_hashed(phone_no,'',1)
t_df = pd.DataFrame({"phone_no": act_phone_no,"hashed_phone_no": hashed_phone_nos})
t_df.to_csv("hashed_test_phone_no.csv",index=False)
print("Saved to hashed_test_phone_no.csv")
