import pandas as pd
import time
import csv
import glob
import os
import re
import TPW_MYOB_link

print("****WELCOME TO TPW-MYOB Upload File Generator****")
print("")
print("Note:Please make sure to export the latest file from TPW website")
print("Note:Keep the file in Export Folder")
print("Note:Find the upload file in the Upload FOLDER")
print("")


# list_of_files = glob.glob('Export File/*') # * means all if need specific format then *.csv
# latest_file = max(list_of_files, key=os.path.getctime)
# print("Using File:   "+latest_file+"    PLEASE CHECK!!")

# data=pd.read_csv(latest_file)
# #data.head()

# data=data[(data["Backorder Date"].isnull()) & (data["Order Status"]=="Pending Ship Confirmation")]
# data=data[["PO #","Item#","Quantity","Wholesale Price"]]
# #data.head()

# def split_it(item):
#     item_name=re.findall('\=\"(.*)\"', item)
#     return item_name[0]

# data['Item#'] = data['Item#'].apply(split_it)
# #data.head()

# #data.shape
# #data.head()

# data.reset_index(inplace = True) 
# #data.head(30)

# data_template=pd.read_csv("TPW_template_file/tpw_template.txt",sep='\t',header=None)
# #data_template.head(30)

# order_count=data.shape[0]
# #order_count

# keep_frame=data_template[2:order_count+2].reset_index()
# del keep_frame["index"]
# #keep_frame.shape

# template_head=data_template[0:2]
# #template_head
# #keep_frame.head(30)
# keep_frame[12]=data["Item#"]
# keep_frame[13]=data["Quantity"]
# keep_frame[14]=data["PO #"]
# keep_frame[15]=data["Wholesale Price"]
# keep_frame[17]=(keep_frame[15]*keep_frame[13])
# keep_frame[26]=(keep_frame[17]/10)
# keep_frame[15] = keep_frame[15].apply(lambda x: "${:.2f}".format((x)))
# keep_frame[17] = keep_frame[17].apply(lambda x: "${:.2f}".format((x)))
# keep_frame[26] = keep_frame[26].apply(lambda x: "${:.2f}".format((x)))
# #keep_frame[26].head()
# #keep_frame[12].head(30)
# upload_data=template_head.append(keep_frame, ignore_index = True) 
# del upload_data[49]
# #upload_data.head()
# timestr = time.strftime("%Y.%m.%d-%H.%M.%S")
# #print(timestr)
# filename="Upload_File/Import_MYOB_"+timestr
# upload_data.to_csv(filename, sep="\t", quoting=csv.QUOTE_NONE,index=False,header=False)

while True:

    TPW_MYOB_link.converterProgram()
    #print("CONVERSION PROCESS COMPLETED PLEASE CHECK!")
    
    while True:
        answer = input('Run again? (y/n): ')
        if answer in ('y', 'n'):
            break
        print('Invalid input.')
    if answer == 'y':
        #cls()
        #print(chr(27) + "[2J")
        os.system("cls")
        reload(TPW_MYOB_link)
        continue
    else:
        print('Goodbye')
        break