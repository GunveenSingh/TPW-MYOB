###Change log###
#Date:22/02/2021
#Issue with PDFs not comming in right order fixed
#invoice number is now posted by date and time and identified using the same
#Date:19/02/2021
#Issue with quantity available vs SOH fixed
#Issue with PDFs not comming in right order fixed
#Future change Zanui FTP fix


import pandas as pd
import time
import csv
import glob
import os
import re
import pdfplumber
from PyPDF2 import PdfFileReader, PdfFileWriter
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pandas import ExcelWriter
import numpy as np
import requests
import warnings
from ftplib import FTP
from pathlib import Path
import socket
from datetime import date
import pysftp

def converterProgram():
    
    try:
        # Step 1 : Create upload File
        list_of_files = glob.glob('Export File/*') # * means all if need specific format then *.csv
        latest_file = max(list_of_files, key=os.path.getctime)
        print("Using File:   "+latest_file+"    PLEASE CHECK!!")
        warnings.filterwarnings("ignore")
        data=pd.read_csv(latest_file)
        #data.head()

        data=data[(data["Backorder Date"].isnull()) & (data["Order Status"]=="Pending Ship Confirmation")]
        data=data[["PO #","Item#","Quantity","Wholesale Price"]]
        #print(data.head())

        def split_it(item):
            item_name=re.findall('\=\"(.*)\"', item)
            #print(type(item_name))
            if len(item_name) == 1:
                #print("in if")
                #print(item_name)
                return item_name[0]
            else:
                #print(item)
                #print("in else")
                return item

        data['Item#'] = data['Item#'].apply(split_it)
        #data.head()

        #data.shape
        #data.head()

        data.reset_index(inplace = True) 
        #data.head(30)

        data_template=pd.read_csv("Program-Files/TPW_template_file/tpw_template.txt",sep='\t',header=None)
        #data_template.head(30)

        order_count=data.shape[0]
        #order_count

        keep_frame=data_template[2:order_count+2].reset_index()
        del keep_frame["index"]
        #keep_frame.shape

        template_head=data_template[0:2]
        #template_head
        #keep_frame.head(30)
        keep_frame[12]=data["Item#"]
        keep_frame[13]=data["Quantity"]
        keep_frame[14]=data["PO #"]
        keep_frame[15]=data["Wholesale Price"]
        keep_frame[17]=(keep_frame[15]*keep_frame[13])
        keep_frame[26]=(keep_frame[17]/10)
        keep_frame[15] = keep_frame[15].apply(lambda x: "${:.2f}".format((x)))
        keep_frame[17] = keep_frame[17].apply(lambda x: "${:.2f}".format((x)))
        keep_frame[26] = keep_frame[26].apply(lambda x: "${:.2f}".format((x)))
        #keep_frame[26].head()
        #keep_frame[12].head(30)
        #Reverse the order of the dataframe to match with the website
        keep_frame=keep_frame.iloc[::-1]
        ###################OLD##########################
        # upload_data=template_head.append(keep_frame, ignore_index = True) 
        # del upload_data[49]
        # #upload_data.head()
        # timestr = time.strftime("%Y.%m.%d-%H.%M.%S")
        # #print(timestr)
        # filename="Upload_File/Import_MYOB_"+timestr
        # upload_data.to_csv(filename, sep="\t", quoting=csv.QUOTE_NONE,index=False,header=False)
        ##############################################
        ######################NEW#######################

        #Get DF with only required columns

        upload_data=keep_frame[[12,13,14,15,17]].reset_index()
        del upload_data['index']
        upload_data.columns = ['Number', 'ShipQuantity','Description','UnitPrice','Total']
        #print("comment line below")
        #print(upload_data.head())

        #Get UIDs of all items
        ######OLD CODE################################
#         #Get all item UID in ascending order

#         url = "https://secure.myob.com/oauth2/v1/authorize/"

#         payload = 'client_id=tdd4dk347sw7nvdnr8tvbwta&client_secret=ZaCSqqus8DrEfC3fr3VvWr8C&grant_type=refresh_token&refresh_token=Pkv5%21IAAAAIKM9DYl342RjLtg3HZcfJn2zxEwrLBSUcjcm0ZRC98HsQAAAAHj3FUWHNzI_wS0ihflLHgVuDVqQNg7GNf4_TY3xUYRNN99_92IDs-h9IvwWiGO_lFPjc9HvWYlOUZtoiid1XC8u4VoZa3AruyasYVCVAhSh_qOi9qZ6zSQvPfqsOFEnFxPC5bQzHcMk7k1Ote7wxPzv8VxGKHHEw0C4sX8a0hnnQS7LGVgfTUvLIejFBQAjRIoD5sz_0lVZJoGnN3PPvU3rnVMICP3x-w7rQpXzoDi3w'
#         headers = {
#           'Content-Type': 'application/x-www-form-urlencoded'
#         }

#         response = requests.request("POST", url, headers=headers, data = payload)

#         #print(response.text.encode('utf8'))
#         json_response = response.json()
#         #print(json_response)
#         token=json_response['access_token']
#         Authorization_replace="Bearer "+token

#         #url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$filter=Number eq '23-2115' & $select='UID'"
#         #url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$filter=Number eq '23-7280'"

#         #url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$filter=startswith(Number, '23-') eq true & $top=1"
#         url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number asc"

#         payload = {}
#         headers = {
#           'x-myobapi-key': 'tdd4dk347sw7nvdnr8tvbwta',
#           'x-myobapi-version': 'v2',
#           'Accept-Encoding': 'gzip,deflate',
#           'Authorization': Authorization_replace
#         }

#         response = requests.request("GET", url, headers=headers, data = payload)
#         json_response = response.json()
#         #print(json_response)
#         item_UID=json_response['Items'][0]['Number']
#         #print(item_UID)
#         item_UID=json_response['Items']
#         #print(item_UID)
#         item_df_aes = pd.DataFrame(columns=['Item', 'UID'])
#         for rows in item_UID:
#             items_uid=rows['UID']
#             items_number=rows['Number']
#             item_df_aes = item_df_aes.append({'Item': items_number, 'UID': items_uid}, ignore_index=True)
#         item_df_aes.head()
#         item_df_aes.shape
#         #print(response.text.encode('utf8'))

#         #Get all item UID in descending order

#         url = "https://secure.myob.com/oauth2/v1/authorize/"

#         payload = 'client_id=tdd4dk347sw7nvdnr8tvbwta&client_secret=ZaCSqqus8DrEfC3fr3VvWr8C&grant_type=refresh_token&refresh_token=Pkv5%21IAAAAIKM9DYl342RjLtg3HZcfJn2zxEwrLBSUcjcm0ZRC98HsQAAAAHj3FUWHNzI_wS0ihflLHgVuDVqQNg7GNf4_TY3xUYRNN99_92IDs-h9IvwWiGO_lFPjc9HvWYlOUZtoiid1XC8u4VoZa3AruyasYVCVAhSh_qOi9qZ6zSQvPfqsOFEnFxPC5bQzHcMk7k1Ote7wxPzv8VxGKHHEw0C4sX8a0hnnQS7LGVgfTUvLIejFBQAjRIoD5sz_0lVZJoGnN3PPvU3rnVMICP3x-w7rQpXzoDi3w'
#         headers = {
#           'Content-Type': 'application/x-www-form-urlencoded'
#         }

#         response = requests.request("POST", url, headers=headers, data = payload)

#         #print(response.text.encode('utf8'))
#         json_response = response.json()
#         #print(json_response)
#         token=json_response['access_token']
#         Authorization_replace="Bearer "+token

#         #url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$filter=Number eq '23-2115' & $select='UID'"
#         #url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$filter=Number eq '23-7280'"

#         #url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$filter=startswith(Number, '23-') eq true & $top=1"
#         url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number desc"

#         payload = {}
#         headers = {
#           'x-myobapi-key': 'tdd4dk347sw7nvdnr8tvbwta',
#           'x-myobapi-version': 'v2',
#           'Accept-Encoding': 'gzip,deflate',
#           'Authorization': Authorization_replace
#         }

#         response = requests.request("GET", url, headers=headers, data = payload)
#         json_response = response.json()
#         #print(json_response)
#         item_UID=json_response['Items'][0]['Number']
#         #print(item_UID)
#         item_UID=json_response['Items']
#         #print(item_UID)
#         item_df_desc = pd.DataFrame(columns=['Item', 'UID'])
#         for rows in item_UID:
#             items_uid=rows['UID']
#             items_number=rows['Number']
#             item_df_desc = item_df_desc.append({'Item': items_number, 'UID': items_uid}, ignore_index=True)
#         # item_df_desc.head()
#         item_df_desc.shape
#         #print(response.text.encode('utf8'))

#         item_df=pd.merge(item_df_aes, item_df_desc, on = ['Item','UID'], how = 'outer')
#         #print(item_df.shape)
#         print("comment this....")
#         print(item_df.head())


#         upload_data.head()
#         upload_data=pd.merge(upload_data, item_df, left_on = ['Number'],right_on=['Item'], how = 'left')
#         #print(item_df.shape)
        ###############################################################################
        ######NEW CODE BELOW#######
        
        url = "https://secure.myob.com/oauth2/v1/authorize/"

        payload = 'client_id=tdd4dk347sw7nvdnr8tvbwta&client_secret=ZaCSqqus8DrEfC3fr3VvWr8C&grant_type=refresh_token&refresh_token=Pkv5%21IAAAAIKM9DYl342RjLtg3HZcfJn2zxEwrLBSUcjcm0ZRC98HsQAAAAHj3FUWHNzI_wS0ihflLHgVuDVqQNg7GNf4_TY3xUYRNN99_92IDs-h9IvwWiGO_lFPjc9HvWYlOUZtoiid1XC8u4VoZa3AruyasYVCVAhSh_qOi9qZ6zSQvPfqsOFEnFxPC5bQzHcMk7k1Ote7wxPzv8VxGKHHEw0C4sX8a0hnnQS7LGVgfTUvLIejFBQAjRIoD5sz_0lVZJoGnN3PPvU3rnVMICP3x-w7rQpXzoDi3w'
        headers = {
          'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data = payload)

        #print(response.text.encode('utf8'))
        json_response = response.json()
        #print(json_response)
        token=json_response['access_token']
        Authorization_replace="Bearer "+token
        
        ###GET ALL ITEMS
        def hasCharacters(inputString):
            return bool(re.search(r'[a-zA-Z]', inputString))

        #Get UIDs of all items
        time.sleep(1)
        #Get all item UID in ascending order

        url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number asc"

        payload = {}
        headers = {
          'x-myobapi-key': 'dze2svn2nwsbrjzvmhwpsq57',
          'x-myobapi-version': 'v2',
          'Accept-Encoding': 'gzip,deflate',
          'Authorization': Authorization_replace
        }

        response = requests.request("GET", url, headers=headers, data = payload)
        #print(response)
        json_response = response.json()
        #print(json_response)
        item_UID=json_response['Items'][0]['Number']
        #print(item_UID)
        item_UID=json_response['Items']
        #print(item_UID)
        item_df_aes = pd.DataFrame(columns=['UID','Item','unitWeight','BinLoc','SOH','baleQTY','Name'])
        for rows in item_UID:
            items_uid=rows['UID']
            items_number=rows['Number']
            items_name=rows['Name']
            #print('items_number',items_number)
            #print('CustomList3',rows['CustomList3'])
            #print('CustomField3',rows['CustomField3'])
            if(rows['CustomList3']!=None):
                if(rows['CustomField3']!=None and hasCharacters(rows['CustomField3']['Value'])==False):
                    if(float(rows['CustomField3']['Value'])!=0):
                        items_weight=float(rows['CustomList3']['Value'])/float(rows['CustomField3']['Value'])
                    else:
                        items_weight=float(rows['CustomList3']['Value'])
                else:
                    items_weight=float(rows['CustomList3']['Value'])
            else:
                items_weight=""
            #print('CustomList1',rows['CustomList1'])
            if(rows['CustomField1']!=None):
                items_location=rows['CustomField1']['Value']
            else:
                items_location=""
            #items_location=rows['CustomList1']['Value']
            #QuantityAvailable has incorrect values when new stock is on water so turned off
            item_SOH=rows['QuantityOnHand']
            #item_SOH=rows['QuantityAvailable']
            if(rows['CustomField3']!=None):
                items_bale_qty=rows['CustomField3']['Value']
            else:
                items_bale_qty=""
            item_df_aes = item_df_aes.append({'UID': items_uid,'Item': items_number,'unitWeight':items_weight,'BinLoc':items_location,'SOH':item_SOH,'baleQTY':items_bale_qty,'Name':items_name}, ignore_index=True)
        item_df_aes.head()
        item_df_aes.shape
        #print(response.text.encode('utf8'))

        #Get all item UID in descending order
        time.sleep(1)
        url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number desc"

        payload = {}
        headers = {
          'x-myobapi-key': 'dze2svn2nwsbrjzvmhwpsq57',
          'x-myobapi-version': 'v2',
          'Accept-Encoding': 'gzip,deflate',
          'Authorization': Authorization_replace
        }

        response = requests.request("GET", url, headers=headers, data = payload)
        json_response = response.json()
        #print(json_response)
        item_UID=json_response['Items'][0]['Number']
        #print(item_UID)
        item_UID=json_response['Items']
        #print(item_UID)
        item_df_desc = pd.DataFrame(columns=['UID','Item','unitWeight','BinLoc','SOH','baleQTY','Name'])
        for rows in item_UID:
            items_uid=rows['UID']
            items_number=rows['Number']
            items_name=rows['Name']
            #print('items_number',items_number)
            #print('CustomList3',rows['CustomList3'])
            #print('CustomField3',rows['CustomField3'])
            if(rows['CustomList3']!=None):
                if(rows['CustomField3']!=None and hasCharacters(rows['CustomField3']['Value'])==False):
                    if(float(rows['CustomField3']['Value'])!=0):
                        items_weight=float(rows['CustomList3']['Value'])/float(rows['CustomField3']['Value'])
                    else:
                        items_weight=float(rows['CustomList3']['Value'])
                else:
                    items_weight=float(rows['CustomList3']['Value'])
            else:
                items_weight=""
            #print('CustomList1',rows['CustomList1'])
            if(rows['CustomField1']!=None):
                items_location=rows['CustomField1']['Value']
            else:
                items_location=""
            #items_location=rows['CustomList1']['Value']
            item_SOH=rows['QuantityOnHand']
            #item_SOH=rows['QuantityAvailable']
            if(rows['CustomField3']!=None):
                items_bale_qty=rows['CustomField3']['Value']
            else:
                items_bale_qty=""
            item_df_desc = item_df_desc.append({'UID': items_uid,'Item': items_number,'unitWeight':items_weight,'BinLoc':items_location,'SOH':item_SOH,'baleQTY':items_bale_qty,'Name':items_name}, ignore_index=True)
        # item_df_desc.head()
        item_df_desc.shape
        #print(response.text.encode('utf8'))

        item_df=pd.merge(item_df_aes, item_df_desc, on = ['UID','Item','unitWeight','BinLoc','SOH','baleQTY','Name'], how = 'outer')
        
        #print(item_df.head())
        
        #print(item_df.shape)
        #item_df.head()
        ######NEW CODE UP#######
     
        #print("comment this....")
        #print(item_df.head())

        upload_data.head()
        upload_data=pd.merge(upload_data, item_df, left_on = ['Number'],right_on=['Item'], how = 'left')
        #print(item_df.shape)
        
        
        #print(upload_data.head())
        
        upload_data['UnitPrice'] = upload_data['UnitPrice'].str.replace('$', '')
        upload_data['Total'] = upload_data['Total'].str.replace('$', '')
        # upload_data.head()

        #Create payload for api call
        timestr = time.strftime("%Y-%m-%dT%H:%M:%S")
        #print(timestr)
        order_timestr=timestr
        #Below lines fix the date issue
        #timestr = time.strftime("%Y-%m-%d")
        payload_start_before_time='{\n\t\"Number\": null,\n\t\"Date\": \"'
        payload_start_after_time='\",\n\t\"CustomerPurchaseOrderNumber\": null,\n\t\"Customer\": {\n\t\t\"UID\": \"ef7313ed-db54-481f-9b7b-94cf225b12c5\",\n\t\t\"Name\": \"TPW Group Services Pty Ltd\",\n\t\t\"DisplayID\": \"WAY1240\",\n\t},\n\t\"ShipToAddress\": \"TPW Group Services Pty Ltd\r\n1A/1-7 Unwins Bridge Rd\r\nSt.Peters  NSW  2044\r\nAustralia\r\n\",\n\t\"Terms\": {\n\t\t\"PaymentIsDue\": \"InAGivenNumberOfDays\",\n\t\t\"DiscountDate\": 0,\n\t\t\"BalanceDueDate\": 30,\n\t\t\"DiscountForEarlyPayment\": 0.0,\n\t\t\"MonthlyChargeForLatePayment\": 2.5,\n\t\t\"DiscountExpiryDate\": null,\n\t\t\"Discount\": 0,\n\t\t\"DueDate\": null,\n\t},\n\t\"IsTaxInclusive\": false,\n\t\"IsReportable\": false,\n\t\"Freight\": 0,\n\t\"FreightTaxCode\": null,\n\t\"Category\": null,\n\t\"Salesperson\": {\n\t\t\"UID\": \"aee62020-39bc-42fd-97c6-a0b9ef5f3ac1\"\n\t,\n\t\t\"Name\": \" NSW T\"\n\t,\n\t\t\"DisplayID\": \"T .5%\"\n\t},\n\t\"Comment\": null,\n\t\"ShippingMethod\": \"Australia Post\",\n\t\"PromisedDate\": null,\n\t\"JournalMemo\": \"Sale; TPW Group Services Pty Ltd\",\n\t\"InvoiceDeliveryStatus\": \"Nothing\",\n\t\"Status\": \"Open\",\n\t\"LastPaymentDate\": null,\n\t\"ForeignCurrency\": null,\n\t\"Lines\": ['
        payload_start=payload_start_before_time+timestr+payload_start_after_time
        
        #working code below with date
        #payload_start='{\n\t\"Number\": null,\n\t\"Date\": \"2021-01-26\",\n\t\"CustomerPurchaseOrderNumber\": null,\n\t\"Customer\": {\n\t\t\"UID\": \"ef7313ed-db54-481f-9b7b-94cf225b12c5\",\n\t\t\"Name\": \"TPW Group Services Pty Ltd\",\n\t\t\"DisplayID\": \"WAY1240\",\n\t},\n\t\"ShipToAddress\": \"TPW Group Services Pty Ltd\r\n1A/1-7 Unwins Bridge Rd\r\nSt.Peters  NSW  2044\r\nAustralia\r\n\",\n\t\"Terms\": {\n\t\t\"PaymentIsDue\": \"InAGivenNumberOfDays\",\n\t\t\"DiscountDate\": 0,\n\t\t\"BalanceDueDate\": 30,\n\t\t\"DiscountForEarlyPayment\": 0.0,\n\t\t\"MonthlyChargeForLatePayment\": 2.5,\n\t\t\"DiscountExpiryDate\": null,\n\t\t\"Discount\": 0,\n\t\t\"DueDate\": null,\n\t},\n\t\"IsTaxInclusive\": false,\n\t\"IsReportable\": false,\n\t\"Freight\": 0,\n\t\"FreightTaxCode\": null,\n\t\"Category\": null,\n\t\"Salesperson\": {\n\t\t\"UID\": \"aee62020-39bc-42fd-97c6-a0b9ef5f3ac1\"\n\t,\n\t\t\"Name\": \" NSW T\"\n\t,\n\t\t\"DisplayID\": \"T .5%\"\n\t},\n\t\"Comment\": null,\n\t\"ShippingMethod\": \"Australia Post\",\n\t\"PromisedDate\": null,\n\t\"JournalMemo\": \"Sale; TPW Group Services Pty Ltd\",\n\t\"InvoiceDeliveryStatus\": \"Nothing\",\n\t\"Status\": \"Open\",\n\t\"LastPaymentDate\": null,\n\t\"ForeignCurrency\": null,\n\t\"Lines\": ['
        
        #working code below without date
        #payload_start='{\n\t\"Number\": null,\n\t\"CustomerPurchaseOrderNumber\": null,\n\t\"Customer\": {\n\t\t\"UID\": \"ef7313ed-db54-481f-9b7b-94cf225b12c5\",\n\t\t\"Name\": \"TPW Group Services Pty Ltd\",\n\t\t\"DisplayID\": \"WAY1240\",\n\t},\n\t\"ShipToAddress\": \"TPW Group Services Pty Ltd\r\n1A/1-7 Unwins Bridge Rd\r\nSt.Peters  NSW  2044\r\nAustralia\r\n\",\n\t\"Terms\": {\n\t\t\"PaymentIsDue\": \"InAGivenNumberOfDays\",\n\t\t\"DiscountDate\": 0,\n\t\t\"BalanceDueDate\": 30,\n\t\t\"DiscountForEarlyPayment\": 0.0,\n\t\t\"MonthlyChargeForLatePayment\": 2.5,\n\t\t\"DiscountExpiryDate\": null,\n\t\t\"Discount\": 0,\n\t\t\"DueDate\": null,\n\t},\n\t\"IsTaxInclusive\": false,\n\t\"IsReportable\": false,\n\t\"Freight\": 0,\n\t\"FreightTaxCode\": null,\n\t\"Category\": null,\n\t\"Salesperson\": {\n\t\t\"UID\": \"aee62020-39bc-42fd-97c6-a0b9ef5f3ac1\"\n\t,\n\t\t\"Name\": \" NSW T\"\n\t,\n\t\t\"DisplayID\": \"T .5%\"\n\t},\n\t\"Comment\": null,\n\t\"ShippingMethod\": \"Australia Post\",\n\t\"PromisedDate\": null,\n\t\"JournalMemo\": \"Sale; TPW Group Services Pty Ltd\",\n\t\"InvoiceDeliveryStatus\": \"Nothing\",\n\t\"Status\": \"Open\",\n\t\"LastPaymentDate\": null,\n\t\"ForeignCurrency\": null,\n\t\"Lines\": ['

        #payload_start=payload_start_before_time+str(timestr)+payload_start_after_time
        payload_end=']\n}'
        every_line_start='{\n\t\t\"Type\": \"Transaction\",\n\t\t\"Job\": null,\n\t\t\"DiscountPercent\": 0,\n\t\t\"TaxCode\": {\n\t\t\t\"UID\": \"33c48787-7e0e-4102-af6c-94aeed662333\",\n\t\t\t\"Code\": \"GST\",\n\t\t},'
        description_start='\n\t\t\"Description\": \"'
        description_end='\",'

        ship_qty_start='\n\t\t\"ShipQuantity\": '

        total_start=',\n\t\t\"Total\": '

        unit_price_start=',\n\t\t\"UnitPrice\": '

        item_start=',\n\t\t\"Item\": {\n\t\t\t\"UID\": \"'
        item_end='\",\n\t\t}\n\t}'

        #Real
        payload=payload_start
        last_item=upload_data.shape[0]-1
        for index, row in upload_data.iterrows():
            if (index!=last_item):
                payload=payload+every_line_start+description_start+row['Description']+description_end+ship_qty_start+str(row['ShipQuantity'])+total_start+str(row['Total'])+unit_price_start+str(row['UnitPrice'])+item_start+row['UID']+item_end+','
            else:
                payload=payload+every_line_start+description_start+row['Description']+description_end+ship_qty_start+str(row['ShipQuantity'])+total_start+str(row['Total'])+unit_price_start+str(row['UnitPrice'])+item_start+row['UID']+item_end
        payload_final=payload+payload_end
        #print("Order posting disaabled!")
        #POST ORDER to MYOB


        #Connection
        url = "https://secure.myob.com/oauth2/v1/authorize/"

        payload = 'client_id=tdd4dk347sw7nvdnr8tvbwta&client_secret=ZaCSqqus8DrEfC3fr3VvWr8C&grant_type=refresh_token&refresh_token=Pkv5%21IAAAAIKM9DYl342RjLtg3HZcfJn2zxEwrLBSUcjcm0ZRC98HsQAAAAHj3FUWHNzI_wS0ihflLHgVuDVqQNg7GNf4_TY3xUYRNN99_92IDs-h9IvwWiGO_lFPjc9HvWYlOUZtoiid1XC8u4VoZa3AruyasYVCVAhSh_qOi9qZ6zSQvPfqsOFEnFxPC5bQzHcMk7k1Ote7wxPzv8VxGKHHEw0C4sX8a0hnnQS7LGVgfTUvLIejFBQAjRIoD5sz_0lVZJoGnN3PPvU3rnVMICP3x-w7rQpXzoDi3w'

        headers = {
          'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data = payload)

        #print(response.text.encode('utf8'))
        json_response = response.json()
        token=json_response['access_token']
        Authorization_replace="Bearer "+token

        #Post Order
        url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Sale/Order/Item"

        #payload = "{\n\t\"Number\": null,\n\t\"Date\": \"2020-10-01T00:00:00\",\n\t\"CustomerPurchaseOrderNumber\": null,\n\t\"Customer\": {\n\t\t\"UID\": \"ef7313ed-db54-481f-9b7b-94cf225b12c5\",\n\t\t\"Name\": \"TPW Group Services Pty Ltd\",\n\t\t\"DisplayID\": \"WAY1240\",\n\t},\n\t\"ShipToAddress\": \"TPW Group Services Pty Ltd\r\n1A/1-7 Unwins Bridge Rd\r\nSt.Peters  NSW  2044\r\nAustralia\r\n\",\n\t\"Terms\": {\n\t\t\"PaymentIsDue\": \"InAGivenNumberOfDays\",\n\t\t\"DiscountDate\": 0,\n\t\t\"BalanceDueDate\": 30,\n\t\t\"DiscountForEarlyPayment\": 0.0,\n\t\t\"MonthlyChargeForLatePayment\": 2.5,\n\t\t\"DiscountExpiryDate\": null,\n\t\t\"Discount\": 0,\n\t\t\"DueDate\": null,\n\t},\n\t\"IsTaxInclusive\": false,\n\t\"IsReportable\": false,\n\t\"Freight\": 0,\n\t\"FreightTaxCode\": null,\n\t\"Category\": null,\n\t\"Salesperson\": {\n\t\t\"UID\": \"aee62020-39bc-42fd-97c6-a0b9ef5f3ac1\"\n\t,\n\t\t\"Name\": \" NSW T\"\n\t,\n\t\t\"DisplayID\": \"T .5%\"\n\t},\n\t\"Comment\": null,\n\t\"ShippingMethod\": \"Australia Post\",\n\t\"PromisedDate\": null,\n\t\"JournalMemo\": \"Sale; TPW Group Services Pty Ltd\",\n\t\"InvoiceDeliveryStatus\": \"Nothing\",\n\t\"Status\": \"Open\",\n\t\"LastPaymentDate\": null,\n\t\"ForeignCurrency\": null,\n\t\"Lines\": [{\n\t\t\"Type\": \"Transaction\",\n\t\t\"Description\": \"TW40841410\",\n\t\t\"ShipQuantity\": 1,\n\t\t\"Total\": 19.00,\n\t\t\"UnitPrice\": 19.00,\n\t\t\"Job\": null,\n\t\t\"DiscountPercent\": 0,\n\t\t\"TaxCode\": {\n\t\t\t\"UID\": \"33c48787-7e0e-4102-af6c-94aeed662333\",\n\t\t\t\"Code\": \"GST\",\n\t\t},\n\t\t\"Item\": {\n\t\t\t\"UID\": \"aa6d6b53-0088-4295-94f1-69c7fc66077a\",\n\t\t\t\"Number\": \"23-8131\",\n\t\t}\n\t},{\n\t\t\"Type\": \"Transaction\",\n\t\t\"Description\": \"TW40841450\",\n\t\t\"ShipQuantity\": 1,\n\t\t\"Total\": 19.00,\n\t\t\"UnitPrice\": 19.00,\n\t\t\"Job\": null,\n\t\t\"DiscountPercent\": 0,\n\t\t\"TaxCode\": {\n\t\t\t\"UID\": \"33c48787-7e0e-4102-af6c-94aeed662333\",\n\t\t\t\"Code\": \"GST\",\n\t\t},\n\t\t\"Item\": {\n\t\t\t\"UID\": \"aa6d6b53-0088-4295-94f1-69c7fc66077a\",\n\t\t\t\"Number\": \"23-8131\",\n\t\t}\n\t}]\n}"
        payload=payload_final
        headers = {
          'x-myobapi-key': 'tdd4dk347sw7nvdnr8tvbwta',
          'x-myobapi-version': 'v2',
          'Accept-Encoding': 'gzip,deflate',
          'Content-Type': 'application/json',
          'Authorization': Authorization_replace
        }
        #print("######ORDER POSTING DEACTIVATED#####")
        
        response = requests.request("POST", url, headers=headers, data = payload)

        #print(response.text.encode('utf8'))
        #file = open("resp_text_log.txt", "w")
        #file.write(response.text)
        #file.close()

        ###################################################
        #print("UPLOAD FILE: "+ filename+ " GENERATED PLEASE CHECK!")
        print("ORDER POSTED TO MYOB PLEASE CHECK!")
        #     except:
        #             print("Failed to generate uploadfile")
        #             print("")
        #             return

        #         # while True:

        #         #     tpwConverter.converterProgram()
        #         #     #print("CONVERSION PROCESS COMPLETED PLEASE CHECK!")

        #         #     while True:
        #         #         answer = input('Run again? (y/n): ')
        #         #         if answer in ('y', 'n'):
        #         #             break
        #         #         print('Invalid input.')
        #         #     if answer == 'y':
        #         #         #cls()
        #         #         #print(chr(27) + "[2J")
        #         #         os.system("cls")
        #         #         reload(doormatsConverter3)
        #         #         continue
        #         #     else:
        #         #         print('Goodbye')
        #         #         break
        #     print("UPLOAD FILE GENERATED PLEASE CHECK!")

        #Step 2: Wait for user to download label PDF
        #Moved logic before generating pick slip
        #print("Please download labels and put in the Labels Folder")
        #input("Press Enter to continue...")

        #Step 3: "bin loc", "stock" and "item name" from stock file
        del data['index']
        data.head()

        #Step 4: Create PICK SLIP Here
        #Using item data received from MYOB up top stored in item_df tabel
        #Preparing stock data
        
        #####OLD CODE TO GET STOCK DATA#####################
#         list_of_stock_files = glob.glob('P:\Stocklist\Program\*') # * means all if need specific format then *.csv
#         latest_stock = max(list_of_stock_files, key=os.path.getctime)
#         print("Using Stock File: ",latest_stock)
#         bin_loc_data = pd.read_excel(latest_stock, sheet_name='Sheet1') #Fetching unit weights of bails from file
#         
#         print(bin_loc_data.head())
        #################################################
    
        bin_loc_data_subset=pd.DataFrame()
        bin_loc_data_subset['Item']=item_df["Item"]
        bin_loc_data_subset['Item Name']=item_df["Name"]
        bin_loc_data_subset['SOH']=item_df["SOH"]
        bin_loc_data_subset['Bin Loc']=item_df["BinLoc"]
        bin_loc_data_subset['Bale']=item_df["baleQTY"]
        #bin_loc_data_subset.head()

        data["Item"]=data.iloc[:,1]
        del data["Item#"]
        #Pickslip dataframe
        Bin_item = pd.merge(data[["Item","Quantity","PO #"]],bin_loc_data_subset[["Item","Item Name","Bin Loc","SOH","Bale"]], how='left',on=["Item"])

        Bin_item["Supplied"] = "."
        #Bin_item.head()

        Bin_item_sorted=Bin_item[["Bin Loc","Item","Quantity","Supplied","PO #","Item Name","SOH","Bale"]]
        Bin_item_sorted[['Quantity', 'Bale']] = Bin_item_sorted[['Quantity', 'Bale']].apply(pd.to_numeric) 
        Bin_item_sorted=Bin_item_sorted.sort_values(["Bin Loc", "PO #","Item"], ascending = (True, True,True))
        Bin_item_sorted.reset_index(inplace = True) 
        del Bin_item_sorted['index']

        #print(Bin_item_sorted.head())
        timestr = time.strftime("%Y.%m.%d-%H.%M.%S")
        filename_pick="Pick Slip/Pick_Slip_"+timestr+".xlsx"
        #Bin_item_sorted.to_excel(filename_pick,index=False,header = False)
        #Bin_item_sorted.head() 
        
        #Notify to check low stock items
        Bin_item_sorted_low_stock=Bin_item_sorted.groupby(["Bin Loc","Item","SOH","Bale"])['Quantity'].agg('sum').reset_index()
        filterLogic=Bin_item_sorted_low_stock['SOH']<11
        Bin_item_sorted_low_stock=Bin_item_sorted_low_stock[filterLogic]
        Bin_item_sorted_low_stock.columns = ["Bin Loc","Item","SOH","Bale","Quantity"]
        print("Checking for low stock items...")
        if(len(Bin_item_sorted_low_stock)>0):
            print("Please double check low stock items below!")
            print(Bin_item_sorted_low_stock.to_string(index=False))
            
        else:
            print("No low level\[SOH<11\] stock found.")

        
        #Wait for user input
        print("Please download labels and put in the Labels Folder")
        input("Press Enter to continue...")
        
        #############New Code Down to fetch invoice number
        
        

        today = date.today()

        # dd/mm/YY
        #today_date = today.strftime("%Y-%m-%d")    
        today_date=order_timestr
        url = "https://secure.myob.com/oauth2/v1/authorize/"

        payload = 'client_id=tdd4dk347sw7nvdnr8tvbwta&client_secret=ZaCSqqus8DrEfC3fr3VvWr8C&grant_type=refresh_token&refresh_token=Pkv5%21IAAAAIKM9DYl342RjLtg3HZcfJn2zxEwrLBSUcjcm0ZRC98HsQAAAAHj3FUWHNzI_wS0ihflLHgVuDVqQNg7GNf4_TY3xUYRNN99_92IDs-h9IvwWiGO_lFPjc9HvWYlOUZtoiid1XC8u4VoZa3AruyasYVCVAhSh_qOi9qZ6zSQvPfqsOFEnFxPC5bQzHcMk7k1Ote7wxPzv8VxGKHHEw0C4sX8a0hnnQS7LGVgfTUvLIejFBQAjRIoD5sz_0lVZJoGnN3PPvU3rnVMICP3x-w7rQpXzoDi3w'
        headers = {
          'Content-Type': 'application/x-www-form-urlencoded'
        }

        response = requests.request("POST", url, headers=headers, data = payload)

        #print(response.text.encode('utf8'))
        json_response = response.json()
        #print(json_response)
        token=json_response['access_token']
        Authorization_replace="Bearer "+token


        ##Check if order
        #url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Sale/Order/Item/?$filter=Number eq"+"'"+"0"+user_invoice+"'"
        #url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Sale/Order/Item/?$filter=Date eq datetime'2021-02-01'"
        #url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Sale/Order/Item/?$filter=Date eq datetime'"+ today_date +"' and Customer/Name eq 'TPW Group Services Pty Ltd'"
        url = "https://ar2.api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Sale/Order/Item/?$filter=Date eq datetime'"+ today_date +"'"
        #print(url)
        payload = {}
        headers = {
          'x-myobapi-key': 'dze2svn2nwsbrjzvmhwpsq57',
          'x-myobapi-version': 'v2',
          'Accept-Encoding': 'gzip,deflate',
          'Authorization': Authorization_replace
        }

        response = requests.request("GET", url, headers=headers, data = payload)
        #print(response)
        json_response_order = response.json()
        #print("json_response_order:")
        #print(json_response_order)
        
        if(len(json_response_order['Items'])>0):
            invoice_number=json_response_order['Items'][0]['Number']
            #As it is posted as an order
            print("Order Number:")
            print(invoice_number)

        #############New Code Up
        timestr_pickslip = time.strftime("%d/%m/%Y")
        pickslip_headder="TPW - INV#"+invoice_number+" - "+timestr_pickslip
        #Pick Slip generated
        #Changed the stock on hand to quantity available
        header1=[]
        header1.append([pickslip_headder,"","","","","","",""])
        df_header1 = pd.DataFrame(header1)
        header2=[]
        header2.append(["Bin Loc","Item","Quantity","Supplied","PO #","Item Name","SOH","Bale"])
        df_header2 = pd.DataFrame(header2)
        combine1=df_header1.append(df_header2, ignore_index = True)
        combine1.columns = ["Bin Loc","Item","Quantity","Supplied","PO #","Item Name","SOH","Bale"]
        
        Bin_item_sorted[["Quantity", "Bale"]] = Bin_item_sorted[["Quantity", "Bale"]].apply(pd.to_numeric)
        
        #aggregating duplicate rows to add quantity
        Bin_item_sorted=Bin_item_sorted.groupby(["Bin Loc","Item","Supplied","PO #","Item Name","SOH","Bale"], as_index=False)['Quantity'].sum()
        Bin_item_sorted=Bin_item_sorted[["Bin Loc","Item","Quantity","Supplied","PO #","Item Name","SOH","Bale"]]
        
        Bin_item_sorted_renamed=Bin_item_sorted
        Bin_item_sorted_renamed.columns = ["Bin Loc","Item","Quantity","Supplied","PO #","Item Name","SOH","Bale"]
        combine2=combine1.append(Bin_item_sorted_renamed, ignore_index = True)
        
        
        #combine2.to_excel(filename_pick,index=False,header=None)
        
        writer = pd.ExcelWriter(filename_pick, engine='xlsxwriter')
        combine2.to_excel(writer, sheet_name='Sheet1',index=False,header=None)
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Merge 1st row and format text
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size':18
            #,'fg_color': 'yellow'
        })
        worksheet.merge_range('A1:H1',pickslip_headder, merge_format)
        
        cell_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size':14})
        #worksheet.set_column('A2:H2', cell_format)
        worksheet.conditional_format( 'A2:H2' , { 'type' : 'no_blanks' , 'format' : cell_format} )
        
        new_format1 = workbook.add_format()
        new_format1.set_align('top')
        #new_format1.set_border(1)
        worksheet.set_column('A:A', 13, new_format1)
        worksheet.set_column('B:B', 9, new_format1)
        worksheet.set_column('C:C', 8, new_format1)
        worksheet.set_column('D:D', 8, new_format1)
        worksheet.set_column('E:E', 12, new_format1)
        worksheet.set_column('F:F', 33, new_format1)
        worksheet.set_column('G:G', 4, new_format1)
        worksheet.set_column('H:H', 4, new_format1)
        
        range_maker="A1:H"+str(combine2.shape[0])
        
        cell_format = workbook.add_format({'border': 1})
        #worksheet.set_column('A2:H2', cell_format)
        worksheet.conditional_format( range_maker , { 'type' : 'no_blanks' , 'format' : cell_format} )
        
        worksheet.print_area(range_maker)
        worksheet.center_horizontally()
        #worksheet.center_vertically()
        worksheet.set_margins(left=0.2, right=0.2, top=0.75, bottom=0.75)
        worksheet.set_print_scale(93)
        worksheet.set_paper(9)
        
        writer.save()
        
        print("Pick Slip Generated....in folder Pick Slip!")
        
        #Fixed to keep the PDF arrangement correct
        Bin_item_sorted["row-id"] = ""
        Bin_item_sorted["row-id"] = np.arange(1000, len(Bin_item_sorted) + 1000)

        #Step 5: Find Multi-Label items and replicate the rows
     
        #print("Pick slip")
        #print(Bin_item_sorted.head(60))
        
        #get rows which require replication
        multilabel_item_df=Bin_item_sorted.where(Bin_item_sorted['Quantity']>Bin_item_sorted['Bale']).dropna()
        #multilabel_item_df[["Quantity", "Bale"]] = multilabel_item_df[["Quantity", "Bale"]].apply(pd.to_numeric)
        
        #remove those rows from main df to avoid duplication
        Bin_item_sorted=pd.concat([Bin_item_sorted, multilabel_item_df]).drop_duplicates(keep=False)#keep=False removes any row that repeats
        #Bin_item_sorted=pd.concat([Bin_item_sorted, multilabel_item_df, multilabel_item_df]).drop_duplicates(keep=False)
        #Bin_item_sorted=pd.concat([Bin_item_sorted, multilabel_item_df]).drop_duplicates(keep="first")
        #print("After removing duplicates!")
        #Bin_item_sorted.head(15)
        #Logic to return number of labels for a given quantity in a order
        #Ex: -(-3//4)=1 and -(-12//4)=3
        multilabel_item_df['no of labels']=-(-multilabel_item_df['Quantity']//multilabel_item_df['Bale'])
        #df['$/hour'] = df['$']/df['hours']
        multilabel_item_df['no of labels']=multilabel_item_df['no of labels'].round(0)
        #multilabel_item_df.head()
        #Logic to replicate rows depending on the 'no of labels' 
        multilabel_item_df = multilabel_item_df.loc[multilabel_item_df.index.repeat(multilabel_item_df['no of labels'])]
        #print(multilabel_item_df.head(60))
        
        for i in range(1,len(multilabel_item_df)):
            #PDF_pages_Sorted_PO['Label_no'].iloc[i] = 1
            if multilabel_item_df['PO #'].iloc[i] == multilabel_item_df['PO #'].iloc[i-1]:
                if multilabel_item_df['Item'].iloc[i] == multilabel_item_df['Item'].iloc[i-1]:
                    multilabel_item_df['row-id'].iloc[i] = multilabel_item_df['row-id'].iloc[i-1]+0.1
        
        
        #print("multilabel_item_df")
        #print(multilabel_item_df.head(15))
        
        del multilabel_item_df["no of labels"]
        Bin_item_sorted=Bin_item_sorted.append(multilabel_item_df)
        
        
        #print(Bin_item_sorted.head(20))
        
        #Commented due to label issues
        #Bin_item_sorted=Bin_item_sorted.sort_values(["Bin Loc", "Item","PO #"], ascending = (True, True,True)).reset_index()
        Bin_item_sorted=Bin_item_sorted.sort_values(["PO #","Item"], ascending = (True, True)).reset_index()
        del Bin_item_sorted["index"]
        
        #print(PDF_pages_Sorted_PO.head(50))
        
        #Adding another column to see the index of label under a PO
        Bin_item_sorted["Label_no"] = 1
        for i in range(1,len(Bin_item_sorted)):
            #PDF_pages_Sorted_PO['Label_no'].iloc[i] = 1
            if Bin_item_sorted['PO #'].iloc[i] == Bin_item_sorted['PO #'].iloc[i-1]:
                #Logic fixed
                Bin_item_sorted['Label_no'].iloc[i] = Bin_item_sorted['Label_no'].iloc[i-1]+1
        
        

        
        

        
        #print("Bin_item_sorted")
        #print(Bin_item_sorted.head(50))
        
        #To match with pickslip
        #print(Bin_item_sorted[Bin_item_sorted["PO #"]=="TW41897386"])
        #Bin_item_sorted=Bin_item_sorted.sort_values(["PO #", "Item"], ascending = (True, True)).reset_index()
        #print(Bin_item_sorted[Bin_item_sorted["PO #"]=="TW41897386"])
        #Below 3 lines not in use
        #Bin_item_sorted_counts = Bin_item_sorted.pivot_table(index=["PO #","Item","Quantity"], aggfunc='size')
        #Bin_item_sorted_counts=Bin_item_sorted_counts.reset_index()
        #Bin_item_sorted_counts.columns = ["PO #","Item","Quantity","Labels_Expected"] 
        #Bin_item_sorted_counts=Bin_item_sorted['is_duplicated'].sum()
        #Bin_item_sorted_counts=Bin_item_sorted.duplicated(subset='PO #', keep='first').sum()
        #Bin_item_sorted_counts=Bin_item_sorted.groupby(Bin_item_sorted.columns.tolist(),as_index=False).size()
       
        #print("Bin_item_sorted_counts")
        #print(Bin_item_sorted_counts.columns)
        #print(Bin_item_sorted_counts.head(50))

        #Step 6: Get the correct sequence of PDF file
        #Pick slip sorted by PO
        #pick_slip_Sorted_PO=Bin_item_sorted.sort_values('PO #').reset_index()
        #del pick_slip_Sorted_PO["index"]
        #pick_slip_Sorted_PO.head()
        #pick_slip_Sorted_PO.head(60)

        #Fetch latest label file
        list_of_labels = glob.glob('Labels/*.pdf') # * means all if need specific format then *.csv
        latest_label = max(list_of_labels, key=os.path.getctime)
        print("Using Labels: ",latest_label)
        #Get order of POs
        with pdfplumber.open(latest_label) as pdf:
            #first_page = pdf.pages[1]
            first_page = pdf.pages
            #Update : imrpoved regex for new labels
            r_exp = re.compile("Order[\s]?[R]?[e]?[f]?\:\s([T][W]\d*)")
            newlist=list()
            for item in first_page:
                label_text=item.extract_text()
                #Update : imrpoved regex for new labels
                PO=re.findall(r'Order[\s]?[R]?[e]?[f]?\:\s([T][W]\d*)', label_text)[0]
                newlist.append(PO)
            #print(newlist)
        #print("Order of labels in the PDF:")
        #print(newlist)

        PO_page_no = pd.DataFrame(newlist)
        PO_page_no.head()
        PO_page_no['pg_no'] = PO_page_no.index + 1
        PO_page_no.columns = ['PO #', 'Pg_no'] 
        #PO_page_no.head(60)

        #PDF sorted by PO number
        PDF_pages_Sorted_PO=PO_page_no.sort_values('PO #').reset_index()
        del PDF_pages_Sorted_PO['index']
        
        
        #print(PDF_pages_Sorted_PO.head(50))
        
        #Add Label_no column to findout how many labels agaist 1 "PO #"
        PDF_pages_Sorted_PO["Label_no"] = 1
        for i in range(1,len(PDF_pages_Sorted_PO)):
            #PDF_pages_Sorted_PO['Label_no'].iloc[i] = 1
            if PDF_pages_Sorted_PO['PO #'].iloc[i] == PDF_pages_Sorted_PO['PO #'].iloc[i-1]:
                #Logic fixed
                PDF_pages_Sorted_PO['Label_no'].iloc[i] = PDF_pages_Sorted_PO['Label_no'].iloc[i-1]+1
                
        
        #print(PDF_pages_Sorted_PO.head(50))
            
        #INCORRECT LOGIC NEED TO REPLACE
        # Bin_item_sorted['PO #'].head()
        #pick_slip_Sorted_PO['PDF_position']=PDF_pages_Sorted_PO['Pg_no']
        #pick_slip_Sorted_PO['PDF_PO']=PDF_pages_Sorted_PO['PO #']
        #print(pick_slip_Sorted_PO.head(60))
        
        #Correct logic
        #Validate if the labels match
        #Changed the join from left to inner to only get common rows. To avoid incorrect water mark issue.
        arrange_pdf_df=pd.merge(Bin_item_sorted, PDF_pages_Sorted_PO, on = ['PO #','Label_no'], how = 'inner')
        #added to match the old code
        arrange_pdf_df['PDF_position']=arrange_pdf_df['Pg_no']
        del arrange_pdf_df['Pg_no']
        #print("Bin_item_sorted:")
        #print(Bin_item_sorted.head(20))
        #print("arrange_pdf_df:")
        #print(arrange_pdf_df.head(20))
        

        #pick_slip_Sorted_PO['PO #'].equals(pick_slip_Sorted_PO['PDF_PO']    

        #PDF arrangements by bin loc
        pick_slip_Sorted_Bin_Loc=arrange_pdf_df.sort_values(["PO #", "Item","Label_no"], ascending = (True, True, True)).reset_index()
        del pick_slip_Sorted_Bin_Loc['index']
        #pick_slip_Sorted_Bin_Loc.head(60)
        
        #Adding column to identify the label of Item under a PO
        pick_slip_Sorted_Bin_Loc["Label_Item_no"] = 1
        for i in range(1,len(pick_slip_Sorted_Bin_Loc)):
            #PDF_pages_Sorted_PO['Label_no'].iloc[i] = 1
            if (pick_slip_Sorted_Bin_Loc['PO #'].iloc[i] == pick_slip_Sorted_Bin_Loc['PO #'].iloc[i-1]):
                if pick_slip_Sorted_Bin_Loc['Item'].iloc[i] == pick_slip_Sorted_Bin_Loc['Item'].iloc[i-1]:
                    #Logic fixed
                    pick_slip_Sorted_Bin_Loc['Label_Item_no'].iloc[i] = pick_slip_Sorted_Bin_Loc['Label_Item_no'].iloc[i-1]+1

        
        pick_slip_Sorted_Bin_Loc["Label_Quantity"] = pick_slip_Sorted_Bin_Loc["Quantity"]
        for i in range(len(pick_slip_Sorted_Bin_Loc)):
            #PDF_pages_Sorted_PO['Label_no'].iloc[i] = 1
            if (i!=len(pick_slip_Sorted_Bin_Loc)-1):
                #print("len(pick_slip_Sorted_Bin_Loc):",len(pick_slip_Sorted_Bin_Loc))
                #print("pick_slip_Sorted_Bin_Loc['PO #'].iloc[i]:",pick_slip_Sorted_Bin_Loc['PO #'].iloc[i])
                #print("pick_slip_Sorted_Bin_Loc['PO #'].iloc[i+1]:",pick_slip_Sorted_Bin_Loc['PO #'].iloc[i+1])
                if (pick_slip_Sorted_Bin_Loc['PO #'].iloc[i] == pick_slip_Sorted_Bin_Loc['PO #'].iloc[i+1]):
                    if pick_slip_Sorted_Bin_Loc['Item'].iloc[i] == pick_slip_Sorted_Bin_Loc['Item'].iloc[i+1]:
                        pick_slip_Sorted_Bin_Loc['Label_Quantity'].iloc[i]=pick_slip_Sorted_Bin_Loc['Bale'].iloc[i]
                else:
                    if (pick_slip_Sorted_Bin_Loc['PO #'].iloc[i] == pick_slip_Sorted_Bin_Loc['PO #'].iloc[i-1]):
                        if pick_slip_Sorted_Bin_Loc['Item'].iloc[i] == pick_slip_Sorted_Bin_Loc['Item'].iloc[i-1]:
                            remaining=(pick_slip_Sorted_Bin_Loc['Bale'].iloc[i]*pick_slip_Sorted_Bin_Loc['Label_Item_no'].iloc[i])-pick_slip_Sorted_Bin_Loc['Quantity'].iloc[i]
                            pick_slip_Sorted_Bin_Loc['Label_Quantity'].iloc[i]=pick_slip_Sorted_Bin_Loc['Bale'].iloc[i]-remaining
                 
            else:
                if (pick_slip_Sorted_Bin_Loc['PO #'].iloc[i] == pick_slip_Sorted_Bin_Loc['PO #'].iloc[i-1]):
                    if pick_slip_Sorted_Bin_Loc['Item'].iloc[i] == pick_slip_Sorted_Bin_Loc['Item'].iloc[i-1]:
                        remaining=(pick_slip_Sorted_Bin_Loc['Bale'].iloc[i]*pick_slip_Sorted_Bin_Loc['Label_Item_no'].iloc[i])-pick_slip_Sorted_Bin_Loc['Quantity'].iloc[i]
                        pick_slip_Sorted_Bin_Loc['Label_Quantity'].iloc[i]=pick_slip_Sorted_Bin_Loc['Bale'].iloc[i]-remaining
        
        pick_slip_Sorted_Bin_Loc["Label_mark"] = pick_slip_Sorted_Bin_Loc["Item"] +" X "+ pick_slip_Sorted_Bin_Loc["Label_Quantity"].astype(str)
        #pick_slip_Sorted_Bin_Loc["Label_mark"].head()
        #print("pick_slip_Sorted_Bin_Loc")
        #print(pick_slip_Sorted_Bin_Loc.head(20))
        #Step 7: Sequence PDF
        pdf_document=latest_label
        output_filename="Program-Files/Sequenced PDF/sequence_generated_"+timestr+".pdf"
        #output_filename = "sequence.pdf"
        pdf = PdfFileReader(pdf_document)
        pdf_writer = PdfFileWriter()
        
        #Fixed to keep the PDF arrangement correct
        #pick_slip_Sorted_Bin_Loc=pick_slip_Sorted_Bin_Loc.sort_values(["Bin Loc","PO #", "Item"], ascending = (True, True, True)).reset_index()
        pick_slip_Sorted_Bin_Loc=pick_slip_Sorted_Bin_Loc.sort_values(["row-id"], ascending = (True)).reset_index()
        for desired_seq_no in pick_slip_Sorted_Bin_Loc['PDF_position']:
            for page in range(pdf.getNumPages()):
                current_page = pdf.getPage(page)
                page=page+1
                if page == desired_seq_no:
                    pdf_writer.addPage(current_page)
                    #print(page)
                #else:
                    #pdf_writer_even.addPage(current_page)

        with open(output_filename, "wb") as out:
             pdf_writer.write(out)
             print("created: ", output_filename)

        #Step 8: Watermark PDF
        #from PyPDF2 import PdfFileWriter, PdfFileReader
        #import io
        #from reportlab.pdfgen import canvas
        #from reportlab.lib.pagesizes import letter
        output = PdfFileWriter()

        counter=0
        for desired_watermark in pick_slip_Sorted_Bin_Loc["Label_mark"]:
            packet = io.BytesIO()
            # create a new PDF with Reportlab
            can = canvas.Canvas(packet, pagesize=letter)
            #can.setFont('Vera', 12)
            can.drawString(120, 20, desired_watermark)
            can.save()

            #move to the beginning of the StringIO buffer
            packet.seek(0)
            new_pdf = PdfFileReader(packet)
            # read your existing PDF
            existing_pdf = PdfFileReader(open(output_filename, "rb"))

            # add the "watermark" (which is the new pdf) on the existing page
            page = existing_pdf.getPage(counter)
            counter=counter+1
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)
        # # finally, write "output" to a real file
        # outputStream = open("destination.pdf", "wb")
        # output.write(outputStream)
        # outputStream.close()
        output_filename="Water Marked PDF/watermarked_"+timestr+".pdf"
        #output_filename = "destination.pdf"
        with open(output_filename, "wb") as out:
             output.write(out)
             print("created: ", output_filename)
            
        ##Find problematic POS and labels
    
        Bin_item_sorted["PO_Label"] = Bin_item_sorted['PO #'].astype(str)+"_" + Bin_item_sorted['Label_no'].astype(str)
        #print("Bin_item_sorted:")
        #print(Bin_item_sorted.head(40))
        PDF_pages_Sorted_PO["PO_Label"] = PDF_pages_Sorted_PO['PO #'].astype(str)+"_" + PDF_pages_Sorted_PO['Label_no'].astype(str)
        #print("PDF_pages_Sorted_PO:")
        #print(PDF_pages_Sorted_PO.head(40))
        # Identify what values are in TableB and not in TableA
        key_diff = set(Bin_item_sorted["PO_Label"]).difference(PDF_pages_Sorted_PO["PO_Label"])
        
        if len(key_diff)>0:
            print("WARNING!!!")
            #print("Extra labels found:")
            print("Labels missing for:")
            print(key_diff)
        
        #where_diff = Bin_item_sorted.PO_Label.isin(key_diff)
        #print(where_diff)
        
        
        key_diff = set(PDF_pages_Sorted_PO["PO_Label"]).difference(Bin_item_sorted["PO_Label"])
        
        if len(key_diff)>0:
            print("WARNING!!!")
            #print("Labels missing for:")
            print("Extra labels found:")
            print(key_diff)
        #where_diff = PDF_pages_Sorted_PO.PO_Label.isin(key_diff)
        #print(where_diff)
        
        print("Preparing STOCK-UPDATE for all E-Tailers.....")
        #####Fetching the stock after processing order
        ###Getting item details from MYOB
        #Get UIDs of all items
        time.sleep(3)
        #Get all item UID in ascending order

        url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number asc"

        payload = {}
        headers = {
          'x-myobapi-key': 'dze2svn2nwsbrjzvmhwpsq57',
          'x-myobapi-version': 'v2',
          'Accept-Encoding': 'gzip,deflate',
          'Authorization': Authorization_replace
        }

        response = requests.request("GET", url, headers=headers, data = payload)
        #print(response)
        json_response = response.json()
        #print(json_response)
        item_UID=json_response['Items'][0]['Number']
        #print(item_UID)
        item_UID=json_response['Items']
        #print(item_UID)
        item_df_aes = pd.DataFrame(columns=['UID','Item','unitWeight','BinLoc','SOH','baleQTY','Name'])
        for rows in item_UID:
            items_uid=rows['UID']
            items_number=rows['Number']
            items_name=rows['Name']
            #print('items_number',items_number)
            #print('CustomList3',rows['CustomList3'])
            #print('CustomField3',rows['CustomField3'])
            if(rows['CustomList3']!=None):
                if(rows['CustomField3']!=None and hasCharacters(rows['CustomField3']['Value'])==False):
                    if(float(rows['CustomField3']['Value'])!=0):
                        items_weight=float(rows['CustomList3']['Value'])/float(rows['CustomField3']['Value'])
                    else:
                        items_weight=float(rows['CustomList3']['Value'])
                else:
                    items_weight=float(rows['CustomList3']['Value'])
            else:
                items_weight=""
            #print('CustomList1',rows['CustomList1'])
            if(rows['CustomField1']!=None):
                items_location=rows['CustomField1']['Value']
            else:
                items_location=""
            #items_location=rows['CustomList1']['Value']
            #item_SOH=rows['QuantityAvailable']
            item_SOH=rows['QuantityOnHand']
            if(rows['CustomField3']!=None):
                items_bale_qty=rows['CustomField3']['Value']
            else:
                items_bale_qty=""
            item_df_aes = item_df_aes.append({'UID': items_uid,'Item': items_number,'unitWeight':items_weight,'BinLoc':items_location,'SOH':item_SOH,'baleQTY':items_bale_qty,'Name':items_name}, ignore_index=True)
        item_df_aes.head()
        item_df_aes.shape
        #print(response.text.encode('utf8'))

        #Get all item UID in descending order
        time.sleep(1)
        url = "https://api.myob.com/accountright/48d8b367-05e9-4bcb-86d4-b7fc464e944e/Inventory/Item/?$top=1000&$filter=startswith(Number, '23-') eq true&$orderby=Number desc"

        payload = {}
        headers = {
          'x-myobapi-key': 'dze2svn2nwsbrjzvmhwpsq57',
          'x-myobapi-version': 'v2',
          'Accept-Encoding': 'gzip,deflate',
          'Authorization': Authorization_replace
        }

        response = requests.request("GET", url, headers=headers, data = payload)
        json_response = response.json()
        #print(json_response)
        item_UID=json_response['Items'][0]['Number']
        #print(item_UID)
        item_UID=json_response['Items']
        #print(item_UID)
        item_df_desc = pd.DataFrame(columns=['UID','Item','unitWeight','BinLoc','SOH','baleQTY','Name'])
        for rows in item_UID:
            items_uid=rows['UID']
            items_number=rows['Number']
            items_name=rows['Name']
            #print('items_number',items_number)
            #print('CustomList3',rows['CustomList3'])
            #print('CustomField3',rows['CustomField3'])
            if(rows['CustomList3']!=None):
                if(rows['CustomField3']!=None and hasCharacters(rows['CustomField3']['Value'])==False):
                    if(float(rows['CustomField3']['Value'])!=0):
                        items_weight=float(rows['CustomList3']['Value'])/float(rows['CustomField3']['Value'])
                    else:
                        items_weight=float(rows['CustomList3']['Value'])
                else:
                    items_weight=float(rows['CustomList3']['Value'])
            else:
                items_weight=""
            #print('CustomList1',rows['CustomList1'])
            if(rows['CustomField1']!=None):
                items_location=rows['CustomField1']['Value']
            else:
                items_location=""
            #items_location=rows['CustomList1']['Value']
            item_SOH=rows['QuantityOnHand']
            #item_SOH=rows['QuantityAvailable']
            if(rows['CustomField3']!=None):
                items_bale_qty=rows['CustomField3']['Value']
            else:
                items_bale_qty=""
            item_df_desc = item_df_desc.append({'UID': items_uid,'Item': items_number,'unitWeight':items_weight,'BinLoc':items_location,'SOH':item_SOH,'baleQTY':items_bale_qty,'Name':items_name}, ignore_index=True)
        # item_df_desc.head()
        item_df_desc.shape
        #print(response.text.encode('utf8'))
        ##This dataframe contains the latest data
        item_df=pd.merge(item_df_aes, item_df_desc, on = ['UID','Item','unitWeight','BinLoc','SOH','baleQTY','Name'], how = 'outer')
        #Matching with old dataframe
        quantity_df=item_df
        quantity_df.columns = ['UID','Item No.','unitWeight','BinLoc','Stock','baleQTY','Name']
        #Reading TPW format file
        list_of_files_tpw = glob.glob('Program-Files/TPW-StockUpdate-Format/*.csv') # * means all if need specific format then *.csv
        latest_file_TPW = max(list_of_files_tpw, key=os.path.getctime)
        #print("Using TPW File:   "+latest_file_TPW+"    PLEASE CHECK!!")
        #csv_file=glob.glob('TPW/*.csv')
        #Save stock file for TPW
        outputPath="P:\Stock-eTailers/TPW/"
        #print(csv_file)
        # if len(csv_file) != 1:
        #     raise ValueError('should be only one txt file in the current directory')
        #TPW_filename = csv_file[0]
        #print('Using TPW template file:',latest_file_TPW)
        # with open(filename, 'r',) as file:
        #     reader = csv.reader(file)
        # #     for row in reader:
        # #         print(row)
        #     #print(reader)
        #     df = pd.DataFrame(reader)
        #     print(df.head(10))

        TPW_df=pd.read_csv(latest_file_TPW)
        
        TPW_stock_update = pd.merge(TPW_df,quantity_df[['Item No.','Stock']],left_on=['product_code'], right_on=['Item No.'], how='left')
        #print(TPW_stock_update.head(10))
        # if TPW_stock_update['Stock'].loc[TPW_stock_update['Stock'] == 'Value'] >= 20:
        #     TPW_stock_update['quantity'] = 20
        # else:
        #     TPW_stock_update['quantity'] = 0
        # print(TPW_stock_update.head(10))
        showroom_adjustment=1
        TPW_stock_update['Stock'] = np.where(TPW_stock_update['Stock']>0, TPW_stock_update['Stock'] - showroom_adjustment , TPW_stock_update['Stock'])
        TPW_stock_update.loc[TPW_stock_update['Stock'] > 199, 'quantity'] = 100
        TPW_stock_update['quantity'] = np.where(TPW_stock_update['Stock'].between(100,199, inclusive = True), 50, TPW_stock_update['quantity'])
        TPW_stock_update['quantity'] = np.where(TPW_stock_update['Stock'].between(20,99, inclusive = True), 19, TPW_stock_update['quantity'])
        #TPW_stock_update.loc[TPW_stock_update['Stock'] >= 20, 'quantity'] = 20
        #TPW_stock_update.loc[TPW_stock_update['Stock'] < 20, 'quantity'] = 0
        TPW_stock_update['quantity'] = np.where(TPW_stock_update['Stock'].between(7, 10, inclusive = True), TPW_stock_update['Stock'], TPW_stock_update['quantity'])
        TPW_stock_update['quantity'] = np.where(TPW_stock_update['Stock'].between(1, 6, inclusive = True), 0, TPW_stock_update['quantity'])
        #TPW_stock_update['quantity'] = TPW_stock_update.Stock.apply(lambda x: 1 if x >= 1000 else 0)
        restriction = 5
        TPW_stock_update['quantity'] = np.where(TPW_stock_update['Stock'].between(10, 19, inclusive = True), TPW_stock_update['Stock'] - restriction , TPW_stock_update['quantity'])




        TPW_stock_update.loc[TPW_stock_update['Stock'].isna(), 'quantity'] = 0
        del TPW_stock_update['Stock']
        del TPW_stock_update['Item No.']
        #print(TPW_stock_update.head(10))
        # tpw_fname=(TPW_filename.split("\\")[1]).split(".")[0]
        # #print(tpw_fname)
        # #To CSV
        # tpw_outputfile="output/"+tpw_fname+".csv"

        timestr = time.strftime("%Y%m%d%H%M%S")
        #print(timestr)
        filename_tpw=outputPath+"TPW_stock_update_"+timestr+".csv"

        TPW_stock_update.to_csv(filename_tpw,index=False)
        print("TPW stock file created....PLEASE CHECK!")
        
        #ATTEMPTING FTP TO TPW
        try:
            print("Attempting FTP to TPW....")
            socket.getaddrinfo('localhost', 8080)

            list_of_StockFiles_tpw = glob.glob('P:/Stock-eTailers/TPW/*.csv') # * means all if need specific format then *.csv
            latest_file_TPW = max(list_of_StockFiles_tpw, key=os.path.getctime)
            File2Send=latest_file_TPW

            file_path = Path(File2Send)

            with FTP('116.213.9.118', "inventory","Invent0ry") as ftp, open(file_path, 'rb') as file:
                    ftp.storbinary(f'STOR {file_path.name}', file)
                    print("SENT STOCK-UPDATE TO TPW!")
            
        except:
            print("Failed to sent stock update to TPW!")
        
        ##Living style stock update
        list_of_files_living = glob.glob('Program-Files/LivingStyles-StockUpdate-Format/*.xlsx') # * means all if need specific format then *.csv
        latest_file_living = max(list_of_files_living, key=os.path.getctime)
        # living_file=glob.glob('Living_Styles/*.xlsx')
        #print(living_file)
        # if len(living_file) != 1:
        #     raise ValueError('should be only one xlsx file in the current directory')
        # living_filename = living_file[0]
        #print('Using Living Style template file:',latest_file_living)
        living_df = pd.read_excel(latest_file_living,sheet_name='Sheet1')
        #print(living_df.head(10))
        #display(living_df)
        #display(quantity_df)
        #print("living_df")
        #print(living_df)
        #print("quantity_df")
        #print(quantity_df)
        Living_stock_update = pd.merge(living_df,quantity_df[['Item No.','Stock']],left_on=['Item #'], right_on=['Item No.'], how='left')
        #print("Living_stock_update")
        #print(Living_stock_update)
        # Living_stock_update.loc[Living_stock_update['Stock'] < 20, 'stock'] = 0
        # Living_stock_update.loc[Living_stock_update['Stock'].isna(), 'stock'] = 0
        showroom_adjustment=1
        Living_stock_update['Stock'] = np.where(Living_stock_update['Stock']>0, Living_stock_update['Stock'] - showroom_adjustment , Living_stock_update['Stock'])
        
        Living_stock_update.loc[Living_stock_update['Stock'] > 199, 'stock'] = 100
        Living_stock_update['stock'] = np.where(Living_stock_update['Stock'].between(100, 199, inclusive = True), 50, Living_stock_update['stock'])
        Living_stock_update['stock'] = np.where(Living_stock_update['Stock'].between(20, 99, inclusive = True), 19, Living_stock_update['stock'])
        #Living_stock_update.loc[Living_stock_update['Stock'] >= 20, 'stock'] = 20
        #TPW_stock_update.loc[TPW_stock_update['Stock'] < 20, 'quantity'] = 0
        Living_stock_update['stock'] = np.where(Living_stock_update['Stock'].between(0, 19, inclusive = True), 0, Living_stock_update['stock'])
        #Living_stock_update['stock'] = np.where(Living_stock_update['Stock'].between(1, 6, inclusive = True), "Check Stock", Living_stock_update['stock'])
        #TPW_stock_update['quantity'] = TPW_stock_update.Stock.apply(lambda x: 1 if x >= 1000 else 0)
        #restriction = 5
        #Living_stock_update['stock'] = np.where(Living_stock_update['Stock'].between(10, 19, inclusive = True), Living_stock_update['Stock'] - restriction , Living_stock_update['stock'])

        del Living_stock_update['Stock']
        del Living_stock_update['Item No.']
        #living_fname=(TPW_filename.split("\\")[1])
        # living_fname=(living_filename.split("\\")[1]).split(".xlsx")[0]
        #To CSV
        # import os
        # funded.to_csv(os.path.join(path,r'green1.csv'))
        # living_outputfile="output/"+living_fname+".csv"
        # Living_stock_update.to_csv(living_outputfile,index=False)
        #To EXCELprint(Bin_item_sorted[Bin_item_sorted["PO #"]=="TW41897386"])
        # living_outputfile="output/"+living_fname+".xlsx"
        outputPath="P:\Stock-eTailers/LivingStyles/"
        #filename_living=outputPath+"ib_stock"+timestr+".xlsx"
        filename_living=outputPath+"ib_stock.xlsx"
        Living_stock_update.to_excel(filename_living,index=False,header = True)
        #print("Conversion Completed!")
        print("LivingStyles stock file created....PLEASE CHECK!")
        
        ###Attempting FTP to LivingStyles####
        try:
            print("Attempting FTP to LivingStyles.......")
            list_of_StockFiles_tpw = glob.glob('P:\Stock-eTailers/LivingStyles/*.xlsx') # * means all if need specific format then *.csv
            latest_file_TPW = max(list_of_StockFiles_tpw, key=os.path.getctime)
            File2Send=latest_file_TPW

            file_path = Path(File2Send)

            with FTP("ftp.livingstyles.com.au","ibaus","Zq5qXn66") as ftp, open(file_path, 'rb') as file:
                    ftp.storbinary(f'STOR {file_path.name}', file)
                    print("SENT STOCK-UPDATE TO Living Styles!")
        except:
            print("Faield to send stock update to LivingStyles!")
        ########################
             
        #Fantastic Furniture stock update
        ##reusing TPW code
        list_of_files_tpw = glob.glob('Program-Files/FantasticFurniture-StockUpdate-Format/*.csv') # * means all if need specific format then *.csv
        latest_file_TPW = max(list_of_files_tpw, key=os.path.getctime)
        
        outputPath="P:/Stock-eTailers/FantasticFurniture/"
        #print(csv_file)
        # if len(csv_file) != 1:
        #     raise ValueError('should be only one txt file in the current directory')
        #TPW_filename = csv_file[0]
        #print('Using Fantastic Furniture template file:',latest_file_TPW)
        # with open(filename, 'r',) as file:
        #     reader = csv.reader(file)
        # #     for row in reader:
        # #         print(row)
        #     #print(reader)
        #     df = pd.DataFrame(reader)
        #     print(df.head(10))

        TPW_df=pd.read_csv(latest_file_TPW)
        
        TPW_stock_update = pd.merge(TPW_df,quantity_df[['Item No.','Stock']],left_on=['sku'], right_on=['Item No.'], how='left')
        #print(TPW_stock_update.head(10))
        # if TPW_stock_update['Stock'].loc[TPW_stock_update['Stock'] == 'Value'] >= 20:
        #     TPW_stock_update['quantity'] = 20
        # else:
        #     TPW_stock_update['quantity'] = 0
        # print(TPW_stock_update.head(10))
        showroom_adjustment=1
        TPW_stock_update['Stock'] = np.where(TPW_stock_update['Stock']>0, TPW_stock_update['Stock'] - showroom_adjustment , TPW_stock_update['Stock'])
        TPW_stock_update.loc[TPW_stock_update['Stock'] > 199, 'quantity_available'] = 100
        TPW_stock_update['quantity_available'] = np.where(TPW_stock_update['Stock'].between(100,199, inclusive = True), 50, TPW_stock_update['quantity_available'])
        TPW_stock_update['quantity_available'] = np.where(TPW_stock_update['Stock'].between(20,99, inclusive = True), 19, TPW_stock_update['quantity_available'])
        #TPW_stock_update.loc[TPW_stock_update['Stock'] >= 20, 'quantity'] = 20
        #TPW_stock_update.loc[TPW_stock_update['Stock'] < 20, 'quantity'] = 0
        #TPW_stock_update['Stock On Hand'] = np.where(TPW_stock_update['Stock'].between(7, 10, inclusive = True), TPW_stock_update['Stock'], TPW_stock_update['Stock On Hand'])
        TPW_stock_update['quantity_available'] = np.where(TPW_stock_update['Stock'].between(0, 19, inclusive = True), 0, TPW_stock_update['quantity_available'])
        TPW_stock_update.loc[TPW_stock_update['quantity_available'] > 0, 'status'] = "in-stock"
        #TPW_stock_update['quantity'] = TPW_stock_update.Stock.apply(lambda x: 1 if x >= 1000 else 0)
        #restriction = 5
        #TPW_stock_update['Stock On Hand'] = np.where(TPW_stock_update['Stock'].between(10, 19, inclusive = True), TPW_stock_update['Stock'] - restriction , TPW_stock_update['Stock On Hand'])




        TPW_stock_update.loc[TPW_stock_update['Stock'].isna(), 'quantity_available'] = 0
        del TPW_stock_update['Stock']
        del TPW_stock_update['Item No.']
        #print(TPW_stock_update.head(10))
        # tpw_fname=(TPW_filename.split("\\")[1]).split(".")[0]
        # #print(tpw_fname)
        # #To CSV
        # tpw_outputfile="output/"+tpw_fname+".csv"

        timestr = time.strftime("%Y%m%d%H%M%S")
        #print(timestr)
        filename_tpw=outputPath+"Inventory_"+timestr+".csv"

        TPW_stock_update.to_csv(filename_tpw,index=False)
        print("Fantastic Furniture stock file created....PLEASE CHECK!")
        
        ###Attempt SFTP to FantasticFurniture###
        try:
            print("Attemptting SFTP to FantasticFurniture.....")
            myHostname = "sftp.dsco.io"
            myUsername = "account1000012469"
            myPassword = "513635"

            cnopts = pysftp.CnOpts()
            cnopts.hostkeys = None

            with pysftp.Connection(host=myHostname, username=myUsername, password=myPassword, port=22,cnopts=cnopts) as sftp:
                #print("Connection succesfully established ... ")

                # Switch to a remote directory
                #sftp.cwd('/var/www/vhosts/')

                # Obtain structure of the remote directory '/var/www/vhosts'
                #directory_structure = sftp.listdir_attr()

                # Print data
                #for attr in directory_structure:
                    #print(attr.filename, attr)

                list_of_StockFiles_tpw = glob.glob('P:/Stock-eTailers/FantasticFurniture/*.csv') # * means all if need specific format then *.csv
                latest_file_TPW = max(list_of_StockFiles_tpw, key=os.path.getctime)
                #File2Send=latest_file_TPW
                localFilePath = latest_file_TPW

                # Define the remote path where the file will be uploaded
                timestr = time.strftime("%Y%m%d%H%M%S")
                remoteFilePath = '/in/Inventory_'+timestr+'.csv'

                sftp.put(localFilePath, remoteFilePath)
                print("SENT STOCK-UPDATE TO Fantastic Furniture!")
                
        except:
            print("Failed to send stock update to Fantastic Furniture!")
        ###############################
        
        
        #Zanui stock update
        ##reusing TPW code
        list_of_files_tpw = glob.glob('Program-Files/Zanui-StockUpdate-Format/*.csv') # * means all if need specific format then *.csv
        latest_file_TPW = max(list_of_files_tpw, key=os.path.getctime)
        
        outputPath="P:/Stock-eTailers/Zanui/"
        #print(csv_file)
        # if len(csv_file) != 1:
        #     raise ValueError('should be only one txt file in the current directory')
        #TPW_filename = csv_file[0]
        #print('Using Zanui template file:',latest_file_TPW)
        # with open(filename, 'r',) as file:
        #     reader = csv.reader(file)
        # #     for row in reader:
        # #         print(row)
        #     #print(reader)
        #     df = pd.DataFrame(reader)
        #     print(df.head(10))

        TPW_df=pd.read_csv(latest_file_TPW)
        
        TPW_stock_update = pd.merge(TPW_df,quantity_df[['Item No.','Stock']],left_on=['Supplier SKU'], right_on=['Item No.'], how='left')
        #print(TPW_stock_update.head(10))
        # if TPW_stock_update['Stock'].loc[TPW_stock_update['Stock'] == 'Value'] >= 20:
        #     TPW_stock_update['quantity'] = 20
        # else:
        #     TPW_stock_update['quantity'] = 0
        # print(TPW_stock_update.head(10))
        showroom_adjustment=1
        TPW_stock_update['Stock'] = np.where(TPW_stock_update['Stock']>0, TPW_stock_update['Stock'] - showroom_adjustment , TPW_stock_update['Stock'])
        TPW_stock_update.loc[TPW_stock_update['Stock'] > 199, 'Stock On Hand'] = 100
        TPW_stock_update['Stock On Hand'] = np.where(TPW_stock_update['Stock'].between(100,199, inclusive = True), 50, TPW_stock_update['Stock On Hand'])
        TPW_stock_update['Stock On Hand'] = np.where(TPW_stock_update['Stock'].between(20,99, inclusive = True), 19, TPW_stock_update['Stock On Hand'])
        #TPW_stock_update.loc[TPW_stock_update['Stock'] >= 20, 'quantity'] = 20
        #TPW_stock_update.loc[TPW_stock_update['Stock'] < 20, 'quantity'] = 0
        #TPW_stock_update['Stock On Hand'] = np.where(TPW_stock_update['Stock'].between(7, 10, inclusive = True), TPW_stock_update['Stock'], TPW_stock_update['Stock On Hand'])
        TPW_stock_update['Stock On Hand'] = np.where(TPW_stock_update['Stock'].between(0, 19, inclusive = True), 0, TPW_stock_update['Stock On Hand'])
        #TPW_stock_update['quantity'] = TPW_stock_update.Stock.apply(lambda x: 1 if x >= 1000 else 0)
        #restriction = 5
        #TPW_stock_update['Stock On Hand'] = np.where(TPW_stock_update['Stock'].between(10, 19, inclusive = True), TPW_stock_update['Stock'] - restriction , TPW_stock_update['Stock On Hand'])




        TPW_stock_update.loc[TPW_stock_update['Stock'].isna(), 'Stock On Hand'] = 0
        del TPW_stock_update['Stock']
        del TPW_stock_update['Item No.']
        #print(TPW_stock_update.head(10))
        # tpw_fname=(TPW_filename.split("\\")[1]).split(".")[0]
        # #print(tpw_fname)
        # #To CSV
        # tpw_outputfile="output/"+tpw_fname+".csv"

        timestr = time.strftime("%Y%m%d%H%M%S")
        #print(timestr)
        filename_tpw=outputPath+"Zanui_stock_update_"+timestr+".csv"

        TPW_stock_update.to_csv(filename_tpw,index=False)
        print("Zanui stock file created....PLEASE CHECK!")
        print("Need to upload stock update to Zanui!")
        
        
#         #ATTEMPTING FTP TO Zanui
#         from ftplib import FTP
#         port_num=21
#         host_str="ftp-production.aws.zanui.com.au"
#         ftp=FTP(host_str)
#         ftp.set_debuglevel(1)
#         ftp.set_pasv(False)
#         ftp.connect(host_str,port_num)
#         ftp.getwelcome()
#         ftp.login("ibaustralia","R2t8y?dh")
#         list_of_StockFiles_tpw = glob.glob('P:\Stock-eTailers/Zanui/*.csv') # * means all if need specific format then *.csv
#         latest_file_TPW = max(list_of_StockFiles_tpw, key=os.path.getctime)
#         file_path = Path(latest_file_TPW)
#         ftp.storbinary(f'STOR {file_path.name}', file_path)
        
        
        
        
    except:
        print("Program ended abruptly!!!")
        print("")
        return
