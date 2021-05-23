#!/usr/bin/env python
__author__ = "Ashutosh Mishra"
__credits__ = ["Ashutosh Mishra"]
__code_name__ = "Outlook_Email_Data_Scraper"
__version__ = "1.0"
__maintainer__ = "Ashutosh Mishra"
__status__ = "Production"

import win32com.client
import os
import pandas as pd
from datetime import datetime as dt
import time
from datetime import date
dir_path = os.path.dirname(os.path.realpath(__file__))

output_dict = {'Subject': [],'Sender_Name': [],'Sender_Address': [],'Body': [],'Date': [],'Time': []}

class Email_fetcher:

    def main(self,dat):
        input_val = dat
        outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
        Inbox = outlook.GetDefaultFolder(6)
        message = Inbox.Items

        for data in message:
            temp_date = data.senton.date()
            if temp_date == str(input_val):

                temp_subject = data.Subject
                temp_sender_name = data.Sender()
                temp_sender_address = data.Sender.Address
                temp_body = data.body
                temp_time = data.senton.time()

                output_dict['Subject'].append(temp_subject)
                output_dict['Sender_Name'].append(temp_sender_name)
                output_dict['Sender_Address'].append(temp_sender_address)
                output_dict['Body'].append(temp_body)
                output_dict['Date'].append(temp_date)
                output_dict['Time'].append(temp_time)

    def write_output(self):
        df = pd.DataFrame(output_dict)
        now = dt.now().strftime("_%d_%b_%y_%I_%M_%p")
        writer = pd.ExcelWriter(dir_path + '\Output\Email_data_Output' + str(now) + '.xlsx')
        df.to_excel(writer, 'Output', index=False)
        writer.save()


if __name__ == "__main__":
    start_time = time.time()
    obj = Email_fetcher()
    val = input('Enter the Date in this format: 2020-08-19')
    while val == '':
        val = input('Enter the Date again in this format: 2020-08-19')
    obj.main(val)
    obj.write_output()
    print(f'Execution time: {(time.time() - start_time) / 60} mins')