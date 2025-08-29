from report_sale_amount import report_daily_v2
from send_email import send_to_gmail

path = r"C:\Users\nmb_m\OneDrive\Desktop\python_office_2025-08-27\data\sales_data"
file_name = report_daily_v2(path)

username = 'suraphop.b@minebea.co.th'
password = 'ktyisawbsxusqtmj'
send_to_gmail(file_name,username,password)