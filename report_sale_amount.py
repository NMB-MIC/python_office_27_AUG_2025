import os
import datetime
import xlwings as xw
import pandas as pd

def extract_month(file_path):
        month_order = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
        'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
    }          
        month_name = file_path.split("\\")[-1].split(".")[0] 
        return month_order[month_name]
    
    
def report_daily(path):
    """create sale amount report by daily for 12 month"""
    xlsx_file_lists = []
    xlxs_file_list_sorts = []
    
    
    now = datetime.datetime.now()
    formatted_time = now.strftime("%Y_%m_%d_%H_%M_%S")
    
    for root,dirs,files in os.walk(path): 
        for name in files:
            if name.endswith(".xlsx"):
                file_path = os.path.join(root,name)
                xlsx_file_lists.append(file_path)

    sorted_file_paths = sorted(xlsx_file_lists, key=extract_month, reverse=True)

    for path in sorted_file_paths:
        xlxs_file_list_sorts.append(path)    
    template = xw.Book()
    app = xw.apps.active  

    sheet_1 = template.sheets["Sheet1"]

    for file in xlxs_file_list_sorts:
        df = pd.read_excel(file)
        pivot = pd.pivot_table(df,index="transaction_date",columns="store",values="amount",aggfunc="sum",margins=True,margins_name="Total")
        file_name = file.split("\\")[-1].split(".")[0]
        template.sheets.add(file_name)
        sheet = template.sheets[file_name]
        sheet["A1"].value = f'SALE AMOUNT REPORT BY DAILY AT {file_name.upper()}'
        sheet["A1"].api.Font.Bold = True
        sheet["A1"].font.size = 15
        sheet["A1"].font.name = "Arial"
        sheet["A1"].font.color = (0,0,255) 
        sheet["A3"].value = pivot

    sheet_1.delete()
    template.save(f"export\sale_amont_by_daily_12_month_{formatted_time}.xlsx")
    template.close() #close workbook
    app.kill() # close app

   
def report_daily_v2(path):
    """create sale amount report by daily for 12 month, add return file name"""
    xlsx_file_lists = []
    xlxs_file_list_sorts = []
    
    
    now = datetime.datetime.now()
    formatted_time = now.strftime("%Y_%m_%d_%H_%M_%S")
    
    for root,dirs,files in os.walk(path): 
        for name in files:
            if name.endswith(".xlsx"):
                file_path = os.path.join(root,name)
                xlsx_file_lists.append(file_path)

    sorted_file_paths = sorted(xlsx_file_lists, key=extract_month, reverse=True)

    for path in sorted_file_paths:
        xlxs_file_list_sorts.append(path)    
    template = xw.Book()
    app = xw.apps.active  

    sheet_1 = template.sheets["Sheet1"]

    for file in xlxs_file_list_sorts:
        df = pd.read_excel(file)
        pivot = pd.pivot_table(df,index="transaction_date",columns="store",values="amount",aggfunc="sum",margins=True,margins_name="Total")
        file_name = file.split("\\")[-1].split(".")[0]
        template.sheets.add(file_name)
        sheet = template.sheets[file_name]
        sheet["A1"].value = f'SALE AMOUNT REPORT BY DAILY AT {file_name.upper()}'
        sheet["A1"].api.Font.Bold = True
        sheet["A1"].font.size = 15
        sheet["A1"].font.name = "Arial"
        sheet["A1"].font.color = (0,0,255) 
        sheet["A3"].value = pivot

    sheet_1.delete()
    file_name = f"export\sale_amont_by_daily_12_month_{formatted_time}.xlsx"
    template.save(file_name)
    template.close() #close workbook
    app.kill() # close app
    return file_name