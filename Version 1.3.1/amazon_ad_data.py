import pandas as pd
import numpy as np
import itertools as it
import tkinter as tk
import subprocess
import os
from tkinter import filedialog
from tkinter import *
from datetime import datetime
from simple_colors import *
import xlsxwriter
import csv


dt = datetime.now()


def main():
    def gui():
        global window
        window = tk.Tk()
        window.geometry("800x500")
        window.title("Search Term Tool")
        description0 = tk.Label(text="Brand Name")
        description1 = tk.Label(text="Branded Terms")
        global entry1, entry2, entry0
        entry0 = tk.Entry(window, width=100)
        entry1 = tk.Entry(window, width=100)
        button1 = tk.Button(window, text="Submit", command=get_branded_info)
        description2 = tk.Label(text="Branded ASINs")
        entry2 = tk.Entry(window, width= 100)
        def browseFiles():
            global st_filename
            st_filename = filedialog.askopenfilename(initialdir = "/",
                                                title = "Select a File",
                                                filetypes = (("CSV files",
                                                                "*.csv"),
                                                            ("all files",
                                                                "*.")
                                                            ))
                                                            
            
            # Change label contents
            label_file_explorer.configure(text="File Opened: "+st_filename)
        

        label_file_explorer = Label(window,
                            text = "File Explorer - Find Search Term Report",
                            width = 100, height = 4,
                            fg = "blue")

        button_explore = Button(window,
                        text = "Browse Files",
                        background = "blue",
                        command = browseFiles)



        label_file_explorer.pack()
        button_explore.pack(pady=50)
        description0.pack()
        entry0.pack()
        description1.pack()
        entry1.pack()
        description2.pack()
        entry2.pack()
        button1.pack(pady=25)
        
  
        # Function for opening the
        # file explorer window
        
  
        # button_exit = Button(window,
        #              text = "Exit",
        #              command = exit)

        
        
        # button_exit.pack()
        
        window.mainloop()


    def get_branded_info():

        global entry1, entry2, entry0, branded_keywords, branded_asins, brand_name
        global e1, e2, e0
        e0 = entry0.get()
        e1 = entry1.get()
        e2 = entry2.get()
        window.destroy()
        

       
        # print(branded_keywords)
        # print(branded_asins)
       

    gui()
    try:
        brand_name = e0
        branded_keywords = e1.lower()
        branded_asins = e2
        branded_keywords = branded_keywords.split(", ")
        branded_asins = branded_asins.split(", ")

        print(branded_keywords)
        print(branded_asins)
        calculate_overall_perfomance_metrics(st_filename, branded_keywords, branded_asins, brand_name)

    except (FileNotFoundError, PermissionError) as error:
        print(type(error).__name__, error, sep=": ")


def calculate_overall_perfomance_metrics(st_filename, branded_keywords, branded_asins, brand_name):
    with open(st_filename, newline='') as f:
        reader = csv.reader(f)
    df = pd.read_csv(st_filename)
    global ad_report_file_path
    st_report_dir = os.path.dirname(st_filename)
    ad_report_dir = st_report_dir
    ad_report_file_path = os.path.join(ad_report_dir, f"{brand_name}_ad_report_{dt}.csv")
    ad_report_file_path_no_ext = os.path.join(ad_report_dir, f"{brand_name}_ad_report_{dt}")
    df2 = df.replace(to_replace=np.nan, value=int(0))
    df2["Spend"].fillna(int(0))
    df2["7 Day Total Sales "].fillna(int(0))
    df2["Targeting"] = df2["Targeting"].astype("str")
    df2["7 Day Total Sales "] = df2["7 Day Total Sales "].fillna(0)
    df2["Spend"] = df2["Spend"].fillna(0)
    df3 = df2.fillna(0)
    df3['7 Day Total Sales '] = df['7 Day Total Sales '].replace(np.nan, 0)
    df3[[ 'Ad Type', 'Customer Search Term', 'Targeting', "Spend", "7 Day Total Sales "]] = df3[['Ad Type', 'Customer Search Term', 'Targeting',"Spend", "7 Day Total Sales "]].astype(str)
    df3['7 Day Total Sales '] = df3['7 Day Total Sales '].str.replace("\$", '')
    df3['Spend'] = df['Spend'].str.replace("\$", '')
    df3['7 Day Total Sales '] = df3['7 Day Total Sales '].str.replace("\,", '')
    df3['Spend'] = df3['Spend'].str.replace("\,", '')

    df3[["Spend", "7 Day Total Sales "]] = df3[["Spend", "7 Day Total Sales "]].apply(pd.to_numeric)
    df3[["Spend", "7 Day Total Sales "]] = df3[["Spend", "7 Day Total Sales "]].astype(float)

    sum_spend = df3["Spend"].sum()
    sum_sales = df3["7 Day Total Sales "].sum()
    # print()
    # print(f"Sum of Spend: ${sum_spend:,.2f}")
    # print(f"Sum of Sales: ${sum_sales:,.2f}")
    acos = sum_spend / sum_sales
    # print(f"ACoS: {acos:,.2f}%")
    
   
    
    br_kw_auto_spend_sum = 0
    br_kw_auto_sales_sum = 0
    nb_kw_auto_spend_sum = 0
    nb_kw_auto_sales_sum = 0
    br_pt_auto_spend_sum = 0 
    br_pt_auto_sales_sum = 0
    nb_pt_auto_spend_sum = 0
    nb_pt_auto_sales_sum = 0
    br_pt_manual_spend_sum = 0
    br_pt_manual_sales_sum = 0
    nb_pt_manual_spend_sum = 0
    nb_pt_manual_sales_sum = 0
    br_kw_manual_spend_sum = 0
    br_kw_manual_sales_sum = 0
    nb_kw_manual_spend_sum = 0
    nb_kw_manual_sales_sum = 0
    nb_pt_manual_spend_sum = 0
    nb_pt_manual_sales_sum = 0
    nb_pt_cat_manual_spend_sum = 0
    nb_pt_cat_manual_sales_sum = 0
    nb_broad_spend = 0
    nb_broad_sales = 0
    nb_phrase_spend = 0
    nb_phrase_sales = 0
    nb_exact_spend = 0
    nb_exact_sales = 0
    br_broad_spend = 0
    br_broad_sales = 0
    br_phrase_spend = 0
    br_phrase_sales = 0
    br_exact_spend = 0
    br_exact_sales = 0
    row_number = 1

    branded_keyword = ""
    branded_asin = ""
    nb_st_list = []
    br_st_list = []
    all_st_list = []
    all_st_kw_list = []
    all_target_list = []
    undefined_array = []

    def create_target_list():
        for target in df3["Targeting"]:
            as_str = str(target)
            all_target_list.append(as_str)
    def create_st_list():
        for st in df3["Customer Search Term"]:
            as_str = str(st)
            all_st_list.append(as_str)

    def create_kw_list():
        for target, st in zip(all_target_list, all_st_list):
            if "asin" not in target and "\*" not in target and "category" not in target:
                if "b0" not in st:
                    all_st_kw_list.append(st)
    

    def segment_st_to_br_or_nb(all_st_list, branded_keywords):
        for st in all_st_list:
            if not any(br_kw in st for br_kw in branded_keywords):
                nb_st_list.append(st)
            else:
                br_st_list.append(st)
                
    if len(branded_keywords) <= 1:
        branded_keyword = str(branded_keywords)
    if len(branded_asins) <= 1:
        branded_asin = str(branded_asins)

    str(branded_keyword)
    str(branded_asin)

    create_target_list()
    create_st_list()
    create_kw_list()
    segment_st_to_br_or_nb(all_st_list, branded_keywords)
    ad_type_str = "Ad Type"

    def group_by_ad_type():
        global sp_spend, sp_sales, sp_acos, sb_spend, sb_sales, sb_acos, sbv_spend, sbv_sales, sd_spend, sd_sales
        ad_type_spend = df3.groupby("Ad Type")["Spend"].sum()
        ad_type_sales = df3.groupby("Ad Type")["7 Day Total Sales "].sum()
        if any(df3['Ad Type'].str.contains('SP')):
            sp_spend = ad_type_spend["SP"]
            sp_sales = ad_type_sales["SP"]
        else:
            sp_spend = 0
            sp_sales = 0
        if any(df3['Ad Type'].str.contains('SB')):
            sb_spend = ad_type_spend["SB"]
            sb_sales = ad_type_sales["SB"]
        else:
            sb_spend = 0
            sb_sales = 0
        if any(df3['Ad Type'].str.contains('SBV')):
            sbv_spend = ad_type_spend["SBV"]
            sbv_sales = ad_type_sales["SBV"]
        else:
            sbv_spend = 0
            sbv_sales = 0
        if any(df3['Ad Type'].str.contains('SD')):
            sd_spend = ad_type_spend["SD"]
            sd_sales = ad_type_sales["SD"]
        else:
            sd_spend = 0
            sd_sales = 0
        
        

    column_names = list(df3.columns)
    if any(name.lower() == ad_type_str.lower() for name in column_names):
        group_by_ad_type()
    else:
        global sp_spend, sp_sales, sp_acos, sb_spend, sb_sales, sb_acos, sbv_spend, sbv_sales, sd_spend, sd_sales
        sp_spend = 0
        sp_sales = 0
        sb_spend = 0
        sb_sales = 0
        sbv_spend = 0
        sbv_sales = 0
        sd_spend = 0
        sd_sales = 0

    
   

  
    if len(branded_keywords) <= 1:
        for row in df3["Targeting"]:
            row.lower()
            

            current_row = df3.iloc[row_number - 1]
            current_row = current_row.replace(to_replace=np.nan, value=int(0))
            search_term = current_row["Customer Search Term"]
            if not isinstance(current_row, str):
                str(row)

            if "*" in row:
                if not isinstance(row, str):
                    row = str(row)

                if "asin" not in row:
                    if str(branded_keyword) in str(search_term):
                        current_row_br_kw_auto_spend = current_row["Spend"]
                        br_kw_auto_spend_sum = br_kw_auto_spend_sum + current_row_br_kw_auto_spend
                        current_row_br_kw_auto_sales = current_row["7 Day Total Sales "]
                        br_kw_auto_sales_sum = br_kw_auto_sales_sum + current_row_br_kw_auto_sales
                    else:
                        current_row_nb_kw_auto_spend = current_row["Spend"]
                        nb_kw_auto_spend_sum = nb_kw_auto_spend_sum + current_row_nb_kw_auto_spend
                        current_row_nb_kw_auto_sales = current_row["7 Day Total Sales "]
                        nb_kw_auto_sales_sum = nb_kw_auto_sales_sum + current_row_nb_kw_auto_sales 
                else:

                    if any(str(br_asin) in str(search_term) for br_asin in branded_asins):
                        current_row_br_pt_auto_spend = current_row["Spend"]
                        br_pt_auto_spend_sum = br_pt_auto_spend_sum + current_row_br_pt_auto_spend
                        current_row_br_pt_auto_sales = current_row["7 Day Total Sales "]
                        br_pt_auto_sales_sum = br_pt_auto_sales_sum + current_row_br_pt_auto_sales
                    else:
                        current_row_nb_pt_auto_spend = current_row["Spend"]
                        nb_pt_auto_spend_sum = nb_pt_auto_spend_sum + current_row_nb_pt_auto_spend
                        current_row_nb_pt_auto_sales = current_row["7 Day Total Sales "]
                        nb_pt_auto_sales_sum = nb_pt_auto_sales_sum + current_row_nb_pt_auto_sales 

            elif "asin" in row:
                if any(str(br_asin) in str(row) for br_asin in branded_asins):
                        current_row_br_pt_manual_spend = current_row["Spend"]
                        br_pt_manual_spend_sum = br_pt_manual_spend_sum + current_row_br_pt_manual_spend
                        current_row_br_pt_manual_sales = current_row["7 Day Total Sales "]
                        br_pt_manual_sales_sum = br_pt_manual_sales_sum + current_row_br_pt_manual_sales
                else:
                        current_row_nb_pt_manual_spend = current_row["Spend"]
                        nb_pt_manual_spend_sum = nb_pt_manual_spend_sum + current_row_nb_pt_manual_spend
                        current_row_nb_pt_manual_sales = current_row["7 Day Total Sales "]
                        nb_pt_manual_sales_sum = nb_pt_manual_sales_sum + current_row_nb_pt_manual_sales 
                
            elif "category=" in row:
                current_row_nb_pt_cat_manual_spend = current_row["Spend"]
                nb_pt_cat_manual_spend_sum = nb_pt_cat_manual_spend_sum + current_row_nb_pt_cat_manual_spend
                current_row_nb_pt_cat_manual_sales = current_row["7 Day Total Sales "]
                nb_pt_cat_manual_sales_sum = nb_pt_cat_manual_sales_sum + current_row_nb_pt_cat_manual_sales

            elif any(str(search_term) == br_st for br_st in br_st_list):
                current_row_br_kw_manual_spend = current_row["Spend"]
                br_kw_manual_spend_sum = br_kw_manual_spend_sum + current_row_br_kw_manual_spend
                current_row_br_kw_manual_sales = current_row["7 Day Total Sales "]
                br_kw_manual_sales_sum = br_kw_manual_sales_sum + current_row_br_kw_manual_sales
                match_type = current_row['Match Type']
                if match_type == "BROAD":
                    br_broad_spend = br_broad_spend + current_row_br_kw_manual_spend
                    br_broad_sales = br_broad_sales + current_row_br_kw_manual_sales
                elif match_type == "PHRASE":
                    br_phrase_spend = br_phrase_spend + current_row_br_kw_manual_spend
                    br_phrase_sales = br_phrase_sales + current_row_br_kw_manual_sales
                elif match_type == "EXACT":
                    br_exact_spend = br_exact_spend + current_row_br_kw_manual_spend
                    br_exact_sales = br_exact_sales + current_row_br_kw_manual_sales


            elif any(str(nb_kw) == str(search_term) for nb_kw in nb_st_list):
                current_row_nb_kw_manual_spend = current_row["Spend"]
                nb_kw_manual_spend_sum = nb_kw_manual_spend_sum + current_row_nb_kw_manual_spend
                current_row_nb_kw_manual_sales = current_row["7 Day Total Sales "]
                nb_kw_manual_sales_sum = nb_kw_manual_sales_sum + current_row_nb_kw_manual_sales
                match_type = current_row['Match Type']
                if match_type == "BROAD":
                    nb_broad_spend = nb_broad_spend + current_row_nb_kw_manual_spend
                    nb_broad_sales = nb_broad_sales + current_row_nb_kw_manual_sales
                elif match_type == "PHRASE":
                    nb_phrase_spend = nb_phrase_spend + current_row_nb_kw_manual_spend
                    nb_phrase_sales = nb_phrase_sales + current_row_nb_kw_manual_sales
                elif match_type == "EXACT":
                    nb_exact_spend = nb_exact_spend + current_row_nb_kw_manual_spend
                    nb_exact_sales = nb_exact_sales + current_row_nb_kw_manual_sales

            elif row == "" or row == 0:
                break

            else:
                undefined_array.append(row)
                
            
            row_number = row_number + 1
    elif len(branded_asins) <= 1:
        for row in df3["Targeting"]:
            row.lower()
            

            current_row = df3.iloc[row_number - 1]
            current_row = current_row.replace(to_replace=np.nan, value=int(0))
            search_term = current_row["Customer Search Term"]
            if not isinstance(current_row, str):
                str(row)

            if "*" in row:
                if not isinstance(row, str):
                    row = str(row)

                if "asin" not in row:
                    if any(br_kw in search_term for br_kw in branded_keywords):
                        current_row_br_kw_auto_spend = current_row["Spend"]
                        br_kw_auto_spend_sum = br_kw_auto_spend_sum + current_row_br_kw_auto_spend
                        current_row_br_kw_auto_sales = current_row["7 Day Total Sales "]
                        br_kw_auto_sales_sum = br_kw_auto_sales_sum + current_row_br_kw_auto_sales
                    else:
                        current_row_nb_kw_auto_spend = current_row["Spend"]
                        nb_kw_auto_spend_sum = nb_kw_auto_spend_sum + current_row_nb_kw_auto_spend
                        current_row_nb_kw_auto_sales = current_row["7 Day Total Sales "]
                        nb_kw_auto_sales_sum = nb_kw_auto_sales_sum + current_row_nb_kw_auto_sales 
                else:

                    if branded_asin in search_term:
                        current_row_br_pt_auto_spend = current_row["Spend"]
                        br_pt_auto_spend_sum = br_pt_auto_spend_sum + current_row_br_pt_auto_spend
                        current_row_br_pt_auto_sales = current_row["7 Day Total Sales "]
                        br_pt_auto_sales_sum = br_pt_auto_sales_sum + current_row_br_pt_auto_sales
                    else:
                        current_row_nb_pt_auto_spend = current_row["Spend"]
                        nb_pt_auto_spend_sum = nb_pt_auto_spend_sum + current_row_nb_pt_auto_spend
                        current_row_nb_pt_auto_sales = current_row["7 Day Total Sales "]
                        nb_pt_auto_sales_sum = nb_pt_auto_sales_sum + current_row_nb_pt_auto_sales 

            elif "asin" in row:
                if branded_asin in search_term:
                        current_row_br_pt_manual_spend = current_row["Spend"]
                        br_pt_manual_spend_sum = br_pt_manual_spend_sum + current_row_br_pt_manual_spend
                        current_row_br_pt_manual_sales = current_row["7 Day Total Sales "]
                        br_pt_manual_sales_sum = br_pt_manual_sales_sum + current_row_br_pt_manual_sales
                else:
                        current_row_nb_pt_manual_spend = current_row["Spend"]
                        nb_pt_manual_spend_sum = nb_pt_manual_spend_sum + current_row_nb_pt_manual_spend
                        current_row_nb_pt_manual_sales = current_row["7 Day Total Sales "]
                        nb_pt_manual_sales_sum = nb_pt_manual_sales_sum + current_row_nb_pt_manual_sales 
                
            elif "category=" in row:
                current_row_nb_pt_cat_manual_spend = current_row["Spend"]
                nb_pt_cat_manual_spend_sum = nb_pt_cat_manual_spend_sum + current_row_nb_pt_cat_manual_spend
                current_row_nb_pt_cat_manual_sales = current_row["7 Day Total Sales "]
                nb_pt_cat_manual_sales_sum = nb_pt_cat_manual_sales_sum + current_row_nb_pt_cat_manual_sales

            elif any(br_kw in search_term for br_kw in branded_keywords) in search_term:
                current_row_br_kw_manual_spend = current_row["Spend"]
                br_kw_manual_spend_sum = br_kw_manual_spend_sum + current_row_br_kw_manual_spend
                current_row_br_kw_manual_sales = current_row["7 Day Total Sales "]
                br_kw_manual_sales_sum = br_kw_manual_sales_sum + current_row_br_kw_manual_sales
                match_type = current_row['Match Type']
                if match_type == "BROAD":
                    br_broad_spend = br_broad_spend + current_row_br_kw_manual_spend
                    br_broad_sales = br_broad_sales + current_row_br_kw_manual_sales
                elif match_type == "PHRASE":
                    br_phrase_spend = br_phrase_spend + current_row_br_kw_manual_spend
                    br_phrase_sales = br_phrase_sales + current_row_br_kw_manual_sales
                elif match_type == "EXACT":
                    br_exact_spend = br_exact_spend + current_row_br_kw_manual_spend
                    br_exact_sales = br_exact_sales + current_row_br_kw_manual_sales
                


            elif any(nb_kw in search_term for nb_kw in nb_st_list):
                current_row_nb_kw_manual_spend = current_row["Spend"]
                nb_kw_manual_spend_sum = nb_kw_manual_spend_sum + current_row_nb_kw_manual_spend
                current_row_nb_kw_manual_sales = current_row["7 Day Total Sales "]
                nb_kw_manual_sales_sum = nb_kw_manual_sales_sum + current_row_nb_kw_manual_sales
                match_type = current_row['Match Type']
                if match_type == "BROAD":
                    nb_broad_spend = nb_broad_spend + current_row_nb_kw_manual_spend
                    nb_broad_sales = nb_broad_sales + current_row_nb_kw_manual_sales
                elif match_type == "PHRASE":
                    nb_phrase_spend = nb_phrase_spend + current_row_nb_kw_manual_spend
                    nb_phrase_sales = nb_phrase_sales + current_row_nb_kw_manual_sales
                elif match_type == "EXACT":
                    nb_exact_spend = nb_exact_spend + current_row_nb_kw_manual_spend
                    nb_exact_sales = nb_exact_sales + current_row_nb_kw_manual_sales

            elif row == "" or row == 0:
                break

            else:
                undefined_array.append(row)
                
            
            row_number = row_number + 1
    
    else: 
        for row in df3["Targeting"]:
            row.lower()
            

            current_row = df3.iloc[row_number - 1]
            current_row = current_row.replace(to_replace=np.nan, value=int(0))
            search_term = current_row["Customer Search Term"]
            if not isinstance(current_row, str):
                str(row)

            if "*" in row:
                if not isinstance(row, str):
                    row = str(row)

                if "asin" not in row:
                    if any(br_kw in search_term for br_kw in branded_keywords):
                        current_row_br_kw_auto_spend = current_row["Spend"]
                        br_kw_auto_spend_sum = br_kw_auto_spend_sum + current_row_br_kw_auto_spend
                        current_row_br_kw_auto_sales = current_row["7 Day Total Sales "]
                        br_kw_auto_sales_sum = br_kw_auto_sales_sum + current_row_br_kw_auto_sales
                    else:
                        current_row_nb_kw_auto_spend = current_row["Spend"]
                        nb_kw_auto_spend_sum = nb_kw_auto_spend_sum + current_row_nb_kw_auto_spend
                        current_row_nb_kw_auto_sales = current_row["7 Day Total Sales "]
                        nb_kw_auto_sales_sum = nb_kw_auto_sales_sum + current_row_nb_kw_auto_sales 
                else:

                    if any(br_asin in search_term for br_asin in branded_asins):
                        current_row_br_pt_auto_spend = current_row["Spend"]
                        br_pt_auto_spend_sum = br_pt_auto_spend_sum + current_row_br_pt_auto_spend
                        current_row_br_pt_auto_sales = current_row["7 Day Total Sales "]
                        br_pt_auto_sales_sum = br_pt_auto_sales_sum + current_row_br_pt_auto_sales
                    else:
                        current_row_nb_pt_auto_spend = current_row["Spend"]
                        nb_pt_auto_spend_sum = nb_pt_auto_spend_sum + current_row_nb_pt_auto_spend
                        current_row_nb_pt_auto_sales = current_row["7 Day Total Sales "]
                        nb_pt_auto_sales_sum = nb_pt_auto_sales_sum + current_row_nb_pt_auto_sales 

            elif "asin" in row:
                if any(str(br_asin) in row for br_asin in branded_asins):
                        current_row_br_pt_manual_spend = current_row["Spend"]
                        br_pt_manual_spend_sum = br_pt_manual_spend_sum + current_row_br_pt_manual_spend
                        current_row_br_pt_manual_sales = current_row["7 Day Total Sales "]
                        br_pt_manual_sales_sum = br_pt_manual_sales_sum + current_row_br_pt_manual_sales
                else:
                        current_row_nb_pt_manual_spend = current_row["Spend"]
                        nb_pt_manual_spend_sum = nb_pt_manual_spend_sum + current_row_nb_pt_manual_spend
                        current_row_nb_pt_manual_sales = current_row["7 Day Total Sales "]
                        nb_pt_manual_sales_sum = nb_pt_manual_sales_sum + current_row_nb_pt_manual_sales 
                
            elif "category=" in row:
                current_row_nb_pt_cat_manual_spend = current_row["Spend"]
                nb_pt_cat_manual_spend_sum = nb_pt_cat_manual_spend_sum + current_row_nb_pt_cat_manual_spend
                current_row_nb_pt_cat_manual_sales = current_row["7 Day Total Sales "]
                nb_pt_cat_manual_sales_sum = nb_pt_cat_manual_sales_sum + current_row_nb_pt_cat_manual_sales

            elif any(br_kw in search_term for br_kw in branded_keywords):
                current_row_br_kw_manual_spend = current_row["Spend"]
                br_kw_manual_spend_sum = br_kw_manual_spend_sum + current_row_br_kw_manual_spend
                current_row_br_kw_manual_sales = current_row["7 Day Total Sales "]
                br_kw_manual_sales_sum = br_kw_manual_sales_sum + current_row_br_kw_manual_sales
                match_type = current_row['Match Type']
                if match_type == "BROAD":
                    br_broad_spend = br_broad_spend + current_row_br_kw_manual_spend
                    br_broad_sales = br_broad_sales + current_row_br_kw_manual_sales
                elif match_type == "PHRASE":
                    br_phrase_spend = br_phrase_spend + current_row_br_kw_manual_spend
                    br_phrase_sales = br_phrase_sales + current_row_br_kw_manual_sales
                elif match_type == "EXACT":
                    br_exact_spend = br_exact_spend + current_row_br_kw_manual_spend
                    br_exact_sales = br_exact_sales + current_row_br_kw_manual_sales


            elif any(nb_kw in search_term for nb_kw in nb_st_list):
                current_row_nb_kw_manual_spend = current_row["Spend"]
                nb_kw_manual_spend_sum = nb_kw_manual_spend_sum + current_row_nb_kw_manual_spend
                current_row_nb_kw_manual_sales = current_row["7 Day Total Sales "]
                nb_kw_manual_sales_sum = nb_kw_manual_sales_sum + current_row_nb_kw_manual_sales 
                match_type = current_row['Match Type']
                if match_type == "BROAD":
                    nb_broad_spend = nb_broad_spend + current_row_nb_kw_manual_spend
                    nb_broad_sales = nb_broad_sales + current_row_nb_kw_manual_sales
                elif match_type == "PHRASE":
                    nb_phrase_spend = nb_phrase_spend + current_row_nb_kw_manual_spend
                    nb_phrase_sales = nb_phrase_sales + current_row_nb_kw_manual_sales
                elif match_type == "EXACT":
                    nb_exact_spend = nb_exact_spend + current_row_nb_kw_manual_spend
                    nb_exact_sales = nb_exact_sales + current_row_nb_kw_manual_sales

            elif row == "" or row == 0:
                break

            else:
                undefined_array.append(row)
                
            
            row_number = row_number + 1


## auto calculations ############################
    auto_spend_sum = br_kw_auto_spend_sum + br_pt_auto_spend_sum + nb_kw_auto_spend_sum + nb_pt_auto_spend_sum
    auto_sales_sum = br_kw_auto_sales_sum + br_pt_auto_sales_sum + nb_kw_auto_sales_sum + nb_pt_auto_sales_sum
    if auto_sales_sum != 0:
        auto_acos = auto_spend_sum/auto_sales_sum
    else:
        auto_acos = 0

    auto_kw_spend = br_kw_auto_spend_sum + nb_kw_auto_spend_sum
    auto_kw_sales = br_kw_auto_sales_sum + nb_kw_auto_sales_sum
    if auto_kw_sales != 0:
        auto_kw_acos = auto_kw_spend / auto_kw_sales
    else:
        auto_kw_acos = 0

    auto_pt_spend = br_pt_auto_spend_sum + nb_pt_auto_spend_sum
    auto_pt_sales = br_pt_auto_sales_sum + nb_pt_auto_sales_sum
    if auto_pt_sales != 0:
        auto_pt_acos = auto_pt_spend / auto_pt_sales
    else:
        auto_pt_acos = 0

    
    nb_auto_spend = nb_kw_auto_spend_sum + nb_pt_auto_spend_sum
    br_auto_spend = br_kw_auto_spend_sum + br_pt_auto_spend_sum
    nb_auto_sales = nb_kw_auto_sales_sum + nb_pt_auto_sales_sum
    br_auto_sales = br_kw_auto_sales_sum + br_pt_auto_sales_sum

    if nb_auto_sales != 0:
        nb_auto_acos = nb_auto_spend / nb_auto_sales
    else:
        nb_auto_acos = 0
    if br_auto_sales != 0:
        br_auto_acos = br_auto_spend / br_auto_sales
    else:
        br_auto_acos = 0

    if br_kw_auto_sales_sum != 0:
        br_kw_auto_acos = br_kw_auto_spend_sum / br_kw_auto_sales_sum
    if br_pt_auto_sales_sum != 0:
        br_pt_auto_acos = br_pt_auto_spend_sum / br_pt_auto_sales_sum
    if nb_kw_auto_sales_sum != 0:
        nb_kw_auto_acos = nb_kw_auto_spend_sum / nb_kw_auto_sales_sum
    if nb_pt_auto_sales_sum != 0:
        nb_pt_auto_acos = nb_pt_auto_spend_sum / nb_pt_auto_sales_sum


## manual calculations #############################

    manual_spend_sum = br_kw_manual_spend_sum + br_pt_manual_spend_sum + nb_kw_manual_spend_sum + nb_pt_manual_spend_sum
    manual_sales_sum = br_kw_manual_sales_sum + br_pt_manual_sales_sum + nb_kw_manual_sales_sum + nb_pt_manual_sales_sum
    if manual_sales_sum != 0:
        manual_acos = manual_spend_sum/manual_sales_sum

    manual_kw_spend = br_kw_manual_spend_sum + nb_kw_manual_spend_sum
    manual_kw_sales = br_kw_manual_sales_sum + nb_kw_manual_sales_sum
    if manual_kw_sales != 0:
        manual_kw_acos = manual_kw_spend / manual_kw_sales
    
    manual_pt_spend = br_pt_manual_spend_sum + nb_pt_manual_spend_sum
    manual_pt_sales = br_pt_manual_sales_sum + nb_pt_manual_sales_sum
    if manual_pt_sales != 0:
        manual_pt_acos = manual_pt_spend / manual_pt_sales
    
    nb_manual_spend = nb_kw_manual_spend_sum + nb_pt_manual_spend_sum
    br_manual_spend = br_kw_manual_spend_sum + br_pt_manual_spend_sum
    nb_manual_sales = nb_kw_manual_sales_sum + nb_pt_manual_sales_sum
    br_manual_sales = br_kw_manual_sales_sum + br_pt_manual_sales_sum

    if nb_manual_sales != 0:
        nb_manual_acos = nb_manual_spend / nb_manual_sales
    if br_manual_sales != 0:
        br_manual_acos = br_manual_spend / br_manual_sales
    if br_kw_manual_sales_sum != 0:
        br_kw_manual_acos = br_kw_manual_spend_sum / br_kw_manual_sales_sum
    if br_pt_manual_sales_sum != 0:
        br_pt_manual_acos = br_pt_manual_spend_sum / br_pt_manual_sales_sum
    if nb_kw_manual_sales_sum != 0:
        nb_kw_manual_acos = nb_kw_manual_spend_sum / nb_kw_manual_sales_sum
    if nb_pt_manual_sales_sum != 0:
        nb_pt_manual_acos = nb_pt_manual_spend_sum / nb_pt_manual_sales_sum



    br_total_spend = br_kw_auto_spend_sum + br_kw_manual_spend_sum + br_pt_auto_spend_sum + br_pt_manual_spend_sum
    br_total_sales = br_kw_auto_sales_sum + br_kw_manual_sales_sum + br_pt_auto_sales_sum + br_pt_manual_sales_sum
    if br_total_sales != 0:
        br_acos = br_total_spend / br_total_sales
    else:
        br_acos = 0

    nb_total_spend = nb_kw_auto_spend_sum + nb_kw_manual_spend_sum + nb_pt_auto_spend_sum + nb_pt_manual_spend_sum + nb_pt_cat_manual_spend_sum
    nb_total_sales = nb_kw_auto_sales_sum + nb_kw_manual_sales_sum+ nb_pt_auto_sales_sum + nb_pt_manual_sales_sum + nb_pt_cat_manual_sales_sum
    if nb_total_sales != 0:
        nb_acos = nb_total_spend / nb_total_sales
    else:
        nb_acos = 0

    br_kw_spend = br_kw_auto_spend_sum + br_kw_manual_spend_sum
    br_kw_sales = br_kw_auto_sales_sum + br_kw_manual_sales_sum
    if br_kw_sales != 0:
        br_kw_acos = br_kw_spend / br_kw_sales
    else:
        br_kw_acos = 0

    br_pt_spend = br_pt_auto_spend_sum + br_pt_manual_spend_sum
    br_pt_sales = br_pt_auto_sales_sum + br_pt_manual_sales_sum
    if br_pt_sales != 0:
        br_pt_acos = br_pt_spend / br_pt_sales
    else:
        br_pt_acos = 0

    nb_kw_spend = nb_kw_auto_spend_sum + nb_kw_manual_spend_sum
    nb_kw_sales = nb_kw_auto_sales_sum + nb_kw_manual_sales_sum    
    if nb_kw_sales != 0:
        nb_kw_acos = nb_kw_spend / nb_kw_sales
    else:
        nb_kw_acos = 0
    nb_pt_spend = nb_pt_auto_spend_sum + nb_pt_manual_spend_sum + nb_pt_cat_manual_spend_sum
    nb_pt_sales = nb_pt_auto_sales_sum + nb_pt_manual_sales_sum + nb_pt_cat_manual_sales_sum
    if nb_pt_sales != 0:
        nb_pt_acos = nb_pt_spend / nb_pt_sales
    else:
        nb_pt_acos = 0

    kw_spend = br_kw_spend + nb_kw_spend
    kw_sales = br_kw_sales + nb_kw_sales
    if kw_sales != 0:
        kw_acos = kw_spend / kw_sales
    else:
        kw_acos = 0

    pt_spend = br_pt_spend + nb_pt_spend
    pt_sales = br_pt_sales + nb_pt_sales
    if pt_sales != 0:
        pt_acos = pt_spend / pt_sales
    else:
        pt_acos = 0

    exact_spend = br_exact_spend + nb_exact_spend
    exact_sales = br_exact_sales + nb_exact_sales
    if exact_sales != 0:
        exact_acos = exact_spend / exact_sales
    else:
        exact_acos = 0
    phrase_spend = br_phrase_spend + nb_phrase_spend
    phrase_sales = br_phrase_sales + nb_phrase_sales
    if phrase_sales != 0:
        phrase_acos = phrase_spend / phrase_sales
    else:
        phrase_acos = 0
    broad_spend = br_broad_spend + nb_broad_spend
    broad_sales = br_broad_sales + nb_broad_sales
    if broad_sales != 0:
        broad_acos = broad_spend / broad_sales
    else:
       broad_acos = 0

    if br_exact_sales != 0:
        br_exact_acos = br_exact_spend / br_exact_sales
    else:
        br_exact_acos = 0
    if br_phrase_sales != 0:
        br_phrase_acos = br_phrase_spend / br_phrase_sales
    else:
        br_phrase_acos = 0
    if br_broad_sales != 0:
        br_broad_acos = br_broad_spend / br_broad_sales
    else:
        br_broad_acos = 0

    if nb_exact_sales != 0:
        nb_exact_acos = nb_exact_spend / nb_exact_sales
    else:
        nb_exact_acos = 0
    if nb_phrase_sales != 0:
        nb_phrase_acos = nb_phrase_spend / nb_phrase_sales
    else:
        nb_phrase_acos = 0
    if nb_broad_sales != 0:
        nb_broad_acos = nb_broad_spend / nb_broad_sales
    else:
        nb_broad_acos = 0


    # if sp_sales == None or sp_spend == None:
    #     sp_sales = 0
    #     sp_spend = 0

    # if sb_sales == None or sb_spend == None:
    #     sb_sales = 0
    #     sb_spend = 0

    # if sbv_sales == None or sbv_spend == None:
    #     sbv_sales = 0
    #     sbv_spend = 0

    # if sd_sales == None or sd_spend == None:
    #     sd_sales = 0
    #     sd_spend = 0

    if sp_sales != 0:
        sp_acos = sp_spend / sp_sales
    else:
        sp_acos = 0
    
    if sb_sales != 0:
        sb_acos = sb_spend / sb_sales
    else:
        sb_acos = 0

    if sbv_sales != 0:
        sbv_acos = sbv_spend / sbv_sales
    else:
        sbv_acos = 0

    if sd_sales != 0:
        sd_acos = sd_spend / sd_sales
    else:
        sd_acos = 0


    
## Calculate the spend % for each ad type if one of them does not have any spend, and for when they do have spend.
    if sp_spend > 0 or sb_spend > 0 or sbv_spend > 0 or sd_spend > 0:
        ad_types_spend_sum = sp_spend + sb_spend + sbv_spend + sd_spend
        ad_types_sales_sum = sp_sales + sb_sales + sbv_sales + sd_sales
        sp_spend_percentage = sp_spend / ad_types_spend_sum
        sb_spend_percentage = sb_spend / ad_types_spend_sum
        sbv_spend_percentage = sbv_spend / ad_types_spend_sum
        sd_spend_percentage = sd_spend / ad_types_spend_sum
        sp_sales_percentage = sp_sales / ad_types_sales_sum
        sb_sales_percentage = sb_sales / ad_types_sales_sum
        sbv_sales_percentage = sbv_sales / ad_types_sales_sum
        sd_sales_percentage = sd_sales / ad_types_sales_sum
    else:
        sp_spend_percentage = 0
        sb_spend_percentage = 0
        sbv_spend_percentage = 0
        sd_spend_percentage = 0
        sp_sales_percentage = 0
        sb_sales_percentage = 0
        sbv_sales_percentage = 0
        sd_sales_percentage = 0


    kw_spend_percentage = kw_spend / sum_spend
    kw_sales_percentage = kw_sales / sum_sales
    pt_spend_percentage = pt_spend / sum_spend
    pt_sales_percentage = pt_sales / sum_sales

    br_spend_percentage = br_total_spend / sum_spend
    br_sales_percentage = br_total_sales / sum_sales
    nb_spend_percentage = nb_total_spend / sum_spend
    nb_sales_percentage = nb_total_sales / sum_sales

    br_kw_spend_percentage = br_kw_spend / sum_spend
    br_kw_sales_percentage = br_kw_sales / sum_sales
    nb_kw_spend_percentage = nb_kw_spend / sum_spend
    nb_kw_sales_percentage = nb_kw_sales / sum_sales
    br_pt_spend_percentage = br_pt_spend / sum_spend
    br_pt_sales_percentage = br_pt_sales / sum_sales
    nb_pt_spend_percentage = nb_pt_spend / sum_spend
    nb_pt_sales_percentage = nb_pt_sales / sum_sales

    broad_spend_percentage = broad_spend / kw_spend
    broad_sales_percentage = broad_sales / kw_sales
    phrase_spend_percentage = phrase_spend / kw_spend
    phrase_sales_percentage = phrase_sales / kw_sales
    exact_spend_percentage = exact_spend / kw_spend
    exact_sales_percentage = exact_sales / kw_sales

    nb_broad_spend_percentage = nb_broad_spend / kw_spend
    nb_broad_sales_percentage = nb_broad_sales / kw_sales
    nb_phrase_spend_percentage = nb_phrase_spend / kw_spend
    nb_phrase_sales_percentage = nb_phrase_sales / kw_sales
    nb_exact_spend_percentage = nb_exact_spend / kw_spend
    nb_exact_sales_percentage = nb_exact_sales / kw_sales

    br_broad_spend_percentage = br_broad_spend / kw_spend
    br_broad_sales_percentage = br_broad_sales / kw_sales
    br_phrase_spend_percentage = br_phrase_spend / kw_spend
    br_phrase_sales_percentage = br_phrase_sales / kw_sales
    br_exact_spend_percentage = br_exact_spend / kw_spend
    br_exact_sales_percentage = br_exact_sales / kw_sales

    

# Define row data in the excel data that will later be used when creating the file ###########################
    row1 = [f"{brand_name} Ad Data"]
    row2 = ["", "Total", "", "", "", "", "" ,"","Ad Type"]
    row3 = ["", "Ad Spend", "Ad Sales", "ACoS", "", "", "", "", "", "Ad Spend", "Ad Sales", "ACoS", "% Spend","% Sales" ]
    row4 = [f"{brand_name} Account", f"${sum_spend:,.2f}",f"${sum_sales:,.2f}", f"{acos:,.2f}%", "",  "", "","", "SP", f"${sp_spend:,.2f}", f"${sp_sales:,.2f}", f"{sp_acos:,.2f}%", f"{sp_spend_percentage:,.2f}%", f"{sp_sales_percentage:,.2f}%"]
    row5 = ["", "", "",  "", "", "", "", "", "SB", f"${sb_spend:,.2f}", f"${sb_sales:,.2f}", f"{sb_acos:,.2f}%", f"{sb_spend_percentage:,.2f}%", f"{sb_sales_percentage:,.2f}%"]
    row6 = ["", "", "",  "", "", "","",  "","SBV", f"${sbv_spend:,.2f}", f"${sbv_sales:,.2f}", f"{sbv_acos:,.2f}%", f"{sbv_spend_percentage:,.2f}%", f"{sbv_sales_percentage:,.2f}%"]
    row7 = ["", "Account KW vs PT", "", "", "", "","", "", "SD", f"${sd_spend:,.2f}", f"${sd_sales:,.2f}", f"{sd_acos:,.2f}%", f"{sd_spend_percentage:,.2f}%", f"{sd_sales_percentage:,.2f}%"]
    row8 = ["", "Ad Spend", "Ad Sales", "ACoS", "% Spend","% Sales"]
    row9 = ["Keyword", f"${kw_spend:,.2f}", f"${kw_sales:,.2f}", f"{kw_acos:,.2f}%", f"{kw_spend_percentage:,.2f}%", f"{kw_sales_percentage:,.2f}%"]
    row10 = ["Product Targeting", f"${pt_spend:,.2f}", f"${pt_sales:,.2f}", f"{pt_acos:,.2f}%", f"{pt_spend_percentage:,.2f}%", f"{pt_sales_percentage:,.2f}%"]
    row11 = [""]
    row12 = [""]
    row13 = ["", "BR vs NB"]
    row14 = ["", "Ad Spend", "Ad Sales", "ACoS", "% Spend","% Sales"]
    row15 = ["Branded", f"${br_total_spend:,.2f}", f"${br_total_sales:,.2f}", f"{br_acos:,.2f}%", f"{br_spend_percentage:,.2f}%", f"{br_sales_percentage:,.2f}%"]
    row16 = ["Non-Branded", f"${nb_total_spend:,.2f}", f"${nb_total_sales:,.2f}", f"{nb_acos:,.2f}%", f"{nb_spend_percentage:,.2f}%", f"{nb_sales_percentage:,.2f}%"]
    row17 = [""]
    row18 = [""]
    row19 = ["", "BR vs NB KW & PT"]
    row20 = ["", "Ad Spend", "Ad Sales", "ACoS", "% Spend","% Sales"]
    row21 = ["Branded KW", f"${br_kw_spend:,.2f}", f"${br_kw_sales:,.2f}", f"{br_kw_acos:,.2f}%", f"{br_kw_spend_percentage:,.2f}%", f"{br_kw_sales_percentage:,.2f}%"]
    row22 = ["Branded PT", f"${br_pt_spend:,.2f}", f"${br_pt_sales:,.2f}", f"{br_pt_acos:,.2f}%", f"{br_pt_spend_percentage:,.2f}%", f"{br_pt_sales_percentage:,.2f}%"]
    row23 = ["Non-Branded KW", f"${nb_kw_spend:,.2f}", f"${nb_kw_sales:,.2f}", f"{nb_kw_acos:,.2f}%", f"{nb_kw_spend_percentage:,.2f}%", f"{nb_kw_sales_percentage:,.2f}%"]
    row24 = ["Non-Branded PT", f"${nb_pt_spend:,.2f}", f"${nb_pt_sales:,.2f}", f"{nb_pt_acos:,.2f}%", f"{nb_pt_spend_percentage:,.2f}%", f"{nb_pt_sales_percentage:,.2f}%"]
    
    # Create a new CSV file and add the data to it ################################

    # with open(ad_report_file_path, "w") as report:
    #     report = csv.writer(report)
    #     report.writerow(row1)
    #     report.writerow(row2)
    #     report.writerow(row3)
    #     report.writerow(row4)
    #     report.writerow(row5)
    #     report.writerow(row6)
    #     report.writerow(row7)
    #     report.writerow(row8)
    #     report.writerow(row9)
    #     report.writerow(row10)
    #     report.writerow(row11)
    #     report.writerow(row12)
    #     report.writerow(row13)
    #     report.writerow(row14)
    #     report.writerow(row15)
    #     report.writerow(row16)
    #     report.writerow(row17)
    #     report.writerow(row18)
    #     report.writerow(row19)
    #     report.writerow(row20)
    #     report.writerow(row21)
    #     report.writerow(row22)
    #     report.writerow(row23)
    #     report.writerow(row24)

        
    # subprocess.call(["open", ad_report_file_path])
    # # open("ad_report.csv")  
    
# Create a new excel file, style it, and add data #######################################
    workbook = xlsxwriter.Workbook(f"{ad_report_file_path_no_ext}.xlsx")
    worksheet = workbook.add_worksheet(f"{brand_name} Ad Data")
    worksheet.set_column(0, 0, 16)
    worksheet.set_column(1, 2, 12)
    worksheet.set_column(3, 7, 8)
    worksheet.set_column(9, 10, 12)
    worksheet.set_column(11, 13, 8)
    a1_format = workbook.add_format(
        {
            "font_size": 14,
            "bold": True,
            
            
        }
    )
    top_left_format = workbook.add_format(
        {
            "font_size": 12,
            "bold": True,
            "bottom": True,
            "right": True
            
            
        }
    )
    column_header_format = workbook.add_format(
        {
            "font_size": 12,
            "bold": True,
            "bottom": True
            
        }
    )
    row_header_format = workbook.add_format(
        {
            "font_size": 12,
            "bold": True,
            "right": True
            
        }
    )
    currency_format = workbook.add_format(
        {
            "num_format": 7
        }
    )

    percentage_format = workbook.add_format(
        {
            "num_format": 10
        }
    )

    worksheet.write('A1', f"{brand_name} Ad Data", a1_format)

# Total Table
    worksheet.write('A3', "Total", top_left_format)
    worksheet.write('A4', f"{brand_name} Account", row_header_format)
    worksheet.write('B3', "Ad Spend", column_header_format)
    worksheet.write('B4', float(sum_spend),currency_format)
    worksheet.write('C3', "Ad Sales", column_header_format)
    worksheet.write('C4', float(sum_sales),currency_format)
    worksheet.write('D3', "ACoS", column_header_format)
    worksheet.write('D4', float(acos), percentage_format)
    
# Ad Type Table
    worksheet.write('I3', "Ad Type", top_left_format)
    worksheet.write('I4', "SP", row_header_format)
    worksheet.write('I5', "SB", row_header_format)
    worksheet.write('I6', "SBV", row_header_format)
    worksheet.write('I7', "SD", row_header_format)
    worksheet.write('J3', "Ad Spend", column_header_format)
    worksheet.write('J4', float(sp_spend),currency_format, )
    worksheet.write('J5', float(sb_spend),currency_format)
    worksheet.write('J6', float(sbv_spend),currency_format)
    worksheet.write('J7', float(sd_spend),currency_format)
    worksheet.write('K3', "Ad Sales", column_header_format)
    worksheet.write('K4', float(sp_sales),currency_format)
    worksheet.write('K5', float(sb_sales),currency_format)
    worksheet.write('K6', float(sbv_sales),currency_format)
    worksheet.write('K7', float(sd_sales),currency_format)
    worksheet.write('L3', "ACoS", column_header_format)
    worksheet.write('L4', float(sp_acos), percentage_format)
    worksheet.write('L5', float(sb_acos), percentage_format)
    worksheet.write('L6', float(sbv_acos), percentage_format)
    worksheet.write('L7', float(sd_acos), percentage_format)
    worksheet.write('M3', "% Spend", column_header_format)
    worksheet.write('M4', float(sp_spend_percentage),percentage_format)
    worksheet.write('M5', float(sb_spend_percentage), percentage_format)
    worksheet.write('M6', float(sbv_spend_percentage), percentage_format)
    worksheet.write('M7', float(sd_spend_percentage), percentage_format)
    worksheet.write('N3', "% Sales", column_header_format)
    worksheet.write('N4', float(sp_sales_percentage),percentage_format)
    worksheet.write('N5', float(sb_sales_percentage), percentage_format)
    worksheet.write('N6', float(sbv_sales_percentage), percentage_format)
    worksheet.write('N7', float(sd_sales_percentage), percentage_format)

    # KW vs PT Table
    worksheet.write('A8', "KW vs PT", top_left_format)
    worksheet.write('A9', "Keyword", row_header_format)
    worksheet.write('A10', "Product Targeting", row_header_format)
    worksheet.write('B8', "Ad Spend", column_header_format)
    worksheet.write('B9', float(kw_spend),currency_format, )
    worksheet.write('B10', float(pt_spend),currency_format)
    worksheet.write('C8', "Ad Sales", column_header_format)
    worksheet.write('C9', float(kw_sales),currency_format)
    worksheet.write('C10', float(pt_sales),currency_format)
    worksheet.write('D8', "ACoS", column_header_format)
    worksheet.write('D9', float(kw_acos), percentage_format)
    worksheet.write('D10', float(pt_acos), percentage_format)
    worksheet.write('E8', "% Spend", column_header_format)
    worksheet.write('E9', float(kw_spend_percentage),percentage_format)
    worksheet.write('E10', float(pt_spend_percentage), percentage_format)
    worksheet.write('F8', "% Sales", column_header_format)
    worksheet.write('F9', float(kw_sales_percentage),percentage_format)
    worksheet.write('F10', float(pt_sales_percentage), percentage_format)

    #Auto vs Manual
    worksheet.write('I10', "Auto vs Manual", top_left_format)
    worksheet.write('I11', "Auto", row_header_format)
    worksheet.write('I12', "Manual", row_header_format)
    worksheet.write('J10', "Ad Spend", column_header_format)
    worksheet.write('J11', float(auto_spend_sum),currency_format, )
    worksheet.write('J12', float(manual_spend_sum),currency_format)
    worksheet.write('K10', "Ad Sales", column_header_format)
    worksheet.write('K11', float(auto_sales_sum),currency_format)
    worksheet.write('K12', float(manual_sales_sum),currency_format)
    worksheet.write('L10', "ACoS", column_header_format)
    worksheet.write('L11', float(auto_acos), percentage_format)
    worksheet.write('L12', float(manual_acos), percentage_format)
    worksheet.write('M10', "% Spend", column_header_format)
    # worksheet.write('M11', float(br_spend_percentage),percentage_format)
    # worksheet.write('M12', float(nb_spend_percentage), percentage_format)
    worksheet.write('N10', "% Sales", column_header_format)
    # worksheet.write('N11', float(br_sales_percentage),percentage_format)
    # worksheet.write('N12', float(nb_sales_percentage), percentage_format)

    # BR vs NB
    worksheet.write('A14', "BR vs NB", top_left_format)
    worksheet.write('A15', "Branded", row_header_format)
    worksheet.write('A16', "Non-Branded", row_header_format)
    worksheet.write('B14', "Ad Spend", column_header_format)
    worksheet.write('B15', float(br_total_spend),currency_format, )
    worksheet.write('B16', float(nb_total_spend),currency_format)
    worksheet.write('C14', "Ad Sales", column_header_format)
    worksheet.write('C15', float(br_total_sales),currency_format)
    worksheet.write('C16', float(nb_total_sales),currency_format)
    worksheet.write('D14', "ACoS", column_header_format)
    worksheet.write('D15', float(br_acos), percentage_format)
    worksheet.write('D16', float(nb_acos), percentage_format)
    worksheet.write('E14', "% Spend", column_header_format)
    worksheet.write('E15', float(br_spend_percentage),percentage_format)
    worksheet.write('E16', float(nb_spend_percentage), percentage_format)
    worksheet.write('F14', "% Sales", column_header_format)
    worksheet.write('F15', float(br_sales_percentage),percentage_format)
    worksheet.write('F16', float(nb_sales_percentage), percentage_format)

     #Auto/Manual - KW/PT
    worksheet.write('I15', "Auto/Manual - KW/PT", top_left_format)
    worksheet.write('I16', "Auto KW", row_header_format)
    worksheet.write('I17', "Auto PT", row_header_format)
    worksheet.write('I18', "Manual KW", row_header_format)
    worksheet.write('I19', "Manual PT", row_header_format)
    worksheet.write('J15', "Ad Spend", column_header_format)
    worksheet.write('J16', float(auto_kw_spend),currency_format, )
    worksheet.write('J17', float(auto_pt_spend),currency_format)
    worksheet.write('J18', float(manual_kw_spend),currency_format, )
    worksheet.write('J19', float(manual_pt_spend),currency_format)
    worksheet.write('K15', "Ad Sales", column_header_format)
    worksheet.write('K16', float(auto_kw_sales),currency_format)
    worksheet.write('K17', float(auto_pt_sales),currency_format)
    worksheet.write('K18', float(manual_kw_sales),currency_format)
    worksheet.write('K19', float(manual_pt_sales),currency_format)
    worksheet.write('L15', "ACoS", column_header_format)
    worksheet.write('L16', float(auto_kw_acos),percentage_format)
    worksheet.write('L17', float(auto_pt_acos),percentage_format)
    worksheet.write('L18', float(manual_kw_acos),percentage_format)
    worksheet.write('L19', float(manual_pt_acos),percentage_format)
    worksheet.write('M15', "% Spend", column_header_format)
    # worksheet.write('M11', float(br_spend_percentage),percentage_format)
    # worksheet.write('M12', float(nb_spend_percentage), percentage_format)
    worksheet.write('N15', "% Sales", column_header_format)
    # worksheet.write('N11', float(br_sales_percentage),percentage_format)
    # worksheet.write('N12', float(nb_sales_percentage), percentage_format)

    # BR vs NB KW & PT
    worksheet.write('A20', "BR vs NB KW & PT", top_left_format)
    worksheet.write('A21', "Branded KW", row_header_format)
    worksheet.write('A22', "Branded PT", row_header_format)
    worksheet.write('A23', "Non-Branded KW", row_header_format)
    worksheet.write('A24', "Non-Branded PT", row_header_format)
    worksheet.write('B20', "Ad Spend", column_header_format)
    worksheet.write('B21', float(br_kw_spend),currency_format, )
    worksheet.write('B22', float(br_pt_spend),currency_format)
    worksheet.write('B23', float(nb_kw_spend),currency_format)
    worksheet.write('B24', float(nb_pt_spend),currency_format)
    worksheet.write('C20', "Ad Sales", column_header_format)
    worksheet.write('C21', float(br_kw_sales),currency_format)
    worksheet.write('C22', float(br_pt_sales),currency_format)
    worksheet.write('C23', float(nb_kw_sales),currency_format)
    worksheet.write('C24', float(nb_pt_sales),currency_format)
    worksheet.write('D20', "ACoS", column_header_format)
    worksheet.write('D21', float(br_kw_acos), percentage_format)
    worksheet.write('D22', float(br_pt_acos), percentage_format)
    worksheet.write('D23', float(nb_kw_acos), percentage_format)
    worksheet.write('D24', float(nb_pt_acos), percentage_format)
    worksheet.write('E20', "% Spend", column_header_format)
    worksheet.write('E21', float(br_kw_spend_percentage),percentage_format)
    worksheet.write('E22', float(br_pt_spend_percentage), percentage_format)
    worksheet.write('E23', float(nb_kw_spend_percentage), percentage_format)
    worksheet.write('E24', float(nb_pt_spend_percentage), percentage_format)
    worksheet.write('F20', "% Sales", column_header_format)
    worksheet.write('F21', float(br_kw_sales_percentage),percentage_format)
    worksheet.write('F22', float(br_pt_sales_percentage), percentage_format)
    worksheet.write('F23', float(nb_kw_sales_percentage), percentage_format)
    worksheet.write('F24', float(nb_pt_sales_percentage), percentage_format)


 # KW Match Type
    worksheet.write('A27', "KW Match Type", top_left_format)
    worksheet.write('A28', "Broad", row_header_format)
    worksheet.write('A29', "Phrase", row_header_format)
    worksheet.write('A30', "Exact", row_header_format)
    worksheet.write('B27', "Ad Spend", column_header_format)
    worksheet.write('B28', float(broad_spend),currency_format, )
    worksheet.write('B29', float(phrase_spend),currency_format)
    worksheet.write('B30', float(exact_spend),currency_format)
    worksheet.write('C27', "Ad Sales", column_header_format)
    worksheet.write('C28', float(broad_sales),currency_format)
    worksheet.write('C29', float(phrase_sales),currency_format)
    worksheet.write('C30', float(exact_sales),currency_format)
    worksheet.write('D27', "ACoS", column_header_format)
    worksheet.write('D28', float(broad_acos), percentage_format)
    worksheet.write('D29', float(phrase_acos), percentage_format)
    worksheet.write('D30', float(exact_acos), percentage_format)
    worksheet.write('E27', "% Spend", column_header_format)
    worksheet.write('E28', float(broad_spend_percentage),percentage_format)
    worksheet.write('E29', float(phrase_spend_percentage), percentage_format)
    worksheet.write('E30', float(exact_spend_percentage), percentage_format)
    worksheet.write('F27', "% Sales", column_header_format)
    worksheet.write('F28', float(broad_sales_percentage),percentage_format)
    worksheet.write('F29', float(phrase_sales_percentage), percentage_format)
    worksheet.write('F30', float(exact_sales_percentage), percentage_format)


    ## BR/NB Match Type
    worksheet.write('A33', "KW Match Type", top_left_format)
    worksheet.write('A34', "NB Broad", row_header_format)
    worksheet.write('A35', "NB Phrase", row_header_format)
    worksheet.write('A36', "NB Exact", row_header_format)
    worksheet.write('A37', "BR Broad", row_header_format)
    worksheet.write('A38', "BR Phrase", row_header_format)
    worksheet.write('A39', "BR Exact", row_header_format)
    worksheet.write('B33', "Ad Spend", column_header_format)
    worksheet.write('B34', float(nb_broad_spend),currency_format, )
    worksheet.write('B35', float(nb_phrase_spend),currency_format)
    worksheet.write('B36', float(nb_exact_spend),currency_format)
    worksheet.write('B37', float(br_broad_spend),currency_format, )
    worksheet.write('B38', float(br_phrase_spend),currency_format)
    worksheet.write('B39', float(br_exact_spend),currency_format)
    worksheet.write('C33', "Ad Sales", column_header_format)
    worksheet.write('C34', float(nb_broad_sales),currency_format)
    worksheet.write('C35', float(nb_phrase_sales),currency_format)
    worksheet.write('C36', float(nb_exact_sales),currency_format)
    worksheet.write('C37', float(br_broad_sales),currency_format)
    worksheet.write('C38', float(br_phrase_sales),currency_format)
    worksheet.write('C39', float(br_exact_sales),currency_format)
    worksheet.write('D33', "ACoS", column_header_format)
    worksheet.write('D34', float(nb_broad_acos), percentage_format)
    worksheet.write('D35', float(nb_phrase_acos), percentage_format)
    worksheet.write('D36', float(nb_exact_acos), percentage_format)
    worksheet.write('D37', float(br_broad_acos), percentage_format)
    worksheet.write('D38', float(br_phrase_acos), percentage_format)
    worksheet.write('D39', float(br_exact_acos), percentage_format)
    worksheet.write('E33', "% Spend", column_header_format)
    worksheet.write('E34', float(nb_broad_spend_percentage), percentage_format)
    worksheet.write('E35', float(nb_phrase_spend_percentage), percentage_format)
    worksheet.write('E36', float(nb_exact_spend_percentage), percentage_format)
    worksheet.write('E37', float(br_broad_spend_percentage), percentage_format)
    worksheet.write('E38', float(br_phrase_spend_percentage), percentage_format)
    worksheet.write('E39', float(br_exact_spend_percentage), percentage_format)
    worksheet.write('F33', "% Sales", column_header_format)
    worksheet.write('F34', float(nb_broad_sales_percentage), percentage_format)
    worksheet.write('F35', float(nb_phrase_sales_percentage), percentage_format)
    worksheet.write('F36', float(nb_exact_sales_percentage), percentage_format)
    worksheet.write('F37', float(br_broad_sales_percentage), percentage_format)
    worksheet.write('F38', float(br_phrase_sales_percentage), percentage_format)
    worksheet.write('F39', float(br_exact_sales_percentage), percentage_format)

    workbook.close()
    

    subprocess.call(["open", f"{ad_report_file_path_no_ext}.xlsx"])

if __name__ == "__main__":
    main()

    