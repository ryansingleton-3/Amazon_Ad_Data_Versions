import csv
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
        global entry1, entry2, entry0
        global e1, e2, e0
        e0 = entry0.get()
        e1 = entry1.get()
        e2 = entry2.get()
        window.destroy()

   
    gui()
    try:
        brand_name = e0
        branded_keywords = e1
        branded_asins = e2
        branded_keywords = branded_keywords.split(", ")
        branded_asins = branded_asins.split(", ")
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
    df2 = df.replace(to_replace=np.nan, value=int(0))
    df2["Spend"].fillna(int(0))
    df2["7 Day Total Sales "].fillna(int(0))
    df2["Targeting"] = df2["Targeting"].astype("str")
    df2["7 Day Total Sales "] = df2["7 Day Total Sales "].fillna(0)
    df2["Spend"] = df2["Spend"].fillna(0)
    df3 = df2.fillna(0)
    df3['7 Day Total Sales '] = df['7 Day Total Sales '].replace(np.nan, 0)
    sum_spend = df3["Spend"].sum()
    sum_sales = df3["7 Day Total Sales "].sum()
    # print()
    # print(f"Sum of Spend: ${sum_spend:,.2f}")
    # print(f"Sum of Sales: ${sum_sales:,.2f}")
    acos = sum_spend / sum_sales * 100
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
        # print(ad_type_spend)
        # print(ad_type_sales)
        sp_spend = ad_type_spend["SP"]
        sp_sales = ad_type_sales["SP"]
        sb_spend = ad_type_spend["SB"]
        sb_sales = ad_type_sales["SB"]
        sbv_spend = ad_type_spend["SBV"]
        sbv_sales = ad_type_sales["SBV"]
        sd_spend = ad_type_spend["SD"]
        sd_sales = ad_type_sales["SD"]
        # print(ad_type_spend)
        # print(ad_type_sales)
        # print("")
        # print(f"SP Spend: {sp_spend}")
        # print(f"SP Sales: {sp_sales}")
        

    column_names = list(df3.columns)
    if any(name.lower() == ad_type_str.lower() for name in column_names):
        group_by_ad_type()
        # column_names.remove(name)
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
            

            current_row = df.iloc[row_number - 1]
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


            elif any(str(nb_kw) == str(search_term) for nb_kw in nb_st_list):
                current_row_nb_kw_manual_spend = current_row["Spend"]
                nb_kw_manual_spend_sum = nb_kw_manual_spend_sum + current_row_nb_kw_manual_spend
                current_row_nb_kw_manual_sales = current_row["7 Day Total Sales "]
                nb_kw_manual_sales_sum = nb_kw_manual_sales_sum + current_row_nb_kw_manual_sales  

            elif row == "" or row == 0:
                break

            else:
                undefined_array.append(row)
                
            
            row_number = row_number + 1
    elif len(branded_asins) <= 1:
        for row in df3["Targeting"]:
            

            current_row = df.iloc[row_number - 1]
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


            elif any(nb_kw in search_term for nb_kw in nb_st_list):
                current_row_nb_kw_manual_spend = current_row["Spend"]
                nb_kw_manual_spend_sum = nb_kw_manual_spend_sum + current_row_nb_kw_manual_spend
                current_row_nb_kw_manual_sales = current_row["7 Day Total Sales "]
                nb_kw_manual_sales_sum = nb_kw_manual_sales_sum + current_row_nb_kw_manual_sales  

            elif row == "" or row == 0:
                break

            else:
                undefined_array.append(row)
                
            
            row_number = row_number + 1
    
    else: 
        for row in df3["Targeting"]:
            

            current_row = df.iloc[row_number - 1]
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


            elif any(nb_kw in search_term for nb_kw in nb_st_list):
                current_row_nb_kw_manual_spend = current_row["Spend"]
                nb_kw_manual_spend_sum = nb_kw_manual_spend_sum + current_row_nb_kw_manual_spend
                current_row_nb_kw_manual_sales = current_row["7 Day Total Sales "]
                nb_kw_manual_sales_sum = nb_kw_manual_sales_sum + current_row_nb_kw_manual_sales  

            elif row == "" or row == 0:
                break

            else:
                undefined_array.append(row)
                
            
            row_number = row_number + 1



    br_total_spend = br_kw_auto_spend_sum + br_kw_manual_spend_sum + br_pt_auto_spend_sum + br_pt_manual_spend_sum
    br_total_sales = br_kw_auto_sales_sum + br_kw_manual_sales_sum + br_pt_auto_sales_sum + br_pt_manual_sales_sum
    if br_total_sales != 0:
        br_acos = br_total_spend / br_total_sales * 100
    else:
        br_acos = 0

    nb_total_spend = nb_kw_auto_spend_sum + nb_kw_manual_spend_sum + nb_pt_auto_spend_sum + nb_pt_manual_spend_sum + nb_pt_cat_manual_spend_sum
    nb_total_sales = nb_kw_auto_sales_sum + nb_kw_manual_sales_sum+ nb_pt_auto_sales_sum + nb_pt_manual_sales_sum + nb_pt_cat_manual_sales_sum
    if nb_total_sales != 0:
        nb_acos = nb_total_spend / nb_total_sales * 100
    else:
        nb_acos = 0

    br_kw_spend = br_kw_auto_spend_sum + br_kw_manual_spend_sum
    br_kw_sales = br_kw_auto_sales_sum + br_kw_manual_sales_sum
    if br_kw_sales != 0:
        br_kw_acos = br_kw_spend / br_kw_sales * 100
    else:
        br_kw_acos = 0

    br_pt_spend = br_pt_auto_spend_sum + br_pt_manual_spend_sum
    br_pt_sales = br_pt_auto_sales_sum + br_pt_manual_sales_sum
    if br_pt_sales != 0:
        br_pt_acos = br_pt_spend / br_pt_sales * 100
    else:
        br_pt_acos = 0

    nb_kw_spend = nb_kw_auto_spend_sum + nb_kw_manual_spend_sum
    nb_kw_sales = nb_kw_auto_sales_sum + nb_kw_manual_sales_sum    
    if nb_kw_sales != 0:
        nb_kw_acos = nb_kw_spend / nb_kw_sales * 100
    else:
        nb_kw_acos = 0
    nb_pt_spend = nb_pt_auto_spend_sum + nb_pt_manual_spend_sum + nb_pt_cat_manual_spend_sum
    nb_pt_sales = nb_pt_auto_sales_sum + nb_pt_manual_sales_sum + nb_pt_cat_manual_sales_sum
    if nb_pt_sales != 0:
        nb_pt_acos = nb_pt_spend / nb_pt_sales * 100
    else:
        nb_pt_acos = 0

    kw_spend = br_kw_spend + nb_kw_spend
    kw_sales = br_kw_sales + nb_kw_sales
    if kw_sales != 0:
        kw_acos = kw_spend / kw_sales * 100
    else:
        kw_acos = 0

    pt_spend = br_pt_spend + nb_pt_spend
    pt_sales = br_pt_sales + nb_pt_sales
    if pt_sales != 0:
        pt_acos = pt_spend / pt_sales * 100
    else:
        pt_acos = 0

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
        sp_acos = sp_spend / sp_sales * 100
    else:
        sp_acos = 0
    
    if sb_sales != 0:
        sb_acos = sb_spend / sb_sales * 100
    else:
        sb_acos = 0

    if sbv_sales != 0:
        sbv_acos = sbv_spend / sbv_sales * 100
    else:
        sbv_acos = 0

    if sd_sales != 0:
        sd_acos = sd_spend / sd_sales * 100
    else:
        sd_acos = 0


    

    if sp_spend > 0 or sb_spend > 0 or sbv_spend > 0 or sd_spend > 0:
        ad_types_spend_sum = sp_spend + sb_spend + sbv_spend + sd_spend
        ad_types_sales_sum = sp_sales + sb_sales + sbv_sales + sd_sales
        sp_spend_percentage = sp_spend / ad_types_spend_sum * 100
        sb_spend_percentage = sb_spend / ad_types_spend_sum * 100
        sbv_spend_percentage = sbv_spend / ad_types_spend_sum * 100
        sd_spend_percentage = sd_spend / ad_types_spend_sum * 100
        sp_sales_percentage = sp_sales / ad_types_sales_sum * 100
        sb_sales_percentage = sb_sales / ad_types_sales_sum * 100
        sbv_sales_percentage = sbv_sales / ad_types_sales_sum * 100
        sd_sales_percentage = sd_sales / ad_types_sales_sum * 100
    else:
        sp_spend_percentage = 0
        sb_spend_percentage = 0
        sbv_spend_percentage = 0
        sd_spend_percentage = 0
        sp_sales_percentage = 0
        sb_sales_percentage = 0
        sbv_sales_percentage = 0
        sd_sales_percentage = 0


    kw_spend_percentage = kw_spend / sum_spend * 100
    kw_sales_percentage = kw_sales / sum_sales * 100
    pt_spend_percentage = pt_spend / sum_spend * 100
    pt_sales_percentage = pt_sales / sum_sales * 100

    br_spend_percentage = br_total_spend / sum_spend * 100
    br_sales_percentage = br_total_sales / sum_sales * 100
    nb_spend_percentage = nb_total_spend / sum_spend * 100
    nb_sales_percentage = nb_total_sales / sum_sales * 100

    br_kw_spend_percentage = br_kw_spend / sum_spend * 100
    br_kw_sales_percentage = br_kw_sales / sum_sales * 100
    nb_kw_spend_percentage = nb_kw_spend / sum_spend * 100
    nb_kw_sales_percentage = nb_kw_sales / sum_sales * 100
    br_pt_spend_percentage = br_pt_spend / sum_spend * 100
    br_pt_sales_percentage = br_pt_sales / sum_sales * 100
    nb_pt_spend_percentage = nb_pt_spend / sum_spend * 100
    nb_pt_sales_percentage = nb_pt_sales / sum_sales * 100
    


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
    
    with open(ad_report_file_path, "w") as report:
        report = csv.writer(report)
        report.writerow(row1)
        report.writerow(row2)
        report.writerow(row3)
        report.writerow(row4)
        report.writerow(row5)
        report.writerow(row6)
        report.writerow(row7)
        report.writerow(row8)
        report.writerow(row9)
        report.writerow(row10)
        report.writerow(row11)
        report.writerow(row12)
        report.writerow(row13)
        report.writerow(row14)
        report.writerow(row15)
        report.writerow(row16)
        report.writerow(row17)
        report.writerow(row18)
        report.writerow(row19)
        report.writerow(row20)
        report.writerow(row21)
        report.writerow(row22)
        report.writerow(row23)
        report.writerow(row24)

        
    subprocess.call(["open", ad_report_file_path])
    # open("ad_report.csv")  

if __name__ == "__main__":
    main()

