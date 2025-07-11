# -*- coding: utf-8 -*-
"""
Created on Mon Jan 22 10:52:59 2024

@author: kurra
"""

import os
import shutil
import pandas as pd
import time
from datetime import datetime
import numpy as np
import logging

# Try to import tkinter for local GUI use
try:
    from tkinter import messagebox
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

def show_info(title, msg):
    if TKINTER_AVAILABLE and os.environ.get("FLASK_ENV") != "production":
        messagebox.showinfo(title, msg)
    else:
        print(f"INFO: {title}: {msg}")

def show_warning(title, msg):
    if TKINTER_AVAILABLE and os.environ.get("FLASK_ENV") != "production":
        messagebox.showwarning(title, msg)
    else:
        print(f"WARNING: {title}: {msg}")

def show_error(title, msg):
    if TKINTER_AVAILABLE and os.environ.get("FLASK_ENV") != "production":
        messagebox.showerror(title, msg)
    else:
        logging.error(f"{title}: {msg}")

def create_and_move_csv(input_folder):
    # Example: show_warning("Script 1 triggered", "And first script is running successfully")
    # Iterate through each subfolder in the input folder
    if os.path.exists(input_folder):
        new_folder_path = os.path.join(input_folder, 'NewFolder')
        os.makedirs(new_folder_path, exist_ok=True)
        for folder_name in os.listdir(input_folder):
            folder_path = os.path.join(input_folder, folder_name)
            # Check if it's a subfolder and not a file
            if os.path.isdir(folder_path):
                print(folder_path)
                # Move all CSV files to the new folder
                for file_name in os.listdir(folder_path):
                    print(file_name)
                    if file_name.endswith('.csv'):
                        source_file_path = os.path.join(folder_path, file_name)
                        destination_file_path = os.path.join(new_folder_path, file_name)
                        # Move the CSV file
                        shutil.move(source_file_path, destination_file_path)
                        print(f"Moved: {file_name} from {folder_name} to NewFolder")
        process_and_merge_csv(input_folder)
        show_info("Success", "Process completed successfully.")
        return "Successfull"
    else:
        show_error("Error", "Invalid input folder.")
        return "Invalid"

def process_and_merge_csv(input_folder):
    current_date = datetime.now().strftime("%Y_%m_%d")
    # Create a new folder inside the input folder for writing information
    input_folder = input_folder+r'\NewFolder'
    info_folder_path = os.path.join(input_folder, 'MasterData')
    if not os.path.exists('Data_files'):
        os.makedirs('Data_files', exist_ok=True)
    # Create an empty DataFrame to store the merged data
    merged_data = pd.DataFrame()
    # Iterate through each CSV file in the input folder
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.csv'):
            file_path = os.path.join(input_folder, file_name)
            # Read the CSV file into a DataFrame
            current_data = pd.read_csv(file_path)
            print(current_data.head())
            #current_data = demeaner(current_data)
            # Extract information from the file name (modify this based on your file naming pattern)
            file_info = file_name.split('_')
            if len(file_info) > 6:
                print(file_info)  # Assumes a naming pattern like "info1_info2_filename.csv"
                bacteria = file_info[1]
                concentration = file_info[2]
                volume = file_info[3]
                slide = file_info[4]
                trailXs = file_info[6].split('.')
                trail=trailXs[0]
                current_data['bacteria'] = bacteria
                current_data['concentration'] = concentration
                current_data['volume'] = volume
                current_data['slide'] = slide
                current_data['trial'] = trail
            else:
                bacteria = file_info[1]
                concentration = file_info[2]
                volume = file_info[3]
                slide = 0
                trailXs = file_info[5].split('.')
                trail=trailXs[0]
                current_data['bacteria'] = bacteria
                current_data['concentration'] = concentration
                current_data['volume'] = volume
                current_data['slide'] = slide
                current_data['trial'] = trail
            merged_data = pd.concat([merged_data, current_data], ignore_index=True)
    # Write the merged data to a new CSV file
    merged_file_path =  'Data_files/raw_data/masterData_'+current_date+'.csv'
    merged_data.to_csv(merged_file_path, index=False)

def demeaner(data):
    df=data
    df2=pd.read_csv(r'Data_files\raw_data\EmptyData_2024_02_29.csv')
    df2=avg_calculator(df2)
    print(df2.shape)
    count=0
    for i in [10,20,30,40,50,60]:
        mean1 = df2['Xpos'].iloc[count:(i*1000)].mean()
        mean2 = df2['Ypos'].iloc[count:(i*1000)].mean()
        df['Xpos'].iloc[count:(i*1000)]=df['Xpos'].iloc[count:(i*1000)]-mean1
        df['Ypos'].iloc[count:(i*1000)]=df['Ypos'].iloc[count:(i*1000)]-mean2
        count=i*1000
    return df

def avg_calculator(data):
    columns = ['Xpos','Ypos','bacteria']
    newD = pd.DataFrame(columns=columns)
    for i in data['bacteria'].unique():
        columns = ['Xpos','Ypos']
        empty_dataset = pd.DataFrame(columns=columns)
        dataX=data[data['bacteria']==i]
        data1 = dataX[dataX["trail"]=="T1"]
        data2 = dataX[dataX["trail"]=="T2"]
        data3 = dataX[dataX["trail"]=="T3"]
        empty_dataset['Xpos']=(np.array(data1['Xpos'])+np.array(data2['Xpos'])+np.array(data3['Xpos']))/3
        empty_dataset['Ypos']=(np.array(data1['Ypos'])+np.array(data2['Ypos'])+np.array(data3['Ypos']))/3
        empty_dataset['bacteria'] = i
        newD=pd.concat([newD,empty_dataset])
    return newD

def pruner(data):
    for col in ['Xpos','Ypos','Pow']:
        means = data[col].mean()
        stds = data[col].std()
        lower_bounds = means - 3 * stds
        upper_bounds = means + 3 * stds
        print(col)
        data[col] = np.where(data[col].between(lower_bounds, upper_bounds), data[col], np.nan)
    return data
