#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Apr  5 15:19:09 2021

@author: jordanlubbers
         jlubbers@usgs.gov

This script will take a folder of .csv files from the output
associated with a either a ThermoFisher iCAP RQ or Agilent 8900 QQQ ICP-MS,
strip the metadata at the top,and place all of the raw counts per second data
in one spreadsheet with columns that denote the individual sample that 
each row (e.g., cycle through the mass range) belongs to and its associated
timestamp.

It uses tkinter to do this in the form of a GUI where the user only has to 
input the name of the output file in the text input area, select the button 
'thermo' or 'agilent' in the "Quad Type" area, press the "preprocess data" button,
and go find their folder with all the individual, raw .csv files.

This spreadsheet is now ready for use in LaserTRAM-DB.



"""

import glob as glob
import os
import re
import tkinter as tk
from tkinter import filedialog, ttk

import numpy as np
import pandas as pd
from dateutil.parser import parse

myfont = "Calibri"

def make_LTspot_ready(file):

        # import data
        # remove the top rows. Double check that your header is the specified
        # amount of rows to be skipped in 'skiprows' argument
        data = pd.read_csv(file, skiprows=13)
        # drop empty column at the end
        data.drop(data.columns[len(data.columns) - 1], axis=1, inplace=True)

        # remove dwell time row beneath header row
        LT_ready = data.dropna()

        return LT_ready



def extract_agilent_metadata(file):
        # import data
        # extract sample name
        # extract timestamp
        # extract data and make headers ready for lasertram
        
        df = pd.read_csv(file, sep = '\t', header = None)
        
        sample = df.iloc[0,0].split('\\')[-1].split('.')[0].replace('_','-')
        
        timestamp = parse(df.iloc[2,0].split(' ')[7] + ' ' + df.iloc[2,0].split(' ')[8])
        
        data = pd.DataFrame([sub.split(",") for sub in df.iloc[3:-1,0]])
        
        header = data.iloc[0,:]
        data = data[1:]
        data.columns = header
        newcols = []
        for s in data.columns.tolist():
            l = re.findall('(\d+|[A-Za-z]+)', s)
            if 'Time' in l:
                newcols.append(l[0])
            else:

                newcols.append(l[1]+l[0])
        data.columns = newcols
        
        return [timestamp, file, sample, data]


def preprocess_data():

    global step
    global step2

    filename = filename_entry.get()

    folder_selected = filedialog.askdirectory()

    # the filepath of the folder with the proper tail to grab all the .csv files from it
    # using glob
    files_in_folder = folder_selected
    
    if quad_text.get() == 'thermo':
        # make_LTspot_file(files_in_folder, filename)

        # all of your file paths from that folder as a list
        infiles = glob.glob("{}/*.csv".format(files_in_folder))
        #step for progress bar
        # finds the sample name in cell A1 in your csv file and saves it
        # to a list
        samples = []
        timestamps = []
        for i in range(0, len(infiles)):
            # gets the top row in your csv and turns it into a pandas series
            top = pd.read_csv(infiles[i], nrows=0)

            # since it is only 1 long it is also the column name
            # extract that as a list
            sample = list(top.columns)

            # turn that list value to a string
            sample = str(sample[0])

            # because its a string it can be split
            # split at : removes the time stamp
            sample = sample.split(":")[0]

            # .strip() removes leading and trailing spaces
            sample = sample.strip()

            # replace middle spaces with _ because spaces are bad
            nospace = sample.replace(" ", "_")

            # append to list
            samples.append(nospace)

            # get the timestamp by splitting the string by the previously
            # designated sample. Also drops the colon in front of the date
            timestamp = top.columns.tolist()[0].split(sample)[1:][0][1:]

            # use dateutil.parser function 'parse' to turn that string into
            # a useable date
            timestamps.append(parse(timestamp))

            step.set(100*i/len(infiles))
            step_text.set(f"retrieve metadata: {i+1}/{len(infiles)}")
            root.update_idletasks()

        # go through and make a list of lists for each file that has its metadata
        # in the form [timestamp,samplename,filepath]
        metadata = []
        for sample, file, timestamp in zip(samples, infiles, timestamps):
            metadata.append([timestamp, sample, file])

        # because the timestamp is first we can now sort this list by time
        # put the ordered samplenames and filepaths in their own list
        ordered_samples = []
        ordered_files = []
        ordered_timestamps = []
        for data in sorted(metadata):
            ordered_timestamps.append(data[0])
            ordered_samples.append(data[1])
            ordered_files.append(data[2])

        # now when you strap them together to prep for LaserTRAM spot analyses, things will be
        # in the correct order

        suffix = "_LT_ready.xlsx"
        outpath = os.path.dirname(infiles[0]) + "/" + filename + suffix

        # concatenate all your data

        # concatenate all your data
        all_data = pd.DataFrame()
        for file, sample, timestamp, i in zip(
            ordered_files, ordered_samples, ordered_timestamps, range(len(infiles))
        ):

            sample_data = make_LTspot_ready(file)
            # insert a column with the header 'sample' and sample name in every row
            sample_data.insert(0, "SampleLabel", sample)
            sample_data.insert(0, "timestamp", timestamp)
            # add blank row at the end
            sample_data.loc[sample_data.iloc[-1].name + 1, :] = np.nan

            # append current iteration
            all_data = pd.concat([all_data,sample_data])
            step2.set(100*i/len(infiles))
            step2_text.set(f"combining files: {i+1}/{len(infiles)}")
            if i == len(infiles) - 1:
                 saving_label = tk.Label(root, text = "Saving spreadsheet...")
                 canvas1.create_window(300,250, window = saving_label)


            root.update_idletasks()
            

        # change time units from seconds to milliseconds
        all_data["Time"] = all_data["Time"].multiply(other=1000)

        # blank dataframe for the blank sheet that is required
        df_blank = pd.DataFrame()

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        # {'strings_to_numbers': True} remvoes 'numbers saved as text error in excel'
        writer = pd.ExcelWriter(
            outpath,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_numbers": True}},
        )

        # Convert the dataframe to an XlsxWriter Excel object.
        all_data.to_excel(writer, sheet_name="Buffer", index=False)
        df_blank.to_excel(writer, sheet_name="Sheet1", index=False)

        # access XlsxWriter objects from the dataframe writer object
        workbook = writer.book
        buffer = writer.sheets["Buffer"]
        sheet1 = writer.sheets["Sheet1"]
        # Close the Pandas Excel writer and output the Excel file.
        writer.close()
        saving_label.configure(text = "Saving spreadsheet...Done!")
        filepath_label = tk.Label(
            root,
            text="Success!! the following was saved in the same folder as the raw data:",
            wraplength=175,
            font=myfont,
        )
        canvas1.create_window(
            300,
            320,
            window=filepath_label,
        )

        filepath_label2 = tk.Label(
            root,
            text="{}{}".format(filename_entry.get(), suffix),
            fg="red",
            font="{} 12 bold".format(myfont),
        )
        canvas1.create_window(300, 380, window=filepath_label2)
        
    elif quad_text.get() == 'agilent':
        # make_LTspot_ready_agilent(files_in_folder,filename)
        infiles = glob.glob('{}/*.csv'.format(files_in_folder))

        metadata = []

        for file,i in zip(infiles,range(len(infiles))):
            # file, timestamp, sample, data
            metadata.append(extract_agilent_metadata(file))
            step.set(100*i / len(infiles))
            step_text.set(f"retrieve metadata: {i+1}/{len(infiles)}")
            root.update_idletasks()
            
        # sort the metadata by timestamp
        # take sorted metadata make dataframe with SampleLabel column and 
        # append to blank all_data dataframe to make one large dataframe
        # with properly ordered data


        all_data = pd.DataFrame()
        suffix = 'LT_ready.xlsx'
        outpath = '{}\{}_{}'.format(os.path.dirname(infiles[0]),filename,suffix)


        for data,i in zip(sorted(metadata), range(len(infiles))):

            sample_data = data[3]
            sample_data.insert(0,'SampleLabel',data[2])
            sample_data.insert(0,'timestamp',data[0])



            all_data = pd.concat([all_data,sample_data])
            step2.set(100*i / len(infiles))
            step2_text.set(f"combining files: {i+1}/{len(infiles)}")
            if i == len(infiles) - 1:
                 saving_label = tk.Label(root, text = "Saving spreadsheet...")
                 canvas1.create_window(300,250, window = saving_label)


            root.update_idletasks()
        
        all_data['Time'] = all_data['Time'].astype('float64')*1000
            
            
        #blank dataframe for the blank sheet that is required
        df_blank = pd.DataFrame()

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        #{'strings_to_numbers': True} remvoes 'numbers saved as text error in excel'
        writer = pd.ExcelWriter(outpath, engine='xlsxwriter',
        engine_kwargs={'options': {'strings_to_numbers': True}}
        )

        # Convert the dataframe to an XlsxWriter Excel object.
        all_data.to_excel(writer, sheet_name='Buffer',index = False)
        df_blank.to_excel(writer,sheet_name = 'Sheet1',index = False)

        #access XlsxWriter objects from the dataframe writer object
        workbook  = writer.book
        buffer = writer.sheets['Buffer']
        sheet1 = writer.sheets['Sheet1']
        #Close the Pandas Excel writer and output the Excel file.
        writer.close() 
        saving_label.configure(text = "Saving spreadsheet...Done!")
        filepath_label = tk.Label(
            root,
            text="Success!! the following was saved in the same folder as the raw data:",
            wraplength=175,
            font=myfont,
        )
        canvas1.create_window(
            300,
            300,
            window=filepath_label,
        )

        filepath_label2 = tk.Label(
            root,
            text="{}{}".format(filename_entry.get(), suffix),
            fg="red",
            font="{} 12 bold".format(myfont),
        )
        canvas1.create_window(300, 340, window=filepath_label2)



root = tk.Tk()
root.title("Make data LaserTRAM ready!")
step = tk.IntVar()
step_text = tk.StringVar()
step2 = tk.IntVar()
step2_text = tk.StringVar()

quad_text = tk.StringVar()
quad_text.set(None)

canvas1 = tk.Canvas(root, width=600, height=400, relief="raised")

filename_entry = tk.Entry(root, width=30, font=myfont)
filename_entry.insert(0, "Name your file here!")
filename_label = tk.Label(root, text="Output filename: ", font=myfont)



thermo_button = tk.Radiobutton(root, text = 'Thermo', value = 'thermo',variable = quad_text)
canvas1.create_window(250, 100, window=thermo_button)
agilent_button = tk.Radiobutton(root, text = 'Agilent', value = 'agilent', variable = quad_text)
canvas1.create_window(350, 100, window=agilent_button)

# quad_entry = tk.Entry(root, width=20, font=myfont)
# quad_entry.insert(0, "'thermo' or 'agilent'?")
quad_label = tk.Label(root, text="Quad type: ", font=myfont)





canvas1.create_window(300, 40, window=filename_entry)
canvas1.create_window(300, 10, window=filename_label)

# canvas1.create_window(300, 100, window=quad_entry)
canvas1.create_window(300, 70, window=quad_label)

pb = ttk.Progressbar(root,orient='horizontal',variable = step)
canvas1.create_window(300, 200,window = pb)

pb_label = tk.Label(root,textvariable=step_text)
canvas1.create_window(170,200, window = pb_label)

pb2 = ttk.Progressbar(root,orient='horizontal',variable = step2)
canvas1.create_window(300, 225,window = pb2)

pb2_label = tk.Label(root,textvariable=step2_text)
canvas1.create_window(170,225, window = pb2_label)

process_button = tk.Button(
    text="preprocess data",
    command=preprocess_data,
    height=2,
    fg="green",
    font="{} 12 bold".format(myfont),
)

canvas1.create_window(300,160, window=process_button)

canvas1.pack()

root.mainloop()



