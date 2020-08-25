# -*- coding: utf-8 -*-
"""
Created by : Nikunj Patel
Credits : Corey Schafer. Thanks Corey for making my life easier, with your fantastic python videos.
Tutorial videos by Corey Schafer : https://www.youtube.com/user/schafer5/search?query=regex

"""
from tkinter import messagebox 
from datetime import datetime
import os
dir_path = os.path.dirname(os.path.realpath(__file__))


#Error handling the Tkinter message box is imported before trying to catch an error to show error message on the messagebox.
try:
    from tkinter import Tk
    Tk().withdraw()
    from tkinter.ttk import *
    from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
    import csv
    import sys
    import re
    import time
    import openpyxl
    
    
    messagebox.showinfo("INFO", '''
    -------------Signal List Generator-------------
    
    This script is designed to create and compile a signle csv file from the multiple template files.
    There are three files used in this script
    1) Parafile which will defined all the variable parameter to be used in the template files.
    2) Template files are the individual files where we define individual template content by passing variable parameter from parafile.
    3) The GeneratedFile file is created by merging all the template file in which all the variable parameter are replaced with their values.
    
    ------------------------------------------------
    
    Made by : Nikunj Patel
    Version : 1.0
    Date : 20-April-2020
    Made with love in Python3.

    ''')

    # Initializing working directory in the dir where the this python script is located.
    # workingdir = os.getcwd()
    # os.chdir(workingdir)
    
    #Selecting parameter csv file where all the variable parameters are saved.
    parafile = askopenfilename(initialdir = dir_path ,title = "Select parameter csv file", filetypes=[('CSV file', '*.csv')])
    
    # Selecting the template directory
    templatedir = askdirectory(initialdir = dir_path , title = "Select directory where template files are located")

    #Prompt to name and save the generated file.
    GeneratedFile = asksaveasfilename(initialdir = dir_path ,title = "Save to signal list file", filetypes=[('CSV file', '*.csv')])

    #Open the generated file and trunct it to 0KB so that we can add new lines. If the file already exit this will help us to clean it.
    f = open(GeneratedFile, "w")
    f.truncate()
    f.close()
    
    starttime = time.process_time()

    with open(parafile, 'r') as paralist:
        
        #Reading parameter csv file and converting it to a dictionary
        parameter = csv.DictReader(paralist)
        
        
        for line in parameter:
            # Grabing the first line i.e all the names of parameters used. The parameters are called fieldnames here.
            fieldnames = line.keys()
            
            #Looking for FileName parameter in the parameter csv file and look for those files in the template directory folder.
            file_name = os.path.join(templatedir,line['FileName'])
            with open(file_name, 'r') as template:
                
                tpl_content = csv.reader(template)
                
                with open(GeneratedFile, 'a', newline='') as new_file:
                    csv_writer = csv.writer(new_file)                
                    
                    # For each line present in the template file. Which will be a list.
                    for eachlines in tpl_content:
                        #For each item in the list.
                        for item in eachlines:
                            # For each field in the fieldnames
                            for field in fieldnames:
                                #If field is equal to item then remove the field and replace it with key of dict.
                                #This if loop will check only direct variable parameter define in the template file.
                                if field == item:
                                    position = eachlines.index(item)
                                    eachlines.pop(position)
                                    # Replacing the item with key value from the the line dictionary
                                    eachlines.insert(position, line.get(field))
                                    
                                #This elif loop will check if the variable parameter are define along with regular cell values. In this loop we will take the value of the cell and only replace the variable parameter with the field key value from the line dictionary.
                                elif field != item :
                                    pattern = re.compile(field)
                                    matches = pattern.finditer(item)
                                    for match in matches:
                                        position = eachlines.index(item)
                                        substitute = re.sub(field, line.get(field), item)
                                        item = substitute                
                                        eachlines.pop(position)
                                        eachlines.insert(position, substitute)      
                                                  
                        csv_writer.writerow(eachlines)
                    #Writing a blank line between each template file.
                    csv_writer.writerow('\n')
                    print('Processing ' + os.path.basename(file_name)+ ' for generating your output file.')
                    
                    
        # print(GeneratedFile + ' file is generated !')
        
    elapsedtime = time.process_time() - starttime
    
    messagebox.showinfo("INFO", GeneratedFile + ' file is generated in ' + str(round(elapsedtime,3)) + ' seconds !!')
except Exception as error:
    messagebox.showerror("Error", error)    
    err = str(error)
    with open (os.path.join(dir_path,'Log.txt'), 'a') as log_file:
    # Writing errors in a log file.
        log_file.write(str(datetime.now()) + '\t' + err + '\n')
