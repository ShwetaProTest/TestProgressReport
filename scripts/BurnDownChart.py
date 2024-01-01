# Reads an Excel file, extracts relevant data, calculates burndown statistics, and plots a graph representing the progress of software development by displaying the
# number of builds remaining each day, given the start and end dates of the project.
# Developer: RHM

print("The Process Starts Executing, Please Wait!..")

import subprocess
import sys
import os
import pandas as pd
import numpy as np
from matplotlib import pyplot as plt, ticker
from datetime import datetime
import datetime as dt
from dateutil.relativedelta import relativedelta
import matplotlib.dates as mdates
import matplotlib as mpl
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from openpyxl import Workbook
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
import time
import img2pdf
import shutil
import logging
import warnings
warnings.filterwarnings("ignore", message="Ignoring invalid distribution.*")

class burndown():
    #Function initializes a GUI application using Tkinter to prompt the user to select a start and end date using Calendar widgets
    def __init__(self):
        print(f"\nscript1:")
        print(f"Please choose the desired start and end dates from the GUI window & Submit. To choose before 90 and after 30 days from the current date, click use default.")
        # Initialize instance variables for start and end dates, and today's date
        self.start_date = None
        self.end_date = None
        self.today = dt.date.today()

        # Calculate the start and end dates
        start_date = self.today- relativedelta(months=1)
        end_date = self.today

        # Create the root window for the GUI application
        self.root = tk.Tk()
        self.root.title("Select Dates")

        # Create a label and Calendar widget for selecting the start date
        self.start_label = ttk.Label(self.root, text="Start Date:")
        self.start_label.pack()
        self.start_cal = Calendar(self.root, selectmode='day',
                                  year=start_date.year,
                                  month=start_date.month,
                                  day=start_date.day)
        self.start_cal.pack()

        # Create a label and Calendar widget for selecting the end date
        self.end_label = ttk.Label(self.root, text="End Date:")
        self.end_label.pack()
        self.end_cal = Calendar(self.root, selectmode='day',
                                year=end_date.year,
                                month=end_date.month,
                                day=end_date.day,
                                selectforeground='white',
                                selectbackground='green')
        self.end_cal.pack()

        # Create a submit button to submit the selected dates and a default button to use today's date as the default
        self.submit_button = ttk.Button(self.root, text="Submit", command=self.submit)
        self.submit_button.pack()
        self.default_button = ttk.Button(self.root, text="Use Default", command=self.use_default)
        self.default_button.pack()

        # Start the GUI application event loop
        self.root.mainloop()

    def submit(self):
        # Get the selected start and end dates from the Calendar widgets
        self.start_date = self.start_cal.selection_get()
        self.end_date = self.end_cal.selection_get()

        # If a start date was not selected, use 90 days before today as the default start date
        if not self.start_date:
            self.start_date = self.today - dt.timedelta(days=180)

        # If an end date was not selected, use 30 days after today as the default end date
        if not self.end_date:
            self.end_date = self.today + dt.timedelta(days=30)

        # Print the selected start and end dates and close the GUI window
        print(f"The user selected Start Date is : {self.start_date}")
        print(f"The user selected End Date is : {self.end_date}")
        self.root.destroy()

    def use_default(self):
        # Set the start date to 90 days before today and the end date to 30 days after today
        self.start_cal.selection_set(self.today - dt.timedelta(days=180))
        self.end_cal.selection_set(self.today + dt.timedelta(days=30))

        # Submit the default dates and close the GUI window
        self.submit()

    def read_file(self,burndown_report,report_pdf):
        excel_file = pd.read_excel(self.filename, sheet_name='Worksheet', usecols="B", header=1, nrows=0)
        build_name = excel_file.columns.values[0]
        logging.info(f"The Build Name of Current Execution is : {build_name}")
        self.Exceldata = pd.read_excel(self.filename, sheet_name='Worksheet', skiprows=4)
        df = pd.DataFrame(self.Exceldata)
        logging.info("######### Build Calculation Started #########")
        # Count the number of test cases in the DataFrame and log the result
        Test_Case = df['Test Case'].count()
        logging.info("The Total Test Case Count for build " + build_name + " is : " + str(Test_Case))
        # Create some empty lists to be used later
        li = []
        dfs = []
        dfss = []

        logging.info(f'The Project Estimated Date is from: {self.start_date} to {self.end_date}\n')
        for i, col in enumerate(df.filter(regex=('Date.*'))):
            # Filter the date column of each build and convert to date only
            df[col] = pd.to_datetime(df[col]).apply(lambda x: x.date())
            # Append the filtered date column to the list li
            li.append(df[col])
            # Concatenate the date columns horizontally (axis=1) into a new DataFrame
            df1 = pd.concat(li, axis=1)
            # Add the 'Test Case' column to the new DataFrame
            df1['Test Case'] = pd.Series(df['Test Case'])
            # Count the occurrences of date in each build, by grouping the date column
            df1 = df1.groupby([col])[col].count().reset_index(name=f'count{i}')

            #Selecting the date range from " + sdate + " to " + edate + " and setting the value to null if the chosen date is absent from the source for each build
            # Select the date range between `start_date` and `end_date`
            idx = pd.date_range(start=self.start_date, end=self.end_date)
            # Create a Series object `s` using the values from `df1['count' + str(i)]`
            # and the index from `df1[col]`, which represents the dates in the current build
            s = pd.Series(df1['count' + str(i)].values, index=df1[col])
            # Set the index of the `s` Series to a DatetimeIndex
            s.index = pd.DatetimeIndex(s.index)
            # Reindex the `s` Series with the `idx` index, filling any missing values with 0
            s = s.reindex(idx, fill_value=0)
            # Create a new DataFrame `df2` with the index of `s` as the 'Date' column
            # and the values of `s` as the 'count' column
            df2 = pd.DataFrame({'Date': s.index, 'count' + str(i): s.values})

            # Subtract the number of executions (`count` column) from the total number of test cases (`Test_Case`)
            df2['Build' + str(i)] = Test_Case - df2['count' + str(i)]
            # Compute the cumulative sum of the `count
            df2['Build' + str(i)] = df2['count' + str(i)].cumsum().mul(-1).add(Test_Case)

            #To prevent extension lines when plotting, remove the initial duplicate values
            remove = lambda x: df2[x].duplicated(keep='first')
            df2.loc[remove(['Build' + str(i)]), ['Build' + str(i)]] = np.nan
            dfs.append(df2)
            df3 = pd.concat(dfs, axis=1)
            df4 = df3.loc[:, ~df3.columns.duplicated()].copy()
            df5 = df4.drop(df4.filter(regex='count.*').columns, axis=1)
            df5 = df5[df5.any(axis=1)]
            df5 = df5.apply(lambda series: series.loc[:series.last_valid_index()].ffill())
            df6 = df5[::-1]
            remove = lambda x: df6[x].duplicated(keep='first')
            df6.loc[remove(['Build' + str(i)]), ['Build' + str(i)]] = np.nan
            dfss.append(df6)
            df7 = pd.concat(dfss, axis=1)
            df8 = df7.loc[:, ~df7.columns.duplicated()].copy()
            df9 = df8.apply(lambda series: series.loc[:series.last_valid_index()].ffill())

        logging.info("Filter all of the source's actual build names and replace them with dataframe column names")
        filter_col = [col for col in df if col.startswith('Build')]
        df9.columns = df9.columns[:1].tolist() + filter_col
        logging.info('The Total Build Columns are : ' + df9.columns.values)
        logging.info('######## Build Calculations are completed ##########' + "\n")

        ############################################################################################

        logging.info("Plot Display ---------")
        fig, ax = plt.subplots(figsize=(15, 8))
        #ax = fig.add_subplot(111)

        ############################################################################################

        logging.info("Set up the plot to show the Build Count value for each day.")
        last_value = None
        for name, series in df9.set_index('Date').items():
            ax.plot(series.index, series.values, marker='*', linestyle='-', ms=5, label=name, zorder=1)
            values = series.values
            for i, txt in enumerate(values):
                val = txt.astype(np.int64)
                if last_value is None or val != last_value or i == len(values)-1:
                    ax.annotate(val, (series.index[i], values[i]), fontsize=7, weight="bold", textcoords="offset points",xytext=(-2, 5))
                    last_value = val

        # for name, series in df9.set_index('Date').items():
        #     ax.plot(series.index, series.values, marker='*', linestyle='-', ms=5, label=name, zorder=1)
        #     values = series.values
        #     for i, txt in enumerate(values):
        #         val = txt.astype(np.int64)
        #         ax.annotate(val, (series.index[i], values[i]), fontsize=7, weight="bold", textcoords="offset points",xytext=(-2, 5))

        ###########################################################################################

        logging.info("The milestone configuration for project estimation")
        cdate = [dt.date.today().strftime('%Y-%m-%d')]
        dates = [self.start_date, self.end_date]
        x1 = [datetime.strptime(x.strftime('%Y-%m-%d'), '%Y-%m-%d') for x in dates]
        y1 = [Test_Case, 0]
        c1 = dt.date.today() + relativedelta(days=1)
        edate1 = dt.datetime.strptime(self.end_date.strftime('%Y-%m-%d'), '%Y-%m-%d').date()
        start_date_str = self.start_date.strftime('%Y-%m-%d')
        end_date_str = self.end_date.strftime('%Y-%m-%d')

        if self.today == edate1:
            plt.plot_date(x1, y1, linestyle='dashed', marker='', color="gray", zorder=0, label='Test Estimation')
            plt.axvspan(x1[0], x1[0] + relativedelta(days=1), alpha=0.4, color='green',
                        label='Milestone Start Date : ' + start_date_str)
            plt.axvspan(dt.date.today(), c1, alpha=0.8, color='red', label='Milestone End Date is today : ' + cdate[0])
        else:
            plt.plot_date(x1, y1, linestyle='dashed', marker='', color="gray", zorder=0, label='Test Estimation')
            plt.axvspan(x1[0], x1[0] + relativedelta(days=1), alpha=0.4, color='green',label='Milestone Start Date : ' + start_date_str)
            plt.axvspan(x1[1], x1[1] + relativedelta(days=1), alpha=0.4, color='red',label='Milestone End Date : ' + end_date_str)
            plt.axvspan(dt.date.today(), c1, alpha=0.5, color='blue', label='Today : ' + cdate[0])
            Date_Position = (Test_Case + Test_Case / 40)
            plt.annotate(cdate[0], xy=(dt.date.today(), Date_Position), fontsize=10, textcoords='offset points',
                         xytext=(10, 40), ha='left',
                         va='top', color="blue", weight="medium",
                         bbox=dict(boxstyle="round", fc=(0.0, 0.0, 1.0, 0.2), ec=(0.0, 0.0, 1.0, 0.1)),
                         transform=plt.gca().transAxes)

    ####################################################################################################################
        # Y axis configuration
        ytiks = 1 if Test_Case < 16 else 2 if Test_Case < 30 else 5 if Test_Case < 60 else 10

        Test_Case_Scale = Test_Case + int(Test_Case / 5)
        plt.yticks(np.arange(0, Test_Case_Scale, ytiks))

        # X axis configuration
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
        ax.xaxis.set_minor_locator(mdates.DayLocator())

        # Annotate important dates
        plt.annotate(self.start_date, xy=(x1[0], y1[0]), fontsize=10, textcoords='offset points', xytext=(10, 30), ha='left',
                     color="green", weight="medium",
                     bbox=dict(boxstyle="round", fc=(0.0, 0.6, 0.0, 0.2), ec=(0.0, 0.6, 0.0, 0.1)), va='top',
                     transform=plt.gca().transAxes)
        plt.annotate(self.end_date, xy=(x1[1], y1[0]), fontsize=10, textcoords='offset points', xytext=(10, 30), ha='left',
                     color="red", weight="medium",
                     bbox=dict(boxstyle="round", fc=(1.0, 0.0, 0.0, 0.2), ec=(1.0, 0.0, 0.0, 0.1)), va='top',
                     transform=plt.gca().transAxes)

    ####################################################################################################################

        logging.info("Configuration to emphasize the weekend")
        xmin, xmax = ax.get_xlim()
        days = np.arange(np.floor(xmin), np.ceil(xmax) + 2)
        weekends = [(dt.weekday() >= 5) | (dt.weekday() == 0) for dt in mdates.num2date(days)]
        ax.fill_between(days, *ax.get_ylim(), where=weekends, facecolor='k', alpha=.12)

    ####################################################################################################################

        logging.info("configuration of the plot's x- and y-axes, labels, titles, colors, and width" + "\n")
        ax.grid(axis='both', which='major', color='k', alpha=.12, linewidth=1.0)
        ax.grid(axis='both', which='minor', color='k', alpha=.1, linewidth=0.5)
        ax.set_xlabel("Date [ Weeks ]", fontweight='bold', fontstyle='oblique')
        ax.set_ylabel("Test remaining [ Test case Count per Build ]", fontweight='bold', fontstyle='oblique')
        plt.title(build_name + " - Test Progress", fontdict={'fontsize': 11, 'fontweight': 'bold'}, color='black',backgroundcolor="whitesmoke", pad=15.0)
        params = {'legend.fontsize': 'small',
                  'axes.labelsize': 8,
                  'axes.titlesize': 'x-large'}
        plt.rcParams.update(params)
        ax.tick_params(axis='y', which='major', length=3, labelsize=9)
        ax.tick_params(axis='x', which='major', length=7, labelsize=9, rotation=25)
        plt.gcf().autofmt_xdate()
        ax.legend(loc='center left', bbox_to_anchor=(0, -0.28))
        ax.set_ylim(int(-Test_Case / 15), Test_Case_Scale)
        fig.tight_layout()

    ####################################################################################################################

        print("Specifing the output file directory to store the executed files")
        #fname = self.filename.split('/')[2]
        fname=os.path.basename(self.filename)
        #output_filename = fname.split('.')[0]
        output_filename=os.path.splitext(fname)[0]
        output = burndown_report
        print('The Output Files are generated in : %s', output)
        df10 = df9[::-1]
        remove = lambda x: df10[x].duplicated(keep='first')
        df10.loc[remove(df10.columns[:1])] = ''
        df11 = df10[::-1]
        output_xlsx_path = output + output_filename + '.xlsx'
        df11.to_excel(output_xlsx_path, index=False)
        print('The Output Data File Name is : %s', output_xlsx_path)

        # Save the image as PNG
        output_image_path = output + output_filename + '.png'
        plt.savefig(output_image_path, dpi=150)
        logging.info('The Output Image File Name is: %s', output_image_path)

        # Convert the PNG image to PDF
        report_output = report_pdf
        output_pdf_path = report_output + build_name + '.pdf'
        with open(output_pdf_path, "wb") as pdf_file:
            pdf_file.write(img2pdf.convert(output_image_path))
        logging.info('The Output PDF File Name is: %s', output_pdf_path)

        # Load the XLSX file
        wb = load_workbook(output_xlsx_path)
        ws = wb.active

        # Add the image to the worksheet
        image = Image(output_image_path)
        image.anchor = 'K1'  # Place the image in cell K1
        ws.add_image(image)

        # Save the modified workbook
        wb.save(output_xlsx_path)
        logging.info('The Result Data & Image File Name is: %s', output_xlsx_path)

    #This function removes log files that are older than a certain number of days in the specified directory
    def remove_old_logs(self,log_dir, days_to_keep):
        now = time.time()
        for filename in os.listdir(log_dir):
            filepath = os.path.join(log_dir, filename)
            if os.path.isfile(filepath):
                # Get the age of the file in days
                age_days = (now - os.path.getmtime(filepath)) / (24 * 3600)
                if age_days > days_to_keep:
                    os.remove(filepath)

    #This function is the main entry point of a script, which contains the logic to process input files,generateXMLfiles,perform calculations,create test reports, and display a plot
    def run_burndown(self,input_folder,log_folder,burndown_report,report_pdf):
        # Remove old log files from the log directory
        self.remove_old_logs(log_folder, 7)  # Alter the log file removal dates from 7 to any required number
        logging.info("Execution of the main function began" + '\n')
        try:
            # Set the input directory
            self.directory = input_folder
            # Set a default value for the input file name
            self.filename = None
            # Loop through all files in the input directory
            for filename in os.listdir(self.directory):
                # Check if the file name matches the required format and file type
                if filename.startswith("resultsTC_") and (filename.endswith(".xlsx") or filename.endswith(".XLSX") or filename.endswith(".xls") or filename.endswith(".XLS")):
                    # Set the input file path
                    self.filename = os.path.join(self.directory, filename)
                    try:
                        # Read the input file
                        print(f"The Burndown process's execution has begun, and the file name is : {self.filename}")
                        self.read_file(burndown_report,report_pdf)
                        print(f"The Burndown process is completed. File Generated in : {burndown_report}")
                    except FileNotFoundError:
                        # Handle error if file not found
                        print(f"File {self.filename} not found")
                    # else:
                    #     plt.show()
            # If no matching file is found in the directory, print a message to the console
            if self.filename is None:
                print("No matching file found in directory")
        except OSError as e:
            print(f"Error accessing directory: {e}")

# val = burndown()
# val.run_burndown()