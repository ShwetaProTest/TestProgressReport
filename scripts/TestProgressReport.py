########################################################################################################
#This script will visually depict the amount of work that needs to be done to complete it with the provided
# projected milestone and the time needed to do so.
#
####################################################################################################

#Import Modules
import os
from datetime import datetime
import datetime as dt
from os.path import join
import tkinter as tk
from tkinter import *
from tkinter import ttk
import numpy as np
import openpyxl
import pandas as pd
import tkcalendar
import xlrd as xlrd
import xlsxwriter
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from dateutil.relativedelta import relativedelta
from matplotlib import pyplot as plt, ticker
import matplotlib.dates as mdates
import matplotlib as mpl
from matplotlib.ticker import MultipleLocator
from numpy import isnan
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from pandas import DataFrame, date_range
from tkcalendar import *
import glob
from pathlib import Path
import xlwt
import warnings
import logging
from logging.handlers import TimedRotatingFileHandler
from TestStatusReport import xmlProcessor

logging.getLogger('matplotlib').setLevel(logging.WARNING)
warnings.filterwarnings("ignore")

# Log Details
LOG_FILENAME = datetime.now().strftime('../log/logfile_%d-%m-%y_%H%M%S.log')
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(filename=LOG_FILENAME, level=logging.INFO)

class burndown():
    def read_file(self):
        excel_file = pd.read_excel(self.filename, sheet_name='Worksheet', usecols="B", header=1, nrows=0)
        build_name = excel_file.columns.values[0]
        logging.info("The Build Name of Current Execution is : " + build_name)
        self.Exceldata = pd.read_excel(self.filename, sheet_name='Worksheet', skiprows=4)
        df = pd.DataFrame(self.Exceldata)
        logging.info("######### Build Calculation Started #########")
        Test_Case = df['Test Case'].count()
        logging.info("The Total Test Case Count for build " + build_name + " is : " + str(Test_Case))
        li = []
        dfs = []
        dfss = []
        sdate = '2022-10-27'
        edate = '2023-01-30'
        logging.info('The Project Estimated Date is from : ' + sdate + ' to ' + edate + '\n')
        for i, col in enumerate(df.filter(regex=('Date.*'))):
            #Filter the date column of each build
            df[col] = pd.to_datetime(df[col]).apply(lambda x: x.date())
            li.append(df[col])
            df1 = pd.concat(li, axis=1)
            df1['Test Case'] = pd.Series(df['Test Case'])
            #Count the occurrences of date in each build, by grouping date column
            df1 = df1.groupby([col])[col].count().reset_index(name='count' + str(i))

            #Selecting the date range from " + sdate + " to " + edate + " and setting the value to null if the chosen date is absent from the source for each build
            idx = pd.date_range(start=sdate, end=edate)
            s = pd.Series(df1['count' + str(i)].values, index=df1[col])
            s.index = pd.DatetimeIndex(s.index)
            s = s.reindex(idx, fill_value=0)
            df2 = pd.DataFrame({'Date': s.index, 'count' + str(i): s.values})

            #Subtract the Total number of test case with the count of execution on each date
            df2['Build' + str(i)] = Test_Case - df2['count' + str(i)]
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
        #logging.info('The Total Build Columns are : ' + df9.columns.values)
        logging.info('######## Build Calculations are completed ##########' + "\n")

        ############################################################################################

        logging.info("Plot Display ---------")
        fig = plt.figure(figsize=(15, 8))
        ax = fig.add_subplot(111)

        ############################################################################################

        logging.info("Set up the plot to show the Build Count value for each day.")
        for name, series in df9.set_index('Date').items():
            ax.plot(series.index, series.values, marker='*', linestyle='-', ms=5, label=name, zorder=1)
            values = series.values
            for i, txt in enumerate(values):
                val = txt.astype(np.int64)
                ax.annotate(val, (series.index[i], values[i]), fontsize=7, weight="bold", textcoords="offset points",xytext=(-2, 5))

        ###########################################################################################

        logging.info("The milestone configuration for project estimation")
        cdate = [datetime.today().strftime('%Y-%m-%d')]
        dates = [sdate, edate]
        x1 = [datetime.strptime(x, '%Y-%m-%d') for x in dates]
        y1 = [Test_Case, 0]

        c1 = dt.date.today() + relativedelta(days=1)
        plt.plot_date(x1, y1, linestyle='dashed', marker='', color="gray", zorder=0, label='Test Estimation')
        plt.axvspan(x1[0], x1[0] + relativedelta(days=1), alpha=0.4, color='green', label='Milestone Start Date')
        plt.axvspan(x1[1], x1[1] + relativedelta(days=1), alpha=0.4, color='red', label='Milestone End Date')
        plt.axvspan(dt.date.today(), c1, alpha=0.2, color='blue', label='Current Day')

        plt.annotate(sdate, xy=(x1[0], y1[0]), fontsize=10, textcoords='offset points', xytext=(-20, 22), ha='left',va='top', transform=plt.gca().transAxes)
        plt.annotate(edate, xy=(x1[1], y1[0]), fontsize=10, textcoords='offset points', xytext=(-20, 20), ha='left',va='top', transform=plt.gca().transAxes)
        plt.annotate(cdate[0], xy=(dt.date.today(), Test_Case), fontsize=10, textcoords='offset points',xytext=(-20, 20), ha='left', va='top', transform=plt.gca().transAxes)

        ##############################################################################################

        logging.info("Configuration to emphasize the weekend")
        xmin, xmax = ax.get_xlim()
        days = np.arange(np.floor(xmin), np.ceil(xmax) + 2)
        weekends = [(dt.weekday() >= 5) | (dt.weekday() == 0) for dt in mdates.num2date(days)]
        ax.fill_between(days, *ax.get_ylim(), where=weekends, facecolor='k', alpha=.12)

        #############################################################################################

        logging.info("configuration of the plot's x- and y-axes, labels, titles, colors, and width" + "\n")
        ax.get_xaxis().set_minor_locator(mpl.ticker.AutoMinorLocator())
        ax.get_yaxis().set_minor_locator(mpl.ticker.AutoMinorLocator())
        ax.grid(b=True, which='major', color='k', alpha=.12, linewidth=1.0)
        ax.grid(b=True, which='minor', color='k', alpha=.1, linewidth=0.5)
        plt.xlabel("Date")
        plt.ylabel("Test case Execution Count")
        plt.title(build_name + " - Burndown Chart", fontdict={'fontsize': 18, 'fontweight': 'medium', })
        params = {'legend.fontsize': 'small',
                  'axes.labelsize': 5,
                  'axes.titlesize': 'x-large'}
        plt.rcParams.update(params)
        ax.legend(loc='lower left')
        plt.gcf().autofmt_xdate()
        fig.tight_layout()

        #################################################################################################

        logging.info("Specifing the output file directory to store the executed files")
        fname = self.filename.split('/')[2]
        output_filename = fname.split('.')[0]
        output = "../Reports/"
        logging.info('The Output Files are generated in : ', output)
        df10 = df9[::-1]
        remove = lambda x: df10[x].duplicated(keep='first')
        df10.loc[remove(df10.columns[:1])] = ''
        df11 = df10[::-1]
        df11.to_excel(output + output_filename + '.xlsx', index=False)
        logging.info('The Output Data File Name is : ', output + output_filename + '.xlsx')
        logging.info('The Output Image File Name is : ', output + output_filename + '.png')
        plt.savefig(output + output_filename, dpi=150)

        wb = openpyxl.load_workbook(output + output_filename + '.xlsx')
        ws = wb.worksheets[0]
        img = openpyxl.drawing.image.Image(output + output_filename + '.png')
        img.anchor = 'k1' #png will get plot in K1 cell of output file
        ws.add_image(img)
        wb.save(output + output_filename + '.xlsx')
        logging.info('The Result Data & Image File Name is : ', output + 'Result_' + output_filename + '.xlsx')
        logging.info('########## Output File Created! Burndown chart Completed.. ##################' + "\n")

##################################################################################################
    # def log_remove(self):
        # #LOG_FILENAME = glob('..log/.log', recursive=True)
        # for file in os.listdir('../log/'):
        #     if file.endswith('.log'):
        #         os.remove('../log/' + file)

    def main(self):
        logging.info("Execution of the main function began" + '\n')
        self.directory = '../Input/'
        logger = logging.getLogger(__name__)
        logger = logging.getLogger('tipper')
        logger.setLevel(logging.INFO)
        pil_logger = logging.getLogger('PIL')
        pil_logger.setLevel(logging.INFO)
        logging.getLogger('requests').setLevel(logging.DEBUG)
        for filename in os.listdir(self.directory):
            if filename.startswith("resultsTC_*") and filename.endswith(".xlsx") or filename.endswith(".XLSX") or filename.endswith(".xls"):
                self.filename = os.path.join(self.directory, filename)
                logging.info('The Process Started Executing for the Filename : ' + self.filename)
                self.read_file()
                logging.info('The Process Completed Executing for the Filename : ' + self.filename + '\n')
                logging.info("The Process Started Executing to create Test status Reports")
                xml_processor = xmlProcessor()
                xml_processor.xml_xl(self.directory)
                xml_processor.build_calc(self.directory,self.Exceldata)
                xml_processor.test_report()
                logging.info("Test status Report Creation Completed..!")
                plt.show()


val = burndown()
val.main()
