#############################################################################################################
#Step 1: Load the xml file to extract the specific element and write the extracted data to excel file
#Step 2: Read the data from the ResultsTC file and produce an individual Excel file for each build
#Step 3: Compare build with the resultsByStatus file to retrieve the matching values and the test status
#Step 4: Plot1 display the total status count
#Step 5: Plot2 display the percentage of completion
#############################################################################################################

#import modules
import os
import re
import pandas as pd
from openpyxl import load_workbook
from pathlib import Path
import itertools
from functools import reduce
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import glob
import xml.etree.ElementTree as ET
from lxml import etree
import plotly.express as px
import matplotlib.ticker as tkr
import logging
import matplotlib.ticker as mtick

################################################################################################################################
#Read the attributes and extract the relevant element from the input xml by converting xml to excel file.
class xmlProcessor():
    def xml_xl(self,directory):
        logging.info('********** xml to excel File Conversion started ********** ')
        parser = etree.XMLParser(remove_blank_text=True)
        inp_name = directory+'Input_xml.xml'
        logging.info('The Input File Locations is : '+ inp_name)
        root = ET.parse(inp_name,parser=parser).getroot()
        data=[]
        for child in root.findall('testsuite'):
            for each in child.findall('.//testcase'):
                version = each.find('.//version').text
                test_id = each.find('.//fullexternalid').text
                Test_Case = each.attrib['name']
                estimated_duration = each.find('.//estimated_exec_duration').text
                imp = each.find('.//importance').text
                for steps in each.findall('.//custom_fields'):
                    for cs in steps.findall('.//custom_field'):
                        last = cs[-1].text
                data.append({'Test_case_ID': test_id,'Test_case_title': Test_Case,'Version':version,'Estimated_exec_[min]':estimated_duration,'Total_steps':last,'Priority':imp})
        df = pd.DataFrame(data)
        S = {'1':'Low','2':'Medium','3':'High'}
        df.columns = df.columns.str.strip()
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        df['Priority'] = df['Priority'].map(S)
        df["Test_case_title"] = df["Test_case_ID"] + ':' + df["Test_case_title"]
        out_fname = directory+'xmltoxl.xlsx'
        df.to_excel(out_fname)
        logging.info('The Excel Output file generated in : ' + out_fname)
        logging.info('********** xml to excel File Conversion Completed **********' + '\n')

#########################################################################################################################
#Create new files with the filtered columns from each build by reading the Result file
    def build_calc(self,directory,Exceldata):
        logging.info("The Input location to fetch resuktsTC file is : " + directory)
        df = pd.DataFrame(Exceldata)
        df = df.filter(regex=('Test Suite|Test Case|^Build|Date.*|Execution duration'))
        self.count = df.columns[df.columns.str.contains('Build')].value_counts().sum()
        dict_of_df = {}
        split_dfs = []
        for i in range(1, self.count + 1):
            dict_of_df["df_{}".format(i)] = df.iloc[:, -1 + (3 * i):(3 * i) + 2]
            dict_of_df["df_{}".format(i)]['Test Case'] = df['Test Case']
            dict_of_df["df_{}".format(i)]['Test Suite'] = df['Test Suite']
            f_n = dict_of_df["df_{}".format(i)].columns[0].replace(' ', '_')
            f_name = f_n.replace(':', '_') + '.xlsx'
            logging.info("The Output file names are : " + f_name)
            dict_of_df["df_{}".format(i)].to_excel('../Input/' + f_name, index=False)
        logging.info("The filtered output files of each build is created in : " + directory+ '\n')

##########################################################################################################
#Obtain the values of Passed, Partially passed, failed, blocked, clarification, not run, and total steps for each build by comparing the status file with each build file.
    def test_report(self):
        logging.info('Test report generation started')
        p = Path('../Input/')
        Build_files = [x for x in p.glob("**/*.xlsx") if x.name.__contains__("Build_")]
        # for filename in os.listdir(p):
        #     if filename.startswith('resultsByStatus_') and filename.endswith(".xls"):
        #         df = pd.read_excel('../Input/'+filename)
        #         fl_name = filename.split('.',1)[0]
        #         df.to_excel('../Input/'+ fl_name + ".xlsx",index=False)
        #         os.remove('../Input/'+filename)
        Status_files = [x for x in p.glob("**/*.xlsx") if x.name.__contains__("resultsByStatus_")]
        df_files = []
        passed = []
        partially_passed = []
        failed = []
        blocked = []
        clarification = []
        not_run = []
        logging.info("Iterate over the status and build files to extract the status of each test case present in the build")
        for files in Build_files:
            for file in Status_files:
                df = pd.read_excel(files)
                df_builds = [col for col in df if col.startswith('Build')]
                df.loc[:,'build'] = df_builds[0]
                df['Build'] = df['build'].str.split(' ',1).str[1]
                df = df.drop('build', axis=1)
                data = load_workbook(file)
                ws = data.active
                val = data['Worksheet']['A1'].value
                status = val.rsplit(' ',2)[0]
                df1 = pd.read_excel(file, skiprows=6)
                # Condition to fetch specific columns if string matches with status
                if 'Partially Passed' in status:
                    df1 = df1[['Test Case','Build','Number of Partially Passed steps']]
                    df_pp = df1.merge(df, on=['Test Case', 'Build'], how="right")
                elif 'Passed' in status:
                    df1 = df1[['Test Case', 'Build','Number of Passed steps']]
                    df_p = df1.merge(df, on=['Test Case', 'Build'], how="right")
                elif 'Failed' in status:
                    df1 = df1[['Test Case', 'Build', 'Number of Failed steps']]
                    df_f = df1.merge(df, on=['Test Case','Build'], how="right")
                elif 'Blocked' in status:
                    df1 = df1[['Test Case', 'Build', 'Number of Blocked steps']]
                    df_b = df1.merge(df, on=['Test Case', 'Build'], how="right")
                elif 'Clarification' in status:
                    df1 = df1[['Test Case', 'Build', 'Number of Clarification steps']]
                    df_c = df1.merge(df, on=['Test Case', 'Build'], how="right")
                elif 'Not Run' in status:
                    df_read = pd.read_excel('../Input/xmltoxl.xlsx', index_col=0)
                    df1 = df_read[['Test_case_title','Total_steps','Estimated_exec_[min]']]
                    df_rn = df1.rename(columns={'Test_case_title':'Test Case'})
                    df_nr = df_rn.merge(df, on=['Test Case'], how="right")
            data_frames = [df_pp,df_p,df_f,df_b,df_c,df_nr]

            #Remove the duplicated columns after merge
            df_merged = reduce(lambda left, right: pd.merge(left, right, on=['Test Case','Build'],how='outer', suffixes=('', '_x,_y,_z')), data_frames)
            df_merged.rename(columns={'Build':'BD'}, inplace=True)
            df_merged.drop(df_merged.filter(regex='BD|_x,_y,_z$').columns, axis=1, inplace=True)
            df_merged['Number of Not Run Steps'] = df_merged['Total_steps'].where(df_merged.iloc[:, 2].str.contains('Not Run'))

            #Rearrange the column names
            df2 = df_merged[df_merged.columns[[5,0,2,3,4,11,6,1,7,8,9,12,10]]]
            df3 = df2.rename(columns={df2.columns[3]: 'Date',df2.columns[4]: 'Execution duration (min)'})

            #File name creation
            fn = df3.columns[2].replace(' ','_')
            fname = fn.replace(':','_')+'.xlsx'
            df3.to_excel('../Reports/'+fname,index=False)

            #Add the status values to a total for each build individually, then append the values to a list based on status
            P_total = df3.iloc[:,6].sum().astype('int32')
            passed.append(P_total)
            PP_total = df3.iloc[:,7].sum().astype('int32')
            partially_passed.append(PP_total)
            F_total = df3.iloc[:, 8].sum().astype('int32')
            failed.append(F_total)
            B_total = df3.iloc[:, 9].sum().astype('int32')
            blocked.append(B_total)
            C_total = df3.iloc[:, 10].sum().astype('int32')
            clarification.append(C_total)
            NR_total = df3.iloc[:, 11].sum().astype('int32')
            not_run.append(NR_total)
        logging.info('Test Report Generation Completed')

#########################################################################################################################

        r = list(range(1, self.count+1))
        opt = '../Reports/'
        logging.info("The output location of report files are : " + opt + '\n')
        li = []
        for root, dirs, files in os.walk(opt):
            for file in files:
                if file.endswith('.xlsx') and file.startswith('Build_'):
                    filenm = file.split('.',1)[0]
                    li.append(filenm)
        # Fetch the file name as build name to add it in plot
        names = tuple(li)

        raw_data = {'passed': passed, 'partiallypassed': partially_passed, 'failed': failed,
                    'blocked': blocked, 'clarification': clarification, 'notrun': not_run}
        df = pd.DataFrame(raw_data)

##########################################################################################################

        # Plot 1
        logging.info('Test status Report Plot Display -----------------')
        logging.info("Plot 1 : Display plot with total status count on each build")

        # Color configuration to display in plot
        colors_list = ['#59bd59','#ffc773','#ff471a','#f2f500','#73abff','#999999']

        ax = df.plot(kind='bar', stacked=True, figsize=(10, 8), rot=0, width=0.8, color=colors_list,edgecolor=None)

        value_list = np.random.randint(0, 99, size = len(names))
        pos_list = np.arange(len(names))
        ax.xaxis.set_major_locator(tkr.FixedLocator((pos_list)))
        ax.xaxis.set_major_formatter(tkr.FixedFormatter((names)))

        for c in ax.containers:
            col = c.get_label()
            labels = [v if v > 20 else '' for v in df[col]]
            ax.bar_label(c, labels=labels, label_type='center', fontweight='normal')
        plt.setp(ax.get_xticklabels(), rotation=10, horizontalalignment='right')
        plt.xlabel("Builds",fontsize=10)
        plt.ylabel("Test coverage in percentage",fontsize=10)
        plt.legend(title='Categories',loc='upper left', bbox_to_anchor=(1,1))
        plt.title('Test Status')

########################################################################################################################
        # Plot 2
        logging.info("Plot 2: Display plot to show values from raw to percentage")
        totals = [i + j + k + l + m + n for i, j, k, l, m, n in zip(df['passed'], df['partiallypassed'], df['failed'], df['blocked'], df['clarification'],df['notrun'])]
        passed = [i / j * 100 for i, j in zip(df['passed'], totals)]
        partiallypassed = [i / j * 100 for i, j in zip(df['partiallypassed'], totals)]
        failed = [i / j * 100 for i, j in zip(df['failed'], totals)]
        blocked = [i / j * 100 for i, j in zip(df['blocked'], totals)]
        clarification = [i / j * 100 for i, j in zip(df['clarification'], totals)]
        notrun = [i / j * 100 for i, j in zip(df['notrun'], totals)]

        barWidth = 0.75
        #fig, ax1 = plt.subplots(figsize=(15,8))
        fig = plt.figure()
        ax1 = fig.add_subplot(111)

        # Create Bars
        ax1.bar(r, passed, color='#59bd59', edgecolor='white', width=barWidth, label="passed", alpha=0.9)
        ax1.bar(r, partiallypassed, bottom=passed, color='#ffc773', edgecolor='white', width=barWidth,label="partiallypassed", alpha=0.9)
        ax1.bar(r, failed, bottom=[i + j for i, j in zip(passed, partiallypassed)], color='#ff471a', edgecolor='white',width=barWidth, label="failed", alpha=0.9)
        ax1.bar(r, blocked, bottom=[i + j + k for i, j, k in zip(passed, partiallypassed, failed)], color='#f2f500',edgecolor='white', width=barWidth, label="blocked", alpha=0.8)
        ax1.bar(r, clarification, bottom=[i + j + k + l for i, j, k, l in zip(passed, partiallypassed, failed, blocked)],color='#73abff', edgecolor='white', width=barWidth, label="clarification", alpha=0.9)
        ax1.bar(r, notrun, bottom=[i + j + k + l + m for i, j, k, l, m in zip(passed, partiallypassed, failed, blocked, clarification)], color='#999999',edgecolor='white', width=barWidth, label="notrun", alpha=0.9)

        # Add percentages as labels
        for i, rect in enumerate(ax1.patches):
            # Find where everything is located
            height = rect.get_height()
            width = rect.get_width()
            x = rect.get_x()
            y = rect.get_y()

            # The height of the bar is the count value and can used as the label
            label_text = f'{height:.0f} %'

            label_x = x + width / 2
            label_y = y + height / 2

            # don't include label if it's equivalently 0
            if height > 0:
                ax1.text(label_x, label_y, label_text, ha='center', va='center')
                ax1.yaxis.set_major_formatter(mtick.PercentFormatter())

        plt.title('Test Status completion in Percentage')
        plt.xticks(r, names,rotation=8)
        plt.xlabel("Builds with test coverage",fontsize=10)
        plt.ylabel("Percent(%)",fontsize=10)
        plt.legend(title='Categories',loc='upper left', bbox_to_anchor=(1,1))

val1 = xmlProcessor()