#This script generates various reports including a Burn Down Chart, Test Specification Report, Build Report, and Test Progress Reports for Test Cases and Test Steps.
# It also logs errors during the execution of the script.
# Developer: RHM

import subprocess
import sys
import os

# Install required packages from requirements.txt file
subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "-r", "requirements.txt"])

import pandas as pd
from datetime import datetime
import datetime as dt
import logging
import warnings
import shutil
import matplotlib.pyplot as plt
from BurnDownChart import burndown
import Supporting_scripts
import TestProgressReport_TestCases
import TestProgressReport_TestSteps
warnings.filterwarnings("ignore", message="Ignoring invalid distribution.*")
warnings.filterwarnings("ignore", category=UserWarning, message=".*-atplotlib.*")

directory = "../Input/"
for filename in os.listdir(directory):
    if filename.startswith("resultsTC_") and (
            filename.endswith(".xlsx") or filename.endswith(".XLSX") or filename.endswith(".xls") or filename.endswith(".XLS")):
        excel_file = pd.read_excel(directory + filename, sheet_name='Worksheet', usecols="B", header=1, nrows=0)
        Test_plan = excel_file.columns.values[0]
        Test_plan_title = Test_plan.replace(" ", "_")

if not os.path.exists('../Reports/'):
    os.makedirs('../Reports/')

if not os.path.exists('../log/'):
    os.makedirs('../log/')

current_datetime = datetime.now().strftime(Test_plan_title +'-%d%m%Y_%H%M%S')

folders = {
    'input_folder': '../Input/',
    'build_input':'../Input/Build_Input/',
    'log_folder': '../log/',
    'burndown_report':'../Reports/' + current_datetime + '/BurnDownReport/',
    'testspecification_report':'../Reports/' + current_datetime + '/TestSpecificationReport/',
    'build_report':'../Reports/' + current_datetime + '/BuildReport/',
    'testcase_report':'../Reports/' + current_datetime + '/TestProgressReport/TestCase/',
    'teststep_report':'../Reports/' + current_datetime + '/TestProgressReport/TestStep/',
    'comparison_report':'../Reports/' + current_datetime + '/ProductComparisonReport/',
    'build_status_report':'../Reports/' + current_datetime + '/BuildStatusReport/'
}

# create folders if they don't exist
for folder in folders.values():
    if not os.path.exists(folder):
        os.makedirs(folder)

# Define the log filename using the current date and time
log_filename = datetime.now().strftime('logfile_%d-%m-%y_%H%M%S.log')
log_path = os.path.join(folders['log_folder'], log_filename)
logging.basicConfig(filename=log_path, level=logging.INFO)

# Color configuration to display in plot
colors = ['#29c233', '#ffad33', '#f23300', '#f5eb00', '#0063f2', '#808080']
#colors = [#green, #orange, #red, #Yellow, #Blue, #gray]

#Define functions to generate reports
def run_burndown_chart():
    try:
        bd = burndown()
        bd.run_burndown(input_folder=folders['input_folder'],log_folder=folders['log_folder'],burndown_report=folders['burndown_report'])
    except Exception as e:
        print("Error generating burndown chart:", e)

def run_supporting_func(current_datetime):
    try:
        Supporting_scripts.Supporting_func(input_folder=folders['input_folder'],xml_folder=folders['testspecification_report'],build_input=folders['build_input'],build_report=folders['build_report'],current_datetime=current_datetime)
    except Exception as e:
        print("Error running supporting function for Supporting file generation report:", e)

def run_progressreport_testcases(colors,current_datetime):
    try:
        TestProgressReport_TestCases.test_case_report(build_input=folders['build_input'],testcase_report=folders['testcase_report'],comparison_report=folders['comparison_report'],build_status_report=folders['build_status_report'],colors=colors,current_datetime=current_datetime)
    except Exception as e:
        print("Error running Test progress Report for Test Cases function:", e)

def run_progressreport_teststeps(colors,current_datetime):
    try:
        TestProgressReport_TestSteps.test_step_report(build_report=folders['build_report'],teststep_report=folders['teststep_report'],comparison_report=folders['comparison_report'],build_status_report=folders['build_status_report'],colors=colors,current_datetime=current_datetime)
    except Exception as e:
        print("Error running Test progress Report for Test Step function:", e)

#Execute the functions to generate reports
if __name__ == '__main__':
    run_burndown_chart()
    run_supporting_func(current_datetime)
    run_progressreport_testcases(colors,current_datetime)
    run_progressreport_teststeps(colors,current_datetime)
    print(f"\nPlease view the graph and close it to complete the script")
    #plt.show()
    print("The script has completed successfully!!.....")
    shutil.rmtree(folders['build_input'])

