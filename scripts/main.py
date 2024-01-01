# This script generates various reports including a Burn Down Chart, Test Specification Report, Build Report, and Test Progress Reports for Test Cases and Test Steps.
# It also logs errors during the execution of the script.
# Developer: RHM
# Created at: 12/07/2023

#import modules
import subprocess
import sys
import os
import pandas as pd
from datetime import datetime
import logging
import warnings
import shutil
import matplotlib.pyplot as plt
import PyPDF2
from BurnDownChart import burndown
import Supporting_scripts
import TestProgressReport_TestCases

# Ignore specific warning messages
warnings.filterwarnings("ignore", message="Ignoring invalid distribution.*")
warnings.filterwarnings("ignore", category=UserWarning, message=".*-atplotlib.*")

class ReportGenerator:
    def __init__(self):
        # Get the current directory of the script
        self.current_directory = os.path.dirname(os.path.abspath(__file__))
        self.requirements_path = os.path.join(self.current_directory, 'requirements.txt')

        # Define the colors for the reports
        self.colors = ['#29c233', '#ffad33', '#f23300', '#f5eb00', '#0063f2', '#808080']
        # colors = [#green, #orange, #red, #Yellow, #Blue, #gray]

        # Define the initial folders with empty values
        self.folders = {
            'input_folder': os.path.abspath(os.path.join(self.current_directory, '../Input/')),
            'build_input': os.path.abspath(os.path.join(self.current_directory, '../Input/Build_Input/')),
            'log_folder': os.path.abspath(os.path.join(self.current_directory, '../log/')),
            'burndown_report': '',
            'testspecification_report': '',
            'build_report': '',
            'testcase_report': '',
            'teststep_report': '',
            'comparison_report': '',
            'build_status_report': '',
            'report_pdf': ''
        }

        # Get the current datetime and log filename
        self.log_filename = datetime.now().strftime('logfile_%d-%m-%y_%H%M%S.log')
        self.log_path = os.path.join(self.folders['log_folder'], self.log_filename)
        self.current_datetime = ''

        # Set the current datetime and update the report folders
        self.set_current_datetime()

    def install_packages(self):
        # Check if the requirements file exists and install the required packages
        if os.path.exists(self.requirements_path):
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", self.requirements_path])
        else:
            # Print an error message specifying the expected folder where the requirements.txt file should be present
            print(f"Error: The requirements file 'requirements.txt' does not exist in the folder: {self.current_directory}. Please make sure the 'requirements.txt' file is present in the correct folder.")

    def create_folders(self):
        # Create necessary folders if they don't exist
        for folder in self.folders.values():
            if not os.path.exists(folder):
                os.makedirs(folder)

    def set_current_datetime(self):
        # Set the current datetime based on the input filename
        directory = os.path.join(self.current_directory, '../Input/')

        # Flag to check if resultsTC file is found
        file_found = False

        for filename in os.listdir(directory):
            if filename.startswith("resultsTC_") and (
                filename.endswith(".xlsx") or filename.endswith(".XLSX") or
                filename.endswith(".xls") or filename.endswith(".XLS")
            ):
                excel_file = pd.read_excel(
                    os.path.join(directory, filename),
                    sheet_name='Worksheet', usecols="B", header=1, nrows=0
                )
                Test_plan = excel_file.columns.values[0]
                Test_plan_title = Test_plan.replace(" ", "_")
                self.current_datetime = datetime.now().strftime(Test_plan_title + '-%d%m%Y_%H%M%S')

                # Update the report folders with the current datetime
                self.folders['burndown_report'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'BurnDownReport/')
                self.folders['testspecification_report'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'TestSpecificationReport/')
                self.folders['build_report'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'BuildReport/')
                self.folders['testcase_report'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'TestProgressReport/TestCase/')
                self.folders['teststep_report'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'TestProgressReport/TestStep/')
                self.folders['comparison_report'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'ProductComparisonReport/')
                self.folders['build_status_report'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'BuildStatusReport/')
                self.folders['report_pdf'] = os.path.join(self.current_directory, '../Reports', self.current_datetime, 'GeneratepdfReport/')

                # Set the flag to True since resultsTC file is found
                file_found = True

                # Break the loop assuming only the first matching filename is needed
                break

        # If resultsTC file is not found, raise an exception
        if not file_found:
            raise FileNotFoundError(f"Error: The 'resultsTC_' file is not found in the: {directory} + folder.")

    def generate_log_file(self):
        # Generate log file using the specified log path
        logging.basicConfig(filename=self.log_path, level=logging.INFO)

    def run_burndown_chart(self):
        try:
            bd = burndown()
            # Run burndown chart generation
            bd.run_burndown(
                input_folder=self.folders['input_folder'],
                log_folder=self.folders['log_folder'],
                burndown_report=self.folders['burndown_report'],
                report_pdf=self.folders['report_pdf']
            )
        except Exception as e:
            print("Error generating burndown chart:", e)

    def run_supporting_func(self):
        try:
            # Run supporting function for generating supporting files
            Supporting_scripts.Supporting_func(
                input_folder=self.folders['input_folder'],
                xml_folder=self.folders['testspecification_report'],
                build_input=self.folders['build_input'],
                build_report=self.folders['build_report'],
                current_datetime=self.current_datetime
            )
        except Exception as e:
            print("Error running supporting function for Supporting file generation report:", e)

    def run_progressreport_testcases(self):
        try:
            # Run test case report generation
            TestProgressReport_TestCases.test_case_report(
                build_input=self.folders['build_input'],
                testspecification_report=self.folders['testspecification_report'],
                testcase_report=self.folders['testcase_report'],
                comparison_report=self.folders['comparison_report'],
                build_status_report=self.folders['build_status_report'],
                colors=self.colors,
                current_datetime=self.current_datetime,
                report_pdf=self.folders['report_pdf']
            )
        except Exception as e:
            print("Error running Test progress Report for Test Cases function:", e)

    def merge_pdfs(self,build_name):
        pdf_files = [
            (file, os.path.getmtime(os.path.join(self.folders['report_pdf'], file)))
            for file in os.listdir(self.folders['report_pdf']) if file.endswith(".pdf")
        ]

        pdf_files.sort(key=lambda x: x[1])

        pdf_writer = PyPDF2.PdfWriter()

        for file, _ in pdf_files:
            path = os.path.join(self.folders['report_pdf'], file)
            pdf_reader = PyPDF2.PdfReader(path)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                pdf_writer.add_page(page)

        output_path = os.path.join(self.folders['report_pdf'], build_name + "_ProgressReport.pdf")

        with open(output_path, 'wb') as output_file:
            pdf_writer.write(output_file)

        for pdf_file, _ in pdf_files:
            if pdf_file != 'Report.pdf':
                file_path = os.path.join(self.folders['report_pdf'], pdf_file)
                os.remove(file_path)

    def generate_reports(self):
        input_folder = self.folders['input_folder']
        for filename in os.listdir(input_folder):
            if filename.startswith("resultsTC_") and (filename.endswith(".xlsx") or filename.endswith(".XLSX") or filename.endswith(".xls") or filename.endswith(".XLS")):
                resultstc_file_path = os.path.join(input_folder, filename)
                excel_file = pd.read_excel(resultstc_file_path, sheet_name='Worksheet', usecols="B", header=1, nrows=0)
                build_name = excel_file.columns.values[0]
        self.create_folders()
        self.generate_log_file()
        self.run_burndown_chart()
        self.run_supporting_func()
        self.run_progressreport_testcases()
        self.merge_pdfs(build_name)
        print("\nPlease view the graph and close it to complete the script")
        plt.show()
        print("The script has completed successfully!!.....")
        shutil.rmtree(self.folders['build_input'])

if __name__ == '__main__':
    report_generator = ReportGenerator()
    report_generator.install_packages()
    report_generator.generate_reports()
