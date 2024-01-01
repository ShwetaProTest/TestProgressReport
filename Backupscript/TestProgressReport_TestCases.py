#A Python script that reads input files, generates test case progress reports, and displays the results in a plot.
# Developer: RHM

import math
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import matplotlib.dates as mdates
from datetime import datetime
from pathlib import Path
import os
import csv
import numpy as np
import matplotlib.ticker as tkr
import matplotlib.ticker as mtick
import matplotlib.patches as mpatches
import warnings
warnings.filterwarnings('ignore')

def process_build_file(file_path, xl_df):
    df = pd.read_excel(file_path)

    # Rename the columns
    df = df.rename(columns={
        df.columns[1]: "Date",
        df.columns[2]: "Execution duration (min)",
        df.columns[3]: "Test_case_title"
    })

    # Get the build name
    build_name = df.iloc[:, 0].name

    # Get the column names that start with 'Build'
    df_builds = df.filter(regex='^Build').columns.tolist()

    # Add a new column for the build
    df['build'] = df_builds[0]
    df['Build'] = df['build'].str.split(' ', 1).str[1]
    df = df.drop(['build', 'Build'], axis=1)

    # Merge with the xl_df dataframe
    merged_df = pd.merge(df, xl_df[['Test_case_title', 'Estimated_exec_[min]']], on='Test_case_title', how='left')

    # Reorder the columns
    merged_df = merged_df[['Test Suite', 'Test_case_title', df_builds[0], 'Date', 'Estimated_exec_[min]', 'Execution duration (min)']]
    merged_df = merged_df.rename(columns={'Test_case_title': 'Test Case','Execution duration (min)': 'Actual_exec_[min]'})

    # Extract the raw result and result columns
    merged_df['Raw Result'] = df[df_builds[0]]
    merged_df['Result'] = merged_df['Raw Result'].str.replace(r'\[v\d+\]', '').str.strip()

    return merged_df

def plot_graph(df_list,build_name,testcase_report,add_prev_values=True,):
    for df in df_list:
        date_col = 'Date' if 'Date' in df.columns else 'Execution Date'
        statuses = ['Passed', 'Partially Passed', 'Failed', 'Blocked', 'Clarification', 'Not Run']
        date_counts = {status: {} for status in statuses}
        Total_test_case_count = df['Test Case'].count()

        # Iterate through each date and count the number of tests for each status
        prev_row = None
        prev_not_run_count = Total_test_case_count

        df[date_col] = pd.to_datetime(df[date_col])

        for date, group in df.groupby(df[date_col].dt.date):
            for status in statuses:
                if status == 'Passed':
                    count = group['Result'].str.count(status).sum() - group['Result'].str.count('Partially Passed').sum()
                elif status == 'Partially Passed':
                    count = group['Result'].str.count(status).sum()
                else:
                    count = group['Result'].str.count(status).sum()

                if status not in date_counts:
                    date_counts[status][date] = 0

                if date not in date_counts[status]:
                    date_counts[status][date] = count
                else:
                    date_counts[status][date] += count

            # Calculate the Not Run count for this date
            if date == df[date_col].min():
                not_run_count = Total_test_case_count - group['Result'].count()
            else:
                status_counts = [date_counts[status][date] for status in statuses]
                not_run_count = prev_not_run_count - sum(status_counts)

            date_counts['Not Run'][date] = not_run_count

            # Add values from previous date
            if add_prev_values and prev_row is not None:
                for status in statuses:
                    if status != 'Not Run':
                        if prev_row in date_counts[status]:
                            date_counts[status][date] += date_counts[status][prev_row]

            prev_row = date
            prev_not_run_count = date_counts['Not Run'][date]

        date_counts_df = pd.DataFrame.from_dict(date_counts)
        date_counts_df.index = pd.to_datetime(date_counts_df.index, format='%Y-%m-%d')

    date_counts_df.to_csv(testcase_report + build_name + '_TestCase.csv')
    return date_counts_df

def configure_plot(build_name):
    plt.figure(figsize=(12, 6))
    plt.xticks(rotation=25)
    plt.gcf().set_size_inches(12, 6)
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    plt.subplots_adjust(right=0.7)
    plt.grid(visible=True, which='major', color='k', alpha=.12, linewidth=1.0,linestyle='--')
    plt.grid(visible=True, which='minor', color='k', alpha=.1, linewidth=0.5,linestyle='--')
    params = {'legend.fontsize': 'small',
              'axes.labelsize': 8,
              'axes.titlesize': 'x-large'}
    plt.rcParams.update(params)
    plt.tick_params(axis='y', which='major', length=3, labelsize=9)
    plt.tick_params(axis='x', which='major', length=7, labelsize=9, rotation=25)
    plt.xlabel('Date',fontsize=10,fontweight='bold', fontstyle='oblique')
    plt.ylabel('Number of Test Cases',fontsize=10,fontweight='bold', fontstyle='oblique')
    plt.title(build_name + " - Test Progress Report [in Test cases]", fontdict={'fontsize': 11, 'fontweight': 'bold'},color='black',backgroundcolor="whitesmoke", pad=15.0)

def annotate_weekends(ax, start, end):
    for i, day in enumerate(pd.date_range(start, end)):
        if day.dayofweek in [5, 6]:
            ax.axvspan(day, day + pd.Timedelta(days=1), alpha=0.1, color='gray')

def format_axis(ax):
    ax.grid(b=True, which='major', color='k', alpha=.12, linewidth=1.0, linestyle='--')
    ax.grid(b=True, which='minor', color='k', alpha=.1, linewidth=0.5, linestyle='--')
    ax.tick_params(axis='y', which='major', length=3, labelsize=8)
    ax.tick_params(axis='x', which='major', length=7, labelsize=8, rotation=8)

def create_build_status_plot(date_counts_df, total_test_cases, build_name, build_status_report,new_labels,colors_dict,colors):
    # Create a figure with two subplots
    fig, (ax1, ax2) = plt.subplots(nrows=1, ncols=2, figsize=(12, 6))
    fig.subplots_adjust(wspace=0.3)

    #Total Number of Test Cases by Status
    ax1.bar(date_counts_df.columns, date_counts_df.iloc[-1], color=[colors_dict[col] for col in date_counts_df.columns])
    ax1.set_xlabel("Test Status", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax1.set_ylabel("Number of Test Cases", fontsize=10,fontweight='bold', fontstyle='oblique')
    format_axis(ax1)
    for i, val in enumerate(date_counts_df.iloc[-1]):
        ax1.text(i, val, int(val), ha='center', va='bottom', fontsize=8)
    ax1.set_xticklabels(new_labels,rotation=25)

    #Percentage of Test Cases by Status
    df_percent = date_counts_df.apply(lambda x: x / total_test_cases * 100)
    ax2.bar(df_percent.columns, df_percent.iloc[-1], color=[colors_dict[col] for col in df_percent.columns])
    ax2.set_xlabel("Test Status", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax2.set_ylabel("Percentage of Test Cases", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax2.set_ylim(0, 100)
    format_axis(ax2)
    for i, val in enumerate(df_percent.iloc[-1]):
        ax2.text(i, val, str(round(val, 2)) + '%', ha='center', va='bottom', fontsize=8)
    ax2.set_xticklabels(new_labels, rotation=25)

    new_handles = []
    for col, label in zip(date_counts_df.columns, new_labels):
        new_handles.append(mpatches.Patch(color=colors_dict[col], label=label))
    new_handles.append(mpatches.Patch(color='white', label=f'Total Test Cases: {int(total_test_cases)}'))
    fig.suptitle(build_name + " - Build Status [Test Case]", fontweight='bold', color='black',backgroundcolor="whitesmoke")
    plt.legend(handles=new_handles, loc='center left', bbox_to_anchor=(1.04, 0.8), fontsize=8)

    plt.savefig(build_status_report + build_name + '_Buildstatus_TestCase.png', bbox_inches='tight')
    fig.tight_layout()
    ####################################################################################################################

def product_comparison_report(df, names, new_labels, total_test_cases, comparison_report, colors):
    # Create figure and axes
    fig, (ax1, ax2) = plt.subplots(nrows=1, ncols=2, figsize=(12, 6))
    fig.subplots_adjust(wspace=0.3)

    barWidth = 0.75
    pos_list = np.arange(len(names))

    df.plot(kind='bar', stacked=True, ax=ax1, rot=0, width=0.8, color=colors, edgecolor=None, legend=None)
    ax1.xaxis.set_major_locator(tkr.FixedLocator((pos_list)))
    ax1.xaxis.set_major_formatter(tkr.FixedFormatter((names)))
    ax1.tick_params(axis='x', labelsize=8)

    for c in ax1.containers:
        col = c.get_label()
        labels = [str(v) if v != 0 else ' ' for v in df[col]]
        #labels = [v if v>5 else '' for v in df[col]]
        ax1.bar_label(c, labels=labels, label_type='center', fontweight='normal', alpha=0.8, fontsize=8)
    plt.setp(ax1.get_xticklabels(), rotation=15, horizontalalignment='right')
    ax1.set_xlabel("Builds", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax1.set_ylabel("Number of Test Cases", fontsize=10,fontweight='bold', fontstyle='oblique')

    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
    ax2.xaxis.set_major_locator(tkr.FixedLocator((pos_list)))
    ax2.xaxis.set_major_formatter(tkr.FixedFormatter((names)))
    ax2.tick_params(axis='x', labelsize=8)

    totals = df.sum(axis=1)
    percentages = df.divide(totals, axis=0) * 100

    for i, label in enumerate(new_labels):
        bottom = percentages.iloc[:, :i].sum(axis=1)
        ax2.bar(pos_list, percentages[label], bottom=bottom, color=colors[i], label=label, alpha=0.8)
        for j, height in enumerate(percentages[label]):
            if height > 0:
                ax2.text(j, bottom[j] + height / 2 + 1, f'{height:.0f} %', ha='center', va='center', fontsize=8)

    plt.setp(ax2.get_xticklabels(), rotation=15, horizontalalignment='right')
    ax2.set_xlabel("Builds", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax2.set_ylabel("Percentage of test Cases", fontsize=10,fontweight='bold', fontstyle='oblique')
    fig.suptitle('Product Progress Report [in Test Cases & percentage(%)]', fontweight='bold', color='black',backgroundcolor="whitesmoke")

    # Add Total Steps to the legend
    handles, labels = ax2.get_legend_handles_labels()
    handles.append(mpatches.Patch(color='white', label=f'Total Test Cases: {int(total_test_cases)}'))
    plt.legend(handles=handles, loc='center left', bbox_to_anchor=(1.04, 0.8), fontsize=8)

    plt.savefig(comparison_report + 'Progress_Report_TestCase.png', bbox_inches='tight')
    print(f"The total Test cases build File contain are : {int(total_test_cases)}")
    print(f"The Progress Report for TestCases files have been created. Files generated in : : {comparison_report}")
    fig.tight_layout()
    ####################################################################################################################

def plot_teststep(df, total_test_cases,legend_labels):
    estimated_time_sum = 0
    actual_time_sum = 0
    Suggested_Estimation_Buffer_Days = 0

    # Calculate the estimated and actual time sums for all builds
    estimated_time_sum += df['Estimated_exec_[min]'].sum() / 60
    actual_time_sum += df['Actual_exec_[min]'].sum() / 60
    total_time_remaining = estimated_time_sum - actual_time_sum

    # Calculation2 - Project the Remaining Estimated time in plot
    estimated_execution_days = math.ceil(estimated_time_sum / 8)
    actual_execution_days = math.ceil(actual_time_sum / 8)
    Remaining_Estimated_Days = estimated_execution_days - actual_execution_days

    # Calculation3 - Conditions to Calculate Remaining Estimated efforts
    if Remaining_Estimated_Days < 5:
        Suggested_Estimation_Buffer_Days += 1
    elif Remaining_Estimated_Days >= 5 and Remaining_Estimated_Days < 10:
        Suggested_Estimation_Buffer_Days += 2
    elif Remaining_Estimated_Days >= 10 and Remaining_Estimated_Days < 20:
        Suggested_Estimation_Buffer_Days += 3
    elif Remaining_Estimated_Days >= 20:
        Suggested_Estimation_Buffer_Days += 5

    # Add the estimated and actual time values to the legend
    legend_labels.extend([
        "",
        f"Total Test Cases - {total_test_cases}",
        "",
        f"Estimated Execution Time - {estimated_time_sum:.2f} hrs",
        f"Actual Execution Time - {actual_time_sum:.2f} hrs",
        f"Remaining Estimation Time - {total_time_remaining:.2f} hrs",
        "",
        f"Estimated Execution Days - {estimated_execution_days} days",
        f"Actual Execution Days - {actual_execution_days} days",
        f"Remaining Estimation Days - {Remaining_Estimated_Days} days",
        "",
        f"Suggested Estimation Buffer Days - {Suggested_Estimation_Buffer_Days} days"
    ])

def test_case_report(build_input,testcase_report,comparison_report,build_status_report,colors,current_datetime):
    print(f"\nscript3:")
    print(f"The Execution has started to create a Progress Report for TestCases. Please wait!..")
    current_dir = os.getcwd()
    parent_dir = os.path.abspath(os.path.join(current_dir, os.pardir))

    # Check if the Build_Input folder exists and if it contains any Build_*.xlsx files
    build_file_path = os.path.join(parent_dir, "Input/Build_Input")
    build_files = [x for x in Path(build_file_path).glob("**/*.xlsx") if x.name.__contains__("Build_")]
    if not build_files:
        raise FileNotFoundError("No Build files found in the Build_Input folder")

    # Check if the TestSpecification_Testlink.xlsx file exists in the TestSpecificationReport folder
    xmltoxl_file_path = os.path.join(parent_dir, 'Reports/' + current_datetime + '/TestSpecificationReport', "TestSpecification_Testlink.xlsx")
    if not os.path.exists(xmltoxl_file_path):
        raise FileNotFoundError("TestSpecification_Testlink.xlsx file not found in the TestSpecificationReport folder")

    # Read the TestSpecification_Testlink.xlsx file into a pandas dataframe
    xl_df = pd.read_excel(xmltoxl_file_path)

    df_list = []
    passed = []
    partially_passed = []
    failed = []
    blocked = []
    clarification = []
    not_run = []
    for i, file_path in enumerate(build_files):
        df = process_build_file(file_path, xl_df)
        df_list.append(df)
        build_name = df.columns.values[2]
        build_name = df.columns.values[2].replace(' ', '_').replace("+", "_").replace(":", "_")
        while '__' in build_name:
            build_name = build_name.replace("__", "_")
        date_counts_df = plot_graph(df_list,build_name,testcase_report)
        total_test_cases = date_counts_df.iloc[-1].sum()
        new_labels = ['Passed', 'Partially Passed', 'Failed', 'Blocked', 'Clarification', 'Not Run']
        # Define the color codes
        colors_dict = {'Passed': colors[0],
                       'Partially Passed': colors[1],
                       'Failed': colors[2],
                       'Blocked': colors[3],
                       'Clarification': colors[4],
                       'Not Run': colors[5]
                       }
        #plot1
        create_build_status_plot(date_counts_df, total_test_cases, build_name, build_status_report,new_labels,colors_dict,colors)

    ####################################################################################################################
        #plot2
        configure_plot(build_name)
        ax = plt.gca()
        annotate_weekends(ax, date_counts_df.index[0], date_counts_df.index[-1])

        # Add extra row to DataFrame
        new_row = {status: 0 for status in new_labels}
        new_row['Not Run'] = total_test_cases
        new_index = [
            (pd.to_datetime(date_counts_df.index, format='%Y-%m-%d').min() - pd.Timedelta(days=1)).strftime('%Y-%m-%d')]
        df_extra = pd.DataFrame(new_row, index=new_index)

        # Concatenate df_extra and df_counts
        date_counts_df = pd.concat([df_extra, date_counts_df])
        date_counts_df.index = pd.to_datetime(date_counts_df.index)
        date_counts_df = date_counts_df.astype(int)

        # Set x-ticks
        xtick_dates = [date_counts_df.index[0], date_counts_df.index[-1]]  # First and last date
        mondays = date_counts_df.index.weekday == 0  # Boolean mask for dates that fall on a Monday
        xtick_dates = xtick_dates + list(date_counts_df.index[mondays])

        # Set x-tick labels
        date_format = '%Y-%m-%d'  # or any other desired format
        xtick_labels = [dt.strftime(date_format) for dt in xtick_dates]
        ax.set_xticks(xtick_dates)
        ax.set_xticklabels(xtick_labels, rotation=45, ha='right')

        max_count = date_counts_df.values.max()
        ax.set_ylim(0, max_count + 1)

        # Y axis configuration
        ytiks = 1 if total_test_cases < 16 else 2 if total_test_cases < 30 else 5 if total_test_cases < 60 else 10

        Test_Case_Scale = total_test_cases + int(total_test_cases / 5)
        plt.yticks(np.arange(0, Test_Case_Scale, ytiks))

        legend_labels = []
        legend_handles = []
        first_occurrence = True  # set to False to use last occurrence instead
        for status in date_counts_df.columns:
            total_count = date_counts_df[status].sum()
            if total_count > 0:
                if status == "Not Run":
                    last_value = int(date_counts_df[status].iloc[-1])
                    legend_labels.append(f"{status} - {last_value}")
                else:
                    if first_occurrence:
                        last_value = int(
                            date_counts_df[status].where(~date_counts_df[status].duplicated(keep='first')).fillna(method='ffill')[-1])
                    else:
                        last_value = int(date_counts_df[status].where(~date_counts_df[status].duplicated(keep='last')).fillna(method='ffill')[-1])
                    legend_labels.append(f"{status} - {last_value}")
                line, = plt.plot(date_counts_df.index, date_counts_df[status], color=colors_dict[status], linestyle='-',marker='*', markersize=5, linewidth=1.0, alpha=1, label=status,zorder=0)
                legend_handles.append(line)

                for i, count in enumerate(date_counts_df[status]):
                    if count > 0:
                        label = str(count)
                        if first_occurrence:
                            if date_counts_df[status].duplicated(keep='first')[i]:
                                label = ""
                        else:
                            if date_counts_df[status].duplicated(keep='last')[i]:
                                label = ""
                        plt.annotate(label, xy=(date_counts_df.index[i], count), fontsize=8,
                                     textcoords='offset points',xytext=(-2, 10), ha='center', va='top')

        plot_teststep(df, total_test_cases, legend_labels)

        for i, label in enumerate(legend_labels[-20:]):
            legend_handles.append(Line2D([0], [0], color='white', lw=2, label=label))

        # Add the legend handle extend to show all the labels
        plt.legend(handles=legend_handles, labels=legend_labels, handlelength=3, handletextpad=1.2,handleheight=0.5,bbox_to_anchor=(1.05, 1))
        plt.savefig(testcase_report + build_name + '_TestCase_plot.png', bbox_inches='tight')

    ####################################################################################################################
        last_values = date_counts_df.iloc[-1].tolist()
        passed.append(last_values[0])
        partially_passed.append(last_values[1])
        failed.append(last_values[2])
        blocked.append(last_values[3])
        clarification.append(last_values[4])
        not_run.append(last_values[5])

    opt = build_input
    li = []
    for root, dirs, files in os.walk(opt):
        for file in files:
            if file.endswith('.xlsx') and file.startswith('Build_'):
                filenm = file.split('.', 1)[0]
                li.append(filenm)
    names = tuple(li)
    df = pd.DataFrame.from_dict({
        'Passed': passed,
        'Partially Passed': partially_passed,
        'Failed': failed,
        'Blocked': blocked,
        'Clarification': clarification,
        'Not Run': not_run
    })
    ####################################################################################################################
    # plot3
    product_comparison_report(df, names, new_labels, total_test_cases, comparison_report, colors)

#test_case_report()
