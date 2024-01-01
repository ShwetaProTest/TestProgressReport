#A Python script that reads input files, generates test steps progress reports, and displays the results in a plot.
# Developer: RHM

import pandas as pd
from pathlib import Path
import os
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import math
from matplotlib.lines import Line2D
import matplotlib.ticker as tkr
import matplotlib.ticker as mtick
import matplotlib.patches as mpatches
import warnings
warnings.filterwarnings('ignore')

def configure_plot(build_name):
    plt.figure(figsize=(12, 6))
    plt.xticks(rotation=25)

    # format x-axis, add grid, and get legend handles
    plt.gcf().set_size_inches(12, 6)
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    plt.subplots_adjust(right=0.7)
    # format x-axis, add grid, and get legend handles
    plt.grid(visible=True, which='major', color='k', alpha=.12, linewidth=1.0, linestyle='--')
    plt.grid(visible=True, which='minor', color='k', alpha=.1, linewidth=0.5, linestyle='--')
    params = {'legend.fontsize': 'small',
              'axes.labelsize': 8,
              'axes.titlesize': 'x-large'}
    plt.rcParams.update(params)
    # Set the title and labels
    plt.title(build_name + " - Test Progress Report [in Test Steps]",fontdict={'fontsize': 11, 'fontweight': 'bold'},color='black',backgroundcolor="whitesmoke", pad=15.0)
    plt.xlabel('Date', fontsize=10,fontweight='bold', fontstyle='oblique')
    plt.ylabel('Number of Test Steps', fontsize=10,fontweight='bold', fontstyle='oblique')

def annotate_weekends(ax, start, end):
    for i, day in enumerate(pd.date_range(start, end)):
        if day.dayofweek in [5, 6]:
            ax.axvspan(day, day + pd.Timedelta(days=1), alpha=0.1, color='gray')

def format_axis(ax):
    ax.grid(b=True, which='major', color='k', alpha=.12, linewidth=1.0, linestyle='--')
    ax.grid(b=True, which='minor', color='k', alpha=.1, linewidth=0.5, linestyle='--')
    ax.tick_params(axis='y', which='major', length=3, labelsize=8)
    ax.tick_params(axis='x', which='major', length=7, labelsize=8, rotation=8)
def create_build_status_plot(df_counts, colors_dict,new_labels,build_name,steps,build_status_report,colors):
    # Create a figure with two subplots
    fig, (ax1, ax2) = plt.subplots(nrows=1, ncols=2, figsize=(12, 6))
    fig.subplots_adjust(wspace=0.3)

    #Total Number of Steps by Status
    ax1.bar(df_counts.columns, df_counts.iloc[-1], color=[colors_dict[col] for col in df_counts.columns])
    ax1.set_xlabel("Test Status", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax1.set_ylabel("Number of Test Steps", fontsize=10,fontweight='bold', fontstyle='oblique')
    format_axis(ax1)
    for i, val in enumerate(df_counts.iloc[-1]):
        ax1.text(i, val, int(val), ha='center', va='bottom', fontsize=8)
    ax1.set_xticklabels(new_labels, rotation=25)

    #Percentage of Steps by Status
    df_percent = df_counts.apply(lambda x: x / steps * 100)
    ax2.bar(df_percent.columns, df_percent.iloc[-1], color=[colors_dict[col] for col in df_percent.columns])
    ax2.set_xlabel("Test Status", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax2.set_ylabel("Percentage of Test Steps", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax2.set_ylim(0, 100)
    format_axis(ax2)
    for i, val in enumerate(df_percent.iloc[-1]):
        ax2.text(i, val, str(round(val, 2)) + '%', ha='center', va='bottom', fontsize=8)
    ax2.set_xticklabels(new_labels, rotation=25)

    # Add Total Steps to the legend
    new_handles = []
    for col, label in zip(df_counts.columns, new_labels):
        new_handles.append(mpatches.Patch(color=colors_dict[col], label=label))
    new_handles.append(mpatches.Patch(color='white', label=f'Total Steps: {int(steps)}'))
    fig.suptitle(build_name + " - Build Status [Test Step]", fontweight='bold', color='black',backgroundcolor="whitesmoke")
    plt.legend(handles=new_handles, loc='center left', bbox_to_anchor=(1.04, 0.8), fontsize=8)

    plt.savefig(build_status_report + build_name + '_Buildstatus_TestStep.png', bbox_inches='tight')
    fig.tight_layout()

def product_comparison_report(df,names,new_labels,steps,comparison_report,colors):
    # Create figure and axes
    fig, (ax1, ax2) = plt.subplots(nrows=1, ncols=2, figsize=(12, 6))
    fig.subplots_adjust(wspace=0.3)

    barWidth = 0.75
    pos_list = np.arange(len(names))

    df.plot(kind='bar', stacked=True, ax=ax1, rot=0, width=0.8, color=colors, edgecolor=None, legend=None)
    ax1.xaxis.set_major_locator(tkr.FixedLocator(pos_list))
    ax1.xaxis.set_major_formatter(tkr.FixedFormatter(names))
    ax1.tick_params(axis='x', labelsize=8)

    for c in ax1.containers:
        col = c.get_label()
        labels = [str(v) if v != 0 else ' ' for v in df[col]]
        #labels = [v if v > 5 else ' ' for v in df[col]]
        ax1.bar_label(c, labels=labels, label_type='center', fontweight='normal', alpha=0.8, fontsize=8)
    plt.setp(ax1.get_xticklabels(), rotation=15, horizontalalignment='right')
    ax1.set_xlabel("Builds", fontsize=10,fontweight='bold', fontstyle='oblique')
    ax1.set_ylabel("Number of Test Steps", fontsize=10,fontweight='bold', fontstyle='oblique')

    ax2.yaxis.set_major_formatter(mtick.PercentFormatter())
    ax2.xaxis.set_major_locator(tkr.FixedLocator(pos_list))
    ax2.xaxis.set_major_formatter(tkr.FixedFormatter(names))
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
    ax2.set_ylabel("Percentage of test Steps", fontsize=10,fontweight='bold', fontstyle='oblique')
    fig.suptitle('Product Progress Report [in Test Steps & percentage(%)]', fontweight='bold', color='black',backgroundcolor="whitesmoke")

    # Add Total Steps to the legend
    handles, labels = ax2.get_legend_handles_labels()
    handles.append(mpatches.Patch(color='white', label=f'Total Steps: {int(steps)}'))
    plt.legend(handles=handles, loc='center left', bbox_to_anchor=(1.04, 0.8), fontsize=8)

    print(f"The total Test steps build File contain are : {int(steps)}")
    print(f"The Progress Report for TestSteps files have been created. Files generated in : : {comparison_report}")

    plt.savefig(comparison_report + 'Progress_Report_TestStep.png', bbox_inches='tight')
    fig.tight_layout()

def plot_teststep(df,steps,legend_labels):
    estimated_time_sum = 0
    actual_time_sum = 0
    Suggested_Estimation_Buffer_Days = 0

    # Calculate the estimated and actual time sums for all builds
    estimated_time_sum += df['Estimated_exec_[min]'].sum() / 60
    actual_time_sum += df['Execution duration (min)'].sum() / 60
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

    legend_labels.extend([
        "",
        f"Total Steps - {int(steps)}",
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

def test_step_report(build_report,teststep_report,comparison_report,build_status_report,colors,current_datetime):
    print(f"\nscript4:")
    print(f"The Execution has started to create a Progress Report for TestSteps. Please wait!..")
    current_dir = os.getcwd()
    parent_dir = os.path.abspath(os.path.join(current_dir, os.pardir))

    # Check if the BuildReport folder exists and if it contains any Build_*.xlsx files
    build_file_path = os.path.join(parent_dir, 'Reports/' + current_datetime + '/BuildReport')
    build_files = [x for x in Path(build_file_path).glob("**/*.xlsx") if x.name.startswith("Build_")]
    if not build_files:
        raise FileNotFoundError("No Build files found in the BuildReport folder")

    all_builds_status_sums = []
    passed = []
    partially_passed = []
    failed = []
    blocked = []
    clarification = []
    not_run = []
    build_counts = []
    for file_path in build_files:
        df = pd.read_excel(file_path)
        build_name = df.columns.values[2]
        build_name = df.columns.values[2].replace(' ', '_').replace("+", "_").replace(":", "_")
        while '__' in build_name:
            build_name = build_name.replace("__", "_")

        # Get the sum of each status on each date
        date_col = [col for col in df.columns if 'Date' in col][0]
        df[date_col] = pd.to_datetime(df[date_col], format='%Y-%m-%d', errors='coerce')
        df[date_col] = df[date_col].dt.strftime('%Y-%m-%d')
        statuses = ['Number of Passed steps', 'Number of Partially Passed steps', 'Number of Failed steps','Number of Blocked steps', 'Number of Clarification steps']
        date_counts = {status: {} for status in statuses + ['Number of Not Run steps']}

        # Calculate the Number of Not Run steps and add values from previous date for all the statuses columns
        prev_date_counts = {status: 0 for status in statuses}
        steps = df['Total_steps'].sum()

        prev_date = pd.to_datetime(df[date_col], format='%Y-%m-%d').min().strftime('%Y-%m-%d')
        date_counts['Number of Not Run steps'][prev_date] = df['Total_steps'].sum()

        for date, group in df.groupby([date_col]):
            total_steps = group['Total_steps'].sum()
            if prev_date in date_counts['Number of Not Run steps']:
                date_counts['Number of Not Run steps'][date] = int(
                    date_counts['Number of Not Run steps'][prev_date] - total_steps)
            prev_date = date

            for status in statuses:
                if status not in date_counts:
                    date_counts[status] = {}
                date_counts[status][date] = prev_date_counts[status] + group[status].sum()
                prev_date_counts[status] = date_counts[status][date]

        #Convert the dictionary to a dataframe
        df_counts = pd.DataFrame.from_dict(date_counts)

        df_counts.to_csv(teststep_report + build_name + '_TestStep.csv')

    ####################################################################################################################
        # Define new labels for x-axis
        new_labels = ['Passed', 'Partially Passed', 'Failed', 'Blocked', 'Clarification', 'Not Run']
        # Define the color codes
        colors_dict = {'Number of Passed steps': colors[0],
                    'Number of Partially Passed steps': colors[1],
                    'Number of Failed steps': colors[2],
                    'Number of Blocked steps': colors[3],
                    'Number of Clarification steps': colors[4],
                    'Number of Not Run steps': colors[5]}

        #plot 1
        create_build_status_plot(df_counts,colors_dict,new_labels,build_name,steps,build_status_report,colors)
    ####################################################################################################################
        #plot2
        # Function call
        configure_plot(build_name)
        annotate_weekends(plt.gca(), df_counts.index[0], df_counts.index[-1])

        # Add extra row to DataFrame
        new_row = {status: 0 for status in statuses}
        new_row['Number of Not Run steps'] = steps
        new_index = [(pd.to_datetime(df_counts.index, format='%Y-%m-%d').min() - pd.Timedelta(days=1)).strftime('%Y-%m-%d')]
        df_extra = pd.DataFrame(new_row, index=new_index)

        # Concatenate df_extra and df_counts
        df_counts = pd.concat([df_extra, df_counts])
        df_counts.index = pd.to_datetime(df_counts.index)
        df_counts = df_counts.astype(int)

        # Plot the graphs using the new index values from df_counts
        legend_labels = []
        legend_handles = []
        first_occurrence = True #set to False to use last occurence instead
        for status in statuses:
            if status in df_counts.columns:
                color = colors_dict[status]
                plt.plot(df_counts.index, df_counts[status], label=f"{status} ({df_counts[status].iloc[-1]})",
                         color=color, linestyle='-', marker='*', markersize=5, linewidth=1.0, alpha=1,zorder=0)
                # Add annotations for each data point
                prev_val = None
                for i, val in enumerate(df_counts[status]):
                    if prev_val != val:
                        prev_val = val
                        if first_occurrence:
                            for j in range(i, len(df_counts[status])):
                                if df_counts[status][j] == val:
                                    occurrence = j
                                    break
                        else:
                            for j in range(len(df_counts[status]) - 1, i - 1, -1):
                                if df_counts[status][j] == val:
                                    occurrence = j
                                    break
                        plt.text(df_counts.index[occurrence], val, str(val), ha='center', va='bottom', fontsize=8)
                legend_labels.append(f"{status} ({df_counts[status].iloc[-1]})")
                legend_handles.append(Line2D([0], [0], color=color, lw=1))
        if 'Number of Not Run steps' in df_counts.columns:
            df_counts['Number of Not Run steps'] = df_counts['Number of Not Run steps'].astype(int)
            color = colors_dict['Number of Not Run steps']
            plt.plot(df_counts.index, df_counts['Number of Not Run steps'],
                     label=f"Number of Not Run steps ({df_counts['Number of Not Run steps'].iloc[-1]})", linestyle='--',
                     color=color,marker='*', markersize=5, linewidth=1.0, alpha=1,zorder=0)
            # Add annotations for each data point
            prev_val = None
            for i, val in enumerate(df_counts['Number of Not Run steps']):
                if prev_val != val:
                    prev_val = val
                    if first_occurrence:
                        for j in range(i, len(df_counts['Number of Not Run steps'])):
                            if df_counts['Number of Not Run steps'][j] == val:
                                occurrence = j
                                break
                    else:
                        for j in range(len(df_counts['Number of Not Run steps']) - 1, i - 1, -1):
                            if df_counts['Number of Not Run steps'][j] == val:
                                occurrence = j
                                break
                    plt.text(df_counts.index[occurrence], val, str(val), ha='center', va='bottom', fontsize=8)
            legend_labels.append(f"Number of Not Run steps ({df_counts['Number of Not Run steps'].iloc[-1]})")
            legend_handles.append(Line2D([0], [0], color=color, lw=1, linestyle='--'))

        plot_teststep(df, steps,legend_labels)

        # Add legend handles
        for i, label in enumerate(legend_labels[-20:]):
            legend_handles.append(Line2D([0], [0], color='white', lw=2, label=label))

        # Add the legend handle extend to show all the labels
        plt.legend(handles=legend_handles, labels=legend_labels, handlelength=3, handletextpad=1.2, handleheight=0.5,
                   bbox_to_anchor=(1.05, 1))

        plt.savefig(teststep_report + build_name + '_TestStep_plot.png', bbox_inches='tight')

    ####################################################################################################################
        P_total = df.iloc[:, 6].sum().astype('int32')
        passed.append(P_total)
        PP_total = df.iloc[:, 7].sum().astype('int32')
        partially_passed.append(PP_total)
        F_total = df.iloc[:, 8].sum().astype('int32')
        failed.append(F_total)
        B_total = df.iloc[:, 9].sum().astype('int32')
        blocked.append(B_total)
        C_total = df.iloc[:, 10].sum().astype('int32')
        clarification.append(C_total)
        NR_total = df.iloc[:, 11].sum().astype('int32')
        not_run.append(NR_total)

    opt = build_report
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
    #plot3
    product_comparison_report(df,names,new_labels,steps,comparison_report,colors)

#test_step_report()



