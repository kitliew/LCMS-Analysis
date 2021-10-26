#! python3
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import openpyxl
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from xlsxwriter.utility import xl_range_abs
import os


# ----------------------------------------------------------------------------------------------------------------------
# INIT files

def result_dir():
    # INIT files
    if not os.path.exists(RESULT_DIRECTORY):
        os.makedirs(RESULT_DIRECTORY)
    # INIT xlsx SHORT_WEIGHT_FILE
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    wb.save(filename=SHORT_WEIGHT_FILE)


# ----------------------------------------------------------------------------------------------------------------------
# Functions

def plot_box_swarm(df, plot_title, x_axis="Sample ID", y_axis="Normalized"):
    """Plot box-plot and swarm plot for data list.

    Args:
        df (list of list): List of lists with data to be plotted.
        y_axis (str): Y- axis label.
        x_axis (list of str): List with labels of x-axis.
        plot_title (str): Plot title.

    """
    sns.set(color_codes=True)

    # add title to plot
    plt.title(plot_title)

    # plot data on Swarm plot and boxplot
    ax = sns.boxplot(x=x_axis, y=y_axis, data=df)
    ax = sns.swarmplot(x=x_axis, y=y_axis, data=df, color=".25")

    # y-axis label
    y_unit = " (pg/mg)"
    y_label = str(y_axis) + y_unit
    ax.set(ylabel=y_label)

    # xticks
    ax.set_xticklabels(ax.get_xticklabels(), rotation=30)

    # write figure file with quality 400 dpi
    img_file = os.path.join(RESULT_DIRECTORY, plot_title + "." + "png")
    plt.savefig(img_file, bbox_inches='tight', dpi=400)

    # cleanup fig
    plt.close()


# ----------------------------------------------------------------------------------------------------------------------
# Data Analysis

def lcms_summary():
    """ Create 3 set of files
        1. short_report_weight.xlsx with the following feature
            1. Concat short report and sample weight.
            2. Calculate Normalized value
            3. Generate Summary Page calculating mean and stdev for each sheet/compound

        2. summary_graph.xlsx with following feature
            1. Summary Page containing bar chart with stdev as error bar (excel generated style)

        3. compounds.png Image files
            1. Showing distribution of data using box plot + swarm plot

    """

    # read weight file
    weight_df = pd.read_csv(WEIGHT_FILE)

    # create dict of dataframes
    df = pd.read_excel(SHORT_REPORT, sheet_name=None, skiprows=3, header=1, na_values="NF")

    summary_grp = []    # to store summary of each sheet
    sheet_list = []     # list containing sheet name
    id_len = 0          # number of Sample ID

    # iterate through dict of dataframes
    for k, v in df.items():

        if k == "Component" or k == "Summary":
            continue

        # update sheet_list
        sheet_list.append(k)

        with pd.ExcelWriter(SHORT_WEIGHT_FILE, mode='a', engine='openpyxl', if_sheet_exists="replace") as writer:

            # merge df at Filename col
            result_df = pd.merge(left=v, right=weight_df, on="Filename")
            result_df["Normalized"] = result_df["Area Ratio"] / result_df["Sample wt"]      # normalized

            # generate plot_box_swarm img
            plot_box_swarm(result_df, k)

            # data for summary page
            summary = result_df.groupby(["Sample ID"]).agg({"Normalized": ["mean", "std"]})
            summary = summary.rename(columns={"Normalized": k})
            summary_grp.append(summary)

            # number of SampleID
            sample_ngrp = len(summary[k])
            if sample_ngrp > id_len:
                id_len = sample_ngrp

            # write processed data to Summary sheet
            result_df.to_excel(writer, sheet_name=k, index=False, header=True, na_rep="NA")

    # combine all the compounds (mean, std)
    summary_result = pd.concat(summary_grp, axis=1)

    with pd.ExcelWriter(SHORT_WEIGHT_FILE, mode='a', engine='openpyxl', if_sheet_exists="replace") as writer:
        summary_result.to_excel(writer, sheet_name="Summary", na_rep="NA")  # the data goes on this sheet

    return summary_result, summary_grp, sheet_list, id_len


def bar_with_stdev(summary_result, summary_grp, sheet_list, id_len):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    sheet_name = "Summary"
    with pd.ExcelWriter(RESULT_SUMMARY_FILE, engine='xlsxwriter') as writer:

        summary_result.to_excel(writer, sheet_name=sheet_name)

        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # loop
        x_row = id_len - 1

        for i in range(len(summary_grp)):
            page = "Summary"
            start_row = 3
            end_row = 3 + id_len - 1
            mean_start_col = (i * 2) + 1
            mean_end_col = mean_start_col
            error_start_col = (i + 1) * 2
            error_end_col = error_start_col
            page_ref = "=" + page + "!"
            error_range = page_ref + str(xl_range_abs(start_row, error_start_col, end_row, error_end_col))

            # Create a chart object.
            chart = workbook.add_chart({'type': 'column'})

            # Configure the series of the chart from the dataframe data.
            chart.set_title({
                'name': str(sheet_list[i])
            })

            chart.add_series({
                'categories': [page, start_row, 0, end_row, 0],
                'values': [page, start_row, mean_start_col, end_row, mean_end_col],
                'y_error_bars': {
                    'type': 'custom',
                    'plus_values': error_range,
                    'minus_values': error_range,
                },
            })

            chart.set_legend({'position': 'none'})

            chart.set_y_axis({'name': 'Normalized (pg/mg)'})

            worksheet.insert_chart('H2', chart)


# ----------------------------------------------------------------------------------------------------------------------
# Main

if __name__ == "__main__":
    Tk().withdraw()
    WEIGHT_FILE = askopenfilename()
    SHORT_REPORT = askopenfilename()
    BASE_DIR = os.path.dirname(SHORT_REPORT)
    RESULT_DIRECTORY = os.path.join(BASE_DIR, "Results")        # dir of result folder
    SHORT_WEIGHT_FILE = os.path.join(RESULT_DIRECTORY, "short_report_weight.xlsx")
    RESULT_SUMMARY_FILE = os.path.join(RESULT_DIRECTORY, "summary_graph.xlsx")

    result_dir()

    summary_result, summary_grp, sheet_list, id_len = lcms_summary()
    bar_with_stdev(summary_result, summary_grp, sheet_list, id_len)

