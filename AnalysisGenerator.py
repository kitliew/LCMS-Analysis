# TODO read the result file as pd.df1
#   each sheet
# TODO read the dry weight file as pd.df2
#   contain column ["Filename", "Sample wt", "IS conc"]
# TODO merge both df
# TODO calculate normalized value

import pandas as pd
import os
from xlsxwriter.utility import xl_range_abs
import seaborn as sns
import matplotlib.pyplot as plt
from tkinter import *
from tkinter.filedialog import askopenfilename


DW_COLUMNS = ["Filename", "Sample wt"]
RESULT_DIRECTORY = os.getcwd()


def combined_df(summary_file, weight_file):

    # to store ['mean', 'stdev'] for each compound/sheet
    summary_grp = []
    # read summary_file.xlsx
    summary_df = pd.read_excel(summary_file, sheet_name=None, skiprows=3, header=1,
                               na_values="NF")   # sheet_name=None select all sheets

    # read dw_file.xlsx
    weight_df = pd.read_excel(weight_file, names=DW_COLUMNS)   # fix columns name

    # create pandas excel writer using xlsxwriter as engine
    writer = pd.ExcelWriter(RESULT_SUMMARY_FILE, engine="xlsxwriter")

    # merge each sheet with the DW sheet
    for sheet in summary_df:

        # current worksheet
        current_df = summary_df[sheet]
        # merge the two df
        merged_df = pd.merge(left=current_df, right=weight_df, on="Filename")

        # create a column to calculate the Normalized value
        merged_df["Normalized value (pg/mg DW)"] = merged_df["Area Ratio"] * IS_CONC * 1000 / merged_df["Sample wt"]

        # write df to worksheet
        merged_df.to_excel(writer, sheet_name=sheet, index=False, na_rep="NA")
        # result_df.to_excel(writer, sheet_name=sheet, index=False, header=True, na_rep="NA")

        # replace old df with the merged df
        summary_df[sheet] = merged_df

        # generate img_plot
        img_strip_box_plot(summary_df[sheet], sheet)

        # groupby ID to calculate mean and stdev
        summary = merged_df.groupby(["Sample ID"], sort=False).agg({"Normalized value (pg/mg DW)": ["mean", "std"]})
        summary = summary.rename(columns={"Normalized value (pg/mg DW)": sheet})
        summary_grp.append(summary)

    # convert groupby to df
    summary_result = pd.concat(summary_grp, axis=1)

    # write the summary df to sheet
    summary_result.to_excel(writer, sheet_name="Summary", na_rep="NA")

    # generate bar graph on summary sheet
    bar_with_stdev(writer, len(summary_df), list(summary_df.keys()), len(summary_grp))


def bar_with_stdev(writer, num_sheet, sheet_list, id_len):
    """
    Create graph using xlsxwriter.
    It is interactive like how one would create graph in excel.
    Type of graph: bar graph with stdev as error bar.
    """

    workbook = writer.book
    worksheet = writer.sheets["Summary"]

    for i in range(num_sheet):
        page = "Summary"
        start_row = 3
        end_row = 3 + id_len - 1  # python format [start:end]
        mean_col = (i * 2) + 1
        error_col = (i + 1) * 2

        # Convert Row-column notation and A1 notation
        page_ref = "=" + page + "!"
        error_range = page_ref + str(xl_range_abs(start_row, error_col, end_row, error_col))

        # Create a chart object.
        chart = workbook.add_chart({'type': 'column'})

        # Configure the series of the chart from the dataframe data.
        chart.set_title({
            'name': str(sheet_list[i])
        })

        # Add series details
        chart.add_series({
            'categories': [page, start_row, 0, end_row, 0],
            'values': [page, start_row, mean_col, end_row, mean_col],
            'y_error_bars': {
                'type': 'custom',
                'plus_values': error_range,
                'minus_values': error_range,
            },
        })

        chart.set_legend({'position': 'none'})

        chart.set_y_axis({'name': 'Normalized (pg/mg)'})

        worksheet.insert_chart('H2', chart)

    writer.save()


def img_strip_box_plot(df, plot_title, x_axis="Sample ID", y_axis="Normalized value (pg/mg DW)"):
    """Generate Box-plot overlap Swarm plot img file.

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
    # y_unit = ""
    y_label = str(y_axis)  # + y_unit
    ax.set(ylabel=y_label)

    # xticks
    ax.set_xticklabels(ax.get_xticklabels(), rotation=30)

    # write figure file with quality 400 dpi
    img_file = os.path.join(RESULT_DIRECTORY, plot_title + "." + "png")
    plt.savefig(img_file, bbox_inches='tight', dpi=400)

    # cleanup fig
    plt.close()


def get_is_conc():
    global e
    global IS_CONC
    IS_CONC = int(e.get())
    root.destroy()


if __name__ == "__main__":

    root = Tk()
    root.title('Internal Standard Amount (ng): ')
    e = Entry(root)
    e.pack()
    e.focus_set()
    b = Button(root, text='Okay', command=get_is_conc)
    b.pack(side='bottom')
    root.mainloop()

    Tk().withdraw()
    WEIGHT_FILE = askopenfilename(title="WEIGHT file", filetypes=(("CSV files", "*.xlsx"), ("all files", "*.*")))
    SHORT_REPORT = askopenfilename(title="REPORT file", filetypes=(("XLSX files", "*.xlsx"), ("all files", "*.*")))
    BASE_DIR = os.path.dirname(SHORT_REPORT)
    RESULT_DIRECTORY = os.path.join(BASE_DIR, "Results")  # dir of result folder
    RESULT_SUMMARY_FILE = os.path.join(RESULT_DIRECTORY, "Results.xlsx")

    if not os.path.exists(RESULT_DIRECTORY):
        os.makedirs(RESULT_DIRECTORY)

    # IS_CONC = int(input("Internal Standard Amount (ng): "))
    combined_df(SHORT_REPORT, WEIGHT_FILE)


