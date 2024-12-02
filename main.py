import tkinter as tk
from tkinter import filedialog, ttk

import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

root = tk.Tk()
root.title("Data Visualization")
root.geometry("1080x720")
root.pack_propagate(False)
root.resizable(0, 0)

file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=50, width=1080, rely=0, relx=0)

button1 = tk.Button(file_frame, text="Browse and Load A File", command=lambda: Load_excel_data())
button1.place(rely=0, relx=0.30)

button2 = tk.Button(file_frame, text="Show Data", command=lambda: ExcelWindow(df))
button2.place(rely=0, relx=0.50)

label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)


def File_dialog():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("csv files", "*.csv*"), ("Excel files", "*.xls*"), ("All Files", "*.*")))
    label_file["text"] = filename
    return None



def Load_excel_data():
    File_dialog()
    file_path = label_file["text"]
    global df
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None
    global data_frame
    data_frame = tk.LabelFrame(root, text="Data visualize section")
    data_frame.place(height=620, width=1080, relx=0, rely=0.1)
    T = tk.Text(data_frame, height=5, width=130)
    T.place(rely=0, relx=0)
    T.insert(tk.END, "Columns in given data are : " + ' '.join(list(df.columns)))
    button3 = tk.Button(data_frame, text="View Summary Statistics", command=lambda: ExcelWindow(
        pd.concat([pd.DataFrame(df.describe().index), df.describe().set_index(pd.DataFrame(df.describe().index).index)], axis=1)
    ))
    button3.place(rely=0.15, relx=0)
    l = tk.Label(data_frame, text="Select one of the plots : ")
    l.place(rely=0.22, relx=0)
    plots = [
        "Bar plot",
        "Pie chart",
        "Histogram",
        "Line plot",
        "Scatter plot",
        "Box plot",
        "Violin plot",
        "Heatmap",
        "KDE plot",
        "Pair plot"
    ]
    plot_type = tk.StringVar()
    drop = tk.OptionMenu(data_frame, plot_type, *plots)
    drop.place(rely=0.22, relx=0.12)
    l1 = tk.Label(data_frame, text="Select Attributes for plot : ")
    l1.place(rely=0.22, relx=0.2)
    xl = tk.Label(data_frame, text="X : ")
    xl.place(rely=0.22, relx=0.4)
    xvar = tk.StringVar()
    dropx = tk.OptionMenu(data_frame, xvar, *list(df.columns))
    dropx.place(rely=0.22, relx=0.42)
    yl = tk.Label(data_frame, text="Y : ")
    yl.place(rely=0.22, relx=0.6)
    yvar = tk.StringVar()
    dropy = tk.OptionMenu(data_frame, yvar, *list(df.columns))
    dropy.place(rely=0.22, relx=0.62)
    button4 = tk.Button(data_frame, text="Plot Data", command=lambda: plotData(plot_type.get(), xvar.get(), yvar.get()))
    button4.place(rely=0.22, relx=0.8)


def plotData(plot, xval, yval):

    plot_window = tk.Toplevel()
    if plot=='Heatmap'or plot=='Pair plot':
        plot_window.title(f"{plot}")
    elif plot=='Pie chart' or plot =='Histogram' or plot=='KDE plot':
        plot_window.title(f"{plot} for {xval}")
    else:
        plot_window.title(f"{plot} for {xval} and {yval}")
    plot_window.geometry("770x680")

    fig = Figure(figsize=(6, 6), dpi=100)
    fig1 = fig.add_subplot(111)

    numeric_df = df.select_dtypes(include=['number'])

    if plot == "Bar plot":
        sns.barplot(data=df, x=xval, y=yval, ax=fig1)
    elif plot == "Pie chart":
        df[xval].value_counts().plot.pie(autopct='%1.1f%%', ax=fig1)
    elif plot == "Histogram":
        sns.histplot(df[xval], kde=False, ax=fig1)
    elif plot == "Line plot":
        sns.lineplot(data=df, x=xval, y=yval, ax=fig1)
    elif plot == "Scatter plot":
        sns.scatterplot(data=df, x=xval, y=yval, ax=fig1)
    elif plot == "Box plot":
        sns.boxplot(data=df, x=xval, y=yval, ax=fig1)
    elif plot == "Violin plot":
        sns.violinplot(data=df, x=xval, y=yval, ax=fig1)
    elif plot == "Heatmap":
        if numeric_df.empty:
            tk.messagebox.showerror("Error", "No numeric data available for heatmap.")
            plot_window.destroy()
            return
        sns.heatmap(numeric_df.corr(), annot=True, ax=fig1)
    elif plot == "KDE plot":
        if xval in numeric_df.columns:
            sns.kdeplot(data=df, x=xval, ax=fig1)
        else:
            tk.messagebox.showerror("Error", f"'{xval}' is not numeric and cannot be used in KDE plot.")
            plot_window.destroy()
            return
    elif plot == "Pair plot":
        if numeric_df.empty:
            tk.messagebox.showerror("Error", "No numeric data available for pair plot.")
            plot_window.destroy()
            return
        sns.pairplot(numeric_df)\

    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=plot_window)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    toolbar = NavigationToolbar2Tk(canvas, plot_window)
    toolbar.update()
    canvas.get_tk_widget().pack()
    canvas.mpl_connect('motion_notify_event', lambda event: hover_event(toolbar,event))


def hover_event(toolbar,event):
        x, y = event.xdata, event.ydata
        
        if x is not None and y is not None:
            y_min, y_max = event.inaxes.get_ylim()
            y_inverted = y_min-y
            toolbar.set_message(f"x={x:.2f}, y={y_inverted:.2f}")




class ExcelWindow(tk.Toplevel):  # Inherits from tk.Toplevel

    def __init__(self, your_dataframe):  # the dataframe you passed through is here
        super().__init__()

        # Frame for TreeView
        frame1 = tk.LabelFrame(self, text="Excel Data")
        frame1.pack(fill="both", expand="true")
        frame1.pack_propagate(0)

        # the size of the window.
        self.geometry("1080x720")
        self.title("Data at " + str(filename))  # the window title

        # This creates your Treeview widget.
        tv1 = ttk.Treeview(frame1)
        tv1.place(relheight=1, relwidth=1)  # set the height and width of the widget to 100% of its container (frame1).

        treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)  # command means update the yaxis view of the widget
        treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)  # command means update the xaxis view of the widget
        tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)  # assign the scrollbars to the Treeview Widget
        treescrollx.pack(side="bottom", fill="x")  # make the scrollbar fill the x axis of the Treeview widget
        treescrolly.pack(side="right", fill="y")  # make the scrollbar fill the y axis of the Treeview widget

        # this loads the dataframe into the treeview widget
        tv1["column"] = list(your_dataframe.columns)
        tv1["show"] = "headings"
        for column in tv1["columns"]:
            tv1.heading(column, text=column) # let the column heading = column name

        df_rows = your_dataframe.to_numpy().tolist() # turns the dataframe into a list of lists
        for row in df_rows:
            tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert


root.mainloop()