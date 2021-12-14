import os
import pandas as pd
import xlwings as xw
from tkinter import *
from tkinter import filedialog

root = Tk()

root.geometry('410x100')
root.title("tidy LI-6800 excel data")

# path entry variable
path = StringVar()
data_start = StringVar()

#  define path select function
def select_path():
    path_ = filedialog.askdirectory()
    path.set(path_)

# main fucntion
def tidy_data():
    list_excel_data=[]
    #!  set the data folder
    dir_get = path.get()
    list_folder = os.listdir(dir_get)
    for i in list_folder:
        if os.path.splitext(i)[1] == ".xlsx":
            #print (i)
            list_excel_data.append(i)
    
    # csv data folder name
    csv_data = "final_csv_data"
    # csv data folder path
    csv_dir = os.path.join(dir_get, csv_data)

    if not os.path.exists(csv_dir):
        try:
            os.mkdir(csv_dir)
        except:
            print("create folder failed")

    # prepare excel and csv names for the output data
    excel_with_formula = [os.path.join(dir_get, i) for i in list_excel_data]
    csv_finally =  [os.path.join(csv_dir, i).replace("xlsx", "csv") for i in list_excel_data]

    for i in range(len(csv_finally)):
    # read all excel files data and convert it to dataframe
        try:
            print("tidy data now：{}......".format(list_excel_data[i]))
            app=xw.App(visible=True,add_book=False)
            wb = app.books.open(excel_with_formula[i])
            # convert excel data to dataframe
            sheet1 = wb.sheets[0].used_range.value
            df = pd.DataFrame(sheet1)
            app.quit()
            # read the obs row no.
            row_obs = df[df.iloc[:, 0] == 'obs'].index.tolist()[0]
            # convert to header
            df_header = df.iloc[row_obs, :].tolist()
            df_data = df.iloc[(row_obs+2):, :]
            df_data.columns = df_header
            # save to csf data directly
            df_data.to_csv(csv_finally[i], index = False)

            print("convert current files successfully，now the no. of data files converted are：**{}** \n".format(i+1))
        except:
            print("convert failed, may be there are temparary files exists")
    print("convert all files successfully, toltal files converted are ：{}".format(len(excel_with_formula)))
    print("close the window to quit")


# 定义工作目录-----------------------------------------

Label(root, text = "Data folder：").grid(row = 0, column = 0)
dir = Entry(root, textvariable = path)
dir.grid(row = 0, column = 1)
Button(root, text="Choose the data folder",command=select_path).grid(row = 0, column = 2)

Button(root, text="Run LI-6800 Excel Converter",command=tidy_data).grid(row = 2, column = 1)

root.mainloop()
