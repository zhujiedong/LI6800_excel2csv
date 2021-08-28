import os
import pandas as pd
import xlwings as xw
from tkinter import *
from tkinter import filedialog


root = Tk()
root.title("LI-6800 数据整理")

# path of your LI-6800 data FOLDER
path = StringVar()
# the header row of your meansurement data
data_start = IntVar()


# function for choose the parent directory where your data folder lies
def select_path():
    path_ = filedialog.askdirectory()
    path.set(path_)

# main function for data tidy
def tidy_data():
    data_folder=[]
    #!  get the path of data and header
    dir_get = path.get() # data folder
    header_start = data_start.get()
    parent_folder = os.path.join(dir_get, "..") # parent folder of data folder
    excel_files = os.listdir(dir_get) # get all the data files
    for i in excel_files:
        if os.path.splitext(i)[1] == ".xlsx":
            print (i)
            data_folder.append(i)
    
    no_formula = "data_without_formula"
    csv_data = "final_csv_data"
    no_form_dir = os.path.join(parent_folder, no_formula)
    csv_dir = os.path.join(parent_folder, csv_data)

    if not os.path.exists(no_form_dir):
        try:
            os.mkdir(no_form_dir)
        except:
            print("create temparay folder failed for excel data without formula")

    if not os.path.exists(csv_dir):
        try:
            os.mkdir(csv_dir)
        except:
            print("create final csv folder failed")

    # prepare excel and csv names for the output data
    excel_no_formula  = [os.path.join(no_form_dir, i) for i in data_folder]
    excel_with_formula = [os.path.join(dir_get, i) for i in data_folder]
    csv_finally =  [os.path.join(csv_dir, i).replace("xlsx", "csv") for i in data_folder]

    for i in range(len(excel_with_formula)):
    # read all the excel file with formula by xlwings
        print("transforming：{}......".format(data_folder[i]))
        app=xw.App(visible=True,add_book=False)
        wb = app.books.open(excel_with_formula[i])
        sheet1 = wb.sheets[0].used_range.value
        # convert all the data to pandas dataframe
        df = pd.DataFrame(sheet1)
        # save excel files without formula
        df.to_excel(excel_no_formula[i], header = False, index = False)
        wb.close()
        app.quit()
        d = pd.read_excel(excel_no_formula[i], header=header_start)
        d = d[1:]
        d.to_csv(csv_finally[i])
        print("success, the number of files transformed until now：{} \n".format(i+1))
    print("completed，the total number of files transformed：**{}**".format(len(excel_with_formula)))
    print("close the Window to quit")



Label(root, text = "LI-6800 data folder：").grid(row = 0, column = 0)
dir = Entry(root, textvariable = path)
dir.grid(row = 0, column = 1)
Button(root, text="Choose folder",command=select_path).grid(row = 0, column = 2)

Label(root, text = "The row of header：").grid(row = 1, column = 0)
data_start_line = Entry(root, textvariable = data_start)
data_start_line.grid(row = 1, column = 1)
data_start_line.delete(0, "end")
data_start_line.insert(0, "13")

Button(root, text="Run program",command=tidy_data).grid(row = 2, column = 1)

root.mainloop()
