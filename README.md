# Pyhthon GUI program for tidy LI-6800 excel data



## Purpose

There are some very helpful packages for dealing with [LI-6800](https://www.licor.com/env/products/photosynthesis/LI-6800/) photosynthesis data, such as `plantecophys`,`photosynthesis` etc. However these packages has a minimum requirement for the data--they assume the data are [Tidy data ](https://cfss.uchicago.edu/notes/tidy-data/#:~:text= To tidy the data%2C the basic approach,may be scattered across multiple rows. More ).  Though there are packages to read the RAW data of LI-6800, such as `RLicor, racir`. They are not user friendly to those who are not familiar with programmings, especially when the user need to correct the leaf area for recomputing purpose. To help those people achieve the first step before data analysis, I wrote a `tkinter` program to help tidy the LI-6800 excel data.

## Dependencies

First you should install python on your computer, there are plenty of guides on the Internet. The only library need to be installed is `xlwings`, you can learn the basics from [xlwings](https://docs.xlwings.org/en/stable/).  

## How to use

If you have all the dependencies installed, you can simply run the program `tidy_LI6800_excel.py`, If you could read Chinese, you can also download an exe file from https://www.aliyundrive.com/s/kJ71VwonJs2 :

[![hlKehV.png](https://z3.ax1x.com/2021/08/27/hlKehV.png)](https://imgtu.com/i/hlKehV)

### Choose your data folder

You should place all your data to a folder, the folder should **only contain LI-6800 data files**. Though the program is meant to tidy the excel files, you could still include the raw data files in it. For example, my data are in a folder called `6800exceldata`, and I put the data folder in a folder called `test6800`, because I just do not want mess up other folders that already have files and folders.

[![hlKZt0.png](https://z3.ax1x.com/2021/08/27/hlKZt0.png)](https://imgtu.com/i/hlKZt0)

### confirm your data header

In the example data, the header is in row 15, but the first row is empty and python starts from zero, so we use 13 here, you must confirm the row with your own data

[![hlKVkq.png](https://z3.ax1x.com/2021/08/27/hlKVkq.png)](https://imgtu.com/i/hlKVkq)





[![hlKk0s.png](https://z3.ax1x.com/2021/08/27/hlKk0s.png)](https://imgtu.com/i/hlKk0s)



Wait until the program finished, 

[![hlKA7n.png](https://z3.ax1x.com/2021/08/27/hlKA7n.png)](https://imgtu.com/i/hlKA7n)

then you can get your data in the `final_csv_data` folder

[![hlKFmj.png](https://z3.ax1x.com/2021/08/27/hlKFmj.png)](https://imgtu.com/i/hlKFmj)

