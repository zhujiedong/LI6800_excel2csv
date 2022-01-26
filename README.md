# Pyhthon GUI program for tidy LI-6800  and LI-6400 excel data


## LI-6800

### Purpose

There are some very helpful packages for dealing with [LI-6800](https://www.licor.com/env/products/photosynthesis/LI-6800/) photosynthesis data, such as `plantecophys`,`photosynthesis` etc. However these packages has a minimum requirement for the data--they assume the data are [Tidy data ](https://cfss.uchicago.edu/notes/tidy-data/#:~:text= To tidy the data%2C the basic approach,may be scattered across multiple rows. More ).  Though there are packages to read the RAW data of LI-6800, such as `RLicor, racir`. They are not user friendly to those who are not familiar with programmings, especially when the user need to correct the leaf area for recomputing purpose. To help those people achieve the first step before data analysis, I wrote a `tkinter` program to help tidy the LI-6800 excel data.

### Dependencies

First you should install python on your computer, there are plenty of guides on the Internet. The only library need to be installed additionally are `xlwings` and `pandas`, you can learn the basics from [xlwings](https://docs.xlwings.org/en/stable/).  

### How to use

click 'Choose the data folder' to choose the file which you store your LI-6800 excel data files, and then click 'Run LI-6800 Excel Converter' to run the program, after it is finished as indicated by the terminal, you can find your csv files in a folder called 'final-csv-data' in the folder of LI-6800 excel data files.

if you can read Chinese, just download a exe program that I already packaged without run the program from beginning.

[LI-6800-CSV](https://www.aliyundrive.com/s/FPcqSoUfYQJ)

## LI-6400

It is similar with LI-6800, but with 2 folders of csv data, as I do not know which is the best way to deal with the remarks.

[gui download](https://www.aliyundrive.com/s/WsUr16rVn9i)