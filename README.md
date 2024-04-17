\# CSV Builder

CSV Builder is a tool which makes creation of csv files from excel sheets easier. The main functionality of this tool is marging two sheets, such that every row from the first sheet is repeated for all rows in the second sheet. The usecase for this logic is when plant data exists on the first sheet (with each row representing a line) and we want to collect all modbus registers (second sheet) for each line.

In addition, some quality of life functionality was added such as separation of csv files based on device name `Hostname Edge` and static addition of data to the device csv file in case a sheet with the same name as the `Hostname Edge` value exists among the sheets.


# Usage Guide

Prepare an excel file with at least two sheets, one for plant data and another for modbus data.
create an `output` directory if it does not exist.
Run `python app.py` inside the project folder.
Necessary args are shown by using the `python app.py --help` command.
Use auto mode in conjuction with default excel file to quickly produce csv sheets.
Or use manual mode and customize the input by providing a unique column name and desired output columns.

Follow the prompt step by step and choose between Automatic mode or Manual mode.
in case you choose Manual make sure you modify `columns_custom.txt` file. This file determines the output of csv files, so you can choose the columns which should exist inside the csv file and the program checks validity of columns. In case a column does not exist in your excel file the program raise an error.

You can find the final csv outputs for each edge device inside the output directory. 

## Manual mode
In Manual mode you can specify the name of desired columns in a .txt format, the program writes only those inside the resulting excel file.
You need to specify a unique column name which will act as your output identifier, in this case `Hostname Edge` as an example.
`python app.py manual sample.xlsx columns_custom.txt "Hostname Edge"`


# Requirements
[x] Python 3.9
`mkdir output`
`pip install "typer[all]"`
`pip install openpyxl pandas rich`
`pip install typing`
