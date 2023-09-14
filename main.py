from io import StringIO
from typing import Annotated
import pandas as pd
import json
from measure_filler import Populator
from utils_cls import MyUtils
from merger import *
import typer
from typing_extensions import Annotated


utils = MyUtils("./")
excel_file_global = ''
app = typer.Typer(rich_markup_mode="rich")


@app.command(epilog="Made with 🍆 by [blue]MOSS[/blue]")
def auto(excelfile: Annotated[str, typer.Argument(help="Excel file name .xlsx")]):
    """
    Combine excel sheets together
    """

    if not excelfile.endswith(".xlsx"):
        excelfile = excelfile + ".xlsx"

    if not utils.checkIfFileExistsInPath(excelfile):
        print(f"file {excelfile} does not exist!")
        raise OSError("File does not exist")
    
    processor(False,excelfile,0,1,None)


@app.command(epilog="Made with 🍆 by [blue]MOSS[/blue]")
def manual(excelfile: Annotated[str, typer.Argument(help="Excel file name .xlsx")], 
           columnsfile: Annotated[str, typer.Argument(help="File containing column names .txt",
                                                       rich_help_panel="Secondary Arguments")] = 'columns_custom.txt'):
    """
    Combine excel sheets, only select specified columns
    """

    if not excelfile.endswith(".xlsx"):
        excelfile = excelfile + ".xlsx"

    if not columnsfile.endswith(".txt"):
        columnsfile = columnsfile + ".txt"

    if not utils.checkIfFileExistsInPath(excelfile):
        print(f"file {columnsfile} does not exist!")
        raise OSError("File does not exist")
    
    if not utils.checkIfFileExistsInPath(columnsfile):
        print(f"file {columnsfile} does not exist!")
        raise OSError("File does not exist")
    
    processor(True,excelfile,0,1,columnsfile)


def get_device_appendix_data(columns,device):
        device_sheet_name = device[0]
        appendix_data = check_static_data_for_sheet(device_sheet_name,excel_file_global)
        if len(appendix_data):
            #print(f"size: {data[0].keys()}")
            res = check_column_header_against_result_column_header(appendix_data[0].keys(),columns)
            if(len(res) > 0):
                print(f"different column name/s found in sheet {device[0]}, columns: {res}")
                exit()

            return appendix_data
        return []


def processor(mode,excel_file,sheet_1,sheet_2,custom_columns_file):
    plant_df = get_excel_dataframe(excel_file,sheet_1)
    plant_columns = plant_df.columns.tolist()

    result = convert_df_to_json(plant_df)
    plantData = json.loads(result)

    modbus_df = get_excel_dataframe(excel_file,sheet_2)
    modbus_columns = modbus_df.columns.tolist()

    result = convert_df_to_json(modbus_df)
    modbusData = json.loads(result)



    if mode:
        print(f"Tables file name = {custom_columns_file}")
        populator = Populator(custom_columns_file)
        column_names = populator.getMeasures()
       
    else:
        column_names = ['measure-id'] + plant_columns + modbus_columns

    
    device_list = create_output_device_list(plant_df)
    for device in device_list:
        device_csv_output = []
        appendix_data = get_device_appendix_data(column_names,device)
        result = cutter(device[1],device[2],appendix_data,plantData,plant_columns,modbusData,modbus_columns)

        for resultItem in result:
            ordered_obj = change_order(resultItem,column_names)
            device_csv_output.append(ordered_obj)


        output_file_name = f"./output/{device[0]}.csv"
        jsonString = json.dumps(device_csv_output)
        strObject = StringIO(jsonString)

        df = pd.read_json(strObject)
        df.to_csv(output_file_name,sep=';', encoding='utf-8', index=False)
        print(f"CSV file {output_file_name} produced successfully!")

    print("CSV processor finished")


if __name__ == "__main__":
    app()