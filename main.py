import pandas as pd
import json
from measure_filler import Populator
from utils_cls import MyUtils
from merger import *

utils = MyUtils("./")
excel_file_global = ''

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
        print(f"Tables file name = {custom_columns_file}.txt")
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
        df = pd.read_json(jsonString)
        df.to_csv(output_file_name,sep=';', encoding='utf-8', index=False)
        print(f"CSV file {output_file_name} produced successfully!")

    print("CSV processor finished")


while True:
    user_input = input("Enter excel file name .xlsx (or 'quit' to exit): ")
    
    if user_input.lower() == 'quit':
        print("Exiting the program.")
        break
    else:
        excel_file = user_input if user_input !='' else 'Sachim09062023'
        excel_file = excel_file + '.xlsx'
        print(f"Excel file = {excel_file}")

        if not utils.checkIfFileExistsInPath(excel_file):
            print(f"file {excel_file} does not exist!")
            raise OSError("File does not exist")
        
        excel_file_global = excel_file
        
        user_input = input("Enter Plant Sheet name [default Plant] (or 'quit' to exit): ")
        if user_input.lower() == 'quit':
            print("Exiting the program.")
            break
        else:
            plant_sheet = user_input if user_input !='' else 'Plant'


        user_input = input("Enter Modbus Sheet name [default Modbus] (or 'quit' to exit): ")
        if user_input.lower() == 'quit':
            print("Exiting the program.")
            break
        else:
            modbus_sheet = user_input if user_input !='' else 'Modbus'


        print(f"\n There are two modes for processing excel file: \n")
        print(f" auto   (A) Automatic mode combines the column headers from first sheet and second sheet together in the result csv")
        print(f" manual (M) Manual mode only selects specified column headers from the first and second sheet and combines them together \n")
        user_input = input("Enter (A) for automatic or (M) for manual [default Automatic] (or 'quit' to exit):")
        if user_input.lower() == 'quit':
          print("Exiting the program.")
          break
        elif user_input.lower() == 'a' or user_input.lower() == '':
            print("Auto mode selected")
            processor(False,excel_file,plant_sheet,modbus_sheet,None)
        else:
            print("Manual mode selected")
            user_input = input("Enter manual columns file name .txt (or 'quit' to exit): ")
            columns_file = user_input if user_input !='' else 'columns_custom'
            columns_file = f"{columns_file}.txt"

            if not utils.checkIfFileExistsInPath(columns_file):
                print(f"file {columns_file} does not exist!")
                raise OSError("File does not exist")
            
            processor(True,excel_file,plant_sheet,modbus_sheet,columns_file)


            

        

   


