from io import StringIO
import pandas as pd
import json


class CsvBuilder:

    _instance = None

    def __new__(cb, *args, **kwargs):
        if cb._instance is None:
            cb._instance = super().__new__(cb)
        return cb._instance

    def __init__(self, excelfile, sheet1, sheet2, diviceIdentifierColumn="Hostname Edge", desiredColumns=[]) -> None:
        self.sheet1Name = sheet1
        self.sheet2Name = sheet2
        self.xlsFile = excelfile
        self.resultArray = {}
        self.appendixData = {}
        self.deviceColumn = diviceIdentifierColumn
        self.device_name_list = []
        self.desiredColumns: [] = desiredColumns
        self.sheet1Data = []
        self.sheet2Data = []
        self.sheet1Columns = []
        self.sheet2Columns = []
        self.dataFrame1 = pd.read_excel(
            self.xlsFile, sheet_name=self.sheet1Name)
        self.dataFrame2 = pd.read_excel(
            self.xlsFile, sheet_name=self.sheet2Name)

    def getDataFrame(self, sheetName) -> pd.DataFrame:
        my_dict = {
            self.sheet1Name: self.dataFrame1,
            self.sheet2Name: self.dataFrame2
        }

        return my_dict[sheetName]

    def getSheetColumns(self, sheetName) -> []:
        df = self.getDataFrame(sheetName=sheetName)
        return df.columns.tolist()

    def get_total_rows(self, sheet_name):
        df = self.getDataFrame(sheet_name)
        num_rows = df.shape[0]
        return num_rows

    def read_excel_row(self, sheet_name, row_index):
        df = self.getDataFrame(sheet_name)
        row = df.iloc[row_index]
        return row

    def convert_df_to_json(self, excel_dataframe_df: pd.DataFrame) -> str:
        json_str = excel_dataframe_df.to_json(orient='records')
        return json_str

    def change_order(self, data: dict, output_columns: []) -> dict:
        filtered_data = [d for d in data if not any(
            key in d for key in output_columns)]

        for item in filtered_data:
            data.pop(item, None)

        sorted_tuple_list = sorted(
            data.items(), key=lambda item: output_columns.index(item[0]))

        result_dict: dict = {key: value for (key, value) in sorted_tuple_list}
        return result_dict

    #############################################
    def create_output_device_list(self) -> None:
        # self.deviceColumn is the unique hostname of a device
        # csv files are created based on this unique name

        excel_dataframe_df = self.getDataFrame(self.sheet1Name)

        column_name = self.deviceColumn

        previous_value = None
        row_number_start = 2
        row_number_end = 0
        number_of_files = 1

        for value in excel_dataframe_df[column_name]:
            row_number_end += 1
            if previous_value is not None and value != previous_value:
                number_of_files += 1
                self.device_name_list.append(
                    (previous_value, row_number_start, row_number_end))
                row_number_start = row_number_end + 1

            previous_value = value

        row_number_end += 1
        self.device_name_list.append(
            (previous_value, row_number_start, row_number_end))

    def set_desired_columns(self) -> None:
        if len(self.desiredColumns) == 0:
            self.desiredColumns = ['measure-id'] + \
                self.getSheetColumns(self.sheet1Name) + \
                self.getSheetColumns(self.sheet2Name)

    def fill_appendix_data(self) -> None:
        self.set_desired_columns()

        for device in self.device_name_list:
            deviceName = device[0]
            appendixData = self.check_static_data_for_sheet(deviceName)
            self.appendixData[deviceName] = appendixData

            if len(self.appendixData[deviceName]) > 0:
                res = self.check_column_header_against_result_column_header(
                    self.appendixData[deviceName][0].keys(), self.desiredColumns)
                # print(self.appendixData[device][0].keys())

                if len(res) > 0:
                    print(
                        f"different column name/s found in sheet {deviceName}, columns: {res}")
                    exit()

    def check_static_data_for_sheet(self, deviceSheetName) -> []:
        try:
            device_df = pd.read_excel(self.xlsFile, sheet_name=deviceSheetName)
            result = self.convert_df_to_json(device_df)
            sheetData = json.loads(result)
            return sheetData
        except:
            return []

    def check_column_header_against_result_column_header(self, input_column_headers, result_column_headers):
        non_existant = []
        for element in input_column_headers:
            if element not in result_column_headers:
                non_existant.append(element)

        return non_existant

    def show_device_list(self):
        print("Listing devices: ")
        for device in self.device_name_list:
            print(f"{device=}")

    def combiner(self) -> None:
        sheet2TotalRows = self.get_total_rows(sheet_name=self.sheet2Name)
        modbusData = []
        for i in range(0, sheet2TotalRows):
            modbusData.append(self.read_excel_row(self.sheet2Name, i))

        for device in self.device_name_list:
            measure_id = 1
            device_name = device[0]
            # point to the start row index of data for a device inside the array
            device_data_start_index = device[1] - 2
            # point to the end row of data for a device inside the array
            device_data_end_index = device[2] - 1

            # For each device check if appendix data sheet exists
            # if it does add it to the result list
            list = []
            appendixData = self.check_static_data_for_sheet(device_name)

            for data in appendixData:
                rowItem = dict({"measure-id": measure_id})
                rowItem.update(data)
                list.append(rowItem)
                measure_id = measure_id + 1

            for i in range(device_data_start_index, device_data_end_index):
                for j in range(0, len(modbusData)):
                    rowItem = dict({"measure-id": measure_id})
                    sheet1Row = self.read_excel_row(self.sheet1Name, i)
                    rowItem.update(sheet1Row)
                    sheet2Row = modbusData[j]
                    rowItem.update(sheet2Row)
                    list.append(rowItem)
                    measure_id = measure_id + 1

            self.resultArray[device_name] = list

    def reorderData(self):

        for device in self.device_name_list:
            device_csv_output = []

            device_name = device[0]
            device_data = self.resultArray[device_name]

            for data in device_data:
                sorted_obj = self.change_order(
                    data, ["measure-id"] + self.desiredColumns)
                device_csv_output.append(sorted_obj)

            self.resultArray[device_name] = device_csv_output

    def createCSV(self):
        for device in self.device_name_list:
            deviceName = device[0]

            output_file_name = f"./output/{deviceName}.csv"

            jsonString = json.dumps(self.resultArray[deviceName], default=str)
            strObject = StringIO(jsonString)

            df = pd.read_json(strObject)
            df.to_csv(output_file_name, sep=';', encoding='utf-8', index=False)
            print(f"CSV file {output_file_name} produced successfully!")


if __name__ == "__main__":

    desColumns = ['Codice Stabilimento', 'Descrizione Stabilimento', 'Edificio', 'Vettore', 'POD', 'Piano', 'Reparto', 'Quadro', 'Descrizione Quadro', 'Sottoquadro', 'Descrizione Sottoquadro', 'Linea', 'Descrizione Linea', 'Area Funzionale ENEA', 'Cod. Funzionale TERA', 'Tipologia Dispositivo',
                  'Codice Modello Dispositivo', 'Taglia Interruttore', 'Tipologia Misura', 'Hostname Edge', 'ID Modbus', 'Tipo Dispositivo', 'mqtt-topic', 'modbus-id', 'modbus-fc', 'modbus-address', 'modbus-format', 'Misura', 'Unita di misura', 'modbus-n_registers', 'mqtt-qos', 'group', 'ime-type']

    builder = CsvBuilder("sample.xlsx", "Plant", "Modbus",
                         "Hostname Edge", desColumns)
    totalRows = builder.get_total_rows("Plant")
    builder.create_output_device_list()
    builder.fill_appendix_data()
    builder.show_device_list()
    builder.combiner()
    builder.reorderData()
    builder.createCSV()
