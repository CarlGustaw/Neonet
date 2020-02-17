import ExcelReader
from RowRecord import RowRecord


class RowMaker:
    data_sheet = ""

    @staticmethod
    def read_excel_file_to_sheet_and_making_object(main_excel_path_name):
        dataSheet = ExcelReader.read_excel_file(main_excel_path_name)
        print("Main excel read successfully")
        list_of_objects = []
        rows_with_bad_dkf = []

        for row_number in range(1, dataSheet.nrows - 1):
            # Check if row contain correct DKF
            if str(dataSheet.row_slice(row_number)[8].value).find("+") != -1 or \
                    dataSheet.row_slice(row_number)[8].value == "" or \
                    dataSheet.row_slice(row_number)[8].value == "brak skanu" or \
                    dataSheet.row_slice(row_number)[8].value == "FS-01WW/00032360/2014 - ZŚW 799/2014" or \
                    dataSheet.row_slice(row_number)[8].value == "ID:379199":
                rows_with_bad_dkf.append(row_number)
            else:
                list_of_objects.append(RowRecord(str(int(dataSheet.row_slice(row_number)[8].value))))

        print("Rows as objects add successfully", " Number of readed rows: ", dataSheet.nrows - 1)
        print("Number of rows with good DKF:  ", len(list_of_objects))
        print("Number of rows with bad DKF:  ", len(rows_with_bad_dkf))
        print()
        return list_of_objects, rows_with_bad_dkf
